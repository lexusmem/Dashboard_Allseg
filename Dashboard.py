import streamlit as st
import pandas as pd
import base64
import logging

# Caminho do arquivo Excel (ajuste conforme necessário para o ambiente de execução)
# Nota: Em um ambiente de produção, considere usar st.file_uploader para permitir que o usuário faça upload do arquivo.
arquivo_excel = r'C:\Users\lexus\Documents\Alseg\Cópia de Precificacao - Copia.xlsx'

# Configura a página para layout amplo
st.set_page_config(layout='wide')


# Dados agrupado de apólices e sinistros:
@st.cache_data
def carregar_e_processar_dados(caminho_arquivo):
    """
    Carrega e processa os dados do arquivo Excel.
    Esta função é cacheada para evitar recarregar e reprocessar os dados
    a cada interação do usuário, tornando a aplicação mais rápida.
    """
    try:
        # Carrega a aba 'apolice_endosso'
        aba_apolice_endosso = pd.read_excel(
            caminho_arquivo, sheet_name='apolice_endosso')

        # Fazer a soma dos prêmios agrupado por apólice:
        soma_por_apolice = aba_apolice_endosso.groupby(
            'cd_apolice')['vl_tarifario_pago'].sum().reset_index()
        soma_por_apolice.rename(columns={
                                'cd_apolice': 'N° Apólice', 'vl_tarifario_pago': 'Soma Prêmio Pago por Apolice'}, inplace=True)

        # Dados adicionais dos dados das apólices:
        colunas_adicionais = [
            'cd_apolice', 'nm_tp_apolice', 'nm_tp_cobranca',
            'nm_regiao_circulacao', 'nm_auto_utilizacao', 'dt_ini_vig_apo',
            'dt_fim_vig_apo', 'nm_uf_cliente', 'nm_cidade', 'nm_estipulante',
            'nm_produto', 'nm_corretor', 'nm_representante'
        ]

        # Selecionar as colunas adicionais, eliminando duplicatas por 'cd_apolice'
        dados_adicionais = aba_apolice_endosso[colunas_adicionais].drop_duplicates(
            subset='cd_apolice')
        dados_adicionais.rename(
            columns={'cd_apolice': 'N° Apólice'}, inplace=True)

        # Merge dos dados de prêmio com os dados adicionais
        premio_com_dados = pd.merge(
            soma_por_apolice, dados_adicionais, on='N° Apólice', how='left')

        # Carrega a aba 'sinistro'
        aba_sinistro = pd.read_excel(caminho_arquivo, sheet_name='sinistro')
        aba_sinistro['Total Sinistro'] = aba_sinistro['vl_sinistro_total'] + aba_sinistro['vl_despesa_total'] + \
            aba_sinistro['vl_honorario_total'] - \
            aba_sinistro['vl_salvado_total']

        # Soma dos sinistros por apólice:
        soma_sinistro_por_apolice = aba_sinistro.groupby(
            'cd_apolice')['Total Sinistro'].sum().reset_index()
        soma_sinistro_por_apolice.rename(columns={
                                         'cd_apolice': 'N° Apólice', 'Total Sinistro': 'Soma Sinistro Por Apolice'}, inplace=True)

        # Merge dos resultados finais
        resultado_final = pd.merge(
            premio_com_dados,
            soma_sinistro_por_apolice,
            on='N° Apólice',
            how='outer'
        )

        # Preenche valores NaN com 0
        resultado_final = resultado_final.fillna(0)

        return resultado_final
    except FileNotFoundError:
        st.error(
            f"Erro: O arquivo '{caminho_arquivo}' não foi encontrado. Por favor, verifique o caminho.")
        return pd.DataFrame()  # Retorna um DataFrame vazio em caso de erro
    except Exception as e:
        st.error(f"Ocorreu um erro ao carregar ou processar os dados: {e}")
        return pd.DataFrame()


# DF com dados de Sinistros:
@st.cache_data
def carregar_e_processar_dados_sinistro(caminho_arquivo):
    """
    Carrega e processa os dados da aba sinistro do arquivo Excel.
    Esta função é cacheada para evitar recarregar e reprocessar os dados
    a cada interação do usuário, tornando a aplicação mais rápida.
    """
    try:
        # Carrega a aba 'sinistro'
        aba_sinistro = pd.read_excel(caminho_arquivo, sheet_name='sinistro')
        aba_sinistro['Total Sinistro'] = aba_sinistro['vl_sinistro_total'] + aba_sinistro['vl_despesa_total'] + \
            aba_sinistro['vl_honorario_total'] - \
            aba_sinistro['vl_salvado_total']

        aba_sinistro.reset_index(drop=True, inplace=True)

        # Soma dos sinistros por apólice:
        dados_de_sinistro = aba_sinistro
        dados_de_sinistro.rename(
            columns={'cd_apolice': 'N° Apólice'}, inplace=True)

        # Preenche valores NaN com 0
        dados_de_sinistro = dados_de_sinistro.fillna(0)

        return dados_de_sinistro
    except FileNotFoundError:
        st.error(
            f"Erro: O arquivo '{caminho_arquivo}' não foi encontrado. Por favor, verifique o caminho.")
        return pd.DataFrame()  # Retorna um DataFrame vazio em caso de erro
    except Exception as e:
        st.error(f"Ocorreu um erro ao carregar ou processar os dados: {e}")
        return pd.DataFrame()

# Função de Formatação de Valores para o padrão Brasileiro


def formatar_valor_br(valor):
    """
    Formata um valor numérico para o padrão monetário brasileiro (R$ X.XXX,XX).
    Lida com valores NaN retornando uma string vazia.
    """
    if pd.isna(valor):
        return ""
    # Formata como float com 2 casas decimais e separador de milhar (padrão US)
    valor_us_format = f"{valor:,.2f}"
    # Inverte os separadores para o padrão brasileiro
    valor_br_format = valor_us_format.replace(
        ",", "X").replace(".", ",").replace("X", ".")
    return valor_br_format


# --- Aplicação Streamlit ---
# Carrega e processa os dados (cacheado para performance)
dados_calculados = carregar_e_processar_dados(arquivo_excel)

# Verifica se os dados foram carregados com sucesso
if dados_calculados.empty:
    st.stop()  # Para a execução se não houver dados

# Cria uma cópia para exibição e cálculos de porcentagem/formatação
dados_exibicao = dados_calculados.copy()

# Cria o percentual de sinistro, tratando divisão por zero
dados_exibicao.loc[:, '% Sin'] = dados_exibicao.apply(
    lambda row: '{:.2%}'.format(
        row['Soma Sinistro Por Apolice'] / row['Soma Prêmio Pago por Apolice'])
    if row['Soma Prêmio Pago por Apolice'] != 0 else '0.00%', axis=1
)

# Formata as colunas para exibição
dados_exibicao.loc[:, 'Soma Prêmio Pago por Apolice'] = dados_exibicao['Soma Prêmio Pago por Apolice'].map(
    formatar_valor_br)
dados_exibicao.loc[:, 'Soma Sinistro Por Apolice'] = dados_exibicao['Soma Sinistro Por Apolice'].map(
    formatar_valor_br)

# Reordenar as colunas para que 'Soma Sinistro Por Apolice' e '% Sin' fiquem nas posições desejadas
colunas = list(dados_exibicao.columns)

# Remove as colunas que vamos inserir manualmente, se existirem
for col in ['Soma Sinistro Por Apolice', '% Sin']:
    if col in colunas:
        colunas.remove(col)

# Insere nas posições desejadas
colunas.insert(2, 'Soma Sinistro Por Apolice')
colunas.insert(3, '% Sin')

# Reordena o DataFrame
dados_exibicao = dados_exibicao[colunas]

# Ordenando por numero da apólice inicialmente
dados_exibicao = dados_exibicao.sort_values('N° Apólice')


# Dados do sinistro
df_sinistros = carregar_e_processar_dados_sinistro(arquivo_excel)
# Verifica se os dados foram carregados com sucesso
if df_sinistros.empty:
    st.stop()  # Para a execução se não houver dados

# '''
# imagem sidebar
#
#
#
#
#
#
# '''


def img_to_base64(image_path):
    """Convert image to base64."""
    try:
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    except Exception as e:
        logging.error(f"Error converting image to base64: {str(e)}")
        return None


# Load and display sidebar image
img_path = r'C:\Users\alex.sousa\Documents\Dados_Sinistros\image\lexus_gemine_II-Photoroom_menor80.png'
img_base64 = img_to_base64(img_path)
if img_base64:
    st.sidebar.markdown(
        # essa função para colocar glowing effect na imagem
        # f'<img src="data:image/png;base64,{img_base64}" class="cover-glow">',
        f'<img src="data:image/png;base64,{img_base64}" style="width: 150px; height: auto; display: block; margin-left: auto; margin-right: auto; margin-top: -40px;margin-bottom: 5px">',
        unsafe_allow_html=True,
    )

# colocar linha embaixo do logo
# st.sidebar.markdown("---")


# '''
# imagem sidebar
#
#
#
#
#
#
# '''


# '''
# para baixo trabalho filtro apólice
#
#
#
#
#
#
# '''

# --- Filtragem dados da Apólice ---
st.sidebar.header('Filtro Apólice')

# Filtro por Apólice - Obtém as apólices únicas
apolices_filtro_apolice = sorted(dados_exibicao['N° Apólice'].unique())

# Define o índice padrão para selectbox
default_index_apolice = 0 if apolices_filtro_apolice else None

apolices_selecionadas_filtro_apolice = st.sidebar.selectbox(
    'Apólice',
    options=apolices_filtro_apolice,
    index=default_index_apolice  # Selecionar o primeiro registro por padrão
)

st.subheader(f'Dados Apólice - {apolices_selecionadas_filtro_apolice}')
dados_filtrados_filtro_apolice = dados_exibicao.copy()
if apolices_selecionadas_filtro_apolice:
    dados_filtrados_filtro_apolice = dados_filtrados_filtro_apolice[
        dados_filtrados_filtro_apolice['N° Apólice'] == apolices_selecionadas_filtro_apolice]

st.sidebar.markdown("---")

# Converte as colunas de volta para numérico para somar
# É importante fazer isso em uma cópia para não afetar a exibição formatada
df_para_filtro_apolice = dados_filtrados_filtro_apolice.copy()
df_para_filtro_apolice['Soma Prêmio Pago por Apolice'] = df_para_filtro_apolice['Soma Prêmio Pago por Apolice'].str.replace(
    '.', '').str.replace(',', '.').astype(float)
df_para_filtro_apolice['Soma Sinistro Por Apolice'] = df_para_filtro_apolice['Soma Sinistro Por Apolice'].str.replace(
    '.', '').str.replace(',', '.').astype(float)

total_premio_filtro_apolice = df_para_filtro_apolice['Soma Prêmio Pago por Apolice'].sum(
)
total_sinistro_filtro_apolice = df_para_filtro_apolice['Soma Sinistro Por Apolice'].sum(
)

# Calcula o percentual de sinistro total
percentual_sinistro_total_filtro_apolice = (
    total_sinistro_filtro_apolice / total_premio_filtro_apolice) if total_premio_filtro_apolice != 0 else 0

col_apl_1, col_apl_2, col_apl_3 = st.columns(3)

with col_apl_1:
    st.metric(label="Total Prêmio Pago",
              value=f"R$ {formatar_valor_br(total_premio_filtro_apolice)}")
with col_apl_2:
    st.metric(label="Total Sinistro",
              value=f"R$ {formatar_valor_br(total_sinistro_filtro_apolice)}")
with col_apl_3:
    st.metric(label="% Sinistro Total",
              value=f"{percentual_sinistro_total_filtro_apolice:.2%}")

# st.subheader('Segurado: ')
# st.caption('Segurado: ')
# st.write('Segurado: ')
# st.text('Segurado: ')
# st.markdown("**Segurado:**")

col_seg_1, col_cor_2, col_rep_3, col_util_4 = st.columns(4)

segurado = list(dados_filtrados_filtro_apolice['nm_estipulante'].unique())
corretor = list(dados_filtrados_filtro_apolice['nm_corretor'].unique())
representante = list(
    dados_filtrados_filtro_apolice['nm_representante'].unique())
utilização = list(
    dados_filtrados_filtro_apolice['nm_auto_utilizacao'].unique())


with col_seg_1:
    st.markdown("<p style='margin-bottom: 0;'>Segurado</p>",
                unsafe_allow_html=True)
    st.markdown(
        f"<h6 style='margin-top: 0; margin-bottom: 0.2rem;'>{segurado[0].title()}</h6>", unsafe_allow_html=True)
with col_cor_2:
    st.markdown("<p style='margin-bottom: 0;'>Corretor</p>",
                unsafe_allow_html=True)
    st.markdown(
        f"<h6 style='margin-top: 0; margin-bottom: 0.2rem;'>{corretor[0].title()}</h6>", unsafe_allow_html=True)
with col_rep_3:
    st.markdown("<p style='margin-bottom: 0;'>Representante</p>",
                unsafe_allow_html=True)
    st.markdown(
        f"<h6 style='margin-top: 0; margin-bottom: 0.2rem;'>{representante[0].title()}</h6>", unsafe_allow_html=True)
with col_util_4:
    st.markdown("<p style='margin-bottom: 0;'>Utilização</p>",
                unsafe_allow_html=True)
    st.markdown(
        f"<h6 style='margin-top: 0; margin-bottom: 0.2rem;'>{utilização[0].title()}</h6>", unsafe_allow_html=True)

st.dataframe(dados_filtrados_filtro_apolice, hide_index=True)

#
#
#
#
# DADOS DO SEGURADO PARA APRESENTAÇÃO
#
#
#
#

st.subheader(f'Dados do Segurado - {segurado[0]}')

dados_apolices_segurado = dados_exibicao.copy()
if apolices_selecionadas_filtro_apolice:
    dados_apolices_segurado = dados_apolices_segurado[
        dados_apolices_segurado['nm_estipulante'] == segurado[0]]


df_pr_sin_segurado = dados_apolices_segurado.copy()

df_pr_sin_segurado['Soma Prêmio Pago por Apolice'] = df_pr_sin_segurado['Soma Prêmio Pago por Apolice'].str.replace(
    '.', '').str.replace(',', '.').astype(float)
df_pr_sin_segurado['Soma Sinistro Por Apolice'] = df_pr_sin_segurado['Soma Sinistro Por Apolice'].str.replace(
    '.', '').str.replace(',', '.').astype(float)

# Dados de sinistro do segurado
df_sinistro_segurado = df_sinistros.loc[df_sinistros['nm_cliente'] == segurado[0]]


total_pr_segurado = df_pr_sin_segurado['Soma Prêmio Pago por Apolice'].sum()
total_sinistro_segurado = df_pr_sin_segurado['Soma Sinistro Por Apolice'].sum()
sinistralidade_segurado = (
    total_sinistro_segurado / total_pr_segurado) if total_pr_segurado != 0 else 0
qtd_apolice_segurado = df_pr_sin_segurado['N° Apólice'].nunique()
qtd_sinistros_segurado = df_sinistro_segurado['nr_sinistro'].nunique()

seg_apl_1, seg_apl_2, seg_apl_3, seg_apl_4, seg_apl_5 = st.columns(5)

with seg_apl_1:
    st.metric(label="Total Prêmio Pago",
              value=f"R$ {formatar_valor_br(total_pr_segurado)}")
with seg_apl_2:
    st.metric(label="Total Sinistro",
              value=f"R$ {formatar_valor_br(total_sinistro_segurado)}")
with seg_apl_3:
    st.metric(label="% Sinistro Total",
              value=f"{sinistralidade_segurado:.2%}")
with seg_apl_4:
    st.metric(label='Qtd. Apolices', value=qtd_apolice_segurado)
with seg_apl_5:
    st.metric(label='Qtd Sinistros', value=qtd_sinistros_segurado)


# dados_apolices_segurado

#
#
#
#
# DADOS DO SEGURADO PARA APRESENTAÇÃO
#
#
#
#

#
#
#
#
# DADOS DE SINISTRO
#
#
#
#

# st.dataframe(df_sinistros, hide_index=True)

st.dataframe(df_sinistro_segurado, hide_index=True)

# dados de sinistro por cobertura por segurado
df_sinistro_segurado_cobertura = df_sinistro_segurado.groupby(df_sinistro_segurado['Cobertura']).agg(
    soma_total_snistro=('Total Sinistro', 'sum'),
    contagem_de_sinistro=('nr_sinistro', 'nunique')
)

df_sinistro_segurado_cobertura

#
#
#
#
# FIM DADOS DE SINISTRO
#
#
#
#


st.divider()

# '''
# para cima trabalho filtro apólice
#
#
#
#
#
#
# '''

# --- Lógica de Filtragem Hierárquica na Sidebar ---
st.sidebar.header('Filtros Dados Gerais')

# 1. Filtro por Representante
# Obtém os representantes únicos da base de dados completa, garantindo que sejam strings
representantes_unicos = sorted(
    dados_exibicao['nm_representante'].astype(str).unique())
representantes_selecionados = st.sidebar.multiselect(
    'Representante(s)',
    options=representantes_unicos,
    default=[]  # Nenhuma seleção padrão
)

# Aplica o filtro de Representante
dados_filtrados_rep = dados_exibicao.copy()
if representantes_selecionados:
    dados_filtrados_rep = dados_filtrados_rep[dados_filtrados_rep['nm_representante'].astype(
        str).isin(representantes_selecionados)]

# 2. Filtro por Corretor (baseado nos dados já filtrados por Representante)
# Obtém os corretores únicos dos dados filtrados por representante, garantindo que sejam strings
corretores_unicos = sorted(
    dados_filtrados_rep['nm_corretor'].astype(str).unique())
corretores_selecionados = st.sidebar.multiselect(
    'Corretor(es)',
    options=corretores_unicos,
    default=[]  # Nenhuma seleção padrão
)

# Aplica o filtro de Corretor
dados_filtrados_corr = dados_filtrados_rep.copy()
if corretores_selecionados:
    dados_filtrados_corr = dados_filtrados_corr[dados_filtrados_corr['nm_corretor'].astype(
        str).isin(corretores_selecionados)]


# 3. Filtro por Segurado (baseado nos dados já filtrados por corretor)
# Obtém os segurados únicos dos dados filtrados por corretor, garantindo que sejam strings
segurados_unicos = sorted(
    dados_filtrados_corr['nm_estipulante'].astype(str).unique())
segurados_selecionados = st.sidebar.multiselect(
    'Segurado(s)',
    options=segurados_unicos,
    default=[]  # Nenhuma seleção padrão
)

# Aplica o filtro de Segurado
dados_filtrados_segurado = dados_filtrados_corr.copy()
if segurados_selecionados:
    dados_filtrados_segurado = dados_filtrados_segurado[dados_filtrados_segurado['nm_estipulante'].astype(
        str).isin(segurados_selecionados)]


# 4. Filtro por Apólice (baseado nos dados já filtrados por Representante e Corretor)
# Obtém as apólices únicas dos dados filtrados por corretor
apolices_unicas = sorted(dados_filtrados_segurado['N° Apólice'].unique())
apolices_selecionadas = st.sidebar.multiselect(
    'Apólice(s)',
    options=apolices_unicas,
    default=[]  # Nenhuma seleção padrão
)

# Aplica o filtro de Apólice
resultado_final_filtrado = dados_filtrados_segurado.copy()
if apolices_selecionadas:
    resultado_final_filtrado = resultado_final_filtrado[resultado_final_filtrado['N° Apólice'].isin(
        apolices_selecionadas)]

# --- Indicadores Chave (KPIs) ---
st.subheader("Dados Gerais")

# Converte as colunas de volta para numérico para somar
# É importante fazer isso em uma cópia para não afetar a exibição formatada
df_para_soma = resultado_final_filtrado.copy()
df_para_soma['Soma Prêmio Pago por Apolice'] = df_para_soma['Soma Prêmio Pago por Apolice'].str.replace(
    '.', '').str.replace(',', '.').astype(float)
df_para_soma['Soma Sinistro Por Apolice'] = df_para_soma['Soma Sinistro Por Apolice'].str.replace(
    '.', '').str.replace(',', '.').astype(float)

total_premio = df_para_soma['Soma Prêmio Pago por Apolice'].sum()
total_sinistro = df_para_soma['Soma Sinistro Por Apolice'].sum()

# Calcula o percentual de sinistro total
percentual_sinistro_total = (
    total_sinistro / total_premio) if total_premio != 0 else 0

col1, col2, col3 = st.columns(3)

with col1:
    st.metric(label="Total Prêmio Pago",
              value=f"R$ {formatar_valor_br(total_premio)}")
with col2:
    st.metric(label="Total Sinistro",
              value=f"R$ {formatar_valor_br(total_sinistro)}")
with col3:
    st.metric(label="% Sinistro Total",
              value=f"{percentual_sinistro_total:.2%}")


# --- Exibição dos Resultados ---
st.subheader("Dados de Sinistros e Prêmios")

if not resultado_final_filtrado.empty:
    st.dataframe(resultado_final_filtrado, hide_index=True)
else:
    st.info("Nenhum dado encontrado com os filtros selecionados.")


# --- Dados de Prêmio e Sinistro por Utilização ---
st.subheader("Dados de Prêmio e Sinistro por Utilização")

if not resultado_final_filtrado.empty:
    # Crie uma cópia do DataFrame filtrado para realizar os cálculos numéricos
    # sem afetar a formatação da tabela principal.
    df_para_groupby = resultado_final_filtrado.copy()

    # Converta as colunas de prêmio e sinistro de volta para numérico (float)
    # antes de realizar a soma, removendo os separadores de milhar e decimal.
    df_para_groupby['Soma Prêmio Pago por Apolice'] = df_para_groupby['Soma Prêmio Pago por Apolice'].str.replace(
        '.', '').str.replace(',', '.').astype(float)
    df_para_groupby['Soma Sinistro Por Apolice'] = df_para_groupby['Soma Sinistro Por Apolice'].str.replace(
        '.', '').str.replace(',', '.').astype(float)

    # Agrupe por 'nm_auto_utilizacao' e some os valores numéricos
    groupby_utilizacao = df_para_groupby.groupby('nm_auto_utilizacao').agg(
        Total_Premio=('Soma Prêmio Pago por Apolice', 'sum'),
        Total_Sinistro=('Soma Sinistro Por Apolice', 'sum')
    ).reset_index()

    # Calcule a % de Sinistralidade para cada grupo
    groupby_utilizacao['% Sinistralidade'] = groupby_utilizacao.apply(
        lambda row: '{:.2%}'.format(
            row['Total_Sinistro'] / row['Total_Premio'])
        if row['Total_Premio'] != 0 else '0.00%', axis=1
    )

    # Formate as colunas de valores para exibição no padrão BR
    groupby_utilizacao['Total_Premio'] = groupby_utilizacao['Total_Premio'].map(
        formatar_valor_br)
    groupby_utilizacao['Total_Sinistro'] = groupby_utilizacao['Total_Sinistro'].map(
        formatar_valor_br)

    # Renomeie a coluna de agrupamento para melhor apresentação
    groupby_utilizacao.rename(
        columns={'nm_auto_utilizacao': 'Utilização'}, inplace=True)

    # Ordene o DataFrame pelo 'Total_Premio' em ordem decrescente
    groupby_utilizacao = groupby_utilizacao.sort_values(
        by='Total_Premio', ascending=False)

    # Exiba o DataFrame agrupado
    st.dataframe(groupby_utilizacao, hide_index=True)
else:
    st.info("Nenhum dado disponível para agrupar por Utilização.")

# Instruções para executar o Streamlit:
# python -m streamlit run 1_dashboard_5_atual.py
# ---
# **Para executar este aplicativo Streamlit:**
# 1. Abra o terminal ou prompt de comando.
# 2. Navegue até o diretório onde você salvou o arquivo.
# 5. Execute o comando: `python -m streamlit run 1_dashboard_4_atual.py`
# Se o Streamlit não estiver instalado, execute: `pip install streamlit pandas openpyxl`
