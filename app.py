import streamlit as st
import pandas as pd
import plotly.express as px
import time

# Função para carregar os dados da planilha
def load_data(file):
    with st.spinner('Carregando os dados...'):
        time.sleep(1)  # Simulação de carregamento para UX
        excel_data = pd.ExcelFile(file)
        df = excel_data.parse('09-09')
    return df

# Função para criar o gráfico de hierarquia com legenda e cores baseadas na função
def create_hierarchy_chart(df):
    # Agrupar na última camada por "COMMON FUNCTION"
    hierarchy_data = df[['COMPANY', 'PROJECT', 'LEAD', 'INCHARGE SUPERVISOR', 'LEADER', 'EMPLOYEE NAME', 'COMMON FUNCTION']]
    hierarchy_data = hierarchy_data.drop_duplicates()

    fig = px.treemap(hierarchy_data,
                     path=['COMPANY', 'PROJECT', 'LEAD', 'INCHARGE SUPERVISOR', 'LEADER', 'COMMON FUNCTION', 'EMPLOYEE NAME'],
                     color='COMMON FUNCTION',
                     color_discrete_sequence=px.colors.qualitative.Bold,
                     title="Hierarquia Organizacional com Funções")
    
    fig.update_layout(margin=dict(t=50, l=25, r=25, b=25), 
                      hovermode="closest",
                      uniformtext_minsize=10, 
                      uniformtext_mode='hide')
    return fig

# Verifica se o arquivo foi carregado
if "file_uploaded" not in st.session_state:
    st.session_state["file_uploaded"] = False

# Página de upload de arquivo
if not st.session_state["file_uploaded"]:
    st.title("Carregar o Arquivo Excel")
    
    # Design mais moderno para o botão de upload
    uploaded_file = st.file_uploader("📁 Faça upload de um arquivo Excel", type="xlsx")

    if uploaded_file is not None:
        st.session_state["file_uploaded"] = True
        st.session_state["data"] = load_data(uploaded_file)
        st.experimental_set_query_params(page="dashboard")

# Página de dashboard com abas de gráfico e tabela
if st.session_state["file_uploaded"]:
    st.title("Dashboard de Hierarquia Organizacional")
    
    # Filtros na barra lateral
    st.sidebar.title("Filtros de Pesquisa")
    st.sidebar.markdown("### Selecione os filtros desejados:")

    df = st.session_state["data"]

    # Filtro por Nome do Funcionário
    employee_name = st.sidebar.text_input("Nome do Funcionário", "")
    if employee_name:
        df = df[df['EMPLOYEE NAME'].str.contains(employee_name, case=False, na=False)]

    # Filtro por Projeto
    project_name = st.sidebar.text_input("Nome do Projeto", "")
    if project_name:
        df = df[df['PROJECT'].str.contains(project_name, case=False, na=False)]

    # Filtro por Empresa
    company_name = st.sidebar.text_input("Nome da Empresa", "")
    if company_name:
        df = df[df['COMPANY'].str.contains(company_name, case=False, na=False)]

    # Filtro por Supervisor Responsável
    supervisor_name = st.sidebar.text_input("Nome do Supervisor", "")
    if supervisor_name:
        df = df[df['INCHARGE SUPERVISOR'].str.contains(supervisor_name, case=False, na=False)]

    # Filtro por Função (MULTI-SELEÇÃO) usando a coluna 'COMMON FUNCTION'
    if not df['COMMON FUNCTION'].isnull().all():
        common_functions = df['COMMON FUNCTION'].unique()
        selected_common_functions = st.sidebar.multiselect("Filtrar por Função (Common Function)", options=common_functions, default=common_functions)
        if selected_common_functions:
            df = df[df['COMMON FUNCTION'].isin(selected_common_functions)]

    # Filtro por Turno (MULTI-SELEÇÃO)
    if not df['SHIFT'].isnull().all():
        shifts = df['SHIFT'].unique()
        selected_shifts = st.sidebar.multiselect("Filtrar por Turno", options=shifts, default=shifts)
        if selected_shifts:
            df = df[df['SHIFT'].isin(selected_shifts)]

    # Filtro por Presença (MULTI-SELEÇÃO) com valores em inglês
    if not df['DAILY ATTENDENCE'].isnull().all():
        attendence_options = df['DAILY ATTENDENCE'].unique()
        selected_attendence = st.sidebar.multiselect("Filtrar por Presença", options=attendence_options, default=attendence_options)
        if selected_attendence:
            df = df[df['DAILY ATTENDENCE'].isin(selected_attendence)]

    # Converter "EMPLOYEE ID" para numérico, ignorando erros, e remover valores NaN
    df['EMPLOYEE ID'] = pd.to_numeric(df['EMPLOYEE ID'], errors='coerce')
    df = df.dropna(subset=['EMPLOYEE ID'])

    # Verifica se há IDs válidos para o slider
    if not df['EMPLOYEE ID'].isnull().all():
        min_id = int(df['EMPLOYEE ID'].min())
        max_id = int(df['EMPLOYEE ID'].max())
        employee_id_range = st.sidebar.slider("Intervalo de ID do Funcionário", min_value=min_id, max_value=max_id, value=(min_id, max_id))
        df = df[(df['EMPLOYEE ID'] >= employee_id_range[0]) & (df['EMPLOYEE ID'] <= employee_id_range[1])]

    # Exibe abas
    tab1, tab2 = st.tabs(["Gráfico de Hierarquia", "Tabela de Dados"])

    with tab1:
        st.write("### Gráfico de Hierarquia com Cores por Função")
        fig = create_hierarchy_chart(df)
        st.plotly_chart(fig, use_container_width=True)

        # Legenda
        st.write("#### Legenda:")
        st.markdown("""
        - **Funções em Cores**:
            - Cada cor no gráfico representa uma função diferente baseada na coluna `COMMON FUNCTION`.
        """)

    with tab2:
        st.write("### Tabela de Dados Filtrados")
        st.dataframe(df)
