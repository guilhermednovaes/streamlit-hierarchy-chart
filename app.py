import streamlit as st
import pandas as pd
import plotly.express as px

# Função para carregar os dados da planilha
def load_data(file):
    # Carrega a planilha
    excel_data = pd.ExcelFile(file)
    # Lê a aba '09-09'
    df = excel_data.parse('09-09')
    return df

# Função para criar o gráfico de hierarquia
def create_hierarchy_chart(df):
    # Filtra as colunas de interesse
    hierarchy_data = df[['COMPANY', 'PROJECT', 'LEAD', 'INCHARGE SUPERVISOR', 'LEADER', 'EMPLOYEE NAME', 'FUNCTION', 'SHIFT', 'DAILY ATTENDENCE']]
    
    # Remove duplicatas
    hierarchy_data = hierarchy_data.drop_duplicates()

    # Cria um gráfico de hierarquia do tipo "treemap"
    fig = px.treemap(hierarchy_data,
                     path=['COMPANY', 'PROJECT', 'LEAD', 'INCHARGE SUPERVISOR', 'LEADER', 'EMPLOYEE NAME'],
                     values=None,  # Sem valores associados, apenas visualização hierárquica
                     title="Hierarquia Organizacional")
    
    # Melhora a aparência do gráfico
    fig.update_layout(margin=dict(t=50, l=25, r=25, b=25))
    
    return fig

# Configuração do site no Streamlit
st.title("Visualização da Hierarquia Organizacional")

# Upload do arquivo Excel
uploaded_file = st.file_uploader("Faça upload de um arquivo Excel", type="xlsx")

if uploaded_file is not None:
    # Carrega os dados
    df = load_data(uploaded_file)
    
    st.write("Exibindo as primeiras linhas da planilha:")
    st.dataframe(df.head())

    # Adicionar filtros para a pesquisa
    st.sidebar.title("Filtros de Pesquisa")
    st.sidebar.markdown("### Selecione os filtros desejados:")

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

    # Filtro por Função (MULTI-SELEÇÃO)
    if not df['FUNCTION'].isnull().all():
        functions = df['FUNCTION'].unique()
        selected_functions = st.sidebar.multiselect("Filtrar por Função", options=functions, default=functions)
        if selected_functions:
            df = df[df['FUNCTION'].isin(selected_functions)]
    
    # Filtro por Turno (MULTI-SELEÇÃO)
    if not df['SHIFT'].isnull().all():
        shifts = df['SHIFT'].unique()
        selected_shifts = st.sidebar.multiselect("Filtrar por Turno", options=shifts, default=shifts)
        if selected_shifts:
            df = df[df['SHIFT'].isin(selected_shifts)]

    # Filtro por Presença (SELEÇÃO ÚNICA)
    if not df['DAILY ATTENDENCE'].isnull().all():
        attendence_status = st.sidebar.selectbox("Filtrar por Presença", options=["Todos", "Presente", "Ausente", "Emprestado"], index=0)
        if attendence_status != "Todos":
            df = df[df['DAILY ATTENDENCE'].str.contains(attendence_status, case=False, na=False)]

    # Converter "EMPLOYEE ID" para numérico, ignorando erros, e remover valores NaN
    df['EMPLOYEE ID'] = pd.to_numeric(df['EMPLOYEE ID'], errors='coerce')
    df = df.dropna(subset=['EMPLOYEE ID'])
    
    # Verifica se há IDs válidos para o slider
    if not df['EMPLOYEE ID'].isnull().all():
        # Filtro por intervalo de ID do Empregado (SLIDER)
        min_id = int(df['EMPLOYEE ID'].min())
        max_id = int(df['EMPLOYEE ID'].max())
        employee_id_range = st.sidebar.slider("Intervalo de ID do Funcionário", min_value=min_id, max_value=max_id, value=(min_id, max_id))
        df = df[(df['EMPLOYEE ID'] >= employee_id_range[0]) & (df['EMPLOYEE ID'] <= employee_id_range[1])]
    else:
        st.sidebar.write("Nenhum ID de Funcionário disponível para o filtro.")

    # Exibe o gráfico de hierarquia se houver dados
    if not df.empty:
        st.write("### Gráfico de Hierarquia")
        fig = create_hierarchy_chart(df)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.write("### Nenhum dado encontrado para os filtros aplicados.")
else:
    st.write("Por favor, faça upload de um arquivo Excel para continuar.")
