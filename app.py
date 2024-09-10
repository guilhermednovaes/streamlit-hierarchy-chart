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
    hierarchy_data = df[['COMPANY', 'PROJECT', 'LEAD', 'INCHARGE SUPERVISOR', 'LEADER', 'EMPLOYEE NAME']]
    
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

    # Cria e exibe o gráfico de hierarquia apenas se houver dados
    if not df.empty:
        st.write("Gráfico de Hierarquia")
        fig = create_hierarchy_chart(df)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.write("Nenhum dado encontrado para os filtros aplicados.")
