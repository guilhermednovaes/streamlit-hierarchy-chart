import streamlit as st
import pandas as pd
import plotly.express as px
import io

# Função para carregar o arquivo MANPOWER.xlsx automaticamente com tratamento de exceção
def load_data(sheet_name='09-09'):
    try:
        file_path = 'MANPOWER.xlsx'
        excel_data = pd.ExcelFile(file_path)
        df = excel_data.parse(sheet_name)
        return df
    except FileNotFoundError:
        st.error("Arquivo não encontrado. Verifique o caminho ou o nome do arquivo.")
        return pd.DataFrame()  # Retorna um DataFrame vazio no caso de erro

# Função para criar o gráfico de hierarquia com ajustes de texto dinâmico e zoom interativo
def create_hierarchy_chart(df, filter_function=None):
    if df.empty:
        st.warning("Nenhum dado disponível para gerar o gráfico.")
        return None

    # Preparar os dados de hierarquia
    hierarchy_data = df[['COMPANY', 'PROJECT', 'LEAD', 'INCHARGE SUPERVISOR', 'LEADER', 'EMPLOYEE NAME', 'COMMON FUNCTION', 'EMPLOYEE ID', 'DAILY ATTENDENCE']]
    hierarchy_data = hierarchy_data.drop_duplicates()

    # Aplicar o filtro de função, se houver
    if filter_function:
        hierarchy_data = hierarchy_data[hierarchy_data['COMMON FUNCTION'].isin(filter_function)]

    # Criar um label customizado para incluir nome, função e status de presença no hover
    hierarchy_data['LABEL'] = (hierarchy_data['EMPLOYEE NAME'] + '<br>' + 
                               'Função: ' + hierarchy_data['COMMON FUNCTION'] + '<br>ID: ' + 
                               hierarchy_data['EMPLOYEE ID'].astype(str) + '<br>Status: ' + hierarchy_data['DAILY ATTENDENCE'])

    # Definir cores dinamicamente
    available_functions = hierarchy_data['COMMON FUNCTION'].unique()
    color_map = {func: px.colors.qualitative.Pastel[i % len(px.colors.qualitative.Pastel)] for i, func in enumerate(available_functions)}

    # Criar o gráfico de hierarquia com zoom interativo e exibição de textos dinâmicos
    fig = px.treemap(
        hierarchy_data,
        path=['COMPANY', 'PROJECT', 'LEAD', 'INCHARGE SUPERVISOR', 'LEADER', 'LABEL'],
        color='COMMON FUNCTION',
        color_discrete_map=color_map,
        title="Hierarquia Organizacional com Funções"
    )

    # Ajustar o layout para melhorar a navegação e legibilidade do texto
    fig.update_layout(
        margin=dict(t=20, l=10, r=10, b=20),
        height=1000,  # Ocupar a tela inteira
        hovermode="closest",
        uniformtext_minsize=10,
        uniformtext_mode='hide',
        paper_bgcolor='rgba(0,0,0,0)',  # Fundo transparente
        plot_bgcolor='#FFFFFF',
        font=dict(family="Arial", size=12, color="#000000")
    )

    # Ajuste de bordas e espaçamento entre os cartões
    fig.update_traces(marker=dict(line=dict(color='#000000', width=0.5)))

    return fig

# Função para converter DataFrame para CSV
def convert_df_to_csv(df):
    return df.to_csv(index=False).encode('utf-8')

# Função para converter DataFrame para Excel
def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Dados Filtrados')
    return output.getvalue()

# Função para aplicar filtros de forma modular
def apply_filters(df, filters):
    if 'function' in filters:
        df = df[df['COMMON FUNCTION'].isin(filters['function'])]
    if 'employee_name' in filters:
        df = df[df['EMPLOYEE NAME'].str.contains(filters['employee_name'], case=False, na=False)]
    if 'project_name' in filters:
        df = df[df['PROJECT'].str.contains(filters['project_name'], case=False, na=False)]
    if 'company_name' in filters:
        df = df[df['COMPANY'].str.contains(filters['company_name'], case=False, na=False)]
    if 'supervisor_name' in filters:
        df = df[df['INCHARGE SUPERVISOR'].str.contains(filters['supervisor_name'], case=False, na=False)]
    if 'shift' in filters:
        df = df[df['SHIFT'].isin(filters['shift'])]
    if 'attendance' in filters:
        df = df[df['DAILY ATTENDENCE'].isin(filters['attendance'])]
    return df

# Carregar os dados automaticamente
df = load_data()

# Verificar se o DataFrame foi carregado corretamente
if df.empty:
    st.stop()

# Página de dashboard com abas de gráfico e tabela
st.title("Dashboard de Hierarquia Organizacional")

# Filtros na barra lateral
st.sidebar.title("Filtros de Pesquisa")
st.sidebar.markdown("### Selecione os filtros desejados:")

# Aplicar filtros usando função modular
filters = {}
filters['function'] = st.sidebar.multiselect("Filtrar por Função", options=df['COMMON FUNCTION'].unique(), default=df['COMMON FUNCTION'].unique())
filters['employee_name'] = st.sidebar.text_input("Nome do Funcionário", "")
filters['project_name'] = st.sidebar.text_input("Nome do Projeto", "")
filters['company_name'] = st.sidebar.text_input("Nome da Empresa", "")
filters['supervisor_name'] = st.sidebar.text_input("Nome do Supervisor", "")
filters['shift'] = st.sidebar.multiselect("Filtrar por Turno", options=df['SHIFT'].unique(), default=df['SHIFT'].unique())
filters['attendance'] = st.sidebar.multiselect("Filtrar por Presença", options=df['DAILY ATTENDENCE'].unique(), default=df['DAILY ATTENDENCE'].unique())

# Aplicar os filtros ao DataFrame
df_filtered = apply_filters(df, filters)

# Exibir abas
tab1, tab2 = st.tabs(["Gráfico de Hierarquia", "Tabela de Dados"])

with tab1:
    st.write("### Gráfico de Hierarquia com Cores por Função")
    fig = create_hierarchy_chart(df_filtered, filter_function=filters['function'])
    if fig:
        st.plotly_chart(fig, use_container_width=True)

with tab2:
    st.write("### Tabela de Dados Filtrados")
    st.dataframe(df_filtered)

    # Botão para baixar dados em CSV
    csv = convert_df_to_csv(df_filtered)
    st.download_button(label="📥 Baixar dados filtrados em CSV", data=csv, file_name='dados_filtrados.csv', mime='text/csv')

    # Botão para baixar dados em Excel
    excel = convert_df_to_excel(df_filtered)
    st.download_button(label="📥 Baixar dados filtrados em Excel", data=excel, file_name='dados_filtrados.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
