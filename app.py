import streamlit as st
import pandas as pd
import plotly.express as px
import io  # Certifique-se de importar o io para trabalhar com o BytesIO

# Função para carregar o arquivo MANPOWER.xlsx automaticamente
def load_data():
    file_path = 'MANPOWER.xlsx'
    excel_data = pd.ExcelFile(file_path)
    df = excel_data.parse('09-09')
    return df

# Função para criar o gráfico de hierarquia com melhorias
def create_hierarchy_chart(df, filter_function=None):
    # Preparar os dados de hierarquia
    hierarchy_data = df[['COMPANY', 'PROJECT', 'LEAD', 'INCHARGE SUPERVISOR', 'LEADER', 'EMPLOYEE NAME', 'COMMON FUNCTION', 'EMPLOYEE ID', 'DAILY ATTENDENCE']]
    hierarchy_data = hierarchy_data.drop_duplicates()

    # Aplicar o filtro de função, se houver
    if filter_function:
        hierarchy_data = hierarchy_data[hierarchy_data['COMMON FUNCTION'].isin(filter_function)]

    # Criar um label customizado para incluir nome, função e status de presença no hover
    hierarchy_data['LABEL'] = hierarchy_data['EMPLOYEE NAME'] + '<br>' + 'Função: ' + hierarchy_data['COMMON FUNCTION'] + '<br>ID: ' + hierarchy_data['EMPLOYEE ID'].astype(str) + '<br>Status: ' + hierarchy_data['DAILY ATTENDENCE']

    # Criar o gráfico de hierarquia com zoom progressivo
    fig = px.treemap(hierarchy_data,
                     path=['COMPANY', 'PROJECT', 'LEAD', 'INCHARGE SUPERVISOR', 'LEADER', 'LABEL'],
                     color='COMMON FUNCTION',  # Cor baseada na função comum
                     color_discrete_sequence=px.colors.qualitative.Safe,  # Cores suaves e harmoniosas
                     title="Hierarquia Organizacional com Funções")

    # Ajustar o layout para otimizar a tela e zoom progressivo
    fig.update_layout(margin=dict(t=20, l=10, r=10, b=10),
                      height=800,  # Ajuste para altura do gráfico, ocupando a tela inteira
                      hovermode="closest",
                      uniformtext_minsize=14,  # Ajustar tamanho do texto
                      uniformtext_mode='hide')

    # Configurar o hover para exibir nome, função e detalhes
    fig.update_traces(hovertemplate='<b>%{label}</b><extra></extra>')

    return fig

# Função para converter DataFrame para CSV
def convert_df_to_csv(df):
    return df.to_csv(index=False).encode('utf-8')

# Função para converter DataFrame para Excel
def convert_df_to_excel(df):
    output = io.BytesIO()  # Corrigido o uso do io
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Dados Filtrados')
    processed_data = output.getvalue()
    return processed_data

# Carregar os dados automaticamente
df = load_data()

# Página de dashboard com abas de gráfico e tabela
st.title("Dashboard de Hierarquia Organizacional")

# Filtros na barra lateral
st.sidebar.title("Filtros de Pesquisa")
st.sidebar.markdown("### Selecione os filtros desejados:")

# Filtro de Função (Modo Foco)
common_functions = df['COMMON FUNCTION'].unique()
selected_common_functions = st.sidebar.multiselect("Filtrar por Função (Common Function)", options=common_functions, default=common_functions)
df_filtered = df[df['COMMON FUNCTION'].isin(selected_common_functions)]

# Filtro por Nome do Funcionário
employee_name = st.sidebar.text_input("Nome do Funcionário", "")
if employee_name:
    df_filtered = df_filtered[df_filtered['EMPLOYEE NAME'].str.contains(employee_name, case=False, na=False)]

# Filtro por Projeto
project_name = st.sidebar.text_input("Nome do Projeto", "")
if project_name:
    df_filtered = df_filtered[df_filtered['PROJECT'].str.contains(project_name, case=False, na=False)]

# Filtro por Empresa
company_name = st.sidebar.text_input("Nome da Empresa", "")
if company_name:
    df_filtered = df_filtered[df_filtered['COMPANY'].str.contains(company_name, case=False, na=False)]

# Filtro por Supervisor Responsável
supervisor_name = st.sidebar.text_input("Nome do Supervisor", "")
if supervisor_name:
    df_filtered = df_filtered[df_filtered['INCHARGE SUPERVISOR'].str.contains(supervisor_name, case=False, na=False)]

# Filtro por Turno (MULTI-SELEÇÃO)
shifts = df_filtered['SHIFT'].unique()
selected_shifts = st.sidebar.multiselect("Filtrar por Turno", options=shifts, default=shifts)
if selected_shifts:
    df_filtered = df_filtered[df_filtered['SHIFT'].isin(selected_shifts)]

# Filtro por Presença (MULTI-SELEÇÃO) com valores em inglês
attendence_options = df_filtered['DAILY ATTENDENCE'].unique()
selected_attendence = st.sidebar.multiselect("Filtrar por Presença", options=attendence_options, default=attendence_options)
if selected_attendence:
    df_filtered = df_filtered[df_filtered['DAILY ATTENDENCE'].isin(selected_attendence)]

# Exibe abas
tab1, tab2 = st.tabs(["Gráfico de Hierarquia", "Tabela de Dados"])

with tab1:
    st.write("### Gráfico de Hierarquia com Cores por Função")
    fig = create_hierarchy_chart(df_filtered, filter_function=selected_common_functions)
    st.plotly_chart(fig, use_container_width=True)

    # Legenda separada
    st.write("#### Legenda de Cores:")
    st.markdown("""
    - **Funções Representadas pelas Cores**:
        - Cada cor representa uma função específica.
    """)

with tab2:
    st.write("### Tabela de Dados Filtrados")
    st.dataframe(df_filtered)

    # Botão para baixar dados em CSV
    csv = convert_df_to_csv(df_filtered)
    st.download_button(
        label="📥 Baixar dados filtrados em CSV",
        data=csv,
        file_name='dados_filtrados.csv',
        mime='text/csv',
    )

    # Botão para baixar dados em Excel
    excel = convert_df_to_excel(df_filtered)
    st.download_button(
        label="📥 Baixar dados filtrados em Excel",
        data=excel,
        file_name='dados_filtrados.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
