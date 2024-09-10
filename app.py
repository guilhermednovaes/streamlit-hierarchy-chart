import streamlit as st
import pandas as pd
import plotly.express as px
import io

# Função para carregar o arquivo MANPOWER.xlsx automaticamente
def load_data():
    file_path = 'MANPOWER.xlsx'
    excel_data = pd.ExcelFile(file_path)
    df = excel_data.parse('09-09')
    return df

# Função para criar o gráfico de hierarquia com nome e função no cartão
def create_hierarchy_chart(df):
    # Preparar os dados de hierarquia
    hierarchy_data = df[['COMPANY', 'PROJECT', 'LEAD', 'INCHARGE SUPERVISOR', 'LEADER', 'EMPLOYEE NAME', 'COMMON FUNCTION']]
    hierarchy_data = hierarchy_data.drop_duplicates()

    # Criar um label customizado para incluir nome e função do funcionário
    hierarchy_data['LABEL'] = hierarchy_data['EMPLOYEE NAME'] + '<br>' + hierarchy_data['COMMON FUNCTION']

    # Criar o gráfico de hierarquia com nome e função
    fig = px.treemap(hierarchy_data,
                     path=['COMPANY', 'PROJECT', 'LEAD', 'INCHARGE SUPERVISOR', 'LEADER', 'LABEL'],
                     color='COMMON FUNCTION',  # Cor baseada na função comum
                     color_discrete_sequence=px.colors.qualitative.Bold,
                     title="Hierarquia Organizacional com Funções")
    
    # Ajustar o layout para otimizar a tela
    fig.update_layout(margin=dict(t=20, l=10, r=10, b=10),
                      height=800,  # Ajuste para altura do gráfico, ocupando a tela inteira
                      hovermode="closest",
                      uniformtext_minsize=14,  # Ajustar tamanho do texto
                      uniformtext_mode='hide')

    # Configurar o hover para exibir apenas nome e função
    fig.update_traces(hovertemplate='<b>%{label}</b><extra></extra>')

    return fig

# Função para converter DataFrame para CSV
def convert_df_to_csv(df):
    return df.to_csv(index=False).encode('utf-8')

# Função para converter DataFrame para Excel
def convert_df_to_excel(df):
    output = io.BytesIO()
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

    # Legenda separada
    st.write("#### Legenda de Cores:")
    st.markdown("""
    - **Funções Representadas pelas Cores**:
        - Azul: LEAD
        - Verde: INCHARGE SUPERVISOR
        - Vermelho: LEADER
        - Amarelo: EMPLOYEE NAME (Função baseada na coluna `COMMON FUNCTION`)
    """)

with tab2:
    st.write("### Tabela de Dados Filtrados")
    st.dataframe(df)

    # Botão para baixar dados em CSV
    csv = convert_df_to_csv(df)
    st.download_button(
        label="📥 Baixar dados filtrados em CSV",
        data=csv,
        file_name='dados_filtrados.csv',
        mime='text/csv',
    )

    # Botão para baixar dados em Excel
    excel = convert_df_to_excel(df)
    st.download_button(
        label="📥 Baixar dados filtrados em Excel",
        data=excel,
        file_name='dados_filtrados.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
