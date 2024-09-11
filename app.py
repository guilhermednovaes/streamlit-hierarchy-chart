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

# Função para truncar texto (limitar número de caracteres)
def truncate_text(text, max_len=20):
    if len(text) > max_len:
        return text[:max_len] + '...'
    return text

# Função para criar o gráfico de hierarquia com melhorias de agrupamento e navegação
def create_hierarchy_chart(df, filter_function=None):
    # Preparar os dados de hierarquia
    hierarchy_data = df[['COMPANY', 'PROJECT', 'LEAD', 'INCHARGE SUPERVISOR', 'LEADER', 'EMPLOYEE NAME', 'COMMON FUNCTION', 'EMPLOYEE ID', 'DAILY ATTENDENCE']]
    hierarchy_data = hierarchy_data.drop_duplicates()

    # Aplicar o filtro de função, se houver
    if filter_function:
        hierarchy_data = hierarchy_data[hierarchy_data['COMMON FUNCTION'].isin(filter_function)]

    # Agrupar por função comum e contar o número de funcionários por função
    hierarchy_data['EMPLOYEE COUNT'] = hierarchy_data.groupby('COMMON FUNCTION')['EMPLOYEE NAME'].transform('count')
    hierarchy_data['LABEL'] = hierarchy_data.apply(
        lambda row: f"{truncate_text(row['EMPLOYEE NAME'])}<br>Função: {row['COMMON FUNCTION']}<br>ID: {row['EMPLOYEE ID']}<br>Status: {row['DAILY ATTENDENCE']}", axis=1)

    # Criar o gráfico de hierarquia com agrupamento por função
    fig = px.sunburst(
        hierarchy_data,
        path=['COMPANY', 'PROJECT', 'LEAD', 'INCHARGE SUPERVISOR', 'LEADER', 'LABEL'],
        color='COMMON FUNCTION',
        color_discrete_map={  # Paleta de cores refinada
            'WELDER': '#5D6D7E',      # Cinza neutro
            'GRINDER': '#AAB7B8',     # Cinza claro
            'PIPE FITTER': '#3498DB', # Azul elegante
            'LEADER': '#2ECC71',      # Verde sutil
            'SUPERVISOR': '#F4D03F',  # Amarelo suave
            'EMPLOYEE': '#F5B7B1'     # Rosa suave
        },
        maxdepth=6,  # Limitar a profundidade de visualização para manter o gráfico mais limpo
        title="Hierarquia Organizacional com Funções"
    )

    # Ajustar o layout para otimizar a navegação e legibilidade do texto
    fig.update_layout(
        margin=dict(t=20, l=10, r=10, b=20),
        height=1000,  # Ocupar a tela inteira
        hovermode="closest",
        uniformtext_minsize=10,  # Texto dinâmico para caber nos cartões menores
        uniformtext_mode='hide',
        paper_bgcolor='rgba(0,0,0,0)',  # Fundo transparente para visual limpo
        plot_bgcolor='#FFFFFF',  # Fundo branco
        font=dict(family="Arial", size=12, color="#000000")  # Fonte elegante e legível
    )

    # Ajuste de bordas e espaçamento entre os cartões
    fig.update_traces(marker=dict(line=dict(color='#000000', width=0.5)))  # Bordas finas

    # Configurar o hover para exibir nome completo, função e detalhes nos cartões menores
    fig.update_traces(hovertemplate='<b>%{label}</b><extra></extra>', textinfo='label+text')

    return fig

# Função para converter DataFrame para CSV
def convert_df_to_csv(df):
    return df.to_csv(index=False).encode('utf-8')

# Função para converter DataFrame para Excel
def convert_df_to_excel(df):
    output = io.BytesIO()  # Certifique-se de importar o io
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

# Filtro por Nome do Funcionário com busca em tempo real
employee_name = st.sidebar.text_input("Nome do Funcionário (busca em tempo real)", "")
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
selected_attendence = st.sidebar.multiselect("Filtrar por Presença", options=attendence_options, default
