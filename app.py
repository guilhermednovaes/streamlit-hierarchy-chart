# app.py

import streamlit as st
import pandas as pd
import plotly.graph_objects as go

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
    
    # Remove duplicatas para evitar repetição no gráfico
    hierarchy_data = hierarchy_data.drop_duplicates()
    
    # Criar a hierarquia de Company até Employee Name
    fig = go.Figure(go.Sunburst(
        labels=hierarchy_data['EMPLOYEE NAME'].tolist() + hierarchy_data['LEADER'].tolist() + 
               hierarchy_data['INCHARGE SUPERVISOR'].tolist() + hierarchy_data['LEAD'].tolist() +
               hierarchy_data['PROJECT'].tolist() + hierarchy_data['COMPANY'].tolist(),
        parents=[''] * len(hierarchy_data),  # No sunburst need for parent relationship
        maxdepth=6  # Até o nível de EMPLOYEE NAME
    ))

    fig.update_layout(margin=dict(t=0, l=0, r=0, b=0))
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

    # Cria e exibe o gráfico de hierarquia
    st.write("Gráfico de Hierarquia")
    fig = create_hierarchy_chart(df)
    st.plotly_chart(fig)
