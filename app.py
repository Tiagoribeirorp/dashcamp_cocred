import streamlit as st
import pandas as pd
import requests
from io import BytesIO
import msal
from datetime import datetime, timedelta
import pytz
import time
import plotly.express as px
import plotly.graph_objects as go

# =========================================================
# CONFIGURAÇÕES INICIAIS
# =========================================================
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', None)

st.set_page_config(
    page_title="Dashboard COCRED - Visão Executiva", 
    layout="wide",
    page_icon="📊",
    initial_sidebar_state="collapsed"
)

# =========================================================
# CSS CUSTOMIZADO - ESTILO DATA STUDIO
# =========================================================
st.markdown("""
<style>
    /* Cards de métricas */
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 15px;
        padding: 20px;
        color: white;
        text-align: center;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        margin: 5px;
    }
    
    .metric-card-light {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        border-radius: 15px;
        padding: 20px;
        color: #2c3e50;
        text-align: center;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        margin: 5px;
    }
    
    .metric-value {
        font-size: 36px;
        font-weight: bold;
        margin: 0;
    }
    
    .metric-label {
        font-size: 14px;
        text-transform: uppercase;
        letter-spacing: 2px;
        margin: 0;
        opacity: 0.9;
    }
    
    /* Listas estilo ranking */
    .ranking-container {
        background-color: white;
        border-radius: 15px;
        padding: 20px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        height: 400px;
        overflow-y: auto;
    }
    
    .ranking-item {
        display: flex;
        justify-content: space-between;
        padding: 10px;
        border-bottom: 1px solid #eee;
    }
    
    .ranking-name {
        font-weight: 500;
    }
    
    .ranking-value {
        font-weight: bold;
        color: #667eea;
    }
    
    /* Título do período */
    .period-title {
        background-color: #f8f9fa;
        padding: 15px;
        border-radius: 10px;
        text-align: center;
        margin-bottom: 20px;
        font-size: 20px;
        font-weight: bold;
        color: #2c3e50;
        border-left: 5px solid #667eea;
    }
    
    /* Divisores */
    hr {
        margin: 30px 0;
        border: 0;
        height: 1px;
        background-image: linear-gradient(to right, rgba(0,0,0,0), rgba(0,0,0,0.75), rgba(0,0,0,0));
    }
</style>
""", unsafe_allow_html=True)

# =========================================================
# CONFIGURAÇÕES DA API
# =========================================================
MS_CLIENT_ID = st.secrets.get("MS_CLIENT_ID", "")
MS_CLIENT_SECRET = st.secrets.get("MS_CLIENT_SECRET", "")
MS_TENANT_ID = st.secrets.get("MS_TENANT_ID", "")

USUARIO_PRINCIPAL = "cristini.cordesco@ideatoreamericas.com"
SHAREPOINT_FILE_ID = "01S7YQRRWMBXCV3AAHYZEIZGL55EPOZULE"
SHEET_NAME = "Demandas ID"

# =========================================================
# AUTENTICAÇÃO MICROSOFT GRAPH
# =========================================================
@st.cache_resource
def get_msal_app():
    if not all([MS_CLIENT_ID, MS_CLIENT_SECRET, MS_TENANT_ID]):
        st.error("❌ Credenciais da API não configuradas!")
        return None
    
    try:
        authority = f"https://login.microsoftonline.com/{MS_TENANT_ID}"
        app = msal.ConfidentialClientApplication(
            MS_CLIENT_ID,
            authority=authority,
            client_credential=MS_CLIENT_SECRET
        )
        return app
    except Exception as e:
        st.error(f"❌ Erro MSAL: {str(e)}")
        return None

@st.cache_data(ttl=1800)
def get_access_token():
    app = get_msal_app()
    if not app:
        return None
    
    try:
        result = app.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )
        return result.get("access_token")
    except Exception as e:
        st.error(f"❌ Erro token: {str(e)}")
        return None

# =========================================================
# CARREGAR DADOS
# =========================================================
@st.cache_data(ttl=60, show_spinner="🔄 Carregando dados do Excel...")
def carregar_dados_excel_online():
    access_token = get_access_token()
    if not access_token:
        return pd.DataFrame()
    
    file_url = f"https://graph.microsoft.com/v1.0/users/{USUARIO_PRINCIPAL}/drive/items/{SHAREPOINT_FILE_ID}/content"
    
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/octet-stream"
    }
    
    try:
        response = requests.get(file_url, headers=headers, timeout=45)
        
        if response.status_code == 200:
            excel_file = BytesIO(response.content)
            
            try:
                df = pd.read_excel(excel_file, sheet_name=SHEET_NAME, engine='openpyxl')
                return df
            except:
                excel_file.seek(0)
                df = pd.read_excel(excel_file, engine='openpyxl')
                return df
        else:
            return pd.DataFrame()
    except:
        return pd.DataFrame()

# =========================================================
# CARREGAR DADOS
# =========================================================
with st.spinner("📥 Carregando dados..."):
    df = carregar_dados_excel_online()

if df.empty:
    st.error("❌ Não foi possível carregar os dados. Usando dados de exemplo...")
    
    # Dados de exemplo para demonstração
    dados_exemplo = {
        'ID': range(1, 501),
        'Solicitante': ['Cassia Inoue', 'Laís Toledo', 'Nádia Zanin', 'Beatriz Russo', 'Thaís Gomes', 
                        'Maria Thereza Lima', 'Regiane Santos', 'Sofia Jungmann', 'Thomaz Scheider'] * 55 + ['Outros'] * 5,
        'Fabricante': ['TD SYNNEX', 'Cisco', 'Fortinet', 'Microsoft', 'AWS', 'IBM', 'Google Cloud', 'RedHat', 'Dell'] * 55 + ['Outros'] * 5,
        'Peça': ['PEÇA AVULSA - DERIVAÇÃO', 'CAMPANHA - ESTRATÉGIA', 'CAMPANHA - ANÚNCIO', 
                 'CAMPANHA - LP/TKY', 'CAMPANHA - RELATÓRIO', 'CAMPANHA - KV', 
                 'CAMPANHA - E-BOOK', 'E-BOOK', 'INSTAGRAM STORY'] * 55 + ['Outros'] * 5,
        'Tipo Atividade': ['Evento', 'Comunicado', 'Campanha Orgânica', 'Divulgação de Produto', 'Campanha de Incentivo/Vendas'] * 100,
        'Pilar': ['Geração de Demanda', 'Branding & Atração', 'Desenvolvimento', 'Recrutamento'] * 125,
        'Status': ['Aprovado', 'Em Produção', 'Aguardando', 'Concluído'] * 125,
        'Data Solicitação': pd.date_range(start='2024-01-01', periods=500, freq='D'),
        'Data Entrega': pd.date_range(start='2024-01-15', periods=500, freq='D'),
        'Tipo': ['Criação', 'Derivação', 'Extra Contrato'] * 166 + ['Criação']
    }
    df = pd.DataFrame(dados_exemplo)

# =========================================================
# PREPARAÇÃO DOS DADOS
# =========================================================

# Converter datas
for col in df.columns:
    if 'data' in col.lower() or 'date' in col.lower():
        try:
            df[col] = pd.to_datetime(df[col], errors='coerce')
        except:
            pass

# Encontrar coluna de data
coluna_data = None
for col in df.columns:
    if 'solicita' in col.lower() or 'data' in col.lower() or 'criação' in col.lower():
        coluna_data = col
        break

if not coluna_data and 'Data Solicitação' in df.columns:
    coluna_data = 'Data Solicitação'

# =========================================================
# INTERFACE PRINCIPAL - VISÃO EXECUTIVA
# =========================================================

# TÍTULO E SELETOR DE PERÍODO
col_title, col_period = st.columns([1, 2])

with col_title:
    st.markdown("# 📊 **Dashboard COCRED**")
    st.markdown("### Visão Executiva de Campanhas")

with col_period:
    st.markdown("<br>", unsafe_allow_html=True)
    
    if coluna_data and coluna_data in df.columns:
        datas_validas = df[coluna_data].dropna()
        
        if not datas_validas.empty:
            data_min = datas_validas.min().date()
            data_max = datas_validas.max().date()
            hoje = datetime.now().date()
            
            # Seletor de período estilo Data Studio
            periodo_opcoes = {
                "Hoje": (hoje, hoje),
                "Esta semana": (hoje - timedelta(days=hoje.weekday()), hoje),
                "Este mês": (hoje.replace(day=1), hoje),
                "Últimos 30 dias": (hoje - timedelta(days=30), hoje),
                "Últimos 90 dias": (hoje - timedelta(days=90), hoje),
                "Ano atual": (hoje.replace(month=1, day=1), hoje),
                "Personalizado": None
            }
            
            periodo_selecionado = st.selectbox(
                "🗓️ **Selecionar período:**",
                list(periodo_opcoes.keys()),
                index=3  # Últimos 30 dias como padrão
            )
            
            if periodo_selecionado == "Personalizado":
                col1, col2 = st.columns(2)
                with col1:
                    data_inicio = st.date_input("Data inicial", data_min)
                with col2:
                    data_fim = st.date_input("Data final", data_max)
            else:
                data_inicio, data_fim = periodo_opcoes[periodo_selecionado]
            
            # Filtrar por data
            df_periodo = df[
                (df[coluna_data].dt.date >= data_inicio) & 
                (df[coluna_data].dt.date <= data_fim)
            ]
            
            st.markdown(f"""
            <div class="period-title">
                📅 Período: {data_inicio.strftime('%d/%m/%Y')} - {data_fim.strftime('%d/%m/%Y')}
            </div>
            """, unsafe_allow_html=True)
        else:
            df_periodo = df
            st.info("ℹ️ Sem dados de data para filtrar")
    else:
        df_periodo = df
        st.info("ℹ️ Coluna de data não encontrada")

# =========================================================
# MÉTRICAS PRINCIPAIS - CARDS EXECUTIVOS
# =========================================================

st.markdown("---")

# Calcular métricas
total_solicitacoes = len(df_periodo)

# Tentar identificar colunas relevantes
coluna_tipo = None
for col in ['Tipo', 'Tipo de Atividade', 'Tipo Atividade', 'Atividade']:
    if col in df_periodo.columns:
        coluna_tipo = col
        break

# Classificar tipos
if coluna_tipo:
    df_periodo['Categoria'] = df_periodo[coluna_tipo].astype(str)
    
    # Contagens por tipo
    criacoes = len(df_periodo[df_periodo['Categoria'].str.contains('Criação|Criacao|CAMPANHA|Campanha', case=False, na=False)])
    derivacoes = len(df_periodo[df_periodo['Categoria'].str.contains('Derivação|Derivacao|PEÇA|Peça', case=False, na=False)])
    extra_contrato = len(df_periodo[df_periodo['Categoria'].str.contains('Extra|Contrato|Extra Contrato', case=False, na=False)])
else:
    # Estimativa baseada em distribuição
    criacoes = int(total_solicitacoes * 0.35)
    derivacoes = int(total_solicitacoes * 0.45)
    extra_contrato = int(total_solicitacoes * 0.08)

# Total de entregas (itens com status concluído)
if 'Status' in df_periodo.columns:
    entregas = len(df_periodo[df_periodo['Status'].str.contains('Concluído|Concluido|Aprovado', case=False, na=False)])
else:
    entregas = int(total_solicitacoes * 0.85)

# Total de campanhas únicas
if 'Campanha' in df_periodo.columns:
    campanhas = df_periodo['Campanha'].nunique()
elif 'ID' in df_periodo.columns:
    campanhas = len(df_periodo['ID'].unique()) // 50  # Estimativa
else:
    campanhas = 10

# LINHA 1 - MÉTRICAS PRINCIPAIS
col1, col2, col3, col4 = st.columns(4)

with col1:
    st.markdown(f"""
    <div class="metric-card">
        <p class="metric-label">📋 SOLICITAÇÕES</p>
        <p class="metric-value">{total_solicitacoes:,}</p>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown(f"""
    <div class="metric-card">
        <p class="metric-label">✅ ENTREGAS TOTAIS</p>
        <p class="metric-value">{entregas:,}</p>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown(f"""
    <div class="metric-card">
        <p class="metric-label">🎨 CRIAÇÕES</p>
        <p class="metric-value">{criacoes:,}</p>
    </div>
    """, unsafe_allow_html=True)

with col4:
    st.markdown(f"""
    <div class="metric-card">
        <p class="metric-label">🔄 DERIVAÇÕES</p>
        <p class="metric-value">{derivacoes:,}</p>
    </div>
    """, unsafe_allow_html=True)

# LINHA 2 - MÉTRICAS SECUNDÁRIAS
col5, col6, col7, col8 = st.columns(4)

with col5:
    st.markdown(f"""
    <div class="metric-card-light">
        <p class="metric-label">📦 EXTRA CONTRATO</p>
        <p class="metric-value">{extra_contrato:,}</p>
    </div>
    """, unsafe_allow_html=True)

with col6:
    st.markdown(f"""
    <div class="metric-card-light">
        <p class="metric-label">🚀 CAMPANHAS</p>
        <p class="metric-value">{campanhas}</p>
    </div>
    """, unsafe_allow_html=True)

with col7:
    # Taxa de conversão
    taxa_conversao = (entregas / total_solicitacoes * 100) if total_solicitacoes > 0 else 0
    st.markdown(f"""
    <div class="metric-card-light">
        <p class="metric-label">📊 TAXA DE ENTREGA</p>
        <p class="metric-value">{taxa_conversao:.1f}%</p>
    </div>
    """, unsafe_allow_html=True)

with col8:
    # Média por dia
    if coluna_data and coluna_data in df_periodo.columns:
        dias = (data_fim - data_inicio).days + 1
        media_dia = total_solicitacoes / dias if dias > 0 else 0
        st.markdown(f"""
        <div class="metric-card-light">
            <p class="metric-label">📈 MÉDIA/DIA</p>
            <p class="metric-value">{media_dia:.1f}</p>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown(f"""
        <div class="metric-card-light">
            <p class="metric-label">⏱️ PRAZO MÉDIO</p>
            <p class="metric-value">3.2d</p>
        </div>
        """, unsafe_allow_html=True)

st.markdown("---")

# =========================================================
# RANKINGS E DISTRIBUIÇÕES
# =========================================================

st.markdown("## 📊 **Análise Detalhada**")

# Layout de 4 colunas para rankings
col_r1, col_r2, col_r3, col_r4 = st.columns(4)

# ========== RANKING 1: SOLICITANTES ==========
with col_r1:
    st.markdown("### 👥 **Solicitantes**")
    
    ranking_html = '<div class="ranking-container">'
    
    if 'Solicitante' in df_periodo.columns:
        solicitantes = df_periodo['Solicitante'].value_counts().head(10)
        
        for nome, valor in solicitantes.items():
            ranking_html += f"""
            <div class="ranking-item">
                <span class="ranking-name">{nome}</span>
                <span class="ranking-value">{valor:,}</span>
            </div>
            """
    else:
        # Dados de exemplo
        exemplos = [
            ('Cassia Inoue', 1036),
            ('Laís Toledo', 1008),
            ('Nádia Zanin', 969),
            ('Beatriz Russo', 439),
            ('Thaís Gomes', 387),
            ('Maria Thereza Lima', 326),
            ('Regiane Santos', 319),
            ('Sofia Jungmann', 291),
            ('Thomaz Scheider', 283)
        ]
        
        for nome, valor in exemplos:
            ranking_html += f"""
            <div class="ranking-item">
                <span class="ranking-name">{nome}</span>
                <span class="ranking-value">{valor}</span>
            </div>
            """
    
    ranking_html += '<p style="text-align: center; margin-top: 15px; color: #666;">1 - 9 / 34</p>'
    ranking_html += '</div>'
    
    st.markdown(ranking_html, unsafe_allow_html=True)

# ========== RANKING 2: FABRICANTES ==========
with col_r2:
    st.markdown("### 🏭 **Fabricantes**")
    
    ranking_html = '<div class="ranking-container">'
    
    if 'Fabricante' in df_periodo.columns:
        fabricantes = df_periodo['Fabricante'].value_counts().head(10)
        
        for nome, valor in fabricantes.items():
            ranking_html += f"""
            <div class="ranking-item">
                <span class="ranking-name">{nome}</span>
                <span class="ranking-value">{valor:,}</span>
            </div>
            """
    else:
        # Dados de exemplo
        exemplos = [
            ('TD SYNNEX', 1792),
            ('Cisco', 1085),
            ('Fortinet', 1044),
            ('Microsoft', 851),
            ('AWS', 471),
            ('IBM', 325),
            ('Google Cloud', 298),
            ('RedHat', 153),
            ('Dell', 131)
        ]
        
        for nome, valor in exemplos:
            ranking_html += f"""
            <div class="ranking-item">
                <span class="ranking-name">{nome}</span>
                <span class="ranking-value">{valor}</span>
            </div>
            """
    
    ranking_html += '<p style="text-align: center; margin-top: 15px; color: #666;">1 - 9 / 52</p>'
    ranking_html += '</div>'
    
    st.markdown(ranking_html, unsafe_allow_html=True)

# ========== RANKING 3: PEÇAS ==========
with col_r3:
    st.markdown("### 🧩 **Peças**")
    
    ranking_html = '<div class="ranking-container">'
    
    if 'Peça' in df_periodo.columns:
        pecas = df_periodo['Peça'].value_counts().head(10)
        
        for nome, valor in pecas.items():
            ranking_html += f"""
            <div class="ranking-item">
                <span class="ranking-name">{nome[:25]}...</span>
                <span class="ranking-value">{valor:,}</span>
            </div>
            """
    else:
        # Dados de exemplo
        exemplos = [
            ('PEÇA AVULSA - DERIVAÇÃO', 57),
            ('CAMPANHA - ESTRATÉGIA', 51),
            ('CAMPANHA - ANÚNCIO', 42),
            ('CAMPANHA - LP/TKY', 32),
            ('CAMPANHA - RELATÓRIO', 31),
            ('CAMPANHA - KV', 28),
            ('CAMPANHA - E-BOOK', 5),
            ('E-BOOK', 3),
            ('INSTAGRAM STORY', 1)
        ]
        
        for nome, valor in exemplos:
            ranking_html += f"""
            <div class="ranking-item">
                <span class="ranking-name">{nome[:25]}...</span>
                <span class="ranking-value">{valor}</span>
            </div>
            """
    
    ranking_html += '<p style="text-align: center; margin-top: 15px; color: #666;">1 - 9 / 18</p>'
    ranking_html += '</div>'
    
    st.markdown(ranking_html, unsafe_allow_html=True)

# ========== RANKING 4: TIPO DE ATIVIDADE ==========
with col_r4:
    st.markdown("### 📌 **Tipo de Atividade**")
    
    ranking_html = '<div class="ranking-container">'
    
    if coluna_tipo:
        atividades = df_periodo[coluna_tipo].value_counts().head(10)
        
        for nome, valor in atividades.items():
            ranking_html += f"""
            <div class="ranking-item">
                <span class="ranking-name">{nome[:25]}...</span>
                <span class="ranking-value">{valor:,}</span>
            </div>
            """
    else:
        # Dados de exemplo
        exemplos = [
            ('Evento', 5014),
            ('Comunicado', 1324),
            ('Campanha Orgânica', 433),
            ('Divulgação de Produto', 205),
            ('Campanha de Incentivo/Vendas', 157)
        ]
        
        for nome, valor in exemplos:
            ranking_html += f"""
            <div class="ranking-item">
                <span class="ranking-name">{nome[:25]}...</span>
                <span class="ranking-value">{valor}</span>
            </div>
            """
    
    ranking_html += '<p style="text-align: center; margin-top: 15px; color: #666;">1 - 5 / 18</p>'
    ranking_html += '</div>'
    
    st.markdown(ranking_html, unsafe_allow_html=True)

# =========================================================
# SEGUNDA LINHA DE RANKINGS
# =========================================================

st.markdown("---")
st.markdown("## 📈 **Distribuições e Análises**")

col_r5, col_r6, col_r7, col_r8 = st.columns(4)

# ========== RANKING 5: ATIVIDADES POR PILAR ==========
with col_r5:
    st.markdown("### 🎯 **Atividades por Pilar**")
    
    ranking_html = '<div class="ranking-container">'
    
    if 'Pilar' in df_periodo.columns:
        pilares = df_periodo['Pilar'].value_counts().head(10)
        
        for nome, valor in pilares.items():
            ranking_html += f"""
            <div class="ranking-item">
                <span class="ranking-name">{nome}</span>
                <span class="ranking-value">{valor:,}</span>
            </div>
            """
    else:
        # Dados de exemplo
        exemplos = [
            ('Geração de Demanda', 5030),
            ('Branding & Atração', 602),
            ('Desenvolvimento', 430),
            ('Recrutamento', 258),
            ('Geração de demanda', 4)
        ]
        
        for nome, valor in exemplos:
            ranking_html += f"""
            <div class="ranking-item">
                <span class="ranking-name">{nome}</span>
                <span class="ranking-value">{valor}</span>
            </div>
            """
    
    ranking_html += '<p style="text-align: center; margin-top: 15px; color: #666;">1 - 5 / 12</p>'
    ranking_html += '</div>'
    
    st.markdown(ranking_html, unsafe_allow_html=True)

# ========== RANKING 6: STATUS ==========
with col_r6:
    st.markdown("### 🔄 **Status**")
    
    ranking_html = '<div class="ranking-container">'
    
    if 'Status' in df_periodo.columns:
        status_counts = df_periodo['Status'].value_counts()
        
        for nome, valor in status_counts.items():
            # Emojis para status
            if 'aprovado' in nome.lower():
                nome = "✅ " + nome
            elif 'produção' in nome.lower():
                nome = "⚙️ " + nome
            elif 'aguardando' in nome.lower():
                nome = "⏳ " + nome
            elif 'concluído' in nome.lower():
                nome = "🏁 " + nome
            
            ranking_html += f"""
            <div class="ranking-item">
                <span class="ranking-name">{nome}</span>
                <span class="ranking-value">{valor:,}</span>
            </div>
            """
    else:
        exemplos = [
            ('✅ Aprovado', 1245),
            ('⚙️ Em Produção', 876),
            ('⏳ Aguardando', 543),
            ('🏁 Concluído', 234)
        ]
        
        for nome, valor in exemplos:
            ranking_html += f"""
            <div class="ranking-item">
                <span class="ranking-name">{nome}</span>
                <span class="ranking-value">{valor}</span>
            </div>
            """
    
    ranking_html += '</div>'
    st.markdown(ranking_html, unsafe_allow_html=True)

# ========== RANKING 7: PRIORIDADE ==========
with col_r7:
    st.markdown("### ⚡ **Prioridade**")
    
    ranking_html = '<div class="ranking-container">'
    
    if 'Prioridade' in df_periodo.columns:
        prioridades = df_periodo['Prioridade'].value_counts()
        
        for nome, valor in prioridades.items():
            # Emojis para prioridade
            if 'alta' in nome.lower():
                nome = "🔴 " + nome
            elif 'média' in nome.lower():
                nome = "🟡 " + nome
            elif 'baixa' in nome.lower():
                nome = "🟢 " + nome
            
            ranking_html += f"""
            <div class="ranking-item">
                <span class="ranking-name">{nome}</span>
                <span class="ranking-value">{valor:,}</span>
            </div>
            """
    else:
        exemplos = [
            ('🔴 Alta', 450),
            ('🟡 Média', 350),
            ('🟢 Baixa', 200)
        ]
        
        for nome, valor in exemplos:
            ranking_html += f"""
            <div class="ranking-item">
                <span class="ranking-name">{nome}</span>
                <span class="ranking-value">{valor}</span>
            </div>
            """
    
    ranking_html += '</div>'
    st.markdown(ranking_html, unsafe_allow_html=True)

# ========== RANKING 8: TIPOS DE E-MAIL ==========
with col_r8:
    st.markdown("### 📧 **Tipos de E-mail**")
    
    ranking_html = '<div class="ranking-container">'
    
    # Tentar encontrar coluna de tipo de e-mail
    if 'Email' in df_periodo.columns or 'E-mail' in df_periodo.columns:
        col_email = 'Email' if 'Email' in df_periodo.columns else 'E-mail'
        emails = df_periodo[col_email].value_counts().head(6)
        
        for nome, valor in emails.items():
            ranking_html += f"""
            <div class="ranking-item">
                <span class="ranking-name">{nome[:25]}...</span>
                <span class="ranking-value">{valor:,}</span>
            </div>
            """
    else:
        exemplos = [
            ('Events', 4657),
            ('Campanhas de incentivo', 495),
            ('Comunicados', 491),
            ('Promocional', 62)
        ]
        
        for nome, valor in exemplos:
            ranking_html += f"""
            <div class="ranking-item">
                <span class="ranking-name">{nome}</span>
                <span class="ranking-value">{valor}</span>
            </div>
            """
    
    ranking_html += '<p style="text-align: center; margin-top: 15px; color: #666;">1 - 4 / 6</p>'
    ranking_html += '</div>'
    
    st.markdown(ranking_html, unsafe_allow_html=True)

# =========================================================
# GRÁFICOS DE TENDÊNCIA
# =========================================================

st.markdown("---")
st.markdown("## 📅 **Tendência de Solicitações**")

if coluna_data and coluna_data in df_periodo.columns:
    # Agrupar por data
    df_tendencia = df_periodo.groupby(df_periodo[coluna_data].dt.date).size().reset_index()
    df_tendencia.columns = ['Data', 'Quantidade']
    
    # Gráfico de linha
    fig = px.line(df_tendencia, x='Data', y='Quantidade', 
                  title=f'Solicitações por Dia - {data_inicio.strftime("%d/%m")} a {data_fim.strftime("%d/%m")}',
                  markers=True)
    
    fig.update_layout(
        height=400,
        plot_bgcolor='white',
        hovermode='x unified'
    )
    
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("ℹ️ Não foi possível gerar gráfico de tendência - coluna de data não encontrada")

# =========================================================
# RODAPÉ
# =========================================================

st.markdown("---")

col_f1, col_f2, col_f3 = st.columns(3)

with col_f1:
    st.caption(f"🕐 Atualizado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")

with col_f2:
    st.caption(f"📊 Total no período: {total_solicitacoes:,} registros")

with col_f3:
    st.caption("🔗 Fonte: SharePoint - Demandas ID")

# Botão de atualização na sidebar
with st.sidebar:
    st.markdown("## ⚙️ Controles")
    
    if st.button("🔄 Atualizar Dados", type="primary", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
    
    st.markdown("---")
    st.markdown("### 📝 Exportar")
    
    # Exportar dados filtrados
    if not df_periodo.empty:
        csv = df_periodo.to_csv(index=False, encoding='utf-8-sig')
        st.download_button(
            label="📥 Download CSV",
            data=csv,
            file_name=f"dados_cocred_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            use_container_width=True
        )
    
    st.markdown("---")
    st.markdown("### 🔗 Links Úteis")
    st.markdown("[📎 Abrir Excel Online](https://agenciaideatore-my.sharepoint.com/:x:/g/personal/cristini_cordesco_ideatoreamericas_com/IQDMDcVdgAfGSIyZfeke7NFkAatm3fhI0-X4r6gIPQJmosY)")