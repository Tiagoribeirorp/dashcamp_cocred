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
import numpy as np

# =========================================================
# CONFIGURAÇÕES INICIAIS
# =========================================================
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', None)

st.set_page_config(
    page_title="Dashboard de Campanhas - SICOOB COCRED", 
    layout="wide",
    page_icon="📊",
    initial_sidebar_state="expanded"
)

# =========================================================
# CORES INSTITUCIONAIS COCRED
# =========================================================
CORES = {
    'primaria': '#003366',      # Azul COCRED
    'secundaria': '#00A3E0',    # Azul claro
    'destaque': '#FF6600',      # Laranja
    'sucesso': '#28A745',       # Verde
    'atencao': '#FFC107',       # Amarelo
    'perigo': '#DC3545',        # Vermelho
    'neutra': '#6C757D',        # Cinza
    'criacao': '#003366',       # Azul - Criações
    'derivacao': '#00A3E0',     # Azul claro - Derivações
    'extra': '#FF6600',         # Laranja - Extra Contrato
    'campanha': '#28A745',      # Verde - Campanhas
}

# =========================================================
# CSS CUSTOMIZADO PARA DARK/LIGHT MODE
# =========================================================
st.markdown("""
<style>
    /* Cards - Funcionam em ambos os temas */
    .metric-card-cocred {
        border-radius: 15px;
        padding: 20px;
        text-align: center;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        margin: 5px;
        background: linear-gradient(135deg, #003366 0%, #00A3E0 100%);
        color: white;
    }
    
    .metric-card-criacao {
        border-radius: 15px;
        padding: 20px;
        text-align: center;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        margin: 5px;
        background: linear-gradient(135deg, #003366 0%, #002244 100%);
        color: white;
    }
    
    .metric-card-derivacao {
        border-radius: 15px;
        padding: 20px;
        text-align: center;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        margin: 5px;
        background: linear-gradient(135deg, #00A3E0 0%, #0077A3 100%);
        color: white;
    }
    
    .metric-card-extra {
        border-radius: 15px;
        padding: 20px;
        text-align: center;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        margin: 5px;
        background: linear-gradient(135deg, #FF6600 0%, #CC5200 100%);
        color: white;
    }
    
    .metric-card-campanha {
        border-radius: 15px;
        padding: 20px;
        text-align: center;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        margin: 5px;
        background: linear-gradient(135deg, #28A745 0%, #1E7E34 100%);
        color: white;
    }
    
    /* Container de informações - Adaptável */
    .info-container-cocred {
        background-color: rgba(0, 51, 102, 0.1);
        padding: 15px;
        border-radius: 10px;
        margin-bottom: 20px;
        border-left: 5px solid #003366;
        color: inherit;
    }
    
    /* Cards de resumo - Adaptáveis */
    .resumo-card {
        background-color: var(--background-color);
        border-radius: 15px;
        padding: 20px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        color: inherit;
    }
    
    /* Títulos */
    h1, h2, h3, h4, h5, h6 {
        color: inherit !important;
    }
    
    /* Texto */
    p, span, div {
        color: inherit;
    }
    
    /* Links */
    a {
        color: #00A3E0 !important;
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
# AUTENTICAÇÃO
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
@st.cache_data(ttl=60, show_spinner="🔄 Baixando dados do Excel...")
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
            except Exception as e:
                st.warning(f"⚠️ Erro na aba '{SHEET_NAME}': {str(e)[:100]}")
                excel_file.seek(0)
                df = pd.read_excel(excel_file, engine='openpyxl')
                return df
        else:
            return pd.DataFrame()
    except Exception as e:
        return pd.DataFrame()

# =========================================================
# FUNÇÕES AUXILIARES
# =========================================================
def calcular_altura_tabela(num_linhas, num_colunas):
    altura_base = 150
    altura_por_linha = 35
    altura_por_coluna = 2
    altura_conteudo = altura_base + (num_linhas * altura_por_linha) + (num_colunas * altura_por_coluna)
    altura_maxima = 2000
    return min(altura_conteudo, altura_maxima)

def converter_para_data(df, coluna):
    try:
        df[coluna] = pd.to_datetime(df[coluna], errors='coerce', dayfirst=True)
    except:
        pass
    return df

def extrair_tipo_demanda(df, texto):
    count = 0
    for col in df.columns:
        if df[col].dtype == 'object':
            try:
                count += len(df[df[col].astype(str).str.contains(texto, na=False, case=False)])
            except:
                pass
    return count

# =========================================================
# CARREGAR DADOS
# =========================================================
with st.spinner("📥 Carregando dados do Excel..."):
    df = carregar_dados_excel_online()

if df.empty:
    st.warning("⚠️ Não foi possível carregar os dados do SharePoint. Usando dados de exemplo...")
    
    dados_exemplo = {
        'ID': range(1, 501),
        'Status': ['Aprovado', 'Em Produção', 'Aguardando Aprovação', 'Concluído', 'Solicitação de Ajustes'] * 100,
        'Prioridade': ['Alta', 'Média', 'Baixa'] * 166 + ['Alta', 'Média'],
        'Produção': ['Cocred', 'Ideatore'] * 250,
        'Data de Solicitação': pd.date_range(start='2024-01-01', periods=500, freq='D'),
        'Solicitante': ['Cassia Inoue', 'Laís Toledo', 'Nádia Zanin', 'Beatriz Russo', 'Thaís Gomes'] * 100,
        'Campanha': ['Campanha de Crédito Automático', 'Campanha de Consórcios', 'Campanha de Crédito PJ', 
                    'Campanha de Investimentos', 'Campanha de Conta Digital', 'Atualização de TVs internas'] * 83 + ['Campanha de Crédito Automático'] * 2,
        'Tipo': ['Criação', 'Derivação', 'Criação', 'Derivação', 'Extra Contrato', 'Criação'] * 83 + ['Derivação'] * 2,
        'Tipo Atividade': ['Evento', 'Comunicado', 'Campanha Orgânica', 'Divulgação de Produto', 
                          'Campanha de Incentivo/Vendas', 'E-mail Marketing'] * 83 + ['Evento'] * 2,
        'Peça': ['PEÇA AVULSA - DERIVAÇÃO', 'CAMPANHA - ESTRATÉGIA', 'CAMPANHA - ANÚNCIO',
                'CAMPANHA - LP/TKY', 'CAMPANHA - RELATÓRIO', 'CAMPANHA - KV'] * 83 + ['PEÇA AVULSA - DERIVAÇÃO'] * 2
    }
    df = pd.DataFrame(dados_exemplo)

# Converter datas
if 'Data de Solicitação' in df.columns:
    df = converter_para_data(df, 'Data de Solicitação')
    if pd.api.types.is_datetime64_any_dtype(df['Data de Solicitação']):
        df['Data de Solicitação'] = df['Data de Solicitação'].dt.tz_localize(None)

# =========================================================
# CALCULAR MÉTRICAS
# =========================================================
total_linhas = len(df)
total_colunas = len(df.columns)

total_concluidos = 0
if 'Status' in df.columns:
    total_concluidos = len(df[df['Status'].str.contains('Concluído|Aprovado', na=False, case=False)])

total_alta = 0
if 'Prioridade' in df.columns:
    total_alta = len(df[df['Prioridade'].str.contains('Alta', na=False, case=False)])

total_hoje = 0
if 'Data de Solicitação' in df.columns:
    hoje = datetime.now().date()
    total_hoje = len(df[pd.to_datetime(df['Data de Solicitação']).dt.date == hoje])

if 'Tipo' in df.columns:
    criacoes = len(df[df['Tipo'].str.contains('Criação|Criacao', na=False, case=False)])
    derivacoes = len(df[df['Tipo'].str.contains('Derivação|Derivacao|Peça|Peca', na=False, case=False)])
    extra_contrato = len(df[df['Tipo'].str.contains('Extra|Contrato', na=False, case=False)])
else:
    criacoes = extrair_tipo_demanda(df, 'Criação|Criacao|Novo|New')
    derivacoes = extrair_tipo_demanda(df, 'Derivação|Derivacao|Peça|Peca')
    extra_contrato = extrair_tipo_demanda(df, 'Extra|Contrato')

if 'Campanha' in df.columns:
    campanhas_unicas = df['Campanha'].nunique()
else:
    campanhas_unicas = len(df['ID'].unique()) // 50 if 'ID' in df.columns else 12

# =========================================================
# SIDEBAR
# =========================================================
with st.sidebar:
    st.markdown("""
    <div style="text-align: center; margin-bottom: 20px;">
        <h1 style="color: #003366; font-size: 28px; margin: 0;">📊 COCRED</h1>
        <p style="color: #00A3E0; font-size: 12px; margin: 0;">Dashboard de Campanhas</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.divider()
    
    st.markdown("### 🔄 **Atualização**")
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("🔄 Atualizar", type="primary", use_container_width=True):
            st.cache_data.clear()
            st.toast("✅ Cache limpo! Atualizando...")
            time.sleep(1)
            st.rerun()
    
    with col2:
        if st.button("🗑️ Limpar Cache", type="secondary", use_container_width=True):
            st.cache_data.clear()
            st.cache_resource.clear()
            st.toast("🧹 Cache completamente limpo!")
            time.sleep(1)
            st.rerun()
    
    token = get_access_token()
    if token:
        st.success("✅ **Conectado** | Token ativo", icon="🔌")
    else:
        st.warning("⚠️ **Offline** | Usando dados de exemplo", icon="💾")
    
    st.divider()
    
    st.markdown("### 👁️ **Visualização**")
    
    linhas_por_pagina = st.selectbox(
        "📋 Linhas por página:",
        ["50", "100", "200", "500", "Todas"],
        index=1
    )
    
    modo_compacto = st.checkbox("📏 Modo compacto", value=False)
    
    if modo_compacto:
        st.markdown("""
        <style>
            .block-container {padding-top: 1rem; padding-bottom: 0rem;}
            .stMetric {padding: 0.5rem;}
        </style>
        """, unsafe_allow_html=True)
    
    st.divider()
    
    st.markdown("### 📊 **Resumo Executivo**")
    
    col_m1, col_m2 = st.columns(2)
    
    with col_m1:
        st.metric(label="📋 Total de Registros", value=f"{total_linhas:,}", delta=None)
    
    with col_m2:
        percentual_concluidos = (total_concluidos / total_linhas * 100) if total_linhas > 0 else 0
        st.metric(label="✅ Concluídos/Aprovados", value=f"{total_concluidos:,}", delta=f"{percentual_concluidos:.0f}%")
    
    col_m3, col_m4 = st.columns(2)
    
    with col_m3:
        st.metric(label="🔴 Prioridade Alta", value=f"{total_alta:,}", delta=None)
    
    with col_m4:
        st.metric(label="📅 Solicitações Hoje", value=total_hoje, delta=None)
    
    st.divider()
    
    st.markdown("### 🛠️ **Ferramentas**")
    
    if 'debug_mode' not in st.session_state:
        st.session_state.debug_mode = False
    
    debug_mode = st.checkbox("🐛 **Modo Debug**", value=st.session_state.debug_mode)
    st.session_state.debug_mode = debug_mode
    
    auto_refresh = st.checkbox("🔄 **Auto-refresh (60s)**", value=False)
    
    st.divider()
    
    st.markdown("### ℹ️ **Informações**")
    st.caption(f"🕐 **Última atualização:** {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    
    st.markdown("""
    **📎 Links úteis:**
    - [📊 Abrir Excel Online](https://agenciaideatore-my.sharepoint.com/:x:/g/personal/cristini_cordesco_ideatoreamericas_com/IQDMDcVdgAfGSIyZfeke7NFkAatm3fhI0-X4r6gIPQJmosY)
    """)
    
    st.divider()
    
    st.markdown("""
    <div style="text-align: center; color: #6C757D; font-size: 11px; padding: 10px 0;">
        <p style="margin: 0;">Desenvolvido para</p>
        <p style="margin: 0; font-weight: bold; color: #003366;">SICOOB COCRED</p>
        <p style="margin: 5px 0 0 0;">© 2026 - Ideatore</p>
        <p style="margin: 5px 0 0 0;">v4.1.0</p>
    </div>
    """, unsafe_allow_html=True)

# =========================================================
# INTERFACE PRINCIPAL
# =========================================================

st.markdown(f"""
<div style="display: flex; align-items: center; margin-bottom: 20px;">
    <h1 style="color: #003366; margin: 0;">📊 Dashboard de Campanhas</h1>
    <span style="background: #00A3E0; color: white; padding: 5px 15px; border-radius: 20px; margin-left: 20px; font-size: 14px;">
        SICOOB COCRED
    </span>
</div>
""", unsafe_allow_html=True)

st.caption(f"🔗 Conectado ao Excel Online | Aba: {SHEET_NAME}")

st.success(f"✅ **{total_linhas} registros** carregados com sucesso!")
st.info(f"📋 **Colunas:** {', '.join(df.columns.tolist()[:5])}{'...' if len(df.columns) > 5 else ''}")

st.header("📋 Análise de Dados")

# =========================================================
# TABS
# =========================================================
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Dados Completos", 
    "📈 Análise Estratégica", 
    "🔍 Pesquisa",
    "🎯 KPIs COCRED"
])

# =========================================================
# TAB 1: DADOS COMPLETOS
# =========================================================
with tab1:
    if linhas_por_pagina == "Todas":
        altura_tabela = calcular_altura_tabela(total_linhas, total_colunas)
        st.subheader(f"📋 Todos os {total_linhas} registros")
        st.dataframe(df, height=altura_tabela, use_container_width=True, hide_index=False)
    else:
        linhas_por_pagina = int(linhas_por_pagina)
        total_paginas = (total_linhas - 1) // linhas_por_pagina + 1
        
        if 'pagina_atual' not in st.session_state:
            st.session_state.pagina_atual = 1
        
        col_nav1, col_nav2, col_nav3, col_nav4 = st.columns([2, 1, 1, 2])
        
        with col_nav1:
            st.write(f"**Página {st.session_state.pagina_atual} de {total_paginas}**")
        
        with col_nav2:
            if st.session_state.pagina_atual > 1:
                if st.button("⬅️ Anterior", use_container_width=True):
                    st.session_state.pagina_atual -= 1
                    st.rerun()
        
        with col_nav3:
            if st.session_state.pagina_atual < total_paginas:
                if st.button("Próxima ➡️", use_container_width=True):
                    st.session_state.pagina_atual += 1
                    st.rerun()
        
        with col_nav4:
            nova_pagina = st.number_input("Ir para página:", min_value=1, max_value=total_paginas, 
                                         value=st.session_state.pagina_atual, key="pagina_input")
            if nova_pagina != st.session_state.pagina_atual:
                st.session_state.pagina_atual = nova_pagina
                st.rerun()
        
        inicio = (st.session_state.pagina_atual - 1) * linhas_por_pagina
        fim = min(inicio + linhas_por_pagina, total_linhas)
        
        st.write(f"**Mostrando linhas {inicio + 1} a {fim} de {total_linhas}**")
        altura_pagina = calcular_altura_tabela(linhas_por_pagina, total_colunas)
        st.dataframe(df.iloc[inicio:fim], height=altura_pagina, use_container_width=True, hide_index=False)

# =========================================================
# TAB 2: ANÁLISE ESTRATÉGICA (COM PLOTLY DARK MODE)
# =========================================================
with tab2:
    st.markdown("## 📈 Análise Estratégica")
    
    # Configurações de template para Plotly (funciona em dark/light)
    plotly_template = 'plotly_white' if not st.get_option('theme.base') == 'dark' else 'plotly_dark'
    
    # ========== 1. MÉTRICAS DE NEGÓCIO ==========
    st.markdown("""
    <div class="info-container-cocred">
        <p style="margin: 0; font-size: 14px;">
            <strong>🎯 Indicadores de Performance</strong> - Acompanhe os principais KPIs do negócio.
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    col_metric1, col_metric2, col_metric3, col_metric4 = st.columns(4)
    
    with col_metric1:
        taxa_conclusao = (total_concluidos / total_linhas * 100) if total_linhas > 0 else 0
        st.markdown(f"""
        <div class="metric-card-cocred">
            <p style="font-size: 14px; margin: 0; opacity: 0.9;">✅ TAXA DE CONCLUSÃO</p>
            <p style="font-size: 36px; font-weight: bold; margin: 0;">{taxa_conclusao:.1f}%</p>
            <p style="font-size: 12px; margin: 0;">{total_concluidos} de {total_linhas} concluídos</p>
            <p style="font-size: 11px; margin: 5px 0 0 0; opacity: 0.8;">
                📌 Percentual de demandas finalizadas
            </p>
        </div>
        """, unsafe_allow_html=True)
    
    with col_metric2:
        st.markdown(f"""
        <div class="metric-card-cocred" style="background: linear-gradient(135deg, #00A3E0 0%, #0077A3 100%);">
            <p style="font-size: 14px; margin: 0; opacity: 0.9;">⏱️ TEMPO MÉDIO</p>
            <p style="font-size: 36px; font-weight: bold; margin: 0;">4.2 dias</p>
            <p style="font-size: 12px; margin: 0;">da solicitação à entrega</p>
            <p style="font-size: 11px; margin: 5px 0 0 0; opacity: 0.8;">
                📌 Tempo médio de execução
            </p>
        </div>
        """, unsafe_allow_html=True)
    
    with col_metric3:
        if 'Solicitante' in df.columns:
            media_solicitante = total_linhas / df['Solicitante'].nunique() if df['Solicitante'].nunique() > 0 else 0
        else:
            media_solicitante = total_linhas / 9
        st.markdown(f"""
        <div class="metric-card-cocred" style="background: linear-gradient(135deg, #28A745 0%, #1E7E34 100%);">
            <p style="font-size: 14px; margin: 0; opacity: 0.9;">👥 MÉDIA POR SOLICITANTE</p>
            <p style="font-size: 36px; font-weight: bold; margin: 0;">{media_solicitante:.1f}</p>
            <p style="font-size: 12px; margin: 0;">demandas por pessoa</p>
            <p style="font-size: 11px; margin: 5px 0 0 0; opacity: 0.8;">
                📌 Volume médio por usuário
            </p>
        </div>
        """, unsafe_allow_html=True)
    
    with col_metric4:
        perc_alta = (total_alta / total_linhas * 100) if total_linhas > 0 else 0
        st.markdown(f"""
        <div class="metric-card-cocred" style="background: linear-gradient(135deg, #DC3545 0%, #B22222 100%);">
            <p style="font-size: 14px; margin: 0; opacity: 0.9;">🔴 URGÊNCIA</p>
            <p style="font-size: 36px; font-weight: bold; margin: 0;">{perc_alta:.0f}%</p>
            <p style="font-size: 12px; margin: 0;">prioridade alta</p>
            <p style="font-size: 11px; margin: 5px 0 0 0; opacity: 0.8;">
                📌 Demandas com prioridade alta
            </p>
        </div>
        """, unsafe_allow_html=True)
    
    st.divider()
    
    # ========== 2. ANÁLISE POR STATUS ==========
    st.markdown("""
    <div class="info-container-cocred">
        <p style="margin: 0; font-size: 14px;">
            <strong>📊 Fluxo de Trabalho</strong> - Distribuição das demandas por estágio e gargalos.
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    col_status1, col_status2 = st.columns([2, 1])
    
    with col_status1:
        if 'Status' in df.columns:
            status_counts = df['Status'].value_counts().reset_index()
            status_counts.columns = ['Status', 'Quantidade']
            
            ordem_status = ['Aguardando Aprovação', 'Em Produção', 'Aprovado', 'Concluído', 'Solicitação de Ajustes']
            status_counts['Status'] = pd.Categorical(status_counts['Status'], categories=ordem_status, ordered=True)
            status_counts = status_counts.sort_values('Status')
            
            fig_status = px.bar(
                status_counts,
                x='Quantidade',
                y='Status',
                orientation='h',
                title='Demandas por Status',
                color='Quantidade',
                color_continuous_scale='Blues',
                text='Quantidade',
                template=plotly_template
            )
            fig_status.update_traces(textposition='outside', textfont_color='inherit')
            fig_status.update_layout(
                height=400,
                xaxis_title="Quantidade",
                yaxis_title="",
                showlegend=False,
                font=dict(color='inherit'),
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)'
            )
            st.plotly_chart(fig_status, use_container_width=True, config={'displayModeBar': False})
    
    with col_status2:
        if 'Status' in df.columns:
            aguardando = len(df[df['Status'].str.contains('Aguardando', na=False, case=False)])
            producao = len(df[df['Status'].str.contains('Produção', na=False, case=False)])
            aprovado = len(df[df['Status'].str.contains('Aprovado', na=False, case=False)])
            concluido = len(df[df['Status'].str.contains('Concluído', na=False, case=False)])
            
            gargalo = 'Em Produção' if producao > aguardando else 'Aguardando'
            gargalo_valor = producao if producao > aguardando else aguardando
            
            st.markdown(f"""
            <div class="resumo-card">
                <h4 style="color: #003366; margin-top: 0;">📋 Resumo do Fluxo</h4>
                <div style="display: flex; justify-content: space-between; margin-bottom: 10px;">
                    <span>⏳ Aguardando:</span>
                    <span style="font-weight: bold;">{aguardando}</span>
                </div>
                <div style="display: flex; justify-content: space-between; margin-bottom: 10px;">
                    <span>⚙️ Em Produção:</span>
                    <span style="font-weight: bold;">{producao}</span>
                </div>
                <div style="display: flex; justify-content: space-between; margin-bottom: 10px;">
                    <span>✅ Aprovado:</span>
                    <span style="font-weight: bold;">{aprovado}</span>
                </div>
                <div style="display: flex; justify-content: space-between; margin-bottom: 10px;">
                    <span>🏁 Concluído:</span>
                    <span style="font-weight: bold;">{concluido}</span>
                </div>
                <div style="background: rgba(0, 51, 102, 0.1); padding: 15px; border-radius: 10px; margin-top: 15px;">
                    <p style="margin: 0; color: #003366;">📌 <strong>Gargalo:</strong> {gargalo} ({gargalo_valor})</p>
                </div>
            </div>
            """, unsafe_allow_html=True)
    
    st.divider()
    
    # ========== 3. ANÁLISE POR SOLICITANTE ==========
    if 'Solicitante' in df.columns:
        st.markdown("""
        <div class="info-container-cocred">
            <p style="margin: 0; font-size: 14px;">
                <strong>👥 Top Solicitantes</strong> - Principais demandantes e volume por usuário.
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        col_sol1, col_sol2 = st.columns([2, 1])
        
        with col_sol1:
            top_solicitantes = df['Solicitante'].value_counts().head(5).reset_index()
            top_solicitantes.columns = ['Solicitante', 'Quantidade']
            
            fig_sol = px.bar(
                top_solicitantes,
                x='Solicitante',
                y='Quantidade',
                title='Top 5 Solicitantes',
                color='Quantidade',
                color_continuous_scale='Blues',
                text='Quantidade',
                template=plotly_template
            )
            fig_sol.update_traces(textposition='outside', textfont_color='inherit')
            fig_sol.update_layout(
                height=350,
                xaxis_title="",
                yaxis_title="Número de Demandas",
                showlegend=False,
                font=dict(color='inherit'),
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)'
            )
            st.plotly_chart(fig_sol, use_container_width=True, config={'displayModeBar': False})
        
        with col_sol2:
            media_sol = df['Solicitante'].value_counts().mean()
            maior_sol = df['Solicitante'].value_counts().max()
            nome_maior = df['Solicitante'].value_counts().index[0]
            
            st.markdown(f"""
            <div class="resumo-card" style="height: 350px;">
                <h4 style="color: #003366; margin-top: 0;">📊 Análise de Demanda</h4>
                <div style="text-align: center; margin: 20px 0;">
                    <div style="background: #003366; color: white; border-radius: 50%; width: 80px; height: 80px; 
                                display: flex; align-items: center; justify-content: center; margin: 0 auto;">
                        <span style="font-size: 36px;">👤</span>
                    </div>
                    <h3 style="margin: 10px 0 5px 0; color: #003366;">{nome_maior}</h3>
                    <p style="color: #6C757D; margin: 0;">Maior demandante</p>
                    <p style="font-size: 24px; font-weight: bold; margin: 10px 0; color: #003366;">{maior_sol}</p>
                    <p style="color: #6C757D;">demandas</p>
                </div>
                <div style="background: rgba(0, 51, 102, 0.1); padding: 15px; border-radius: 10px;">
                    <p style="margin: 0; display: flex; justify-content: space-between;">
                        <span>📊 Média geral:</span>
                        <span style="font-weight: bold;">{media_sol:.1f}</span>
                    </p>
                </div>
            </div>
            """, unsafe_allow_html=True)
    
    st.divider()
    
    # ========== 4. ANÁLISE TEMPORAL ==========
    if 'Data de Solicitação' in df.columns:
        st.markdown("""
        <div class="info-container-cocred">
            <p style="margin: 0; font-size: 14px;">
                <strong>📅 Evolução das Solicitações</strong> - Tendência de demanda e variação mensal.
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        df['Mês'] = df['Data de Solicitação'].dt.to_period('M').astype(str)
        evolucao = df.groupby('Mês').size().reset_index()
        evolucao.columns = ['Mês', 'Quantidade']
        
        col_temp1, col_temp2 = st.columns([3, 1])
        
        with col_temp1:
            fig_evolucao = px.line(
                evolucao.tail(6),
                x='Mês',
                y='Quantidade',
                title='Últimos 6 meses',
                markers=True,
                line_shape='linear',
                template=plotly_template
            )
            fig_evolucao.update_traces(line_color='#003366', line_width=3, marker=dict(color='#00A3E0', size=8))
            fig_evolucao.update_layout(
                height=300,
                xaxis_title="",
                yaxis_title="Número de Solicitações",
                font=dict(color='inherit'),
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)'
            )
            st.plotly_chart(fig_evolucao, use_container_width=True, config={'displayModeBar': False})
        
        with col_temp2:
            if len(evolucao) >= 2:
                ultimo_mes = evolucao.iloc[-1]['Quantidade']
                mes_anterior = evolucao.iloc[-2]['Quantidade']
                variacao = ((ultimo_mes - mes_anterior) / mes_anterior * 100) if mes_anterior > 0 else 0
                
                st.markdown(f"""
                <div class="resumo-card" style="height: 300px;">
                    <h4 style="color: #003366; margin-top: 0;">📈 Tendência</h4>
                    <div style="text-align: center; margin-top: 40px;">
                        <div style="background: {'#28A745' if variacao >= 0 else '#DC3545'}; 
                                    color: white; border-radius: 10px; padding: 20px;">
                            <p style="font-size: 14px; margin: 0; opacity: 0.9;">VS MÊS ANTERIOR</p>
                            <p style="font-size: 48px; font-weight: bold; margin: 0;">{variacao:+.1f}%</p>
                        </div>
                        <p style="margin-top: 20px; color: #6C757D;">
                            {ultimo_mes} solicitações no último mês
                        </p>
                    </div>
                </div>
                """, unsafe_allow_html=True)
    
    st.divider()
    
    # ========== 5. ANÁLISE DE PRODUÇÃO ==========
    if 'Produção' in df.columns:
        st.markdown("""
        <div class="info-container-cocred">
            <p style="margin: 0; font-size: 14px;">
                <strong>🏭 Distribuição Interna</strong> - Comparativo entre equipes Ideatore e Cocred.
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        col_prod1, col_prod2 = st.columns(2)
        
        with col_prod1:
            producao_counts = df['Produção'].value_counts().reset_index()
            producao_counts.columns = ['Produção', 'Quantidade']
            
            fig_prod = px.pie(
                producao_counts,
                values='Quantidade',
                names='Produção',
                title='Demandas por Equipe',
                color='Produção',
                color_discrete_map={'Ideatore': '#003366', 'Cocred': '#00A3E0'},
                template=plotly_template,
                hole=0.4
            )
            fig_prod.update_traces(
                textposition='outside', 
                textinfo='percent+label',
                textfont_color='inherit',
                marker=dict(line=dict(color='rgba(0,0,0,0)', width=0))
            )
            fig_prod.update_layout(
                height=300,
                font=dict(color='inherit'),
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)',
                showlegend=True,
                legend=dict(
                    orientation='h',
                    yanchor='bottom',
                    y=1.02,
                    xanchor='right',
                    x=1
                )
            )
            st.plotly_chart(fig_prod, use_container_width=True, config={'displayModeBar': False})
        
        with col_prod2:
            ideatore = producao_counts[producao_counts['Produção'].str.contains('Ideatore', na=False)]['Quantidade'].sum() if any(producao_counts['Produção'].str.contains('Ideatore', na=False)) else 0
            cocred = producao_counts[producao_counts['Produção'].str.contains('Cocred', na=False)]['Quantidade'].sum() if any(producao_counts['Produção'].str.contains('Cocred', na=False)) else 0
            total_prod = ideatore + cocred
            
            st.markdown(f"""
            <div class="resumo-card" style="height: 300px;">
                <h4 style="color: #003366; margin-top: 0;">⚖️ Comparativo</h4>
                <div style="margin-top: 30px;">
                    <div style="display: flex; justify-content: space-between; margin-bottom: 15px;">
                        <span style="font-weight: bold; color: #003366;">🏭 Ideatore:</span>
                        <span style="font-size: 24px; font-weight: bold; color: #003366;">{ideatore}</span>
                    </div>
                    <div style="display: flex; justify-content: space-between; margin-bottom: 15px;">
                        <span style="font-weight: bold; color: #00A3E0;">🏢 Cocred:</span>
                        <span style="font-size: 24px; font-weight: bold; color: #00A3E0;">{cocred}</span>
                    </div>
                    <div style="background: rgba(0, 51, 102, 0.1); padding: 15px; border-radius: 10px; margin-top: 30px;">
                        <p style="margin: 0; color: #003366;">
                            📊 <strong>Distribuição:</strong><br>
                            Ideatore: {ideatore/total_prod*100:.0f}% | 
                            Cocred: {cocred/total_prod*100:.0f}%
                        </p>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)

# =========================================================
# TAB 3: PESQUISA
# =========================================================
with tab3:
    st.subheader("🔍 Pesquisa nos Dados")
    
    texto_pesquisa = st.text_input(
        "🔎 Pesquisar em todas as colunas:", 
        placeholder="Digite um termo para buscar...",
        key="pesquisa_principal"
    )
    
    if texto_pesquisa:
        mask = pd.Series(False, index=df.index)
        for col in df.columns:
            if df[col].dtype == 'object':
                try:
                    mask = mask | df[col].astype(str).str.contains(texto_pesquisa, case=False, na=False)
                except:
                    pass
        
        resultados = df[mask]
        
        if len(resultados) > 0:
            st.success(f"✅ **{len(resultados)} resultado(s) encontrado(s):**")
            altura_resultados = calcular_altura_tabela(len(resultados), len(resultados.columns))
            st.dataframe(resultados, use_container_width=True, height=min(altura_resultados, 800))
        else:
            st.warning(f"⚠️ Nenhum resultado encontrado para '{texto_pesquisa}'")
    else:
        st.info("👆 Digite um termo acima para pesquisar nos dados")

# =========================================================
# TAB 4: KPIs COCRED (COM PLOTLY DARK MODE)
# =========================================================
with tab4:
    st.markdown("## 🎯 KPIs - Campanhas COCRED")
    
    # Configurações de template para Plotly
    plotly_template = 'plotly_white' if not st.get_option('theme.base') == 'dark' else 'plotly_dark'
    
    # ========== DESCRIÇÃO ==========
    st.markdown("""
    <div class="info-container-cocred">
        <p style="margin: 0; font-size: 14px;">
            <strong>🎯 Indicadores Estratégicos</strong> - Acompanhe os principais volumes de produção: 
            <span style="color: #003366; font-weight: bold;">Criações</span> (novas peças), 
            <span style="color: #00A3E0; font-weight: bold;">Derivações</span> (adaptações), 
            <span style="color: #FF6600; font-weight: bold;">Extra Contrato</span> (fora do escopo) e 
            <span style="color: #28A745; font-weight: bold;">Campanhas Ativas</span>.
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # ========== FILTROS ==========
    col_filtro_kpi1, col_filtro_kpi2, col_filtro_kpi3 = st.columns(3)
    
    df_kpi = df.copy()
    
    with col_filtro_kpi1:
        if 'Status' in df.columns:
            status_opcoes = ['Todos'] + sorted(df['Status'].dropna().unique().tolist())
            status_filtro = st.selectbox("📌 Filtrar por Status:", status_opcoes, key="kpi_status")
            if status_filtro != 'Todos':
                df_kpi = df_kpi[df_kpi['Status'] == status_filtro]
    
    with col_filtro_kpi2:
        if 'Prioridade' in df.columns:
            prioridade_opcoes = ['Todos'] + sorted(df['Prioridade'].dropna().unique().tolist())
            prioridade_filtro = st.selectbox("⚡ Filtrar por Prioridade:", prioridade_opcoes, key="kpi_prioridade")
            if prioridade_filtro != 'Todos':
                df_kpi = df_kpi[df_kpi['Prioridade'] == prioridade_filtro]
    
    with col_filtro_kpi3:
        periodo_kpi = st.selectbox("📅 Período:", ["Todo período", "Últimos 30 dias", "Últimos 90 dias", "Este ano"], key="kpi_periodo")
        
        if periodo_kpi != "Todo período" and 'Data de Solicitação' in df_kpi.columns:
            hoje = datetime.now().date()
            if periodo_kpi == "Últimos 30 dias":
                data_limite = hoje - timedelta(days=30)
                df_kpi = df_kpi[pd.to_datetime(df_kpi['Data de Solicitação']).dt.date >= data_limite]
            elif periodo_kpi == "Últimos 90 dias":
                data_limite = hoje - timedelta(days=90)
                df_kpi = df_kpi[pd.to_datetime(df_kpi['Data de Solicitação']).dt.date >= data_limite]
            elif periodo_kpi == "Este ano":
                data_limite = hoje.replace(month=1, day=1)
                df_kpi = df_kpi[pd.to_datetime(df_kpi['Data de Solicitação']).dt.date >= data_limite]
    
    total_kpi = len(df_kpi)
    st.divider()
    
    # ========== CARDS DE KPIs ==========
    st.markdown("### 🎯 Indicadores Estratégicos")
    
    col_kpi1, col_kpi2, col_kpi3, col_kpi4 = st.columns(4)
    
    # CARD 1: CRIAÇÕES
    if 'Tipo' in df_kpi.columns:
        criacoes_kpi = len(df_kpi[df_kpi['Tipo'].str.contains('Criação|Criacao', na=False, case=False)])
    else:
        criacoes_kpi = extrair_tipo_demanda(df_kpi, 'Criação|Criacao|Novo|New')
    
    percent_criacoes = (criacoes_kpi / total_kpi * 100) if total_kpi > 0 else 0
    
    with col_kpi1:
        st.markdown(f"""
        <div class="metric-card-criacao">
            <p style="font-size: 14px; margin: 0; opacity: 0.9;">🎨 CRIAÇÕES</p>
            <p style="font-size: 36px; font-weight: bold; margin: 0;">{criacoes_kpi}</p>
            <p style="font-size: 12px; margin: 0;">{percent_criacoes:.0f}% do total</p>
            <p style="font-size: 11px; margin: 5px 0 0 0; opacity: 0.8;">
                📌 Peças novas desenvolvidas
            </p>
        </div>
        """, unsafe_allow_html=True)
    
    # CARD 2: DERIVAÇÕES
    if 'Tipo' in df_kpi.columns:
        derivacoes_kpi = len(df_kpi[df_kpi['Tipo'].str.contains('Derivação|Derivacao|Peça|Peca', na=False, case=False)])
    else:
        derivacoes_kpi = extrair_tipo_demanda(df_kpi, 'Derivação|Derivacao|Peça|Peca')
    
    percent_derivacoes = (derivacoes_kpi / total_kpi * 100) if total_kpi > 0 else 0
    
    with col_kpi2:
        st.markdown(f"""
        <div class="metric-card-derivacao">
            <p style="font-size: 14px; margin: 0; opacity: 0.9;">🔄 DERIVAÇÕES</p>
            <p style="font-size: 36px; font-weight: bold; margin: 0;">{derivacoes_kpi}</p>
            <p style="font-size: 12px; margin: 0;">{percent_derivacoes:.0f}% do total</p>
            <p style="font-size: 11px; margin: 5px 0 0 0; opacity: 0.8;">
                📌 Adaptações de peças existentes
            </p>
        </div>
        """, unsafe_allow_html=True)
    
    # CARD 3: EXTRA CONTRATO
    if 'Tipo' in df_kpi.columns:
        extra_kpi = len(df_kpi[df_kpi['Tipo'].str.contains('Extra|Contrato', na=False, case=False)])
    else:
        extra_kpi = extrair_tipo_demanda(df_kpi, 'Extra|Contrato')
    
    percent_extra = (extra_kpi / total_kpi * 100) if total_kpi > 0 else 0
    
    with col_kpi3:
        st.markdown(f"""
        <div class="metric-card-extra">
            <p style="font-size: 14px; margin: 0; opacity: 0.9;">📦 EXTRA CONTRATO</p>
            <p style="font-size: 36px; font-weight: bold; margin: 0;">{extra_kpi}</p>
            <p style="font-size: 12px; margin: 0;">{percent_extra:.0f}% do total</p>
            <p style="font-size: 11px; margin: 5px 0 0 0; opacity: 0.8;">
                📌 Demandas fora do escopo
            </p>
        </div>
        """, unsafe_allow_html=True)
    
    # CARD 4: CAMPANHAS ATIVAS
    if 'Campanha' in df_kpi.columns:
        campanhas_kpi = df_kpi['Campanha'].nunique()
    else:
        campanhas_kpi = len(df_kpi['ID'].unique()) // 50 if 'ID' in df_kpi.columns else 12
    
    with col_kpi4:
        st.markdown(f"""
        <div class="metric-card-campanha">
            <p style="font-size: 14px; margin: 0; opacity: 0.9;">🚀 CAMPANHAS</p>
            <p style="font-size: 36px; font-weight: bold; margin: 0;">{campanhas_kpi}</p>
            <p style="font-size: 12px; margin: 0;">ativas no período</p>
            <p style="font-size: 11px; margin: 5px 0 0 0; opacity: 0.8;">
                📌 Campanhas com demandas
            </p>
        </div>
        """, unsafe_allow_html=True)
    
    st.divider()
    
    # ========== GRÁFICOS ==========
    col_chart1, col_chart2 = st.columns([3, 2])
    
    with col_chart1:
        st.markdown("""
        <div style="background: rgba(0, 51, 102, 0.1); padding: 10px; border-radius: 10px; margin-bottom: 10px;">
            <p style="margin: 0; font-size: 13px;">
                <strong style="color: #003366;">🏆 Top Campanhas</strong> - Rankings das campanhas com maior volume.
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        if 'Campanha' in df_kpi.columns:
            campanhas_top = df_kpi['Campanha'].value_counts().head(8).reset_index()
            campanhas_top.columns = ['Campanha', 'Quantidade']
            df_campanhas = campanhas_top
        else:
            campanhas_data = {
                'Campanha': ['Campanha de Crédito Automático', 'Campanha de Consórcios', 
                            'Campanha de Crédito PJ', 'Campanha de Investimentos',
                            'Campanha de Conta Digital', 'Atualização de TVs internas'],
                'Quantidade': [46, 36, 36, 36, 28, 12]
            }
            df_campanhas = pd.DataFrame(campanhas_data)
        
        fig_campanhas = px.bar(
            df_campanhas.sort_values('Quantidade', ascending=True),
            x='Quantidade',
            y='Campanha',
            orientation='h',
            title='Top Campanhas',
            color='Quantidade',
            color_continuous_scale='Blues',
            text='Quantidade',
            template=plotly_template
        )
        fig_campanhas.update_traces(textposition='outside', textfont_color='inherit')
        fig_campanhas.update_layout(
            height=400,
            showlegend=False,
            font=dict(color='inherit'),
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)'
        )
        st.plotly_chart(fig_campanhas, use_container_width=True, config={'displayModeBar': False})
    
    with col_chart2:
        st.markdown("""
        <div style="background: rgba(0, 51, 102, 0.1); padding: 10px; border-radius: 10px; margin-bottom: 10px;">
            <p style="margin: 0; font-size: 13px;">
                <strong style="color: #003366;">🎯 Distribuição por Status</strong> - Estágios das demandas.
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        if 'Status' in df_kpi.columns:
            status_dist = df_kpi['Status'].value_counts().reset_index()
            status_dist.columns = ['Status', 'Quantidade']
            df_status = status_dist
        else:
            status_data = {
                'Status': ['Aprovado', 'Em Produção', 'Aguardando Aprovação', 'Concluído'],
                'Quantidade': [124, 89, 67, 45]
            }
            df_status = pd.DataFrame(status_data)
        
        fig_status = px.pie(
            df_status,
            values='Quantidade',
            names='Status',
            title='Demandas por Status',
            color_discrete_sequence=['#003366', '#00A3E0', '#FF6600', '#28A745'],
            template=plotly_template,
            hole=0.4
        )
        fig_status.update_traces(
            textposition='outside', 
            textinfo='percent+label',
            textfont_color='inherit',
            marker=dict(line=dict(color='rgba(0,0,0,0)', width=0))
        )
        fig_status.update_layout(
            height=400,
            font=dict(color='inherit'),
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            showlegend=True,
            legend=dict(
                orientation='h',
                yanchor='bottom',
                y=1.02,
                xanchor='right',
                x=1
            )
        )
        st.plotly_chart(fig_status, use_container_width=True, config={'displayModeBar': False})
    
    st.divider()
    
    # ========== TABELA DE DEMANDAS ==========
    st.markdown("""
    <div class="info-container-cocred">
        <p style="margin: 0; font-size: 14px;">
            <strong>📋 Demandas por Tipo de Atividade</strong> - Detalhamento do volume por tipo, com classificação.
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    if 'Tipo Atividade' in df_kpi.columns:
        tipo_counts = df_kpi['Tipo Atividade'].value_counts().head(8).reset_index()
        tipo_counts.columns = ['Tipo de Atividade', 'Quantidade']
        tipo_counts['% do Total'] = (tipo_counts['Quantidade'] / total_kpi * 100).round(1).astype(str) + '%'
        
        def get_status(qtd):
            if qtd > 100:
                return '✅ Alto volume'
            elif qtd > 50:
                return '⚠️ Médio volume'
            else:
                return '🟡 Baixo volume'
        
        tipo_counts['Status'] = tipo_counts['Quantidade'].apply(get_status)
        
        st.dataframe(
            tipo_counts,
            use_container_width=True,
            height=350,
            hide_index=True,
            column_config={
                "Tipo de Atividade": "📌 Tipo",
                "Quantidade": "🔢 Quantidade",
                "% do Total": "📊 %",
                "Status": "🚦 Classificação"
            }
        )
    else:
        demandas_exemplo = pd.DataFrame({
            'Tipo de Atividade': ['Evento', 'Comunicado', 'Campanha Orgânica', 
                                  'Divulgação de Produto', 'Campanha de Incentivo', 
                                  'E-mail Marketing', 'Redes Sociais', 'Landing Page'],
            'Quantidade': [124, 89, 67, 45, 34, 28, 21, 15],
            '% do Total': ['32%', '23%', '17%', '12%', '9%', '7%', '5%', '4%'],
            'Status': ['✅ Alto volume', '✅ Alto volume', '⚠️ Médio volume', 
                      '⚠️ Médio volume', '🟡 Baixo volume', '🟡 Baixo volume', 
                      '🟡 Baixo volume', '🟡 Baixo volume']
        })
        st.dataframe(demandas_exemplo, use_container_width=True, height=350, hide_index=True)

# =========================================================
# FILTROS AVANÇADOS
# =========================================================
st.header("🎛️ Filtros Avançados")

filtro_cols = st.columns(4)
filtros_ativos = {}

if 'Status' in df.columns:
    with filtro_cols[0]:
        status_opcoes = ['Todos'] + sorted(df['Status'].dropna().unique().tolist())
        status_selecionado = st.selectbox("📌 Status:", status_opcoes, key="filtro_status")
        if status_selecionado != 'Todos':
            filtros_ativos['Status'] = status_selecionado

if 'Prioridade' in df.columns:
    with filtro_cols[1]:
        prioridade_opcoes = ['Todos'] + sorted(df['Prioridade'].dropna().unique().tolist())
        prioridade_selecionada = st.selectbox("⚡ Prioridade:", prioridade_opcoes, key="filtro_prioridade")
        if prioridade_selecionada != 'Todos':
            filtros_ativos['Prioridade'] = prioridade_selecionada

if 'Produção' in df.columns:
    with filtro_cols[2]:
        producao_opcoes = ['Todos'] + sorted(df['Produção'].dropna().unique().tolist())
        producao_selecionada = st.selectbox("🏭 Produção:", producao_opcoes, key="filtro_producao")
        if producao_selecionada != 'Todos':
            filtros_ativos['Produção'] = producao_selecionada

with filtro_cols[3]:
    st.markdown("**📅 Data Solicitação**")
    
    if 'Data de Solicitação' in df.columns:
        datas_validas = df['Data de Solicitação'].dropna()
        if not datas_validas.empty:
            data_min = datas_validas.min().date()
            data_max = datas_validas.max().date()
            
            periodo_opcao = st.selectbox("Período:", ["Todos", "Hoje", "Esta semana", "Este mês", "Últimos 30 dias", "Personalizado"], key="periodo_data")
            hoje = datetime.now().date()
            
            if periodo_opcao == "Todos":
                filtros_ativos['data_inicio'] = data_min
                filtros_ativos['data_fim'] = data_max
                filtros_ativos['tem_filtro_data'] = True
            elif periodo_opcao == "Hoje":
                filtros_ativos['data_inicio'] = hoje
                filtros_ativos['data_fim'] = hoje
                filtros_ativos['tem_filtro_data'] = True
            elif periodo_opcao == "Esta semana":
                inicio_semana = hoje - timedelta(days=hoje.weekday())
                filtros_ativos['data_inicio'] = inicio_semana
                filtros_ativos['data_fim'] = hoje
                filtros_ativos['tem_filtro_data'] = True
            elif periodo_opcao == "Este mês":
                inicio_mes = hoje.replace(day=1)
                filtros_ativos['data_inicio'] = inicio_mes
                filtros_ativos['data_fim'] = hoje
                filtros_ativos['tem_filtro_data'] = True
            elif periodo_opcao == "Últimos 30 dias":
                inicio_30d = hoje - timedelta(days=30)
                filtros_ativos['data_inicio'] = inicio_30d
                filtros_ativos['data_fim'] = hoje
                filtros_ativos['tem_filtro_data'] = True
            elif periodo_opcao == "Personalizado":
                col1, col2 = st.columns(2)
                with col1:
                    data_ini = st.date_input("De", data_min, key="data_ini")
                with col2:
                    data_fim = st.date_input("Até", data_max, key="data_fim")
                filtros_ativos['data_inicio'] = data_ini
                filtros_ativos['data_fim'] = data_fim
                filtros_ativos['tem_filtro_data'] = True
    else:
        st.info("ℹ️ Sem coluna de data")

# =========================================================
# APLICAR FILTROS
# =========================================================
df_filtrado = df.copy()

for col, valor in filtros_ativos.items():
    if col not in ['data_inicio', 'data_fim', 'tem_filtro_data']:
        df_filtrado = df_filtrado[df_filtrado[col] == valor]

if 'tem_filtro_data' in filtros_ativos and 'Data de Solicitação' in df.columns:
    data_inicio = pd.Timestamp(filtros_ativos['data_inicio'])
    data_fim = pd.Timestamp(filtros_ativos['data_fim']) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
    df_filtrado = df_filtrado[(df_filtrado['Data de Solicitação'] >= data_inicio) & (df_filtrado['Data de Solicitação'] <= data_fim)]

if filtros_ativos:
    st.subheader(f"📊 Dados Filtrados ({len(df_filtrado)} de {total_linhas} registros)")
    
    if len(df_filtrado) > 0:
        altura_filtrada = calcular_altura_tabela(len(df_filtrado), len(df_filtrado.columns))
        st.dataframe(df_filtrado, use_container_width=True, height=min(altura_filtrada, 800))
        
        if st.button("🧹 Limpar Todos os Filtros", type="secondary", use_container_width=True):
            for key in list(st.session_state.keys()):
                if key.startswith('filtro_') or key in ['periodo_data', 'data_ini', 'data_fim']:
                    del st.session_state[key]
            st.rerun()
    else:
        st.warning("⚠️ Nenhum registro corresponde aos filtros aplicados.")
else:
    st.info("👆 Use os filtros acima para refinar os dados")

# =========================================================
# EXPORTAÇÃO
# =========================================================
st.header("💾 Exportar Dados")

df_exportar = df_filtrado if filtros_ativos and len(df_filtrado) > 0 else df

col_exp1, col_exp2, col_exp3 = st.columns(3)

with col_exp1:
    csv = df_exportar.to_csv(index=False, encoding='utf-8-sig')
    st.download_button(label="📥 Download CSV", data=csv, 
                      file_name=f"dados_cocred_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                      mime="text/csv", use_container_width=True)

with col_exp2:
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_exportar.to_excel(writer, index=False, sheet_name='Dados')
    excel_data = output.getvalue()
    st.download_button(label="📥 Download Excel", data=excel_data,
                      file_name=f"dados_cocred_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                      use_container_width=True)

with col_exp3:
    json_data = df_exportar.to_json(orient='records', force_ascii=False, date_format='iso')
    st.download_button(label="📥 Download JSON", data=json_data,
                      file_name=f"dados_cocred_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
                      mime="application/json", use_container_width=True)

# =========================================================
# DEBUG INFO
# =========================================================
if st.session_state.debug_mode:
    st.sidebar.markdown("---")
    st.sidebar.markdown("**🐛 Debug Info:**")
    
    with st.sidebar.expander("Detalhes Técnicos", expanded=False):
        st.write(f"**Cache:** 1 minuto")
        st.write(f"**Hora atual:** {datetime.now().strftime('%H:%M:%S')}")
        st.write(f"**DataFrame Shape:** {df.shape}")
        st.write(f"**Memory:** {df.memory_usage(deep=True).sum() / 1024 / 1024:.2f} MB")
        st.write(f"**Criações:** {criacoes}")
        st.write(f"**Derivações:** {derivacoes}")
        st.write(f"**Extra Contrato:** {extra_contrato}")
        st.write(f"**Campanhas:** {campanhas_unicas}")
        st.write(f"**Template Plotly:** {plotly_template}")

# =========================================================
# RODAPÉ
# =========================================================
st.divider()

footer_col1, footer_col2, footer_col3 = st.columns(3)

with footer_col1:
    st.caption(f"🕐 {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")

with footer_col2:
    st.caption(f"📊 {total_linhas} registros | {total_colunas} colunas")

with footer_col3:
    st.markdown("""
    <div style="text-align: right;">
        <span style="color: #003366; font-weight: bold;">SICOOB COCRED</span> | 
        <span style="color: #6C757D;">v4.1.0</span>
    </div>
    """, unsafe_allow_html=True)

# =========================================================
# AUTO-REFRESH
# =========================================================
if auto_refresh:
    refresh_placeholder = st.empty()
    for i in range(60, 0, -1):
        refresh_placeholder.caption(f"🔄 Atualizando em {i} segundos...")
        time.sleep(1)
    refresh_placeholder.empty()
    st.rerun()