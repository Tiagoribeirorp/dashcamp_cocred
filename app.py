import streamlit as st
import pandas as pd
import requests
from io import BytesIO
import msal
from datetime import datetime
import pytz
import time

# =========================================================
# CONFIGURAÇÕES DA API
# =========================================================
st.set_page_config(page_title="Dashboard de Campanhas - SICOOB COCRED", layout="wide")

# 1. CREDENCIAIS DA API
MS_CLIENT_ID = st.secrets.get("MS_CLIENT_ID", "")
MS_CLIENT_SECRET = st.secrets.get("MS_CLIENT_SECRET", "")
MS_TENANT_ID = st.secrets.get("MS_TENANT_ID", "")

# 2. INFORMAÇÕES DO EXCEL (CONFIGURADO CORRETAMENTE!)
USUARIO_PRINCIPAL = "cristini.cordesco@ideatoreamericas.com"
SHAREPOINT_FILE_ID = "01S7YQRRWMBXCV3AAHYZEIZGL55EPOZULE"
SHEET_NAME = "Demandas ID"

# =========================================================
# 1. AUTENTICAÇÃO MICROSOFT GRAPH
# =========================================================
@st.cache_resource
def get_msal_app():
    """Configura a aplicação MSAL"""
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

@st.cache_data(ttl=1800)  # 30 minutos
def get_access_token():
    """Obtém token de acesso"""
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
# 2. CARREGAR DADOS (VERSÃO OTIMIZADA)
# =========================================================
@st.cache_data(ttl=60, show_spinner="🔄 Baixando dados do Excel...")  # APENAS 1 MINUTO!
def carregar_dados_excel_online():
    """Carrega dados da aba 'Demandas ID' com cache curto"""
    
    access_token = get_access_token()
    if not access_token:
        st.error("❌ Token não disponível")
        return pd.DataFrame()
    
    file_url = f"https://graph.microsoft.com/v1.0/users/{USUARIO_PRINCIPAL}/drive/items/{SHAREPOINT_FILE_ID}/content"
    
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/octet-stream"
    }
    
    try:
        # Baixar arquivo
        response = requests.get(file_url, headers=headers, timeout=45)
        
        if response.status_code == 200:
            # Ler Excel
            excel_file = BytesIO(response.content)
            
            # DEBUG: Mostrar tamanho
            if st.session_state.get('debug_mode', False):
                st.sidebar.info(f"📦 Arquivo: {len(response.content):,} bytes")
            
            # Ler aba específica
            try:
                df = pd.read_excel(excel_file, sheet_name=SHEET_NAME, engine='openpyxl')
                
                # DEBUG: Mostrar informações
                if st.session_state.get('debug_mode', False):
                    st.sidebar.success(f"✅ {len(df)} linhas carregadas")
                
                return df
                
            except Exception as e:
                # Tentar primeira aba
                st.warning(f"⚠️ Erro na aba '{SHEET_NAME}': {str(e)[:100]}")
                excel_file.seek(0)
                df = pd.read_excel(excel_file, engine='openpyxl')
                return df
                
        else:
            st.error(f"❌ Erro {response.status_code}")
            return pd.DataFrame()
            
    except Exception as e:
        st.error(f"❌ Erro: {str(e)}")
        return pd.DataFrame()

# =========================================================
# 3. INTERFACE PRINCIPAL
# =========================================================

# Título
st.title("📊 Dashboard de Campanhas – SICOOB COCRED")
st.caption(f"🔗 Conectado ao Excel Online | Aba: {SHEET_NAME} | Última atualização: {datetime.now().strftime('%H:%M:%S')}")

# Sidebar
st.sidebar.header("⚙️ Controles")

# Controle de debug
if 'debug_mode' not in st.session_state:
    st.session_state.debug_mode = False

st.session_state.debug_mode = st.sidebar.checkbox("🐛 Modo Debug", value=st.session_state.debug_mode)

# Botão de atualização FORÇADA
if st.sidebar.button("🔄 ATUALIZAR AGORA (Forçar)", type="primary", use_container_width=True):
    st.cache_data.clear()
    st.rerun()

# Status
st.sidebar.markdown("---")
st.sidebar.markdown("**📊 Status:**")

# Testar conexão
if st.sidebar.button("🔍 Testar Conexão", use_container_width=True):
    token = get_access_token()
    if token:
        st.sidebar.success("✅ API Conectada")
    else:
        st.sidebar.error("❌ API Offline")

# Link para Excel
st.sidebar.markdown("---")
st.sidebar.markdown("**📝 Editar Excel:**")
st.sidebar.markdown(f"""
[✏️ Abrir no Excel Online](https://agenciaideatore-my.sharepoint.com/:x:/g/personal/cristini_cordesco_ideatoreamericas_com/IQDMDcVdgAfGSIyZfeke7NFkAatm3fhI0-X4r6gIPQJmosY)

**Lembre-se:**
1. Edite e **SALVE** (Ctrl+S)
2. Clique em **"ATUALIZAR AGORA"**
3. Dados atualizam em **1 minuto**
""")

# =========================================================
# 4. CARREGAR E MOSTRAR DADOS
# =========================================================

# Carregar dados
with st.spinner("📥 Carregando dados do Excel..."):
    df = carregar_dados_excel_online()

# Verificar se tem dados
if df.empty:
    st.error("❌ Nenhum dado carregado")
    st.stop()

# Mostrar contador REAL
total_linhas = len(df)
total_colunas = len(df.columns)

st.success(f"✅ **{total_linhas} registros** carregados com sucesso!")
st.info(f"📋 **Colunas:** {', '.join(df.columns.tolist()[:5])}{'...' if len(df.columns) > 5 else ''}")

# =========================================================
# 5. VISUALIZAÇÃO COMPLETA DOS DADOS
# =========================================================

st.header("📋 Dados Completos")

# Configurar para mostrar TUDO
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

# Opções de visualização
tab1, tab2, tab3 = st.tabs(["📊 Dados Completos", "📈 Estatísticas", "🔍 Pesquisa"])

with tab1:
    # ALTURA DINÂMICA baseada no número de linhas
    altura_tabela = min(800, 200 + (total_linhas * 25))
    
    st.subheader(f"Todos os {total_linhas} registros")
    st.dataframe(df, height=altura_tabela, use_container_width=True)
    
    # Contadores
    col_count1, col_count2, col_count3 = st.columns(3)
    with col_count1:
        st.metric("Linhas", total_linhas)
    with col_count2:
        st.metric("Colunas", total_colunas)
    with col_count3:
        ultima_data = df['Data de Solicitação'].max() if 'Data de Solicitação' in df.columns else "N/A"
        st.metric("Última Solicitação", 
                 ultima_data.strftime('%d/%m/%Y') if hasattr(ultima_data, 'strftime') else ultima_data)

with tab2:
    # Estatísticas
    st.subheader("📈 Estatísticas dos Dados")
    
    col_stat1, col_stat2 = st.columns(2)
    
    with col_stat1:
        st.write("**Resumo Numérico:**")
        st.dataframe(df.describe(), use_container_width=True)
    
    with col_stat2:
        st.write("**Informações das Colunas:**")
        info_df = pd.DataFrame({
            'Coluna': df.columns,
            'Tipo': df.dtypes.astype(str),
            'Únicos': [df[col].nunique() for col in df.columns],
            'Nulos': [df[col].isnull().sum() for col in df.columns],
            '% Preenchido': [f"{(1 - df[col].isnull().sum() / total_linhas) * 100:.1f}%" 
                           for col in df.columns]
        })
        st.dataframe(info_df, use_container_width=True, height=400)
    
    # Distribuição por colunas importantes
    st.subheader("📊 Distribuições")
    
    cols_dist = st.columns(2)
    
    # Status
    if 'Status' in df.columns:
        with cols_dist[0]:
            st.write("**Distribuição por Status:**")
            status_counts = df['Status'].value_counts()
            st.bar_chart(status_counts)
    
    # Prioridade
    if 'Prioridade' in df.columns:
        with cols_dist[1]:
            st.write("**Distribuição por Prioridade:**")
            prioridade_counts = df['Prioridade'].value_counts()
            st.bar_chart(prioridade_counts)

with tab3:
    # Pesquisa e filtros
    st.subheader("🔍 Pesquisa nos Dados")
    
    # Pesquisa por texto
    texto_pesquisa = st.text_input("🔎 Pesquisar em todas as colunas:", placeholder="Digite um termo para buscar...")
    
    if texto_pesquisa:
        # Criar máscara de pesquisa
        mask = pd.Series(False, index=df.index)
        for col in df.columns:
            if df[col].dtype == 'object':  # Apenas colunas de texto
                try:
                    mask = mask | df[col].astype(str).str.contains(texto_pesquisa, case=False, na=False)
                except:
                    pass
        
        resultados = df[mask]
        st.write(f"**{len(resultados)} resultado(s) encontrado(s):**")
        st.dataframe(resultados, use_container_width=True, height=300)
    else:
        st.info("Digite um termo acima para pesquisar nos dados")

# =========================================================
# 6. FILTROS INTERATIVOS
# =========================================================

st.header("🎛️ Filtros Avançados")

# Criar filtros dinâmicos
filtro_cols = st.columns(3)

filtros_ativos = {}

# Filtro 1: Status
if 'Status' in df.columns:
    with filtro_cols[0]:
        status_opcoes = ['Todos'] + sorted(df['Status'].dropna().unique().tolist())
        status_selecionado = st.selectbox("Status:", status_opcoes)
        if status_selecionado != 'Todos':
            filtros_ativos['Status'] = status_selecionado

# Filtro 2: Prioridade
if 'Prioridade' in df.columns:
    with filtro_cols[1]:
        prioridade_opcoes = ['Todos'] + sorted(df['Prioridade'].dropna().unique().tolist())
        prioridade_selecionada = st.selectbox("Prioridade:", prioridade_opcoes)
        if prioridade_selecionada != 'Todos':
            filtros_ativos['Prioridade'] = prioridade_selecionada

# Filtro 3: Produção
if 'Produção' in df.columns:
    with filtro_cols[2]:
        producao_opcoes = ['Todos'] + sorted(df['Produção'].dropna().unique().tolist())
        producao_selecionada = st.selectbox("Produção:", producao_opcoes)
        if producao_selecionada != 'Todos':
            filtros_ativos['Produção'] = producao_selecionada

# Aplicar filtros
df_filtrado = df.copy()
for col, valor in filtros_ativos.items():
    df_filtrado = df_filtrado[df_filtrado[col] == valor]

# Mostrar dados filtrados
if filtros_ativos:
    st.subheader(f"📊 Dados Filtrados ({len(df_filtrado)} de {total_linhas} registros)")
    st.dataframe(df_filtrado, use_container_width=True, height=400)
    
    # Botão para limpar filtros
    if st.button("🧹 Limpar Filtros"):
        st.rerun()
else:
    st.info("👆 Use os filtros acima para refinar os dados")

# =========================================================
# 7. ANÁLISE DE PRAZOS (se houver coluna)
# =========================================================

if 'Prazo em dias' in df.columns:
    st.header("⏱️ Análise de Prazos")
    
    # Processar prazos
    df['Prazo em dias'] = df['Prazo em dias'].astype(str).str.strip()
    
    # Classificar
    def classificar_prazo(x):
        if pd.isna(x) or str(x).lower() == 'nan':
            return "Sem prazo"
        elif 'encerrado' in str(x).lower():
            return "Prazo encerrado"
        else:
            try:
                dias = int(float(str(x)))
                if dias < 0:
                    return "Atrasado"
                elif dias <= 3:
                    return "Urgente (≤3 dias)"
                elif dias <= 7:
                    return "Próxima semana"
                else:
                    return "Em prazo"
            except:
                return str(x)
    
    df['Situação do Prazo'] = df['Prazo em dias'].apply(classificar_prazo)
    
    # Mostrar distribuição
    col_prazo1, col_prazo2 = st.columns(2)
    
    with col_prazo1:
        situacao_counts = df['Situação do Prazo'].value_counts()
        st.bar_chart(situacao_counts)
    
    with col_prazo2:
        st.write("**Distribuição:**")
        for situacao, count in situacao_counts.items():
            st.write(f"• {situacao}: {count} ({count/total_linhas*100:.1f}%)")

# =========================================================
# 8. EXPORTAÇÃO
# =========================================================

st.header("💾 Exportar Dados")

col_exp1, col_exp2, col_exp3 = st.columns(3)

with col_exp1:
    # CSV
    csv = df.to_csv(index=False, encoding='utf-8-sig')
    st.download_button(
        label="📥 Download CSV",
        data=csv,
        file_name=f"dados_cocred_{datetime.now().strftime('%Y%m%d')}.csv",
        mime="text/csv",
        use_container_width=True
    )

with col_exp2:
    # Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Dados')
    excel_data = output.getvalue()
    
    st.download_button(
        label="📥 Download Excel",
        data=excel_data,
        file_name=f"dados_cocred_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

with col_exp3:
    # JSON
    json_data = df.to_json(orient='records', force_ascii=False)
    st.download_button(
        label="📥 Download JSON",
        data=json_data,
        file_name=f"dados_cocred_{datetime.now().strftime('%Y%m%d')}.json",
        mime="application/json",
        use_container_width=True
    )

# =========================================================
# 9. DEBUG INFO (apenas se ativado)
# =========================================================

if st.session_state.debug_mode:
    st.sidebar.markdown("---")
    st.sidebar.markdown("**🐛 Debug Info:**")
    
    with st.sidebar.expander("Detalhes Técnicos"):
        st.write(f"**Cache:** 1 minuto")
        st.write(f"**Hora atual:** {datetime.now().strftime('%H:%M:%S')}")
        
        token = get_access_token()
        if token:
            st.success(f"Token: ...{token[-10:]}")
        
        st.write(f"**DataFrame Info:**")
        st.write(f"- Shape: {df.shape}")
        st.write(f"- Memory: {df.memory_usage(deep=True).sum() / 1024 / 1024:.2f} MB")
        
        # Mostrar primeiras e últimas linhas
        st.write("**Primeiras 3 linhas:**")
        st.dataframe(df.head(3))
        
        st.write("**Últimas 3 linhas:**")
        st.dataframe(df.tail(3))

# =========================================================
# 10. RODAPÉ
# =========================================================

st.divider()

footer_col1, footer_col2, footer_col3 = st.columns(3)

with footer_col1:
    st.caption(f"🕐 {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")

with footer_col2:
    st.caption(f"📊 {total_linhas} registros | {total_colunas} colunas")

with footer_col3:
    st.caption("🔄 Atualiza a cada 1 minuto")

# =========================================================
# 11. AUTO-REFRESH (opcional)
# =========================================================

# Auto-refresh a cada 60 segundos (opcional)
# Comente se não quiser auto-refresh
auto_refresh = st.sidebar.checkbox("🔄 Auto-refresh (60s)", value=False)

if auto_refresh:
    time.sleep(60)
    st.rerun()