import streamlit as st
import pandas as pd
import requests
from io import BytesIO
import msal
from datetime import datetime
import pytz

# =========================================================
# CONFIGURAÇÕES DA API (ATUALIZE AQUI!)
# =========================================================
st.set_page_config(page_title="Dashboard de Campanhas - SICOOB COCRED", layout="wide")

# 1. SUAS CREDENCIAIS DA GRAPH API (do Azure AD)
MS_CLIENT_ID = st.secrets.get("MS_CLIENT_ID", "")        # Application ID
MS_CLIENT_SECRET = st.secrets.get("MS_CLIENT_SECRET", "") # Secret VALUE (o valor, não o ID!)
MS_TENANT_ID = st.secrets.get("MS_TENANT_ID", "")        # Directory ID

# 2. INFORMAÇÕES DO EXCEL ONLINE (CONFIGURAÇÃO CORRETA!)
USUARIO_PRINCIPAL = "cristini.cordesco@ideatoreamericas.com"  # ← USUÁRIO COM PONTO!
SHAREPOINT_FILE_ID = "01S7YQRRWMBXCV3AAHYZEIZGL55EPOZULE"     # ← NOVO FILE ID CORRETO
SHEET_NAME = "Demandas ID"  # ← NOME DA ABA

# =========================================================
# 1. AUTENTICAÇÃO MICROSOFT GRAPH
# =========================================================
@st.cache_resource
def get_msal_app():
    """Configura a aplicação MSAL com suas credenciais"""
    if not all([MS_CLIENT_ID, MS_CLIENT_SECRET, MS_TENANT_ID]):
        st.error("❌ Credenciais da API não configuradas!")
        st.info("""
        Configure no Streamlit Cloud:
        Settings → Secrets → Adicione:
        ```
        MS_CLIENT_ID = "seu-application-id"
        MS_CLIENT_SECRET = "seu-secret-value"  # O VALOR, não o ID!
        MS_TENANT_ID = "seu-tenant-id"
        ```
        """)
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
        st.error(f"❌ Erro ao configurar MSAL: {str(e)}")
        return None

@st.cache_data(ttl=3500)  # Token válido por ~1 hora
def get_access_token():
    """Obtém access token para Microsoft Graph"""
    app = get_msal_app()
    if not app:
        return None
    
    try:
        result = app.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )
        
        if "access_token" in result:
            return result["access_token"]
        else:
            error_msg = result.get("error_description", "Erro desconhecido")
            st.error(f"❌ Falha na autenticação: {error_msg}")
            return None
    except Exception as e:
        st.error(f"❌ Erro ao obter token: {str(e)}")
        return None

# =========================================================
# 2. CARREGAR DADOS DO EXCEL ONLINE (FUNÇÃO CORRIGIDA)
# =========================================================
@st.cache_data(ttl=300)  # Cache de 5 minutos para os dados
def carregar_dados_excel_online():
    """Carrega dados da aba 'Demandas ID' do Excel Online"""
    
    access_token = get_access_token()
    if not access_token:
        st.error("❌ Não foi possível obter token de acesso")
        return pd.DataFrame()
    
    # URL CORRETA para acessar o arquivo via Microsoft Graph
    file_url = f"https://graph.microsoft.com/v1.0/users/{USUARIO_PRINCIPAL}/drive/items/{SHAREPOINT_FILE_ID}/content"
    
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/octet-stream"
    }
    
    try:
        with st.spinner("🔄 Conectando ao Excel Online..."):
            # Baixar o arquivo Excel
            response = requests.get(file_url, headers=headers, timeout=30)
        
        if response.status_code == 200:
            # Ler o arquivo Excel
            excel_file = BytesIO(response.content)
            
            # Tentar ler a aba específica "Demandas ID"
            try:
                df = pd.read_excel(excel_file, sheet_name=SHEET_NAME, engine='openpyxl')
                st.sidebar.success(f"✅ Aba '{SHEET_NAME}' carregada")
            except Exception as e:
                st.sidebar.warning(f"⚠️ Não encontrei aba '{SHEET_NAME}'. Tentando primeira aba...")
                df = pd.read_excel(excel_file, engine='openpyxl')
            
            # Verificar se carregou dados
            if df.empty:
                st.error(f"❌ O arquivo está vazia ou não contém dados.")
                return pd.DataFrame()
            
            # Pegar informações do arquivo
            metadata_url = f"https://graph.microsoft.com/v1.0/users/{USUARIO_PRINCIPAL}/drive/items/{SHAREPOINT_FILE_ID}"
            meta_response = requests.get(metadata_url, headers=headers)
            
            if meta_response.status_code == 200:
                metadata = meta_response.json()
                last_modified = metadata.get('lastModifiedDateTime', '')
                
                if last_modified:
                    # Converter para horário Brasil
                    dt = datetime.fromisoformat(last_modified.replace('Z', '+00:00'))
                    dt_brazil = dt.astimezone(pytz.timezone('America/Sao_Paulo'))
                    
                    # Mostrar informações de atualização
                    st.sidebar.caption(f"📅 Última atualização: {dt_brazil.strftime('%d/%m/%Y %H:%M')}")
                    
                    # Mostrar quem modificou
                    modified_by = metadata.get('lastModifiedBy', {}).get('user', {}).get('displayName', '')
                    if modified_by:
                        st.sidebar.caption(f"👤 Por: {modified_by}")
            
            # Informações do arquivo
            st.sidebar.caption(f"📊 {len(df)} registros × {len(df.columns)} colunas")
            
            return df
            
        elif response.status_code == 404:
            st.error("❌ Arquivo não encontrado no OneDrive")
            st.info(f"""
            **Verifique:**
            1. File ID: {SHAREPOINT_FILE_ID}
            2. Usuário: {USUARIO_PRINCIPAL}
            3. O arquivo ainda existe no OneDrive
            """)
            
        elif response.status_code == 403:
            st.error("❌ Permissão negada")
            st.info("""
            **Solução:**
            1. Verifique se o app tem permissão **"Files.Read.All"**
            2. Confirme que deu **"Admin Consent"** no Azure AD
            3. App Registration → API permissions → Files.Read.All
            """)
            
        elif response.status_code == 401:
            st.error("❌ Token expirado ou inválido")
            st.cache_data.clear()  # Limpar cache para novo token
            
        else:
            st.error(f"❌ Erro HTTP {response.status_code}")
            st.text(f"Detalhes: {response.text[:200]}")
        
        return pd.DataFrame()
        
    except requests.exceptions.Timeout:
        st.error("⏱️ Timeout - Verifique sua conexão com a internet")
        return pd.DataFrame()
        
    except Exception as e:
        st.error(f"❌ Erro inesperado: {str(e)}")
        return pd.DataFrame()

# =========================================================
# 3. INTERFACE STREAMLIT
# =========================================================

# Título principal
st.title("📊 Dashboard de Campanhas – SICOOB COCRED")
st.caption(f"🔗 Conectado ao Excel Online | Aba: {SHEET_NAME}")

# Sidebar - Controles
st.sidebar.header("⚙️ Controles")

# Botão de atualização
if st.sidebar.button("🔄 Atualizar agora", width='stretch', type="primary"):
    st.cache_data.clear()
    st.rerun()

# Status da conexão
st.sidebar.markdown("---")
st.sidebar.markdown("**🔗 Status da Conexão:**")

# Testar conexão
if st.sidebar.button("🔍 Testar Conexão API", width='stretch'):
    token = get_access_token()
    if token:
        st.sidebar.success("✅ API: Conectada")
        st.sidebar.code(f"Token: ...{token[-10:]}")
    else:
        st.sidebar.error("❌ API: Falha na conexão")

# Link para editar
st.sidebar.markdown("---")
st.sidebar.markdown("**📝 Editar planilha:**")
st.sidebar.markdown(f"""
[✏️ Abrir no Excel Online](https://agenciaideatore-my.sharepoint.com/:x:/g/personal/cristini_cordesco_ideatoreamericas_com/IQDMDcVdgAfGSIyZfeke7NFkAatm3fhI0-X4r6gIPQJmosY?e=R0o2FK)

**Instruções:**
1. Edite na aba **"{SHEET_NAME}"**
2. Salve (Ctrl+S)
3. Dashboard atualiza em 5min
4. Ou clique em "Atualizar agora"
""")

# =========================================================
# 4. CARREGAR DADOS
# =========================================================

# Carregar dados do Excel Online
df = carregar_dados_excel_online()

# Verificar se carregou
if df.empty:
    st.error("""
    ❌ **Não foi possível carregar os dados**
    
    **Possíveis causas:**
    1. Credenciais da API não configuradas
    2. Arquivo não encontrado no OneDrive
    3. Permissões insuficientes
    4. Aba '{SHEET_NAME}' não existe
    """)
    
    # Mostrar configuração necessária
    with st.expander("🔧 Configuração necessária"):
        st.markdown(f"""
        ### 1. Configure as Secrets no Streamlit Cloud:
        ```toml
        MS_CLIENT_ID = "2b3245ac-e6f7-4f70-beee-f78f5f31598e"
        MS_CLIENT_SECRET = "sua-chave-secreta-aqui"
        MS_TENANT_ID = "46d481f9-b227-467f-8b1a-b46734313c90"
        ```
        
        ### 2. Verifique no Azure AD:
        - App tem permissão **Files.Read.All**
        - **Admin Consent** foi dado
        - Client secret está ativo
        
        ### 3. Verifique o OneDrive:
        - Usuário: **{USUARIO_PRINCIPAL}**
        - File ID: **{SHAREPOINT_FILE_ID}**
        - Aba: **{SHEET_NAME}**
        """)
    
    # Fallback: Upload manual
    st.warning("⚠️ Enquanto isso, use upload manual:")
    uploaded_file = st.file_uploader("📤 Upload do Excel", type=["xlsx", "xls"])
    
    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file, sheet_name=SHEET_NAME, engine='openpyxl')
            st.success("✅ Dados carregados manualmente")
        except:
            df = pd.read_excel(uploaded_file, engine='openpyxl')
            st.warning("⚠️ Usando primeira aba do arquivo")
    else:
        st.stop()

# =========================================================
# 5. PROCESSAMENTO DOS DADOS
# =========================================================

# Exemplo de tratamento - AJUSTE CONFORME SUA PLANILHA
st.header("📈 Análise dos Dados")

# Mostrar dataframe
st.subheader("Dados Brutos")
st.dataframe(df, width='stretch', height=400)

# Estatísticas básicas
st.subheader("📊 Estatísticas")
col1, col2, col3 = st.columns(3)

with col1:
    st.metric("Total de Registros", len(df))

with col2:
    st.metric("Total de Colunas", len(df.columns))

with col3:
    # Verificar se há coluna de data
    date_cols = [col for col in df.columns if 'data' in col.lower() or 'date' in col.lower()]
    if date_cols:
        try:
            latest_date = pd.to_datetime(df[date_cols[0]]).max()
            st.metric("Data Mais Recente", latest_date.strftime('%d/%m/%Y'))
        except:
            st.metric("Amostra", "5 registros")

# Processamento específico para "Prazo em dias" (se existir)
if "Prazo em dias" in df.columns:
    st.subheader("⏱️ Análise de Prazos")
    
    # Converter para string e limpar
    df["Prazo em dias"] = df["Prazo em dias"].astype(str).str.strip()
    
    # Classificar situação do prazo
    df["Situação do Prazo"] = df["Prazo em dias"].apply(
        lambda x: "Prazo encerrado" if "encerrado" in x.lower() else "Em prazo"
    )
    
    # Tentar converter para numérico
    df["Prazo em dias"] = pd.to_numeric(df["Prazo em dias"], errors="coerce")
    
    # Mostrar distribuição
    if not df["Situação do Prazo"].empty:
        situacao_counts = df["Situação do Prazo"].value_counts()
        st.bar_chart(situacao_counts)

# Verificar outras colunas importantes
st.subheader("🔍 Colunas Disponíveis")

# Listar todas as colunas
cols = st.columns(3)
for i, col_name in enumerate(df.columns):
    with cols[i % 3]:
        with st.expander(f"**{col_name}**"):
            st.write(f"Tipo: {df[col_name].dtype}")
            st.write(f"Valores únicos: {df[col_name].nunique()}")
            st.write(f"Valores nulos: {df[col_name].isnull().sum()}")
            
            # Mostrar amostra
            if df[col_name].dtype == 'object':
                st.write("Amostra:", df[col_name].head(5).tolist())

# =========================================================
# 6. FILTROS INTERATIVOS
# =========================================================
st.header("🎛️ Filtros")

# Filtro por colunas específicas (se existirem)
filtro_cols = st.columns(3)

# Coluna 1: Filtro por tipo (se houver coluna 'Tipo' ou similar)
tipo_cols = [col for col in df.columns if 'tipo' in col.lower() or 'categoria' in col.lower()]
if tipo_cols:
    with filtro_cols[0]:
        tipos = df[tipo_cols[0]].dropna().unique()
        selected_tipos = st.multiselect(f"Filtrar por {tipo_cols[0]}", options=tipos)
        if selected_tipos:
            df = df[df[tipo_cols[0]].isin(selected_tipos)]

# Coluna 2: Filtro por status (se houver coluna 'Status' ou similar)
status_cols = [col for col in df.columns if 'status' in col.lower() or 'situação' in col.lower()]
if status_cols:
    with filtro_cols[1]:
        statuses = df[status_cols[0]].dropna().unique()
        selected_status = st.multiselect(f"Filtrar por {status_cols[0]}", options=statuses)
        if selected_status:
            df = df[df[status_cols[0]].isin(selected_status)]

# Coluna 3: Filtro por data (se houver coluna de data)
date_cols = [col for col in df.columns if 'data' in col.lower() or 'date' in col.lower()]
if date_cols:
    with filtro_cols[2]:
        try:
            df[date_cols[0]] = pd.to_datetime(df[date_cols[0]], errors='coerce')
            min_date = df[date_cols[0]].min()
            max_date = df[date_cols[0]].max()
            
            if pd.notna(min_date) and pd.notna(max_date):
                date_range = st.date_input(
                    f"Filtrar por {date_cols[0]}",
                    value=(min_date, max_date),
                    min_value=min_date,
                    max_value=max_date
                )
                
                if len(date_range) == 2:
                    start_date, end_date = date_range
                    df = df[(df[date_cols[0]] >= pd.Timestamp(start_date)) & 
                           (df[date_cols[0]] <= pd.Timestamp(end_date))]
        except:
            pass

# Mostrar dados filtrados
st.subheader("Dados Filtrados")
st.dataframe(df, width='stretch', height=300)

# =========================================================
# 7. EXPORTAÇÃO DE DADOS
# =========================================================
st.header("💾 Exportar Dados")

col_export1, col_export2 = st.columns(2)

with col_export1:
    # Exportar para CSV
    csv = df.to_csv(index=False).encode('utf-8-sig')
    st.download_button(
        label="📥 Download CSV",
        data=csv,
        file_name=f"dados_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
        mime="text/csv",
        width='stretch'
    )

with col_export2:
    # Exportar para Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Dados')
    excel_data = output.getvalue()
    
    st.download_button(
        label="📥 Download Excel",
        data=excel_data,
        file_name=f"dados_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width='stretch'
    )

# =========================================================
# 8. RODAPÉ COM INFORMAÇÕES
# =========================================================
st.divider()

col_footer1, col_footer2, col_footer3 = st.columns(3)

with col_footer1:
    st.caption(f"📅 {datetime.now().strftime('%d/%m/%Y %H:%M')}")

with col_footer2:
    st.caption("🔄 Atualização automática a cada 5min")

with col_footer3:
    st.caption(f"📊 {len(df)} registros | Aba: {SHEET_NAME}")

# =========================================================
# 9. CONFIGURAÇÃO DAS SECRETS (instruções)
# =========================================================
with st.sidebar.expander("⚙️ Configurar Secrets", expanded=False):
    st.markdown("""
    ### No Streamlit Cloud:
    
    1. Vá em **Settings**
    2. Clique em **Secrets**
    3. Cole:
    ```toml
    MS_CLIENT_ID = "2b3245ac-e6f7-4f70-beee-f78f5f31598e"
    MS_CLIENT_SECRET = "sua-chave-secreta-aqui"
    MS_TENANT_ID = "46d481f9-b227-467f-8b1a-b46734313c90"
    ```
    
    ### Como obter as credenciais:
    1. **MS_CLIENT_ID**: Application ID do Azure AD
    2. **MS_CLIENT_SECRET**: VALUE do client secret (não o ID!)
    3. **MS_TENANT_ID**: Directory ID do Azure AD
    
    ### Permissões necessárias no Azure AD:
    - Files.Read.All (para ler arquivos do OneDrive)
    - User.Read (permissão básica)
    - Sites.Read.All (opcional, para SharePoint)
    """)
    
    st.markdown("---")
    st.markdown("**🔧 Configuração atual:**")
    st.code(f"""
    Usuário: {USUARIO_PRINCIPAL}
    File ID: {SHAREPOINT_FILE_ID}
    Aba: {SHEET_NAME}
    """)

# =========================================================
# 10. MODO DEBUG (apenas para desenvolvimento)
# =========================================================
if st.sidebar.checkbox("🐛 Modo Debug", value=False):
    with st.sidebar.expander("Informações de Debug"):
        st.write("**Configurações:**")
        st.json({
            "MS_CLIENT_ID": MS_CLIENT_ID[:8] + "..." if MS_CLIENT_ID else "Não configurado",
            "MS_TENANT_ID": MS_TENANT_ID[:8] + "..." if MS_TENANT_ID else "Não configurado",
            "USUARIO_PRINCIPAL": USUARIO_PRINCIPAL,
            "SHAREPOINT_FILE_ID": SHAREPOINT_FILE_ID,
            "SHEET_NAME": SHEET_NAME
        })
        
        if not df.empty:
            st.write("**Informações do DataFrame:**")
            st.write(f"- Shape: {df.shape}")
            st.write(f"- Colunas: {list(df.columns)}")
            st.write(f"- Tipos de dados: {df.dtypes.to_dict()}")
            
            # Testar token
            token = get_access_token()
            if token:
                st.success(f"Token ativo: ...{token[-10:]}")
            else:
                st.error("Token não disponível")