import streamlit as st
import pandas as pd
import requests
from io import BytesIO
import msal
from datetime import datetime
import pytz
import time

# =========================================================
# CONFIGURAÇÕES DA API (AJUSTE AQUI!)
# =========================================================
st.set_page_config(page_title="Dashboard de Campanhas - SICOOB COCRED", layout="wide")

# 1. SUAS CREDENCIAIS DA GRAPH API (do Azure AD)
MS_CLIENT_ID = st.secrets.get("MS_CLIENT_ID", "")        # Application ID
MS_CLIENT_SECRET = st.secrets.get("MS_CLIENT_SECRET", "") # Secret VALUE
MS_TENANT_ID = st.secrets.get("MS_TENANT_ID", "")        # Directory ID

# 2. INFORMAÇÕES DO SEU EXCEL ONLINE
SHAREPOINT_FILE_ID = "IQDMDcVdgAfGSIyZfeke7NFkAatm3fhI0-X4r6gIPQJmosY"  # ID do arquivo
SHEET_NAME = "Demandas ID"  # ← NOME DA ABA QUE VOCÊ MENCIONOU!

# 3. SITE DO SHAREPOINT (do seu link)
SHAREPOINT_SITE = "agenciaideatore.sharepoint.com"
SHAREPOINT_SITE_PATH = "/personal/cristini_cordesco_ideatoreamericas_com"

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
# 2. CARREGAR DADOS DO EXCEL ONLINE
# =========================================================
@st.cache_data(ttl=300)  # Cache de 5 minutos para os dados
def carregar_dados_excel_online():
    """Carrega dados da aba 'Demandas ID' do Excel Online"""
    
    access_token = get_access_token()
    if not access_token:
        return pd.DataFrame()
    
    # URL para baixar o arquivo Excel
    file_url = f"https://graph.microsoft.com/v1.0/drives/root/items/{SHAREPOINT_FILE_ID}/content"
    
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
            except Exception as e:
                st.warning(f"⚠️ Não encontrei aba '{SHEET_NAME}'. Tentando primeira aba...")
                df = pd.read_excel(excel_file, engine='openpyxl')
            
            # Verificar se carregou dados
            if df.empty:
                st.error(f"❌ A aba '{SHEET_NAME}' está vazia ou não encontrada.")
                return pd.DataFrame()
            
            # Pegar informações do arquivo
            metadata_url = f"https://graph.microsoft.com/v1.0/drives/root/items/{SHAREPOINT_FILE_ID}"
            meta_response = requests.get(metadata_url, headers=headers)
            
            if meta_response.status_code == 200:
                metadata = meta_response.json()
                last_modified = metadata.get('lastModifiedDateTime', '')
                
                if last_modified:
                    # Converter para horário Brasil
                    dt = datetime.fromisoformat(last_modified.replace('Z', '+00:00'))
                    dt_brazil = dt.astimezone(pytz.timezone('America/Sao_Paulo'))
                    
                    # Mostrar no sidebar
                    st.sidebar.success(f"✅ Conectado: {SHEET_NAME}")
                    st.sidebar.caption(f"📅 Última atualização: {dt_brazil.strftime('%d/%m %H:%M')}")
                    
                    # Mostrar quem modificou
                    modified_by = metadata.get('lastModifiedBy', {}).get('user', {}).get('displayName', '')
                    if modified_by:
                        st.sidebar.caption(f"👤 Por: {modified_by}")
            
            st.sidebar.caption(f"📊 {len(df)} registros carregados")
            
            return df
            
        elif response.status_code == 404:
            st.error("❌ Arquivo não encontrado no SharePoint")
            st.info(f"Verifique o File ID: {SHAREPOINT_FILE_ID}")
            
        elif response.status_code == 403:
            st.error("❌ Permissão negada")
            st.info("""
            **Solução:**
            1. Verifique se o app tem permissão "Files.Read.All"
            2. Confirme que deu "Admin Consent" no Azure AD
            """)
            
        elif response.status_code == 401:
            st.error("❌ Token expirado")
            st.cache_data.clear()  # Limpar cache para novo token
            
        else:
            st.error(f"❌ Erro HTTP {response.status_code}")
            st.text(f"Resposta: {response.text[:200]}")
        
        return pd.DataFrame()
        
    except requests.exceptions.Timeout:
        st.error("⏱️ Timeout - Verifique sua conexão")
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

# Botão de atualização - CORRIGIDO AQUI (1ª ocorrência)
if st.sidebar.button("🔄 Atualizar agora", width='stretch', type="primary"):  # <-- CORREÇÃO
    st.cache_data.clear()
    st.rerun()

# Status da conexão
st.sidebar.markdown("---")
st.sidebar.markdown("**🔗 Status da Conexão:**")

# Testar conexão - CORRIGIDO AQUI (2ª ocorrência)
if st.sidebar.button("🔍 Testar Conexão API", width='stretch'):  # <-- CORREÇÃO
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
    2. Arquivo não encontrado no SharePoint
    3. Permissões insuficientes
    4. Aba '{SHEET_NAME}' não existe
    """)
    
    # Mostrar configuração necessária
    with st.expander("🔧 Configuração necessária"):
        st.markdown("""
        ### 1. Configure as Secrets no Streamlit Cloud:
        ```toml
        MS_CLIENT_ID = "{seu-application-id}"
        MS_CLIENT_SECRET = "{seu-secret-value}"
        MS_TENANT_ID = "{seu-tenant-id}"
        ```
        
        ### 2. Verifique no Azure AD:
        - App tem permissão **Files.Read.All**
        - **Admin Consent** foi dado
        - Client secret está ativo
        
        ### 3. Verifique o Excel Online:
        - Arquivo existe no link acima
        - Aba se chama **"{SHEET_NAME}"**
        - Você tem acesso ao arquivo
        """)
    
    # Fallback: Upload manual
    st.warning("⚠️ Enquanto isso, use upload manual:")
    
    # Uploader - CORRIGIDO AQUI (3ª ocorrência, se houver)
    # Verificando se há mais botões ou componentes com use_container_width
    # Parece que não há no uploader, mas se houver um botão aqui:
    
    uploaded_file = st.file_uploader("📤 Upload do Excel", type=["xlsx"])
    
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
# 5. SEU PROCESSAMENTO ORIGINAL (MANTENHA SEU CÓDIGO AQUI!)
# =========================================================
# COLE TODO O SEU CÓDIGO DE PROCESSAMENTO A PARTIR DAQUI

# Exemplo do SEU tratamento (substitua pelo seu real):
if "Prazo em dias" in df.columns:
    df["Prazo em dias"] = df["Prazo em dias"].astype(str).str.strip()
    
    df["Situação do Prazo"] = df["Prazo em dias"].apply(
        lambda x: "Prazo encerrado" if "encerrado" in x.lower() else "Em prazo"
    )
    
    df["Prazo em dias"] = pd.to_numeric(df["Prazo em dias"], errors="coerce")

# ... Continue com TODO o seu código restante ...

# ATENÇÃO: Se você tiver mais botões ou componentes Streamlit no seu código de processamento,
# verifique e substitua use_container_width por width='stretch' ou width='content'

# =========================================================
# 6. RODAPÉ COM INFORMAÇÕES
# =========================================================
st.divider()

col1, col2, col3 = st.columns(3)

with col1:
    st.caption(f"📅 {datetime.now().strftime('%d/%m/%Y %H:%M')}")

with col2:
    st.caption("🔄 Atualização automática a cada 5min")

with col3:
    st.caption(f"📊 {len(df)} registros | Aba: {SHEET_NAME}")

# =========================================================
# 7. CONFIGURAÇÃO DAS SECRETS (instruções)
# =========================================================
with st.sidebar.expander("⚙️ Configurar Secrets", expanded=False):
    st.markdown("""
    ### No Streamlit Cloud:
    
    1. Vá em **Settings**
    2. Clique em **Secrets**
    3. Cole:
    ```toml
    MS_CLIENT_ID = "seu-application-id"
    MS_CLIENT_SECRET = "seu-secret-value"
    MS_TENANT_ID = "seu-tenant-id"
    ```
    
    ### Como obter:
    - **MS_CLIENT_ID**: Application ID do Azure AD
    - **MS_CLIENT_SECRET**: VALUE do client secret
    - **MS_TENANT_ID**: Directory ID do Azure AD
    """)