# testar_file_id_correto.py
import os
import requests
import msal
from dotenv import load_dotenv
import pandas as pd
from io import BytesIO

load_dotenv()

# Configurações
MS_CLIENT_ID = os.getenv("MS_CLIENT_ID")
MS_CLIENT_SECRET = os.getenv("MS_CLIENT_SECRET")
MS_TENANT_ID = os.getenv("MS_TENANT_ID")

# Dados corretos encontrados
USUARIO_PRINCIPAL = "cristini.cordesco@ideatoreamericas.com"
FILE_ID = "01S7YQRRWMBXCV3AAHYZEIZGL55EPOZULE"
SHEET_NAME = "Demandas ID"

def get_token():
    """Obtém token usando client credentials"""
    authority = f"https://login.microsoftonline.com/{MS_TENANT_ID}"
    app = msal.ConfidentialClientApplication(
        MS_CLIENT_ID,
        authority=authority,
        client_credential=MS_CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    return result.get("access_token")

def testar_acesso_completo():
    """Testa acesso completo ao arquivo Excel"""
    print("=" * 80)
    print("🧪 TESTE COMPLETO DO FILE ID ENCONTRADO")
    print("=" * 80)
    
    print(f"👤 Usuário: {USUARIO_PRINCIPAL}")
    print(f"🆔 File ID: {FILE_ID}")
    print(f"📊 Aba: {SHEET_NAME}")
    
    # 1. Obter token
    print("\n🎫 1. Obtendo token...")
    token = get_token()
    if not token:
        print("❌ Falha ao obter token")
        return False
    
    print(f"✅ Token obtido")
    
    # 2. Testar acesso ao arquivo
    print("\n📂 2. Testando acesso ao arquivo...")
    url_metadata = f"https://graph.microsoft.com/v1.0/users/{USUARIO_PRINCIPAL}/drive/items/{FILE_ID}"
    
    headers = {"Authorization": f"Bearer {token}"}
    
    try:
        response = requests.get(url_metadata, headers=headers, timeout=30)
        
        if response.status_code == 200:
            data = response.json()
            print(f"✅ Arquivo encontrado!")
            print(f"   📄 Nome: {data.get('name')}")
            print(f"   🆔 ID: {data.get('id')}")
            print(f"   📊 Tamanho: {int(data.get('size', 0)) / 1024:.1f} KB")
            print(f"   🔗 URL: {data.get('webUrl', 'N/A')}")
            print(f"   📅 Modificado: {data.get('lastModifiedDateTime', 'N/A')}")
            
            # 3. Testar download do conteúdo
            print("\n⬇️  3. Testando download do conteúdo...")
            url_content = f"{url_metadata}/content"
            
            content_response = requests.get(url_content, headers=headers, timeout=30)
            
            if content_response.status_code == 200:
                print(f"✅ Conteúdo baixado com sucesso!")
                print(f"   Content-Length: {len(content_response.content)} bytes")
                print(f"   Content-Type: {content_response.headers.get('Content-Type', 'N/A')}")
                
                # 4. Testar leitura do Excel
                print("\n📊 4. Testando leitura do Excel...")
                
                try:
                    excel_file = BytesIO(content_response.content)
                    
                    # Ler a aba específica
                    try:
                        df = pd.read_excel(excel_file, sheet_name=SHEET_NAME, engine='openpyxl')
                        print(f"✅ Aba '{SHEET_NAME}' lida com sucesso!")
                        print(f"   📈 {len(df)} linhas")
                        print(f"   📋 {len(df.columns)} colunas")
                        
                        # Mostrar primeiras colunas
                        print(f"\n   🏷️  Colunas encontradas:")
                        for col in df.columns[:10]:  # Mostra até 10 colunas
                            print(f"      - {col}")
                        if len(df.columns) > 10:
                            print(f"      ... e mais {len(df.columns) - 10} colunas")
                        
                        # Mostrar amostra dos dados
                        print(f"\n   👁️  Amostra dos dados (primeiras 3 linhas):")
                        print(df.head(3).to_string(max_cols=5, max_rows=3))
                        
                        return True
                        
                    except Exception as e_aba:
                        print(f"⚠️  Não encontrei aba '{SHEET_NAME}': {str(e_aba)}")
                        
                        # Tentar ler primeira aba
                        print("🔄 Tentando primeira aba...")
                        excel_file.seek(0)
                        df = pd.read_excel(excel_file, engine='openpyxl')
                        
                        print(f"✅ Primeira aba lida com sucesso!")
                        print(f"   📈 {len(df)} linhas")
                        print(f"   📋 {len(df.columns)} colunas")
                        
                        # Mostrar abas disponíveis
                        excel_file.seek(0)
                        xl = pd.ExcelFile(excel_file)
                        print(f"\n   📑 Abas disponíveis no arquivo:")
                        for sheet in xl.sheet_names:
                            print(f"      - {sheet}")
                        
                        return True
                        
                except Exception as e_excel:
                    print(f"❌ Erro ao ler Excel: {str(e_excel)}")
                    return False
                    
            else:
                print(f"❌ Erro no download: {content_response.status_code}")
                return False
                
        else:
            print(f"❌ Erro ao acessar arquivo: {response.status_code}")
            print(f"   Resposta: {response.text[:200]}")
            return False
            
    except Exception as e:
        print(f"❌ Exception: {str(e)}")
        return False

def gerar_codigo_app():
    """Gera o código atualizado para o app.py"""
    print("\n" + "=" * 80)
    print("📝 CÓDIGO ATUALIZADO PARA SEU APP.PY")
    print("=" * 80)
    
    codigo = f'''
# =========================================================
# CONFIGURAÇÕES DA API (ATUALIZE ESTAS LINHAS!)
# =========================================================

# 1. SUAS CREDENCIAIS DA GRAPH API
MS_CLIENT_ID = st.secrets.get("MS_CLIENT_ID", "")
MS_CLIENT_SECRET = st.secrets.get("MS_CLIENT_SECRET", "")
MS_TENANT_ID = st.secrets.get("MS_TENANT_ID", "")

# 2. INFORMAÇÕES DO EXCEL ONLINE (CONFIGURAÇÃO CORRETA!)
USUARIO_PRINCIPAL = "cristini.cordesco@ideatoreamericas.com"  # ← USUÁRIO COM PONTO!
SHAREPOINT_FILE_ID = "01S7YQRRWMBXCV3AAHYZEIZGL55EPOZULE"  # ← NOVO FILE ID CORRETO
SHEET_NAME = "Demandas ID"  # ← NOME DA ABA

# =========================================================
# FUNÇÃO ATUALIZADA - SUBSTITUA NO SEU APP.PY
# =========================================================
@st.cache_data(ttl=300)
def carregar_dados_excel_online():
    """Carrega dados da aba 'Demandas ID' do Excel Online"""
    
    access_token = get_access_token()
    if not access_token:
        return pd.DataFrame()
    
    # URL CORRETA para acessar o arquivo
    file_url = f"https://graph.microsoft.com/v1.0/users/{{USUARIO_PRINCIPAL}}/drive/items/{{SHAREPOINT_FILE_ID}}/content"
    
    headers = {{
        "Authorization": f"Bearer {{access_token}}",
        "Accept": "application/octet-stream"
    }}
    
    try:
        with st.spinner("🔄 Conectando ao Excel Online..."):
            response = requests.get(file_url, headers=headers, timeout=30)
        
        if response.status_code == 200:
            excel_file = BytesIO(response.content)
            
            # Tentar ler a aba específica
            try:
                df = pd.read_excel(excel_file, sheet_name=SHEET_NAME, engine='openpyxl')
            except Exception as e:
                st.warning(f"⚠️ Não encontrei aba '{{SHEET_NAME}}'. Tentando primeira aba...")
                df = pd.read_excel(excel_file, engine='openpyxl')
            
            if df.empty:
                st.error(f"❌ A aba '{{SHEET_NAME}}' está vazia ou não encontrada.")
                return pd.DataFrame()
            
            # Informações de sucesso
            st.sidebar.success(f"✅ Conectado ao Excel Online")
            st.sidebar.caption(f"📄 Arquivo: {{df.shape[0]}} linhas × {{df.shape[1]}} colunas")
            
            return df
            
        elif response.status_code == 404:
            st.error("❌ Arquivo não encontrado")
            st.info(f"Verifique: 1) File ID, 2) Usuário '{{USUARIO_PRINCIPAL}}'")
            
        elif response.status_code == 403:
            st.error("❌ Permissão negada")
            st.info("Verifique as permissões 'Files.Read.All' no Azure AD")
            
        elif response.status_code == 401:
            st.error("❌ Token expirado")
            st.cache_data.clear()
            
        else:
            st.error(f"❌ Erro HTTP {{response.status_code}}")
        
        return pd.DataFrame()
        
    except Exception as e:
        st.error(f"❌ Erro inesperado: {{str(e)}}")
        return pd.DataFrame()
'''
    
    print(codigo)
    
    print("\n" + "=" * 80)
    print("🔄 INSTRUÇÕES PARA ATUALIZAR:")
    print("=" * 80)
    print("""
1. Abra seu arquivo app.py
2. Localize as configurações no início (linhas ~20-30)
3. Substitua por:
   - USUARIO_PRINCIPAL = "cristini.cordesco@ideatoreamericas.com"
   - SHAREPOINT_FILE_ID = "01S7YQRRWMBXCV3AAHYZEIZGL55EPOZULE"
4. Localize a função carregar_dados_excel_online()
5. Substitua pela função acima
6. Salve e execute: streamlit run app.py
    """)

def main():
    """Função principal"""
    print("🚀 CONFIGURAÇÃO FINAL - DASHCAMP COCRED")
    print("=" * 80)
    
    # Testar o acesso
    sucesso = testar_acesso_completo()
    
    if sucesso:
        print("\n" + "=" * 80)
        print("🎉 🎉 🎉 TUDO FUNCIONANDO PERFEITAMENTE! 🎉 🎉 🎉")
        print("=" * 80)
        print("\n✅ Conexão com Microsoft Graph: OK")
        print("✅ Acesso ao arquivo Excel: OK")
        print("✅ Leitura da aba/planilha: OK")
        print("✅ Download do conteúdo: OK")
        
        # Gerar código para atualização
        gerar_codigo_app()
        
        # Teste final
        print("\n" + "=" * 80)
        print("🧪 TESTE FINAL RÁPIDO")
        print("=" * 80)
        print("Execute este comando para testar o app completo:")
        print("\nstreamlit run app.py")
        print("\nO dashboard deve carregar automaticamente os dados!")
        
    else:
        print("\n" + "=" * 80)
        print("❌ AINDA COM PROBLEMAS")
        print("=" * 80)
        print("\nVerifique:")
        print("1. Credenciais no arquivo .env estão corretas")
        print("2. Permissões no Azure AD: Files.Read.All")
        print("3. Admin Consent foi dado")
        print("4. O arquivo ainda existe no OneDrive")

if __name__ == "__main__":
    main()