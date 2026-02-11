# listar_arquivos_drive.py
import os
import requests
import msal
from dotenv import load_dotenv

load_dotenv()

# Configurações
MS_CLIENT_ID = os.getenv("MS_CLIENT_ID")
MS_CLIENT_SECRET = os.getenv("MS_CLIENT_SECRET")
MS_TENANT_ID = os.getenv("MS_TENANT_ID")

# Usuário correto (com PONTO!)
USUARIO = "cristini.cordesco@ideatoreamericas.com"

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

def listar_todos_arquivos(token):
    """Lista todos os arquivos do drive do usuário"""
    print("=" * 80)
    print(f"📁 LISTANDO ARQUIVOS DE: {USUARIO}")
    print("=" * 80)
    
    # URL para listar arquivos da raiz
    url = f"https://graph.microsoft.com/v1.0/users/{USUARIO}/drive/root/children"
    
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json"
    }
    
    try:
        response = requests.get(url, headers=headers, timeout=30)
        
        if response.status_code == 200:
            data = response.json()
            items = data.get('value', [])
            
            if not items:
                print("📭 Pasta vazia")
                return []
            
            print(f"✅ Encontrados {len(items)} itens na raiz:\n")
            
            arquivos_excel = []
            
            for i, item in enumerate(items, 1):
                nome = item.get('name', 'Sem nome')
                item_id = item.get('id')
                tamanho = int(item.get('size', 0)) / 1024
                tipo = "📁 PASTA" if 'folder' in item else "📄 ARQUIVO"
                
                print(f"{i:3d}. {tipo} {nome}")
                print(f"     🆔 ID: {item_id}")
                print(f"     📊 Tamanho: {tamanho:.1f} KB")
                print(f"     📅 Modificado: {item.get('lastModifiedDateTime', 'N/A')}")
                
                # Verificar se é Excel
                if nome.lower().endswith(('.xlsx', '.xls', '.xlsm')):
                    print(f"     📊 ✅ É um arquivo Excel!")
                    arquivos_excel.append((nome, item_id))
                
                print()
            
            return arquivos_excel
            
        elif response.status_code == 404:
            print("❌ Drive ou pasta não encontrada")
        elif response.status_code == 403:
            print("❌ Permissão negada")
        else:
            print(f"❌ Erro {response.status_code}: {response.text[:200]}")
            
    except Exception as e:
        print(f"❌ Exception: {str(e)}")
    
    return []

def buscar_arquivo_por_nome(token, nome_arquivo):
    """Busca um arquivo específico pelo nome"""
    print("\n" + "=" * 80)
    print(f"🔍 BUSCANDO ARQUIVO: {nome_arquivo}")
    print("=" * 80)
    
    # URL de busca
    url = f"https://graph.microsoft.com/v1.0/users/{USUARIO}/drive/root/search(q='{nome_arquivo}')"
    
    headers = {"Authorization": f"Bearer {token}"}
    
    try:
        response = requests.get(url, headers=headers, timeout=30)
        
        if response.status_code == 200:
            data = response.json()
            items = data.get('value', [])
            
            if items:
                print(f"✅ Encontrado(s) {len(items)} resultado(s):\n")
                
                for i, item in enumerate(items, 1):
                    print(f"{i}. 📄 {item.get('name')}")
                    print(f"   🆔 ID: {item.get('id')}")
                    print(f"   📍 Caminho: {item.get('parentReference', {}).get('path', 'N/A')}")
                    print(f"   🔗 URL: {item.get('webUrl', 'N/A')}")
                    print(f"   📊 Tamanho: {int(item.get('size', 0)) / 1024:.1f} KB")
                    print()
            else:
                print(f"❌ Nenhum resultado para '{nome_arquivo}'")
                
        else:
            print(f"❌ Erro na busca: {response.status_code}")
            
    except Exception as e:
        print(f"❌ Exception: {str(e)}")

def testar_acesso_arquivo(token, file_id):
    """Testa o acesso a um arquivo específico"""
    print("\n" + "=" * 80)
    print(f"🧪 TESTANDO ACESSO AO ARQUIVO")
    print("=" * 80)
    
    # URL para acessar o arquivo
    url = f"https://graph.microsoft.com/v1.0/users/{USUARIO}/drive/items/{file_id}"
    
    headers = {"Authorization": f"Bearer {token}"}
    
    try:
        response = requests.get(url, headers=headers, timeout=30)
        
        if response.status_code == 200:
            data = response.json()
            print(f"✅ Arquivo acessível!")
            print(f"   📄 Nome: {data.get('name')}")
            print(f"   🆔 ID: {data.get('id')}")
            print(f"   📊 Tamanho: {int(data.get('size', 0)) / 1024:.1f} KB")
            print(f"   🔗 URL: {data.get('webUrl', 'N/A')}")
            print(f"   📅 Modificado: {data.get('lastModifiedDateTime', 'N/A')}")
            
            # Testar download do conteúdo
            content_url = f"{url}/content"
            print(f"\n🔄 Testando download do conteúdo...")
            
            content_response = requests.get(content_url, headers=headers, timeout=30, stream=True)
            
            if content_response.status_code == 200:
                print(f"✅ Conteúdo acessível para download!")
                print(f"   Content-Type: {content_response.headers.get('Content-Type', 'N/A')}")
                print(f"   Content-Length: {int(content_response.headers.get('Content-Length', 0)) / 1024:.1f} KB")
                return True
            else:
                print(f"❌ Erro no download: {content_response.status_code}")
                return False
                
        else:
            print(f"❌ Erro ao acessar arquivo: {response.status_code}")
            return False
            
    except Exception as e:
        print(f"❌ Exception: {str(e)}")
        return False

def main():
    """Função principal"""
    print("🚀 ENCONTRAR ARQUIVO EXCEL NO DRIVE")
    print("=" * 80)
    
    # 1. Obter token
    print("\n🎫 Obtendo token...")
    token = get_token()
    if not token:
        print("❌ Falha ao obter token")
        return
    
    print(f"✅ Token obtido")
    
    # 2. Listar todos os arquivos
    arquivos_excel = listar_todos_arquivos(token)
    
    # 3. Se encontrou Excel, testar acesso
    if arquivos_excel:
        print("\n" + "=" * 80)
        print("📊 ARQUIVOS EXCEL ENCONTRADOS:")
        print("=" * 80)
        
        for nome, file_id in arquivos_excel:
            print(f"\n🧪 Testando: {nome}")
            sucesso = testar_acesso_arquivo(token, file_id)
            
            if sucesso:
                print(f"\n🎯 ARQUIVO CORRETO PROVÁVEL!")
                print(f"   Use este File ID no app.py: {file_id}")
                break
    
    # 4. Buscar por nome específico
    print("\n" + "=" * 80)
    print("🔎 BUSCA POR NOMES ESPECÍFICOS")
    print("=" * 80)
    
    nomes_possiveis = [
        "dashboard_cocred.xlsx",
        "cocred.xlsx",
        "campanhas.xlsx",
        "demandas.xlsx",
        "sicoob.xlsx",
        "dashcamp.xlsx",
    ]
    
    for nome in nomes_possiveis:
        buscar_arquivo_por_nome(token, nome)
    
    # 5. Instruções finais
    print("\n" + "=" * 80)
    print("📝 CONFIGURAÇÃO FINAL DO APP.PY")
    print("=" * 80)
    
    print(f"""
1. NO SEU app.py, ATUALIZE:

# Linha ~26 (configurações)
USUARIO_PRINCIPAL = "cristini.cordesco@ideatoreamericas.com"  # ← COM PONTO!
SHAREPOINT_FILE_ID = "COLE_O_FILE_ID_AQUI"  # ← ID do arquivo Excel
SHEET_NAME = "Demandas ID"

2. NA FUNÇÃO carregar_dados_excel_online(), use:

file_url = f"https://graph.microsoft.com/v1.0/users/{{USUARIO_PRINCIPAL}}/drive/items/{{SHAREPOINT_FILE_ID}}/content"

3. VERIFIQUE as permissões no Azure AD:
   - Files.Read.All ✅
   - User.Read ✅
   - Admin Consent dado ✅

4. Execute o app:
   streamlit run app.py
    """)

if __name__ == "__main__":
    main()