"""
Script de diagnóstico para API Microsoft Graph
Versão aprimorada para debug
"""

import os
import sys
from dotenv import load_dotenv
import requests
import msal
from datetime import datetime
import json

load_dotenv()

# Configurações
MS_CLIENT_ID = os.getenv("MS_CLIENT_ID", "")
MS_CLIENT_SECRET = os.getenv("MS_CLIENT_SECRET", "")
MS_TENANT_ID = os.getenv("MS_TENANT_ID", "")
SHAREPOINT_FILE_ID = "IQDMDcVdgAfGSIyZfeke7NFkAatm3fhI0-X4r6gIPQJmosY"

def get_token():
    """Obtém token de acesso"""
    authority = f"https://login.microsoftonline.com/{MS_TENANT_ID}"
    app = msal.ConfidentialClientApplication(
        MS_CLIENT_ID,
        authority=authority,
        client_credential=MS_CLIENT_SECRET
    )
    
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    return result.get("access_token")

def testar_me_endpoint(token):
    """Testa o endpoint /me para verificar permissões"""
    print("\n" + "=" * 60)
    print("👤 TESTANDO ENDPOINT /me (verifica permissões)")
    print("=" * 60)
    
    url = "https://graph.microsoft.com/v1.0/me"
    headers = {"Authorization": f"Bearer {token}"}
    
    try:
        response = requests.get(url, headers=headers, timeout=30)
        print(f"Código: {response.status_code}")
        
        if response.status_code == 200:
            data = response.json()
            print(f"✅ Usuário autenticado: {data.get('userPrincipalName', 'N/A')}")
            print(f"   ID: {data.get('id', 'N/A')}")
            print(f"   Nome: {data.get('displayName', 'N/A')}")
            return True
        elif response.status_code == 403:
            print("❌ Permissão insuficiente para acessar /me")
            print("   Adicione permissão 'User.Read' no Azure AD")
            return False
        else:
            print(f"❌ Erro: {response.text}")
            return False
    except Exception as e:
        print(f"❌ Exception: {str(e)}")
        return False

def testar_sites(token):
    """Lista sites disponíveis para verificar o site correto"""
    print("\n" + "=" * 60)
    print("🏢 BUSCANDO SITES DISPONÍVEIS")
    print("=" * 60)
    
    # Tentar diferentes endpoints de sites
    endpoints = [
        "https://graph.microsoft.com/v1.0/sites/root",
        "https://graph.microsoft.com/v1.0/sites?search=*",
        "https://graph.microsoft.com/v1.0/me/drive/root/children",
    ]
    
    headers = {"Authorization": f"Bearer {token}"}
    
    for url in endpoints:
        print(f"\n🔄 Tentando: {url}")
        try:
            response = requests.get(url, headers=headers, timeout=30)
            if response.status_code == 200:
                data = response.json()
                
                if "value" in data:
                    items = data["value"]
                    print(f"✅ Encontrados {len(items)} itens:")
                    for item in items[:5]:  # Mostra só os 5 primeiros
                        print(f"   - {item.get('name', 'Sem nome')}: {item.get('id', 'Sem ID')}")
                elif "id" in data:
                    print(f"✅ Site root: {data.get('webUrl', 'N/A')}")
                    print(f"   ID: {data.get('id', 'N/A')}")
                else:
                    print(f"📦 Resposta: {json.dumps(data, indent=2)[:500]}...")
            else:
                print(f"❌ Código {response.status_code}: {response.text[:200]}")
        except Exception as e:
            print(f"⚠️  Erro: {str(e)}")

def testar_drive_items(token):
    """Testa acessar itens do drive"""
    print("\n" + "=" * 60)
    print("📁 TESTANDO ACESSO AO DRIVE")
    print("=" * 60)
    
    # Primeiro, descobrir qual é o drive correto
    urls = [
        # Tenta acessar diretamente pelo ID do item
        f"https://graph.microsoft.com/v1.0/drives/root/items/{SHAREPOINT_FILE_ID}",
        # Tenta no meu OneDrive
        f"https://graph.microsoft.com/v1.0/me/drive/items/{SHAREPOINT_FILE_ID}",
        # Lista itens do root do drive
        "https://graph.microsoft.com/v1.0/me/drive/root/children",
    ]
    
    headers = {"Authorization": f"Bearer {token}"}
    
    for url in urls:
        print(f"\n🔄 Testando URL: {url}")
        try:
            response = requests.get(url, headers=headers, timeout=30)
            print(f"📊 Status: {response.status_code}")
            
            if response.status_code == 200:
                data = response.json()
                if "value" in data:  # É uma lista
                    items = data["value"]
                    print(f"✅ Encontrados {len(items)} arquivos:")
                    for item in items[:3]:
                        print(f"   📄 {item.get('name')} (ID: {item.get('id')})")
                else:  # É um item individual
                    print(f"✅ Arquivo encontrado:")
                    print(f"   Nome: {data.get('name')}")
                    print(f"   ID: {data.get('id')}")
                    print(f"   URL: {data.get('webUrl', 'N/A')}")
                    print(f"   Drive ID: {data.get('parentReference', {}).get('driveId', 'N/A')}")
                    return data
            elif response.status_code == 404:
                print("❌ Item não encontrado (404)")
            elif response.status_code == 400:
                print("❌ Requisição inválida (400)")
                print(f"   Detalhes: {response.text[:200]}")
            else:
                print(f"❌ Erro: {response.text[:200]}")
                
        except Exception as e:
            print(f"⚠️  Exception: {str(e)}")
    
    return None

def descobrir_estrutura_sharepoint(token):
    """Tenta descobrir a estrutura correta do SharePoint"""
    print("\n" + "=" * 60)
    print("🔍 DESCUBRINDO ESTRUTURA DO SHAREPOINT")
    print("=" * 60)
    
    # O File ID parece ser muito longo para um ID padrão do Graph
    # Vamos tentar usar o caminho completo do arquivo
    
    # Seu arquivo está em: /personal/cristini_cordesco_ideatoreamericas_com
    # O formato correto pode ser:
    
    site_path = "personal.sharepoint.com"
    user_email = "cristini_cordesco@ideatoreamericas.com"
    file_path = f"/personal/cristini_cordesco_ideatoreamericas_com/Documents/dashboard_cocred.xlsx"
    
    print("📝 Informações do seu link:")
    print(f"   Site: agenciaideatore.sharepoint.com")
    print(f"   Caminho: {SHAREPOINT_SITE_PATH}")
    print(f"   File ID: {SHAREPOINT_FILE_ID}")
    
    print("\n🔄 Tentando formatos alternativos...")
    
    # Formato 1: Acessar pelo site e caminho
    site_hostname = "agenciaideatore.sharepoint.com"
    site_path_encoded = "/personal/cristini_cordesco_ideatoreamericas_com"
    
    # Primeiro precisamos obter o site ID
    site_url = f"https://graph.microsoft.com/v1.0/sites/{site_hostname}:{site_path_encoded}"
    
    headers = {"Authorization": f"Bearer {token}"}
    
    print(f"\n1. Buscando site: {site_url}")
    try:
        response = requests.get(site_url, headers=headers, timeout=30)
        if response.status_code == 200:
            site_data = response.json()
            print(f"✅ Site encontrado!")
            print(f"   Site ID: {site_data.get('id')}")
            print(f"   Nome: {site_data.get('name')}")
            print(f"   Web URL: {site_data.get('webUrl')}")
            
            # Agora buscar drives deste site
            drives_url = f"{site_url}/drives"
            print(f"\n2. Buscando drives: {drives_url}")
            
            drives_response = requests.get(drives_url, headers=headers, timeout=30)
            if drives_response.status_code == 200:
                drives_data = drives_response.json()
                drives = drives_data.get('value', [])
                print(f"✅ Encontrados {len(drives)} drives:")
                for drive in drives:
                    print(f"   🚗 {drive.get('name')} (ID: {drive.get('id')})")
                    
                    # Listar alguns itens de cada drive
                    items_url = f"https://graph.microsoft.com/v1.0/drives/{drive.get('id')}/root/children"
                    items_response = requests.get(items_url, headers=headers, timeout=30)
                    if items_response.status_code == 200:
                        items_data = items_response.json()
                        items = items_data.get('value', [])
                        print(f"     📁 {len(items)} itens")
                        for item in items[:2]:  # Mostra 2 itens por drive
                            print(f"       📄 {item.get('name')}")
                    
            return site_data
        else:
            print(f"❌ Site não encontrado: {response.status_code}")
            print(f"   {response.text[:200]}")
    except Exception as e:
        print(f"⚠️  Erro ao buscar site: {str(e)}")
    
    return None

def main():
    """Função principal"""
    print("\n" + "=" * 60)
    print("🔧 DIAGNÓSTICO DETALHADO - API MICROSOFT GRAPH")
    print("=" * 60)
    
    # 1. Obter token
    print("\n1. 🎫 OBTENDO TOKEN...")
    token = get_token()
    if not token:
        print("❌ Falha ao obter token")
        return
    
    print("✅ Token obtido")
    
    # 2. Testar permissões básicas
    testar_me_endpoint(token)
    
    # 3. Tentar acessar arquivo de diferentes formas
    testar_drive_items(token)
    
    # 4. Descobrir estrutura do SharePoint
    descobrir_estrutura_sharepoint(token)
    
    # 5. Sugestões finais
    print("\n" + "=" * 60)
    print("💡 SUGESTÕES PARA CORRIGIR")
    print("=" * 60)
    
    print("""
1. **Verifique o File ID correto:**
   - O File ID parece muito longo
   - IDs do Graph API geralmente são mais curtos
   - Abra o arquivo no navegador e inspecione a URL

2. **Use o URL completo do arquivo:**
   - Seu link: https://agenciaideatore-my.sharepoint.com/:x:/g/personal/cristini_cordesco_ideatoreamericas_com/IQDMDcVdgAfGSIyZfeke7NFkAatm3fhI0-X4r6gIPQJmosY
   - Tente usar o caminho completo em vez do ID

3. **Verifique as permissões:**
   - App Registration → API permissions
   - Adicione: Files.Read.All, Sites.Read.All
   - Clique em "Grant admin consent"

4. **Teste no Graph Explorer:**
   - Acesse: https://developer.microsoft.com/graph/graph-explorer
   - Faça login com sua conta
   - Tente acessar: /me/drive/root/children
   - Veja qual é o ID correto do seu arquivo
    """)

if __name__ == "__main__":
    main()