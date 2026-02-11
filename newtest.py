# diagnostico_excel.py
import os
import requests
import msal
import pandas as pd
from io import BytesIO
from dotenv import load_dotenv
from datetime import datetime
import pytz

load_dotenv()

# Configurações
MS_CLIENT_ID = os.getenv("MS_CLIENT_ID")
MS_CLIENT_SECRET = os.getenv("MS_CLIENT_SECRET")
MS_TENANT_ID = os.getenv("MS_TENANT_ID")
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

def diagnostico_completo():
    """Diagnóstico completo do problema"""
    print("=" * 80)
    print("🔍 DIAGNÓSTICO - DADOS NÃO APARECENDO NO DASH")
    print("=" * 80)
    
    # 1. Verificar token
    print("\n1️⃣  VERIFICANDO TOKEN...")
    token = get_token()
    if not token:
        print("❌ Falha ao obter token")
        return
    
    print(f"✅ Token obtido: ...{token[-10:]}")
    
    # 2. Verificar acesso ao arquivo
    print("\n2️⃣  VERIFICANDO ACESSO AO ARQUIVO...")
    url_metadata = f"https://graph.microsoft.com/v1.0/users/{USUARIO_PRINCIPAL}/drive/items/{FILE_ID}"
    headers = {"Authorization": f"Bearer {token}"}
    
    try:
        response = requests.get(url_metadata, headers=headers, timeout=30)
        
        if response.status_code == 200:
            metadata = response.json()
            print(f"✅ Arquivo encontrado!")
            print(f"   📄 Nome: {metadata.get('name')}")
            print(f"   📊 Tamanho: {int(metadata.get('size', 0)) / 1024:.1f} KB")
            print(f"   📅 Última modificação: {metadata.get('lastModifiedDateTime')}")
            
            # Verificar se é realmente um arquivo Excel
            mime_type = metadata.get('file', {}).get('mimeType', '')
            if 'spreadsheet' in mime_type.lower() or 'excel' in mime_type.lower():
                print(f"   ✅ É um arquivo Excel: {mime_type}")
            else:
                print(f"   ⚠️  Tipo de arquivo inesperado: {mime_type}")
                
        else:
            print(f"❌ Erro {response.status_code}: {response.text[:200]}")
            return
            
    except Exception as e:
        print(f"❌ Exception: {str(e)}")
        return
    
    # 3. Baixar e analisar o conteúdo
    print("\n3️⃣  ANALISANDO CONTEÚDO DO ARQUIVO...")
    url_content = f"{url_metadata}/content"
    
    try:
        response = requests.get(url_content, headers=headers, timeout=30)
        
        if response.status_code == 200:
            print(f"✅ Conteúdo baixado: {len(response.content)} bytes")
            
            # Salvar para análise
            with open('temp_downloaded_file.xlsx', 'wb') as f:
                f.write(response.content)
            print(f"   💾 Salvo como 'temp_downloaded_file.xlsx' para análise")
            
            # Ler o arquivo
            excel_file = BytesIO(response.content)
            
            # 3.1 Verificar todas as abas
            print("\n   📑 LISTANDO TODAS AS ABAS...")
            try:
                xl = pd.ExcelFile(excel_file, engine='openpyxl')
                sheet_names = xl.sheet_names
                print(f"   ✅ {len(sheet_names)} aba(s) encontrada(s):")
                for i, sheet in enumerate(sheet_names, 1):
                    print(f"      {i}. {sheet}")
                    
                # Verificar se a aba "Demandas ID" existe
                if SHEET_NAME in sheet_names:
                    print(f"\n   ✅ Aba '{SHEET_NAME}' ENCONTRADA!")
                else:
                    print(f"\n   ❌ Aba '{SHEET_NAME}' NÃO encontrada!")
                    print(f"   Abas disponíveis: {sheet_names}")
                    
            except Exception as e:
                print(f"   ❌ Erro ao ler abas: {str(e)}")
            
            # 3.2 Ler a aba específica
            print(f"\n   📖 LENDO ABA '{SHEET_NAME}'...")
            excel_file.seek(0)  # Resetar ponteiro
            
            try:
                df = pd.read_excel(excel_file, sheet_name=SHEET_NAME, engine='openpyxl')
                print(f"   ✅ Aba '{SHEET_NAME}' lida com sucesso!")
                print(f"   📊 Formato: {df.shape[0]} linhas × {df.shape[1]} colunas")
                
                # Mostrar informações detalhadas
                print(f"\n   🔍 INFORMAÇÕES DETALHADAS:")
                print(f"      - Memória usada: {df.memory_usage(deep=True).sum() / 1024 / 1024:.2f} MB")
                print(f"      - Colunas: {list(df.columns)}")
                print(f"      - Tipos de dados:")
                for col, dtype in df.dtypes.items():
                    print(f"        • {col}: {dtype}")
                
                # Mostrar primeiras e últimas linhas
                print(f"\n   📋 PRIMEIRAS 5 LINHAS:")
                print(df.head().to_string())
                
                print(f"\n   📋 ÚLTIMAS 5 LINHAS:")
                print(df.tail().to_string())
                
                # Verificar dados recentes
                print(f"\n   ⏰ VERIFICANDO DADOS RECENTES...")
                
                # Procurar por colunas de data
                date_columns = []
                for col in df.columns:
                    try:
                        # Tentar converter para datetime
                        sample = df[col].dropna().head(5)
                        if len(sample) > 0:
                            pd.to_datetime(sample, errors='raise')
                            date_columns.append(col)
                    except:
                        pass
                
                if date_columns:
                    print(f"      Colunas de data encontradas: {date_columns}")
                    for date_col in date_columns[:2]:  # Verificar até 2 colunas
                        try:
                            df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
                            latest = df[date_col].max()
                            if pd.notna(latest):
                                print(f"      Última data em '{date_col}': {latest}")
                        except:
                            pass
                
                # 3.3 Comparar com o que você espera
                print(f"\n   🎯 COMPARAÇÃO COM EXPECTATIVAS:")
                
                # O que você inseriou recentemente? (você precisa me dizer)
                print("""
                **Perguntas para diagnóstico:**
                1. Quantas linhas você ESPERA ver? ______
                2. Quantas colunas você ESPERA ver? ______
                3. Qual a última linha que você adicionou? ______
                4. Há alguma coluna específica com novos dados? ______
                """)
                
                # 3.4 Verificar cache
                print(f"\n   🗂️  VERIFICANDO CACHE:")
                print("""
                O app usa cache de 5 minutos. Possíveis problemas:
                
                1. **Cache antigo**: Aguarde 5 minutos ou clique em "Atualizar agora"
                2. **Cache do Streamlit**: Ctrl+C no terminal e execute novamente
                3. **Cache do navegador**: Ctrl+F5 para forçar atualização
                """)
                
            except Exception as e:
                print(f"   ❌ Erro ao ler aba '{SHEET_NAME}': {str(e)}")
                
                # Tentar ler primeira aba
                print(f"\n   🔄 TENTANDO PRIMEIRA ABA...")
                excel_file.seek(0)
                try:
                    df = pd.read_excel(excel_file, sheet_name=0, engine='openpyxl')
                    print(f"   ✅ Primeira aba lida: {df.shape[0]}×{df.shape[1]}")
                    print(f"   Nome da aba: {xl.sheet_names[0] if 'xl' in locals() else 'Desconhecido'}")
                    print(f"\n   Primeiras linhas:")
                    print(df.head().to_string())
                except Exception as e2:
                    print(f"   ❌ Erro ao ler primeira aba: {str(e2)}")
                    
        else:
            print(f"❌ Erro ao baixar conteúdo: {response.status_code}")
            
    except Exception as e:
        print(f"❌ Exception ao baixar: {str(e)}")
    
    # 4. Verificar permissões e configurações
    print("\n4️⃣  VERIFICANDO CONFIGURAÇÕES...")
    print(f"""
    Configuração atual:
    - Usuário: {USUARIO_PRINCIPAL}
    - File ID: {FILE_ID}
    - Aba: {SHEET_NAME}
    
    **Possíveis problemas:**
    
    1. 📍 **Aba errada**: 
       - Verifique o nome EXATO da aba no Excel
       - É "{SHEET_NAME}"? Ou tem espaço diferente?
    
    2. ⏰ **Cache ativo**:
       - O app tem cache de 5 minutos
       - Clique em "Atualizar agora" no sidebar
       - Ou aguarde 5 minutos
    
    3. 🔄 **Arquivo não salvo**:
       - Você salvou o Excel depois de adicionar dados? (Ctrl+S)
       - Verifique data da última modificação acima
    
    4. 📂 **Arquivo diferente**:
       - Talvez o File ID não seja do arquivo correto
       - Verifique se está editando o mesmo arquivo
    
    5. 👁️ **Filtros ativos**:
       - O dashboard tem filtros que podem estar ocultando dados
       - Verifique se há filtros aplicados
    """)

def testar_app_local():
    """Testa o app localmente para ver se funciona"""
    print("\n" + "=" * 80)
    print("🧪 TESTANDO APP LOCALMENTE")
    print("=" * 80)
    
    # Simular o que o app faz
    token = get_token()
    if not token:
        return
    
    url = f"https://graph.microsoft.com/v1.0/users/{USUARIO_PRINCIPAL}/drive/items/{FILE_ID}/content"
    headers = {"Authorization": f"Bearer {token}"}
    
    try:
        response = requests.get(url, headers=headers, timeout=30)
        
        if response.status_code == 200:
            excel_file = BytesIO(response.content)
            
            # Ler a aba
            try:
                df = pd.read_excel(excel_file, sheet_name=SHEET_NAME, engine='openpyxl')
                print(f"✅ App funcionando localmente!")
                print(f"   Linhas: {len(df)}")
                print(f"   Colunas: {len(df.columns)}")
                print(f"   Última atualização no app: AGORA")
                
                # Mostrar diferença com arquivo salvo
                if os.path.exists('temp_downloaded_file.xlsx'):
                    df_salvo = pd.read_excel('temp_downloaded_file.xlsx', sheet_name=SHEET_NAME, engine='openpyxl')
                    if len(df) != len(df_salvo):
                        print(f"⚠️  Diferença: app={len(df)} vs salvo={len(df_salvo)} linhas")
                    else:
                        print(f"✅ Mesmo número de linhas: {len(df)}")
                        
            except Exception as e:
                print(f"❌ Erro no app: {str(e)}")
                
        else:
            print(f"❌ Erro no app: {response.status_code}")
            
    except Exception as e:
        print(f"❌ Exception no app: {str(e)}")

def main():
    """Função principal"""
    print("🚀 DIAGNÓSTICO - DADOS NÃO VISÍVEIS NO DASHBOARD")
    print("=" * 80)
    
    # Verificar credenciais
    if not all([MS_CLIENT_ID, MS_CLIENT_SECRET, MS_TENANT_ID]):
        print("❌ Credenciais não configuradas no .env")
        return
    
    # Executar diagnóstico
    diagnostico_completo()
    
    # Testar app local
    testar_app_local()
    
    # Instruções
    print("\n" + "=" * 80)
    print("🎯 SOLUÇÕES PARA TESTAR:")
    print("=" * 80)
    print("""
    1. **FORÇAR ATUALIZAÇÃO IMEDIATA:**
       - No sidebar do app, clique em "🔄 Atualizar agora"
       - Isso limpa o cache e recarrega os dados
    
    2. **VERIFICAR ABA CORRETA:**
       - Abra o Excel Online
       - Confirme o nome EXATO da aba
       - Pode ser "Demandas ID", "Demandas_ID", "Demandas-ID", etc.
    
    3. **VERIFICAR SALVAMENTO:**
       - No Excel, pressione Ctrl+S
       - Espere alguns segundos
       - Atualize o dashboard
    
    4. **TESTE DIRETO NO TERMINAL:**
       python diagnostico_excel.py
       (Este script mostra o que está sendo baixado)
    
    5. **VERIFICAR FILTROS:**
       - No dashboard, verifique se há filtros aplicados
       - Remova todos os filtros para ver todos os dados
    
    6. **MODIFICAR CACHE (Streamlit Cloud):**
       - Settings → Advanced → Clear cache
       - Ou edite o app para mudar @st.cache_data(ttl=60) ← 1 minuto
    """)
    
    print("\n⚠️  **Responda estas perguntas para ajudar:**")
    print("""
    1. Você salvou o Excel depois de adicionar os dados? (S/N)
    2. Quantos minutos se passaram desde que salvou?
    3. Quantas linhas você ESPERA ver no total?
    4. As linhas antigas aparecem? Só as novas não?
    5. Você clicou em "Atualizar agora" no sidebar?
    """)

if __name__ == "__main__":
    main()