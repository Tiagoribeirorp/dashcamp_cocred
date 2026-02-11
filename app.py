import streamlit as st
import pandas as pd
import requests
from io import BytesIO
import msal
from datetime import datetime, timedelta
import pytz
import time

# =========================================================
# CONFIGURAÇÕES INICIAIS
# =========================================================
# Configurar pandas para mostrar TUDO
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', None)

st.set_page_config(
    page_title="Dashboard de Campanhas - SICOOB COCRED", 
    layout="wide",
    page_icon="📊"
)

# =========================================================
# CONFIGURAÇÕES DA API
# =========================================================

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
# 3. FUNÇÕES AUXILIARES
# =========================================================
def calcular_altura_tabela(num_linhas, num_colunas):
    """Calcula altura ideal para a tabela"""
    altura_base = 150  # pixels para cabeçalhos e controles
    altura_por_linha = 35  # pixels por linha
    altura_por_coluna = 2  # pixels extras por coluna
    
    # Altura baseada no conteúdo
    altura_conteudo = altura_base + (num_linhas * altura_por_linha) + (num_colunas * altura_por_coluna)
    
    # Limitar a um máximo razoável para performance
    altura_maxima = 2000  # 2000px = ~53 linhas visíveis de uma vez
    
    return min(altura_conteudo, altura_maxima)

def converter_para_data(df, coluna):
    """Converte coluna para datetime se possível"""
    try:
        df[coluna] = pd.to_datetime(df[coluna], errors='coerce', dayfirst=True)
    except:
        pass
    return df

# =========================================================
# 4. INTERFACE PRINCIPAL
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

# Configurações de visualização
st.sidebar.header("👁️ Visualização")
linhas_por_pagina = st.sidebar.selectbox(
    "Linhas por página:", 
    ["50", "100", "200", "500", "Todas"],
    index=1
)

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
# 5. CARREGAR E MOSTRAR DADOS
# =========================================================

# Carregar dados
with st.spinner("📥 Carregando dados do Excel..."):
    df = carregar_dados_excel_online()

# Verificar se tem dados
if df.empty:
    st.error("❌ Nenhum dado carregado")
    st.stop()

# Converter coluna de data de solicitação se existir
if 'Data de Solicitação' in df.columns:
    df = converter_para_data(df, 'Data de Solicitação')
    # Remover timezone se houver
    if pd.api.types.is_datetime64_any_dtype(df['Data de Solicitação']):
        df['Data de Solicitação'] = df['Data de Solicitação'].dt.tz_localize(None)

# Mostrar contador REAL
total_linhas = len(df)
total_colunas = len(df.columns)

st.success(f"✅ **{total_linhas} registros** carregados com sucesso!")
st.info(f"📋 **Colunas:** {', '.join(df.columns.tolist()[:5])}{'...' if len(df.columns) > 5 else ''}")

# =========================================================
# 6. VISUALIZAÇÃO COMPLETA DOS DADOS (COM PAGINAÇÃO)
# =========================================================

st.header("📋 Dados Completos")

# Opções de visualização
tab1, tab2, tab3 = st.tabs(["📊 Dados Completos", "📈 Estatísticas", "🔍 Pesquisa"])

with tab1:
    if linhas_por_pagina == "Todas":
        # Mostrar TODAS as linhas de uma vez
        altura_tabela = calcular_altura_tabela(total_linhas, total_colunas)
        
        st.subheader(f"📋 Todos os {total_linhas} registros")
        
        # Mostrar dataframe completo
        st.dataframe(
            df,
            height=altura_tabela,
            use_container_width=True,
            hide_index=False,
            column_config=None
        )
        
        if altura_tabela >= 2000:
            linhas_visiveis = int((2000 - 150) / 35)
            st.info(f"ℹ️ Mostrando {linhas_visiveis} de {total_linhas} linhas por vez. Use o scroll para navegar.")
        
    else:
        # Paginação manual
        linhas_por_pagina = int(linhas_por_pagina)
        total_paginas = (total_linhas - 1) // linhas_por_pagina + 1
        
        # Inicializar página na session_state
        if 'pagina_atual' not in st.session_state:
            st.session_state.pagina_atual = 1
        
        # Controles de navegação
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
            # Seletor de página direto
            nova_pagina = st.number_input(
                "Ir para página:", 
                min_value=1, 
                max_value=total_paginas, 
                value=st.session_state.pagina_atual,
                key="pagina_input"
            )
            if nova_pagina != st.session_state.pagina_atual:
                st.session_state.pagina_atual = nova_pagina
                st.rerun()
        
        # Calcular índices
        inicio = (st.session_state.pagina_atual - 1) * linhas_por_pagina
        fim = min(inicio + linhas_por_pagina, total_linhas)
        
        st.write(f"**Mostrando linhas {inicio + 1} a {fim} de {total_linhas}**")
        
        # Mostrar dataframe paginado
        altura_pagina = calcular_altura_tabela(linhas_por_pagina, total_colunas)
        
        st.dataframe(
            df.iloc[inicio:fim],
            height=altura_pagina,
            use_container_width=True,
            hide_index=False
        )
    
    # Contadores
    col_count1, col_count2, col_count3 = st.columns(3)
    with col_count1:
        st.metric("📈 Total de Linhas", total_linhas)
    with col_count2:
        st.metric("📊 Total de Colunas", total_colunas)
    with col_count3:
        if 'Data de Solicitação' in df.columns:
            ultima_data = df['Data de Solicitação'].max()
            if pd.notna(ultima_data) and hasattr(ultima_data, 'strftime'):
                st.metric("📅 Última Solicitação", ultima_data.strftime('%d/%m/%Y'))
            else:
                st.metric("📅 Última Solicitação", "N/A")
        else:
            st.metric("📅 Última Atualização", datetime.now().strftime('%d/%m/%Y'))

with tab2:
    # Estatísticas
    st.subheader("📈 Estatísticas dos Dados")
    
    col_stat1, col_stat2 = st.columns(2)
    
    with col_stat1:
        st.write("**Resumo Numérico:**")
        # Filtrar apenas colunas numéricas
        colunas_numericas = df.select_dtypes(include=['number']).columns
        if len(colunas_numericas) > 0:
            st.dataframe(df[colunas_numericas].describe(), use_container_width=True, height=300)
        else:
            st.info("ℹ️ Não há colunas numéricas para análise estatística.")
    
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
    texto_pesquisa = st.text_input(
        "🔎 Pesquisar em todas as colunas:", 
        placeholder="Digite um termo para buscar...",
        key="pesquisa_principal"
    )
    
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
        
        if len(resultados) > 0:
            st.success(f"✅ **{len(resultados)} resultado(s) encontrado(s):**")
            
            # Altura dinâmica para resultados
            altura_resultados = calcular_altura_tabela(len(resultados), len(resultados.columns))
            
            st.dataframe(
                resultados, 
                use_container_width=True, 
                height=min(altura_resultados, 800)
            )
            
            # Botão para exportar resultados
            if st.button("📥 Exportar Resultados", key="export_resultados"):
                csv = resultados.to_csv(index=False, encoding='utf-8-sig')
                st.download_button(
                    label="📥 Download CSV dos Resultados",
                    data=csv,
                    file_name=f"pesquisa_{texto_pesquisa}_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                    mime="text/csv"
                )
        else:
            st.warning(f"⚠️ Nenhum resultado encontrado para '{texto_pesquisa}'")
    else:
        st.info("👆 Digite um termo acima para pesquisar nos dados")

# =========================================================
# 7. FILTROS AVANÇADOS (COM FILTRO DE DATA)
# =========================================================

st.header("🎛️ Filtros Avançados")

# Criar layout de 4 colunas para acomodar o filtro de data
filtro_cols = st.columns(4)

filtros_ativos = {}

# Filtro 1: Status
if 'Status' in df.columns:
    with filtro_cols[0]:
        status_opcoes = ['Todos'] + sorted(df['Status'].dropna().unique().tolist())
        status_selecionado = st.selectbox("📌 Status:", status_opcoes, key="filtro_status")
        if status_selecionado != 'Todos':
            filtros_ativos['Status'] = status_selecionado

# Filtro 2: Prioridade
if 'Prioridade' in df.columns:
    with filtro_cols[1]:
        prioridade_opcoes = ['Todos'] + sorted(df['Prioridade'].dropna().unique().tolist())
        prioridade_selecionada = st.selectbox("⚡ Prioridade:", prioridade_opcoes, key="filtro_prioridade")
        if prioridade_selecionada != 'Todos':
            filtros_ativos['Prioridade'] = prioridade_selecionada

# Filtro 3: Produção
if 'Produção' in df.columns:
    with filtro_cols[2]:
        producao_opcoes = ['Todos'] + sorted(df['Produção'].dropna().unique().tolist())
        producao_selecionada = st.selectbox("🏭 Produção:", producao_opcoes, key="filtro_producao")
        if producao_selecionada != 'Todos':
            filtros_ativos['Produção'] = producao_selecionada

# ========== FILTRO DE DATA DE SOLICITAÇÃO ==========
with filtro_cols[3]:
    st.markdown("**📅 Data Solicitação**")
    
    # Verificar se existe coluna de data
    if 'Data de Solicitação' in df.columns:
        # Garantir que é datetime
        if not pd.api.types.is_datetime64_any_dtype(df['Data de Solicitação']):
            df['Data de Solicitação'] = pd.to_datetime(df['Data de Solicitação'], errors='coerce')
        
        # Remover datas nulas
        datas_validas = df['Data de Solicitação'].dropna()
        
        if not datas_validas.empty:
            data_min = datas_validas.min().date()
            data_max = datas_validas.max().date()
            
            # Opções de período rápido
            periodo_opcao = st.selectbox(
                "Período:",
                ["Todos", "Hoje", "Esta semana", "Este mês", "Últimos 30 dias", "Personalizado"],
                key="periodo_data"
            )
            
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

# Aplicar filtros categóricos
for col, valor in filtros_ativos.items():
    if col not in ['data_inicio', 'data_fim', 'tem_filtro_data']:
        df_filtrado = df_filtrado[df_filtrado[col] == valor]

# Aplicar filtro de data
if 'tem_filtro_data' in filtros_ativos and 'Data de Solicitação' in df.columns:
    data_inicio = pd.Timestamp(filtros_ativos['data_inicio'])
    data_fim = pd.Timestamp(filtros_ativos['data_fim']) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
    
    df_filtrado = df_filtrado[
        (df_filtrado['Data de Solicitação'] >= data_inicio) & 
        (df_filtrado['Data de Solicitação'] <= data_fim)
    ]

# Mostrar resultados dos filtros
if filtros_ativos:
    st.subheader(f"📊 Dados Filtrados ({len(df_filtrado)} de {total_linhas} registros)")
    
    if len(df_filtrado) > 0:
        altura_filtrada = calcular_altura_tabela(len(df_filtrado), len(df_filtrado.columns))
        
        st.dataframe(
            df_filtrado, 
            use_container_width=True, 
            height=min(altura_filtrada, 800)
        )
        
        # Estatísticas dos filtros
        col_filt1, col_filt2, col_filt3 = st.columns(3)
        
        with col_filt1:
            st.metric("📈 Registros Filtrados", len(df_filtrado))
        
        with col_filt2:
            porcentagem = (len(df_filtrado) / total_linhas * 100) if total_linhas > 0 else 0
            st.metric("📊 % do Total", f"{porcentagem:.1f}%")
        
        with col_filt3:
            if 'tem_filtro_data' in filtros_ativos:
                st.metric("📅 Período", 
                         f"{filtros_ativos['data_inicio'].strftime('%d/%m')} a {filtros_ativos['data_fim'].strftime('%d/%m')}")
        
        # Botão para limpar filtros
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
# 8. EXPORTAÇÃO (COM DADOS FILTRADOS)
# =========================================================

st.header("💾 Exportar Dados")

# Dados para exportação (usar filtrados se existirem)
df_exportar = df_filtrado if filtros_ativos and len(df_filtrado) > 0 else df

col_exp1, col_exp2, col_exp3 = st.columns(3)

with col_exp1:
    # CSV
    csv = df_exportar.to_csv(index=False, encoding='utf-8-sig')
    st.download_button(
        label="📥 Download CSV",
        data=csv,
        file_name=f"dados_cocred_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
        mime="text/csv",
        use_container_width=True,
        help="Baixar dados em formato CSV"
    )

with col_exp2:
    # Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_exportar.to_excel(writer, index=False, sheet_name='Dados')
        # Adicionar aba de resumo
        resumo = pd.DataFrame({
            'Métrica': ['Total Registros', 'Total Colunas', 'Data Exportação', 'Filtros Aplicados'],
            'Valor': [len(df_exportar), len(df_exportar.columns), 
                     datetime.now().strftime('%d/%m/%Y %H:%M'),
                     'Sim' if filtros_ativos else 'Não']
        })
        resumo.to_excel(writer, index=False, sheet_name='Resumo')
    
    excel_data = output.getvalue()
    
    st.download_button(
        label="📥 Download Excel",
        data=excel_data,
        file_name=f"dados_cocred_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        help="Baixar dados em formato Excel com abas"
    )

with col_exp3:
    # JSON
    json_data = df_exportar.to_json(orient='records', force_ascii=False, date_format='iso')
    st.download_button(
        label="📥 Download JSON",
        data=json_data,
        file_name=f"dados_cocred_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
        mime="application/json",
        use_container_width=True,
        help="Baixar dados em formato JSON"
    )

# =========================================================
# 9. DEBUG INFO (apenas se ativado)
# =========================================================

if st.session_state.debug_mode:
    st.sidebar.markdown("---")
    st.sidebar.markdown("**🐛 Debug Info:**")
    
    with st.sidebar.expander("Detalhes Técnicos", expanded=False):
        st.write(f"**Cache:** 1 minuto")
        st.write(f"**Hora atual:** {datetime.now().strftime('%H:%M:%S')}")
        
        token = get_access_token()
        if token:
            st.success(f"✅ Token: ...{token[-10:]}")
        else:
            st.error("❌ Token não disponível")
        
        st.write(f"**DataFrame Info:**")
        st.write(f"- Shape: {df.shape}")
        st.write(f"- Memory: {df.memory_usage(deep=True).sum() / 1024 / 1024:.2f} MB")
        st.write(f"- Colunas: {list(df.columns)}")
        
        if 'Data de Solicitação' in df.columns:
            st.write(f"**Data de Solicitação:**")
            st.write(f"- Tipo: {df['Data de Solicitação'].dtype}")
            st.write(f"- Mínimo: {df['Data de Solicitação'].min()}")
            st.write(f"- Máximo: {df['Data de Solicitação'].max()}")
            st.write(f"- Nulos: {df['Data de Solicitação'].isnull().sum()}")
        
        # Mostrar primeiras e últimas linhas
        st.write("**Amostra dos Dados:**")
        
        tab_debug1, tab_debug2 = st.tabs(["Primeiras 5", "Últimas 5"])
        
        with tab_debug1:
            st.dataframe(df.head(5), use_container_width=True)
        
        with tab_debug2:
            st.dataframe(df.tail(5), use_container_width=True)

# =========================================================
# 10. RODAPÉ
# =========================================================

st.divider()

footer_col1, footer_col2, footer_col3 = st.columns(3)

with footer_col1:
    st.caption(f"🕐 {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")

with footer_col2:
    st.caption(f"📊 {total_linhas} registros | {total_colunas} colunas")
    if filtros_ativos and len(df_filtrado) > 0:
        st.caption(f"🎯 Filtrados: {len(df_filtrado)} registros")

with footer_col3:
    st.caption("🔄 Atualiza a cada 1 minuto | 📧 cristini.cordesco@ideatoreamericas.com")

# =========================================================
# 11. AUTO-REFRESH (opcional)
# =========================================================

# Auto-refresh a cada 60 segundos (opcional)
auto_refresh = st.sidebar.checkbox("🔄 Auto-refresh (60s)", value=False, key="auto_refresh")

if auto_refresh:
    refresh_placeholder = st.empty()
    for i in range(60, 0, -1):
        refresh_placeholder.caption(f"🔄 Atualizando em {i} segundos...")
        time.sleep(1)
    refresh_placeholder.empty()
    st.rerun()