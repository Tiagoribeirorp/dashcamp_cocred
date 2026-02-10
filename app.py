import streamlit as st
import pandas as pd
import openpyxl

st.set_page_config(page_title="Dashboard de Campanhas - SICOOB COCRED", layout="wide")

# =========================================================
# CARREGAMENTO E TRATAMENTO DOS DADOS
# =========================================================
@st.cache_data
def carregar_dados():
    df = pd.read_excel("jobs.xlsx", engine='openpyxl')

    # -------------------------
    # Tratamento do prazo
    # -------------------------
    if "Prazo em dias" in df.columns:
        df["Prazo em dias"] = (
            df["Prazo em dias"]
            .astype(str)
            .str.strip()
        )

        # Situação do prazo
        df["Situação do Prazo"] = df["Prazo em dias"].apply(
            lambda x: "Prazo encerrado"
            if "encerrado" in x.lower()
            else "Em prazo"
        )

        # Converter números
        df["Prazo em dias"] = pd.to_numeric(
            df["Prazo em dias"], errors="coerce"
        )

    # -------------------------
    # Faixa de prazo (checkbox)
    # -------------------------
    def classificar_faixa(row):
        if row["Situação do Prazo"] == "Prazo encerrado":
            return "Prazo encerrado"
        if pd.isna(row["Prazo em dias"]):
            return "Sem prazo"
        if row["Prazo em dias"] <= 0:
            return "Prazo encerrado"
        elif row["Prazo em dias"] <= 5:
            return "1 a 5 dias"
        elif row["Prazo em dias"] <= 10:
            return "6 a 10 dias"
        else:
            return "Acima de 10 dias"

    df["Faixa de Prazo"] = df.apply(classificar_faixa, axis=1)

    # -------------------------
    # Semáforo
    # -------------------------
    def classificar_semaforo(row):
        if row["Faixa de Prazo"] == "Prazo encerrado":
            return "Atrasado"
        elif row["Faixa de Prazo"] == "1 a 5 dias":
            return "Atenção"
        else:
            return "No prazo"

    df["Semáforo"] = df.apply(classificar_semaforo, axis=1)

    return df


df = carregar_dados()
df = df.dropna(subset=["Campanha ou Ação", "Status Operacional"])

# =========================================================
# TÍTULO
# =========================================================
st.title("📊 Dashboard de Campanhas – SICOOB COCRED")

# =========================================================
# LEGENDA
# =========================================================
with st.expander("📌 Legendas e critérios"):
    st.markdown("""
**Semáforo de Prazo**
- 🟢 **No prazo:** mais de 5 dias
- 🟡 **Atenção:** 1 a 5 dias
- 🔴 **Atrasado:** prazo encerrado ou vencido

**Faixas de Prazo**
- Prazo encerrado
- 1 a 5 dias
- 6 a 10 dias
- Acima de 10 dias
""")

# =========================================================
# FILTROS (SIDEBAR)
# =========================================================
st.sidebar.header("Filtros")
st.sidebar.caption("Os dados são atualizados automaticamente conforme o Excel.")

df_filtrado = df.copy()

# -------------------------
# Filtro por faixa de prazo
# -------------------------
st.sidebar.subheader("Prazo")
st.sidebar.caption("Filtra jobs por faixa de prazo.")

faixas_ordem = [
    "Prazo encerrado",
    "1 a 5 dias",
    "6 a 10 dias",
    "Acima de 10 dias"
]

faixas_disponiveis = df["Faixa de Prazo"].unique()
faixas_sel = []

for faixa in faixas_ordem:
    if faixa in faixas_disponiveis:
        marcado = st.sidebar.checkbox(
            faixa, value=True, key=f"faixa_{faixa}"
        )
        if marcado:
            faixas_sel.append(faixa)

df_filtrado = df_filtrado[df_filtrado["Faixa de Prazo"].isin(faixas_sel)]

# -------------------------
# Função genérica checkbox
# -------------------------
def filtro_checkbox(coluna, titulo, legenda):
    valores = sorted(df[coluna].dropna().unique())
    selecionados = []

    st.sidebar.subheader(titulo)
    st.sidebar.caption(legenda)

    for valor in valores:
        marcado = st.sidebar.checkbox(
            str(valor), value=True, key=f"{coluna}_{valor}"
        )
        if marcado:
            selecionados.append(valor)

    return selecionados


# Prioridade
if "Prioridade" in df.columns:
    prioridade_sel = filtro_checkbox(
        "Prioridade", "Prioridade", "Nível de urgência do job."
    )
    df_filtrado = df_filtrado[df_filtrado["Prioridade"].isin(prioridade_sel)]

# Produção
if "Produção" in df.columns:
    producao_sel = filtro_checkbox(
        "Produção", "Produção", "Tipo ou canal de produção."
    )
    df_filtrado = df_filtrado[df_filtrado["Produção"].isin(producao_sel)]

# Status
status_sel = filtro_checkbox(
    "Status Operacional", "Status", "Status atual do job."
)
df_filtrado = df_filtrado[df_filtrado["Status Operacional"].isin(status_sel)]

# =========================================================
# ALERTA DE ATRASO
# =========================================================
atrasados = df_filtrado[df_filtrado["Semáforo"] == "Atrasado"]
if len(atrasados) > 0:
    st.error(f"⚠️ {len(atrasados)} job(s) com prazo encerrado.")

# =========================================================
# RESUMO GERAL (CARDS)
# =========================================================
st.subheader("Resumo Geral")
st.caption("Total de jobs por status operacional.")

# Cores que funcionam em light/dark mode
cores_status = {
    "Aprovado": "#00A859",        # Verde SICOOB
    "Em Produção": "#007A3D",     # Verde escuro
    "Aguardando": "#7ED957",      # Verde claro
}

def cor_status(nome):
    nome = nome.lower()
    if "aprovado" in nome:
        return cores_status["Aprovado"]
    if "produção" in nome:
        return cores_status["Em Produção"]
    if "aguardando" in nome:
        return cores_status["Aguardando"]
    return "#6B7280"

resumo_geral = (
    df_filtrado["Status Operacional"]
    .value_counts()
    .reset_index()
)
resumo_geral.columns = ["Status", "Quantidade"]

cols = st.columns(len(resumo_geral))
for i, row in resumo_geral.iterrows():
    cor = cor_status(row["Status"])
    cols[i].markdown(
        f"""
        <div style="
            background:{cor};
            padding:20px;
            border-radius:12px;
            text-align:center;
            color:white;
            font-weight:bold;
        ">
            <div style="font-size:16px;">{row['Status']}</div>
            <div style="font-size:34px;">{int(row['Quantidade'])}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

st.divider()

# =========================================================
# RESUMO POR CAMPANHA (ESTILO EXCEL)
# =========================================================
st.subheader("Resumo por Campanha")
st.caption("Campanhas com atraso aparecem no topo.")

tabela_resumo = pd.pivot_table(
    df_filtrado,
    index="Campanha ou Ação",
    columns="Status Operacional",
    aggfunc="size",
    fill_value=0
)

tabela_resumo["Total"] = tabela_resumo.sum(axis=1)

campanhas_atrasadas = (
    df_filtrado[df_filtrado["Semáforo"] == "Atrasado"]
    ["Campanha ou Ação"]
    .unique()
)

tabela_resumo["Atrasada"] = tabela_resumo.index.isin(campanhas_atrasadas)
tabela_resumo = tabela_resumo.sort_values(
    by=["Atrasada", "Total"],
    ascending=[False, False]
).reset_index()

# CORREÇÃO: Cores que funcionam em light/dark mode
def destacar_campanha(row):
    if row["Atrasada"]:
        # Vermelho claro no light, vermelho escuro no dark
        return ["background-color: rgba(254, 202, 202, 0.3)"] * len(row)
    return [""] * len(row)

st.dataframe(
    tabela_resumo.style.apply(destacar_campanha, axis=1),
    use_container_width=True
)

st.divider()

# =========================================================
# TABELA DETALHADA - CORRIGIDA PARA LIGHT/DARK
# =========================================================
st.subheader("Detalhamento dos Jobs")
st.caption("Dados completos conforme filtros aplicados.")

# Função com cores adaptativas para light/dark mode
def destacar_semaforo(row):
    # Usar transparência para funcionar em ambos os modos
    if row["Semáforo"] == "Atrasado":
        # Vermelho com transparência
        return ["background-color: rgba(254, 202, 202, 0.3)"] * len(row)
    elif row["Semáforo"] == "Atenção":
        # Amarelo com transparência
        return ["background-color: rgba(255, 243, 205, 0.5)"] * len(row)
    elif row["Semáforo"] == "No prazo":
        # Verde com transparência
        return ["background-color: rgba(209, 231, 221, 0.4)"] * len(row)
    return [""] * len(row)

# Configurar o DataFrame com melhor formatação
styled_df = df_filtrado.style.apply(destacar_semaforo, axis=1)

# Adicionar formatação condicional para números
if "Prazo em dias" in df_filtrado.columns:
    styled_df = styled_df.format({
        "Prazo em dias": "{:.0f}",
    }, na_rep="N/A")

# Configurações para melhor visualização
st.dataframe(
    styled_df,
    use_container_width=True,
    height=600,  # Altura fixa com scroll
    column_config={
        "Prazo em dias": st.column_config.NumberColumn(
            "Prazo (dias)",
            help="Prazo em dias para conclusão",
            format="%d"
        ),
        "Prioridade": st.column_config.TextColumn(
            "Prioridade",
            help="Nível de prioridade"
        ),
        "Status Operacional": st.column_config.TextColumn(
            "Status",
            help="Status operacional atual"
        ),
        "Semáforo": st.column_config.TextColumn(
            "Situação",
            help="Situação do prazo: Atrasado, Atenção ou No prazo"
        )
    },
    hide_index=True  # Oculta o índice numérico
)

# =========================================================
# ESTILO CSS ADAPTATIVO PARA LIGHT/DARK MODE
# =========================================================
st.markdown("""
<style>
    /* Estilos para a tabela que funcionam em light/dark mode */
    .stDataFrame {
        border: 1px solid var(--border-color);
        border-radius: 8px;
    }
    
    /* Garantir contraste de texto */
    .stDataFrame [data-testid="stDataFrame"] {
        color: var(--text-color) !important;
    }
    
    /* Headers da tabela */
    .stDataFrame th {
        background-color: var(--background-color) !important;
        color: var(--text-color) !important;
        font-weight: bold !important;
    }
    
    /* Células da tabela */
    .stDataFrame td {
        color: var(--text-color) !important;
        border-color: var(--border-color) !important;
    }
    
    /* Linhas alternadas (zebra striping) */
    .stDataFrame tr:nth-child(even) {
        background-color: rgba(0, 0, 0, 0.02) !important;
    }
    
    /* Hover nas linhas */
    .stDataFrame tr:hover {
        background-color: rgba(0, 0, 0, 0.05) !important;
    }
</style>
""", unsafe_allow_html=True)

# =========================================================
# BOTÃO PARA ALTERNAR VISUALIZAÇÃO (OPCIONAL)
# =========================================================
with st.expander("⚙️ Configurações de visualização"):
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("🔄 Atualizar dados"):
            st.cache_data.clear()
            st.rerun()
    
    with col2:
        st.info("""
        **Dicas:**
        - Clique nos cabeçalhos para ordenar
        - Use Ctrl+F para buscar na tabela
        - Role para ver todas as colunas
        """)
    
    # Estatísticas rápidas
    st.caption(f"📊 Mostrando {len(df_filtrado)} de {len(df)} registros")