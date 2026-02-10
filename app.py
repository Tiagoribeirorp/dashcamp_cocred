import streamlit as st
import pandas as pd

st.set_page_config(page_title="Dashboard de Campanhas", layout="wide")

# ---------------------------
# Carregar dados
# ---------------------------
@st.cache_data
def carregar_dados():
    df = pd.read_excel("jobs.xlsx")

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
            df["Prazo em dias"],
            errors="coerce"
        )

    # Semáforo
    def classificar_semaforo(row):
        if row["Situação do Prazo"] == "Prazo encerrado":
            return "Atrasado"
        if pd.isna(row["Prazo em dias"]):
            return "No prazo"
        if row["Prazo em dias"] <= 0:
            return "Atrasado"
        elif row["Prazo em dias"] <= 5:
            return "Atenção"
        else:
            return "No prazo"

    df["Semáforo"] = df.apply(classificar_semaforo, axis=1)

    return df


df = carregar_dados()
df = df.dropna(subset=["Campanha ou Ação", "Status Operacional"])

st.title("📊 Dashboard de Campanhas - SICOOB COCRED")

# ---------------------------
# LEGENDA DO SEMÁFORO
# ---------------------------
with st.expander("Legenda do Semáforo de Prazo"):
    st.markdown("""
    **Classificação automática dos prazos:**

    - 🟢 **No prazo:** mais de 5 dias restantes
    - 🟡 **Atenção:** entre 1 e 5 dias restantes
    - 🔴 **Atrasado:** prazo encerrado ou vencido
    """)

# ---------------------------
# FILTROS NA SIDEBAR
# ---------------------------
st.sidebar.header("Filtros")

st.sidebar.markdown("""
**Como usar os filtros:**
- Desmarque as opções que não deseja visualizar.
- Os dados do dashboard serão atualizados automaticamente.
""")

df_filtrado = df.copy()

# Situação do prazo
if "Situação do Prazo" in df.columns:
    situacoes = sorted(df["Situação do Prazo"].unique())
    situacao_sel = []

    st.sidebar.subheader("Situação do Prazo")
    st.sidebar.caption("Filtra jobs por prazo ativo ou encerrado.")

    for s in situacoes:
        marcado = st.sidebar.checkbox(s, value=True, key=f"prazo_{s}")
        if marcado:
            situacao_sel.append(s)

    df_filtrado = df_filtrado[
        df_filtrado["Situação do Prazo"].isin(situacao_sel)
    ]

# Prazo numérico
df_prazo = df_filtrado.dropna(subset=["Prazo em dias"])
if not df_prazo.empty:
    prazo_min = int(df_prazo["Prazo em dias"].min())
    prazo_max = int(df_prazo["Prazo em dias"].max())

    st.sidebar.subheader("Prazo em dias")
    st.sidebar.caption("Filtra jobs pelo número de dias restantes.")

    prazo_sel = st.sidebar.slider(
        "Intervalo de prazo",
        prazo_min,
        prazo_max,
        (prazo_min, prazo_max)
    )

    df_filtrado = df_filtrado[
        (df_filtrado["Prazo em dias"].isna()) |
        (
            (df_filtrado["Prazo em dias"] >= prazo_sel[0]) &
            (df_filtrado["Prazo em dias"] <= prazo_sel[1])
        )
    ]

# Função de filtro checkbox
def filtro_checkbox(coluna, titulo, legenda):
    valores = sorted(df[coluna].dropna().unique())
    selecionados = []

    st.sidebar.subheader(titulo)
    st.sidebar.caption(legenda)

    for valor in valores:
        marcado = st.sidebar.checkbox(
            str(valor),
            value=True,
            key=f"{coluna}_{valor}"
        )
        if marcado:
            selecionados.append(valor)

    return selecionados


# Prioridade
if "Prioridade" in df.columns:
    prioridade_sel = filtro_checkbox(
        "Prioridade",
        "Prioridade",
        "Filtra jobs por nível de urgência."
    )
    df_filtrado = df_filtrado[
        df_filtrado["Prioridade"].isin(prioridade_sel)
    ]

# Produção
if "Produção" in df.columns:
    producao_sel = filtro_checkbox(
        "Produção",
        "Produção",
        "Filtra por tipo de produção ou canal."
    )
    df_filtrado = df_filtrado[
        df_filtrado["Produção"].isin(producao_sel)
    ]

# Status
status_sel = filtro_checkbox(
    "Status Operacional",
    "Status",
    "Filtra pelo status atual do job."
)
df_filtrado = df_filtrado[
    df_filtrado["Status Operacional"].isin(status_sel)
]

# ---------------------------
# ALERTA DE ATRASO
# ---------------------------
atrasados = df_filtrado[df_filtrado["Semáforo"] == "Atrasado"]

if len(atrasados) > 0:
    st.error(f"⚠️ {len(atrasados)} job(s) em atraso.")

# ---------------------------
# RESUMO GERAL (CARDS COLORIDOS)
# ---------------------------
st.subheader("Resumo Geral")
st.caption("Quantidade total de jobs por status operacional.")

cores_status = {
    "Aprovado": "#00A859",
    "Em Produção": "#007A3D",
    "Aguardando": "#7ED957",
    "Reprovado": "#4B5563",
}

def cor_status(nome):
    nome = nome.lower()
    if "aprovado" in nome:
        return cores_status["Aprovado"]
    if "produção" in nome:
        return cores_status["Em Produção"]
    if "aguardando" in nome:
        return cores_status["Aguardando"]
    return "#4B5563"

resumo_geral = (
    df_filtrado["Status Operacional"]
    .value_counts()
    .reset_index()
)
resumo_geral.columns = ["Status", "Quantidade"]

if not resumo_geral.empty:
    cols = st.columns(len(resumo_geral))
    for idx, row in resumo_geral.iterrows():
        cor = cor_status(row["Status"])
        cols[idx].markdown(
            f"""
            <div style="
                background:{cor};
                padding:20px;
                border-radius:12px;
                text-align:center;
                color:white;
                font-weight:bold;
            ">
                <div style="font-size:18px;">
                    {row['Status']}
                </div>
                <div style="font-size:36px;">
                    {int(row['Quantidade'])}
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

st.divider()

# ---------------------------
# TABELA RESUMO POR CAMPANHA
# ---------------------------
st.subheader("Resumo por Campanha")
st.caption("Visão consolidada dos status dentro de cada campanha.")

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
)

tabela_resumo = tabela_resumo.reset_index()

def destacar_campanhas(row):
    if row["Atrasada"]:
        return ["background-color: #fecaca"] * len(row)
    return [""] * len(row)

st.dataframe(
    tabela_resumo.style.apply(destacar_campanhas, axis=1),
    use_container_width=True
)

st.divider()

# ---------------------------
# TABELA DETALHADA
# ---------------------------
st.subheader("Detalhamento dos Jobs")
st.caption("Lista completa dos jobs conforme filtros aplicados.")

def destacar_semaforo(row):
    if row["Semáforo"] == "Atrasado":
        return ["background-color: #fecaca"] * len(row)
    elif row["Semáforo"] == "Atenção":
        return ["background-color: #dcfce7"] * len(row)
    else:
        return [""] * len(row)

st.dataframe(
    df_filtrado.style.apply(destacar_semaforo, axis=1),
    use_container_width=True
)
