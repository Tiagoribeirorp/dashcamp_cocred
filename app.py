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
# FILTRO POR INTERVALO DE DIAS (SLIDER) - NOVO!
# -------------------------
st.sidebar.subheader("Intervalo de Prazo (dias)")
st.sidebar.caption("Selecione o intervalo mínimo e máximo de dias.")

if "Prazo em dias" in df.columns:
    # Encontrar valores mínimo e máximo (ignorando NaN e negativos)
    prazos_validos = df["Prazo em dias"].dropna()
    prazos_validos = prazos_validos[prazos_validos >= 0]
    
    if not prazos_validos.empty:
        min_val = int(prazos_validos.min())
        max_val = int(prazos_validos.max())
        
        # Garantir que max seja maior que min
        if min_val == max_val:
            max_val = min_val + 1
        
        # Slider de intervalo
        intervalo = st.sidebar.slider(
            "Selecione o intervalo:",
            min_value=min_val,
            max_value=max_val,
            value=(min_val, max_val),
            help=f"Filtrar jobs com prazo entre X e Y dias. Disponível: {min_val} a {max_val} dias"
        )
        
        # Aplicar filtro de intervalo
        min_dias, max_dias = intervalo
        
        # Filtrar por intervalo (inclui os limites)
        mask_intervalo = (df_filtrado["Prazo em dias"] >= min_dias) & (df_filtrado["Prazo em dias"] <= max_dias)
        
        # Também incluir os "Prazo encerrado" se o usuário quiser ver
        incluir_atrasados = st.sidebar.checkbox(
            "Incluir prazos encerrados/vencidos", 
            value=True,
            help="Mostrar também jobs com prazo já vencido"
        )
        
        if incluir_atrasados:
            # Incluir prazos encerrados
            mask_atrasados = (df_filtrado["Faixa de Prazo"] == "Prazo encerrado")
            df_filtrado = df_filtrado[mask_intervalo | mask_atrasados]
        else:
            # Apenas o intervalo selecionado
            df_filtrado = df_filtrado[mask_intervalo]
        
        # Mostrar estatísticas
        st.sidebar.info(f"**Intervalo selecionado:** {min_dias} a {max_dias} dias")
        
    else:
        st.sidebar.warning("Não há prazos válidos para filtrar")
else:
    st.sidebar.warning("Coluna 'Prazo em dias' não encontrada")

# -------------------------
# FILTRO POR FAIXA (OPCIONAL - mantém como checkbox)
# -------------------------
st.sidebar.subheader("Filtro por Situação")
st.sidebar.caption("Filtre por situação específica do prazo.")

# Criar checkboxes para as situações
situacoes_disponiveis = ["No prazo", "Atenção", "Atrasado"]
situacoes_selecionadas = []

for situacao in situacoes_disponiveis:
    if situacao in df_filtrado["Semáforo"].unique():
        marcado = st.sidebar.checkbox(
            situacao, 
            value=True, 
            key=f"situacao_{situacao}",
            help=f"Mostrar jobs com situação: {situacao}"
        )
        if marcado:
            situacoes_selecionadas.append(situacao)

# Aplicar filtro de situação
if situacoes_selecionadas:
    df_filtrado = df_filtrado[df_filtrado["Semáforo"].isin(situacoes_selecionadas)]
else:
    st.sidebar.warning("Selecione pelo menos uma situação")

# -------------------------
# Função genérica checkbox (MANTIDA para outros filtros)
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

def destacar_campanha(row):
    if row["Atrasada"]:
        return ["background-color:#FECACA"] * len(row)
    return [""] * len(row)

st.dataframe(
    tabela_resumo.style.apply(destacar_campanha, axis=1),
    use_container_width=True
)

st.divider()

# =========================================================
# TABELA DETALHADA
# =========================================================
st.subheader("Detalhamento dos Jobs")
st.caption("Dados completos conforme filtros aplicados.")

def destacar_semaforo(row):
    if row["Semáforo"] == "Atrasado":
        return ["background-color:#FECACA"] * len(row)
    elif row["Semáforo"] == "Atenção":
        return ["background-color:#DCFCE7"] * len(row)
    return [""] * len(row)

st.dataframe(
    df_filtrado.style.apply(destacar_semaforo, axis=1),
    use_container_width=True
)

# =========================================================
# RESUMO DO FILTRO APLICADO
# =========================================================
with st.sidebar.expander("📊 Resumo do Filtro", expanded=False):
    if "Prazo em dias" in df.columns:
        st.write(f"**Jobs no intervalo:** {len(df_filtrado)}")
        if not df_filtrado.empty:
            st.write(f"**Prazo médio:** {df_filtrado['Prazo em dias'].mean():.1f} dias")
            st.write(f"**Prazo mínimo:** {df_filtrado['Prazo em dias'].min():.0f} dias")
            st.write(f"**Prazo máximo:** {df_filtrado['Prazo em dias'].max():.0f} dias")
    
    st.write(f"**Total filtrado:** {len(df_filtrado)} de {len(df)}")