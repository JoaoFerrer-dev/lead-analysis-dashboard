import os
import pandas as pd
import streamlit as st

# =====================
# CONFIGURACAO
# =====================
PASTA = r"C:\Users\joaov\Documents\DESAFIO ESTAGIO"
ARQUIVO = os.path.join(PASTA, "01_Analise_Dados.xlsx")
ABAS_PREFERIDAS = ["BASE_ANALISE", "BASE"]

st.set_page_config(layout="wide")

# =====================
# FUNCOES
# =====================
@st.cache_data
def carregar_base(path: str) -> tuple[pd.DataFrame, str]:
    xls = pd.ExcelFile(path)
    aba = None
    for s in ABAS_PREFERIDAS:
        if s in xls.sheet_names:
            aba = s
            break
    if aba is None:
        raise ValueError(
            f"Nenhuma das abas {ABAS_PREFERIDAS} foi encontrada. Abas disponiveis: {xls.sheet_names}"
        )

    df = pd.read_excel(path, sheet_name=aba)
    df.columns = [str(c).strip() for c in df.columns]

    # Normaliza possiveis nomes de coluna de sub-origem
    if "SUB-ORIGEM" in df.columns and "SUB_ORIGEM" not in df.columns:
        df = df.rename(columns={"SUB-ORIGEM": "SUB_ORIGEM"})
    if "SUB ORIGEM" in df.columns and "SUB_ORIGEM" not in df.columns:
        df = df.rename(columns={"SUB ORIGEM": "SUB_ORIGEM"})

    return df, aba


def preparar_df(df: pd.DataFrame) -> pd.DataFrame:
    obrigatorias = ["LEAD_ID", "VENDIDO"]
    faltando = [c for c in obrigatorias if c not in df.columns]
    if faltando:
        raise ValueError(f"Faltam colunas obrigatorias na base: {faltando}")

    df = df.copy()

    # Flag de conversao (cliente)
    df["cliente"] = df["VENDIDO"].apply(lambda x: 1 if str(x).strip().upper() == "SIM" else 0)

    # Padroniza dimensoes e trata 0/vazio como "Nao informado"
    for col in ["ORIGEM", "SUB_ORIGEM", "MERCADO", "PORTE", "OBJETIVO", "LOCAL"]:
        if col in df.columns:
            df[col] = df[col].fillna("Nao informado").astype(str).str.strip()
            df[col] = df[col].replace(
                {"0": "Nao informado", "": "Nao informado", "nan": "Nao informado", "None": "Nao informado"}
            )

    return df


def tabela_conversao(df: pd.DataFrame, dim: str) -> pd.DataFrame:
    """Retorna Leads, Clientes e Conversao (%) por dimensao."""
    if dim not in df.columns:
        return pd.DataFrame({"Aviso": [f"Coluna '{dim}' nao existe na base."]})

    t = df.groupby(dim).agg(
        Leads=("LEAD_ID", "count"),
        Clientes=("cliente", "sum"),
    )
    t["Conversao (%)"] = (t["Clientes"] / t["Leads"] * 100).round(2)
    t = t.sort_values(["Conversao (%)", "Leads"], ascending=[False, False])
    return t


def aplicar_filtros(df: pd.DataFrame, filtros: dict) -> pd.DataFrame:
    df_f = df.copy()
    for col, valores in filtros.items():
        if valores and col in df_f.columns:
            df_f = df_f[df_f[col].isin(valores)]
    return df_f


def multiselect_sidebar(df: pd.DataFrame, col: str, label: str):
    if col not in df.columns:
        return []
    opcoes = sorted(df[col].dropna().astype(str).unique().tolist())
    return st.sidebar.multiselect(label, opcoes, default=[])


# =====================
# CARREGAR DADOS
# =====================
st.title("Desafio Nology — Conversao de Leads")

try:
    df_raw, aba_usada = carregar_base(ARQUIVO)
    df = preparar_df(df_raw)
except Exception as e:
    st.error(f"Erro ao carregar a base: {e}")
    st.stop()

st.caption(f"Arquivo: {ARQUIVO} | Aba utilizada: {aba_usada}")

# =====================
# SIDEBAR (FILTROS)
# =====================
st.sidebar.header("Filtros (opcional)")
st.sidebar.write("Filtre para explorar segmentos especificos.")

f_origem = multiselect_sidebar(df, "ORIGEM", "ORIGEM")
f_mercado = multiselect_sidebar(df, "MERCADO", "MERCADO")
f_porte = multiselect_sidebar(df, "PORTE", "PORTE")
f_objetivo = multiselect_sidebar(df, "OBJETIVO", "OBJETIVO")

# slider de volume minimo (como no video)
vol_min = st.sidebar.slider("Volume minimo (evitar amostra pequena)", 5, 500, 30, 5)

# Aplica filtros
filtros = {
    "ORIGEM": f_origem,
    "MERCADO": f_mercado,
    "PORTE": f_porte,
    "OBJETIVO": f_objetivo,
}
df_f = aplicar_filtros(df, filtros)

# =====================
# KPIs (REAGEM AO FILTRO)
# =====================
total_leads = int(len(df_f))
total_clientes = int(df_f["cliente"].sum()) if total_leads else 0
taxa = (total_clientes / total_leads * 100) if total_leads else 0.0

k1, k2, k3 = st.columns(3)
k1.metric("Total de Leads", f"{total_leads:,}".replace(",", "."))
k2.metric("Clientes Gerados", f"{total_clientes:,}".replace(",", "."))
k3.metric("Taxa de Conversao", f"{taxa:.2f}%")

st.divider()

# =====================
# ABAS (SLIDES)
# =====================
tabs = st.tabs(
    [
        "Slide 1 — Contexto",
        "Slide 2 — Metodologia",
        "Slide 3 — Insights",
        "Slide 4 — Perfil Ideal",
        "Slide 5 — Recomendacoes",
        "Anexo — Tabelas completas",
        "Extra — CAC simulado (opcional)",
    ]
)

with tabs[0]:
    st.subheader("Slide 1 — Contexto do Desafio")
    st.write("""
- Base de dados com informacoes de leads e vendas
- Objetivo: identificar caracteristicas com maior chance de conversao
- Apoiar decisoes do time de marketing
""")
    st.info("Resumo: entender quais perfis convertem melhor e transformar isso em recomendacoes praticas.")
    st.caption("Estrutura: dados -> insights -> perfil ideal -> recomendacoes.")

with tabs[1]:
    st.subheader("Slide 2 — Metodologia")
    st.write("""
- Consolidacao das tabelas em uma base unica (BASE_ANALISE)
- Criacao de uma flag de conversao (cliente: 1 para SIM, 0 para NAO)
- Analise por segmentacao (Origem, Mercado, Porte e Objetivo)
- Calculo de conversao: Clientes / Leads
- Comparacao considerando conversao e volume (evitar conclusoes por amostra pequena)
""")

with tabs[2]:
    st.subheader("Slide 3 — Insights")
    st.write("""
- O objetivo aqui e comparar desempenho entre segmentos.
- A leitura recomendada e sempre considerar:
  1) Taxa de conversao
  2) Volume de leads (para evitar distorcao por poucos registros)
""")

    cA, cB = st.columns(2)

    with cA:
        st.write("Conversao por Origem")
        if "ORIGEM" in df_f.columns:
            t = tabela_conversao(df_f, "ORIGEM")
            if "Conversao (%)" in t.columns and len(t) > 0:
                st.bar_chart(t["Conversao (%)"])
            st.dataframe(t, use_container_width=True)
        else:
            st.info("Coluna ORIGEM nao encontrada.")

    with cB:
        st.write("Conversao por Mercado")
        if "MERCADO" in df_f.columns:
            t = tabela_conversao(df_f, "MERCADO")
            if "Conversao (%)" in t.columns and len(t) > 0:
                st.bar_chart(t["Conversao (%)"])
            st.dataframe(t, use_container_width=True)
        else:
            st.info("Coluna MERCADO nao encontrada.")

with tabs[3]:
    st.subheader("Slide 4 — Perfil Ideal de Lead")
    st.write("""
Com base nos segmentos com melhor desempenho (conversao + volume), um perfil ideal pode ser descrito como:

- Origem: canais com melhor equilibrio entre volume e conversao (ex.: Google/Organico quando aparecerem como destaques)
- Mercado: segmentos com maior taxa de conversao identificados na analise
- Porte: portes com maior taxa e volume relevante
- Objetivo: objetivos com maior taxa e volume relevante

Observacao: a definicao final do perfil ideal deve priorizar grupos com volume minimo relevante.
""")

with tabs[4]:
    st.subheader("Slide 5 — Recomendacoes")
    st.write("""
Recomendacoes praticas para maximizar conversao e eficiencia de investimento:

- Priorizar investimento nos canais com melhor equilibrio (conversao e volume)
- Direcionar campanhas para mercados/portes/objetivos com melhor desempenho
- Reduzir ou reavaliar segmentos com alto volume e baixa conversao
- Melhorar qualidade do cadastro (reduzir "Nao informado") para permitir segmentacao mais precisa
""")

with tabs[5]:
    st.subheader("Anexo — Tabelas completas (para conferencia)")
    st.write("Tabelas completas calculadas com base nos filtros selecionados.")

    if "ORIGEM" in df_f.columns:
        st.write("ORIGEM")
        st.dataframe(tabela_conversao(df_f, "ORIGEM"), use_container_width=True)

    if "MERCADO" in df_f.columns:
        st.write("MERCADO")
        st.dataframe(tabela_conversao(df_f, "MERCADO"), use_container_width=True)

    if "PORTE" in df_f.columns:
        st.write("PORTE")
        st.dataframe(tabela_conversao(df_f, "PORTE"), use_container_width=True)

    if "OBJETIVO" in df_f.columns:
        st.write("OBJETIVO")
        st.dataframe(tabela_conversao(df_f, "OBJETIVO"), use_container_width=True)

    # Sub-origem se existir
    if "SUB_ORIGEM" in df_f.columns:
        st.write("SUB_ORIGEM")
        st.dataframe(tabela_conversao(df_f, "SUB_ORIGEM"), use_container_width=True)

with tabs[6]:
    st.subheader("Extra — CAC simulado (opcional)")
    st.write("""
Este bloco calcula CAC simulado usando um CPL ficticio por mercado.
Serve apenas como analise complementar, pois custos reais nao foram fornecidos.
""")

    if "MERCADO" not in df_f.columns:
        st.info("Coluna MERCADO nao encontrada; CAC simulado nao sera calculado.")
    else:
        custos_por_mercado = {
            "Servicos": 25,
            "Empreendedor": 20,
            "Varejo": 30,
            "Franqueadora": 45,
            "Imobiliaria/Incorporadora/Construtora": 35,
            "Marketing/Publicidade/Consultoria": 40,
            "Tecnologia": 50,
            "Industria": 45,
            "Nao informado": 20,
        }

        tmp = df_f.copy()
        tmp["cpl_simulado"] = tmp["MERCADO"].astype(str).map(custos_por_mercado).fillna(30)
        tmp["custo_simulado"] = tmp["cpl_simulado"]

        leads = tmp.groupby("MERCADO")["LEAD_ID"].count().rename("Leads")
        clientes = tmp.groupby("MERCADO")["cliente"].sum().rename("Clientes")
        conversao = (tmp.groupby("MERCADO")["cliente"].mean() * 100).round(2).rename("Conversao (%)")
        gasto = tmp.groupby("MERCADO")["custo_simulado"].sum().rename("Gasto (simulado)")
        cac = (gasto / clientes.clip(lower=1)).round(2).rename("CAC (simulado)")

        tabela_cac = pd.concat([leads, clientes, conversao, gasto, cac], axis=1).sort_values("CAC (simulado)")
        st.dataframe(tabela_cac, use_container_width=True)

        st.write("CAC (simulado) por Mercado")
        st.bar_chart(tabela_cac["CAC (simulado)"])

        st.caption(
            f"Observacao: use o volume minimo da sidebar ({vol_min}) para interpretar resultados com mais seguranca."
        )
