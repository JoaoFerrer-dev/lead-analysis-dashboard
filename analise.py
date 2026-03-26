import os
import pandas as pd

# =====================
# CONFIGURACAO
# =====================
PASTA = r"C:\Users\joaov\Documents\DESAFIO ESTAGIO"
ARQUIVO = os.path.join(PASTA, "01_Analise_Dados.xlsx")
ABAS_PREFERIDAS = ["BASE_ANALISE", "BASE"]
SAIDA = os.path.join(PASTA, "resultado_final.xlsx")

# =====================
# FUNCOES
# =====================
def escolher_aba(path: str) -> str:
    xls = pd.ExcelFile(path)
    for s in ABAS_PREFERIDAS:
        if s in xls.sheet_names:
            return s
    raise ValueError(
        f"Nenhuma das abas {ABAS_PREFERIDAS} foi encontrada. Abas disponiveis: {xls.sheet_names}"
    )

def preparar_df(df: pd.DataFrame) -> pd.DataFrame:
    obrigatorias = ["LEAD_ID", "VENDIDO"]
    faltando = [c for c in obrigatorias if c not in df.columns]
    if faltando:
        raise ValueError(f"Faltam colunas obrigatorias: {faltando}")

    df = df.copy()

    # Normaliza possiveis nomes de coluna de sub-origem
    if "SUB-ORIGEM" in df.columns and "SUB_ORIGEM" not in df.columns:
        df = df.rename(columns={"SUB-ORIGEM": "SUB_ORIGEM"})
    if "SUB ORIGEM" in df.columns and "SUB_ORIGEM" not in df.columns:
        df = df.rename(columns={"SUB ORIGEM": "SUB_ORIGEM"})

    # Flag conversao
    df["cliente"] = df["VENDIDO"].apply(lambda x: 1 if str(x).strip().upper() == "SIM" else 0)

    # Padroniza dimensoes
    for col in ["ORIGEM", "SUB_ORIGEM", "MERCADO", "PORTE", "OBJETIVO", "LOCAL"]:
        if col in df.columns:
            df[col] = df[col].fillna("Nao informado").astype(str).str.strip()
            df[col] = df[col].replace(
                {"0": "Nao informado", "": "Nao informado", "nan": "Nao informado", "None": "Nao informado"}
            )

    return df

def tabela_conversao(df: pd.DataFrame, dim: str) -> pd.DataFrame:
    if dim not in df.columns:
        return pd.DataFrame({"Aviso": [f"Coluna '{dim}' nao encontrada na base."]})

    t = df.groupby(dim).agg(
        Leads=("LEAD_ID", "count"),
        Clientes=("cliente", "sum"),
    )
    t["Conversao (%)"] = (t["Clientes"] / t["Leads"] * 100).round(2)
    t = t.sort_values(["Conversao (%)", "Leads"], ascending=[False, False])
    return t

def cac_simulado(df: pd.DataFrame) -> pd.DataFrame:
    if "MERCADO" not in df.columns:
        return pd.DataFrame({"Aviso": ["Coluna 'MERCADO' nao encontrada; CAC simulado nao calculado."]})

    custos = {
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

    tmp = df.copy()
    tmp["cpl_simulado"] = tmp["MERCADO"].astype(str).map(custos).fillna(30)
    tmp["custo_simulado"] = tmp["cpl_simulado"]

    leads = tmp.groupby("MERCADO")["LEAD_ID"].count().rename("Leads")
    clientes = tmp.groupby("MERCADO")["cliente"].sum().rename("Clientes")
    conversao = (tmp.groupby("MERCADO")["cliente"].mean() * 100).round(2).rename("Conversao (%)")
    gasto = tmp.groupby("MERCADO")["custo_simulado"].sum().rename("Gasto (simulado)")
    cac = (gasto / clientes.clip(lower=1)).round(2).rename("CAC (simulado)")

    out = pd.concat([leads, clientes, conversao, gasto, cac], axis=1)
    out = out.sort_values(["CAC (simulado)", "Conversao (%)"], ascending=[True, False])
    return out

# =====================
# EXECUCAO
# =====================
def main():
    aba = escolher_aba(ARQUIVO)
    df = pd.read_excel(ARQUIVO, sheet_name=aba)
    df.columns = [str(c).strip() for c in df.columns]
    df = preparar_df(df)

    t_origem = tabela_conversao(df, "ORIGEM")
    t_mercado = tabela_conversao(df, "MERCADO")
    t_porte = tabela_conversao(df, "PORTE")
    t_objetivo = tabela_conversao(df, "OBJETIVO")
    t_sub = tabela_conversao(df, "SUB_ORIGEM") if "SUB_ORIGEM" in df.columns else pd.DataFrame(
        {"Aviso": ["Coluna 'SUB_ORIGEM' nao encontrada; analise nao gerada."]}
    )
    t_cac = cac_simulado(df)

    with pd.ExcelWriter(SAIDA, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="BASE_USADA", index=False)
        t_origem.to_excel(writer, sheet_name="01_Analise_Origem")
        t_mercado.to_excel(writer, sheet_name="02_Analise_Mercado")
        t_porte.to_excel(writer, sheet_name="03_Analise_Porte")
        t_objetivo.to_excel(writer, sheet_name="04_Analise_Objetivo")
        t_sub.to_excel(writer, sheet_name="05_Analise_SubOrigem")
        t_cac.to_excel(writer, sheet_name="06_CAC_Simulado_Mercado")

    print("Analise concluida.")
    print(f"Arquivo de entrada: {ARQUIVO}")
    print(f"Aba utilizada: {aba}")
    print(f"Arquivo gerado: {SAIDA}")

if __name__ == "__main__":
    main()
