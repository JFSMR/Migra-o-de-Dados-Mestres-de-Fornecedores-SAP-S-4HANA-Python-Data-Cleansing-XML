from itertools import islice
import os
import xml.etree.ElementTree as ET
import pandas as pd

PASTA_ATUAL = os.path.dirname(os.path.abspath(__file__))
RAIZ_PROJETO = os.path.dirname(PASTA_ATUAL)

CAMINHO_BASE_LIMPA = os.path.join(
    RAIZ_PROJETO, "BASE_DADOS_LIMPA", "Base_Limpa_BP_Fornecedor.xlsx"
)
CAMINHO_TEMPLATE_XML = os.path.join(PASTA_ATUAL, "PT_Fornecedor.xml")
CAMINHO_SAIDA_XML = os.path.join(PASTA_ATUAL, "PT_Fornecedor_LTMC_V3_FINAL.xml")

IDIOMA_CORRESPONDENCIA = "PT"
AGRUPAMENTO_PADRAO_SAP = "0001"

df = pd.read_excel(CAMINHO_BASE_LIMPA, dtype=str).fillna("")
df.columns = df.columns.str.strip()

# Garante estritamente 3 dígitos ('001' a '040') sem nenhum prefixo 900
df["PARTNER (Codigo BP)"] = [str(i + 1).zfill(3) for i in range(len(df))]

df["AGRUPAMENTO"] = (
    df["AGRUPAMENTO"]
    .astype(str)
    .str.strip()
    .replace({"": AGRUPAMENTO_PADRAO_SAP, "nan": AGRUPAMENTO_PADRAO_SAP, "None": AGRUPAMENTO_PADRAO_SAP})
    .str.zfill(4)
)

df["POST_CODE1 (CEP SP)"] = (
    df["POST_CODE1 (CEP SP)"]
    .str.replace(r"\D", "", regex=True)
    .str.zfill(8)
    .str.replace(r"^(\d{5})(\d{3})$", r"\1-\2", regex=True)
)

df["BANKL (Chave Banco)"] = (
    df["BANKL (Chave Banco)"].str.strip().replace({"237212345": "23721234"})
)

df["TAXNUM (CNPJ)"] = (
    df["TAXNUM (CNPJ)"].str.replace(r"\D", "", regex=True).str.zfill(14)
)

primeiro = next(iter(df["PARTNER (Codigo BP)"]), "")
ultimo = next(reversed(df["PARTNER (Codigo BP)"].tolist()), "")
print(f"Gerando XML com {len(df)} registros: de '{primeiro}' até '{ultimo}'")

tree = ET.parse(CAMINHO_TEMPLATE_XML)
root = tree.getroot()

NS_SS = "urn:schemas-microsoft-com:office:spreadsheet"
ns = {"ss": NS_SS}

ET.register_namespace("", NS_SS)
ET.register_namespace("o", "urn:schemas-microsoft-com:office:office")
ET.register_namespace("x", "urn:schemas-microsoft-com:office:excel")
ET.register_namespace("dt", "uuid:C2F41010-65B3-11d1-A29F-00AA00C14882")
ET.register_namespace("ss", NS_SS)
ET.register_namespace("html", "http://w3.org")


def popular_aba_ltmc(nome_aba, registros):
    worksheet = next(
        (ws for ws in root.findall("ss:Worksheet", ns) if ws.attrib.get(f"{{{NS_SS}}}Name") == nome_aba),
        None,
    )
    table = worksheet.find("ss:Table", ns)
    rows = table.findall("ss:Row", ns)

    header_row = next(islice(rows, 4, 5))
    col_map = {}
    col_idx = 1
    for cell in header_row.findall("ss:Cell", ns):
        idx_attr = cell.attrib.get(f"{{{NS_SS}}}Index")
        if idx_attr:
            col_idx = int(idx_attr)
        data = cell.find("ss:Data", ns)
        if data is not None and data.text:
            col_map[data.text.strip()] = col_idx
        col_idx += 1

    for old_row in islice(rows, 8, None):
        table.remove(old_row)

    for reg in registros:
        row_elem = ET.SubElement(table, f"{{{NS_SS}}}Row")
        row_elem.set(f"{{{NS_SS}}}AutoFitHeight", "0")

        celulas_ordenadas = sorted(
            [(col_map[k], str(v).strip()) for k, v in reg.items() if str(v).strip() and k in col_map]
        )

        ultimo_idx = 0
        for col_pos, valor_str in celulas_ordenadas:
            cell_elem = ET.SubElement(row_elem, f"{{{NS_SS}}}Cell")
            if col_pos != ultimo_idx + 1:
                cell_elem.set(f"{{{NS_SS}}}Index", str(col_pos))
            data_elem = ET.SubElement(cell_elem, f"{{{NS_SS}}}Data")
            data_elem.set(f"{{{NS_SS}}}Type", "String")
            data_elem.text = valor_str
            ultimo_idx = col_pos

    table.set(f"{{{NS_SS}}}ExpandedRowCount", str(8 + len(registros)))
    print(f"[OK] Aba '{nome_aba}': {len(registros)} linhas.")


col_region = "REGION (UF)" if "REGION (UF)" in df.columns else "REGION"

popular_aba_ltmc("Dados gerais", [
    {
        "LIFNR": r["PARTNER (Codigo BP)"],
        "BU_GROUP": r["AGRUPAMENTO"],
        "NAME_FIRST": r["NAME_ORG1 (Razao Social)"][:40],
        "SORTL": r["BU_SORT1 (Termo Pesquisa)"][:20],
        "STREET": r["STREET (Rua)"][:60],
        "HOUSE_NUM1": r["HOUSE_NUM1 (Numero)"][:10],
        "POST_CODE1": r["POST_CODE1 (CEP SP)"],
        "CITY1": r["CITY1 (Cidade SP)"][:40],
        "COUNTRY": r["COUNTRY"],
        "REGION": r[col_region],
        "LANGU_CORR": IDIOMA_CORRESPONDENCIA,
    }
    for _, r in df.iterrows()
])

funcoes_bp = []
for _, r in df.iterrows():
    if r["FUNCAO_1"]:
        funcoes_bp.append({"LIFNR": r["PARTNER (Codigo BP)"], "BP_ROLE": r["FUNCAO_1"]})
    if r["FUNCAO_2"]:
        funcoes_bp.append({"LIFNR": r["PARTNER (Codigo BP)"], "BP_ROLE": r["FUNCAO_2"]})
popular_aba_ltmc("Funções do parceiro de negócio", funcoes_bp)

popular_aba_ltmc("Dados da empresa", [
    {
        "LIFNR": r["PARTNER (Codigo BP)"],
        "BUKRS": r["BUKRS (Empresa)"],
        "AKONT": r["AKONT (Conta Conciliacao)"],
        "ZTERM1": r["ZTERM (Cond Pagamento)"],
    }
    for _, r in df.iterrows()
])

popular_aba_ltmc("Dados de organização de compra", [
    {
        "LIFNR": r["PARTNER (Codigo BP)"],
        "EKORG": r["EKORG (Org Compras)"],
        "WAERS": r["WAERS (Moeda)"],
        "ZTERM": r["ZTERM (Cond Pagamento)"],
    }
    for _, r in df.iterrows()
])

popular_aba_ltmc("Números de identificação fiscal", [
    {
        "LIFNR": r["PARTNER (Codigo BP)"],
        "TAXTYPE": r["TAXTYPE"],
        "TAXNUM": r["TAXNUM (CNPJ)"],
    }
    for _, r in df.iterrows()
])

popular_aba_ltmc("Detalhes bancários", [
    {
        "LIFNR": r["PARTNER (Codigo BP)"],
        "BANKS": r["COUNTRY"],
        "BANKL": r["BANKL (Chave Banco)"],
        "BANKN": r["BANKN (Conta Bancaria)"],
        "BKONT": r["BKONT (Chave Controle)"][:2],
        "KOINH": r["NAME_ORG1 (Razao Social)"][:60],
    }
    for _, r in df.iterrows()
    if r["BANKL (Chave Banco)"] and r["BANKN (Conta Bancaria)"]
])

tree.write(CAMINHO_SAIDA_XML, encoding="utf-8", xml_declaration=True)

with open(CAMINHO_SAIDA_XML, "r", encoding="utf-8") as f:
    conteudo = f.read()

conteudo = conteudo.replace(
    "<?xml version='1.0' encoding='utf-8'?>",
    '<?xml version="1.0"?>\n<?mso-application progid="Excel.Sheet"?>',
    1,
)

with open(CAMINHO_SAIDA_XML, "w", encoding="utf-8") as f:
    f.write(conteudo)

print(f"\n Arquivo novo criado com sucesso: {CAMINHO_SAIDA_XML}")
