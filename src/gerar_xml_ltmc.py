"""
LTMC Supplier BP Migration - SAP S/4HANA
Gera o arquivo XML no formato Spreadsheet do LTMC a partir da base limpa.

Uso:
    python src/gerar_xml_ltmc.py
"""

from __future__ import annotations

from itertools import islice
from pathlib import Path
import xml.etree.ElementTree as ET
import logging
import sys

import pandas as pd

# ---------------------------------------------------------------------------
# Configuração
# ---------------------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(levelname)s | %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger(__name__)

BASE_DIR = Path(__file__).resolve().parent
ROOT_DIR = BASE_DIR.parent

PATH_CLEAN = ROOT_DIR / "data" / "processed" / "Base_Limpa_BP_Fornecedor.xlsx"
PATH_TEMPLATE = ROOT_DIR / "templates" / "PT_Fornecedor.xml"
PATH_OUTPUT = ROOT_DIR / "output" / "PT_Fornecedor_LTMC_V3_FINAL.xml"

DEFAULT_GROUPING = "0001"
DEFAULT_LANGUAGE = "PT"

# Limites de tamanho de campo do SAP (evita erro de truncamento no LTMC)
MAX_LEN = {
    "NAME_FIRST": 40,
    "SORTL": 20,
    "STREET": 60,
    "HOUSE_NUM1": 10,
    "CITY1": 40,
    "KOINH": 60,
    "BKONT": 2,
}

NS_SS = "urn:schemas-microsoft-com:office:spreadsheet"
ns = {"ss": NS_SS}


# ---------------------------------------------------------------------------
# Funções auxiliares
# ---------------------------------------------------------------------------

def load_and_prepare_data(path: Path) -> pd.DataFrame:
    """Carrega a planilha limpa e aplica as regras de formatação SAP."""
    if not path.exists():
        raise FileNotFoundError(f"Base limpa não encontrada: {path}")

    df = pd.read_excel(path, dtype=str).fillna("")
    df.columns = df.columns.str.strip()

    # Códigos BP sequenciais (001, 002, ...) – remove qualquer prefixo 900 legado
    df["PARTNER (Codigo BP)"] = [str(i + 1).zfill(3) for i in range(len(df))]

    # Agrupamento padrão
    df["AGRUPAMENTO"] = (
        df["AGRUPAMENTO"]
        .astype(str)
        .str.strip()
        .replace({"": DEFAULT_GROUPING, "nan": DEFAULT_GROUPING, "None": DEFAULT_GROUPING})
        .str.zfill(4)
    )

    # CEP brasileiro (somente números + hífen)
    df["POST_CODE1 (CEP SP)"] = (
        df["POST_CODE1 (CEP SP)"]
        .str.replace(r"\D", "", regex=True)
        .str.zfill(8)
        .str.replace(r"^(\d{5})(\d{3})$", r"\1-\2", regex=True)
    )

    # Correção pontual de chave bancária (documentar origem se necessário)
    df["BANKL (Chave Banco)"] = (
        df["BANKL (Chave Banco)"].str.strip().replace({"237212345": "23721234"})
    )

    # CNPJ somente dígitos
    df["TAXNUM (CNPJ)"] = (
        df["TAXNUM (CNPJ)"].str.replace(r"\D", "", regex=True).str.zfill(14)
    )

    return df


def popular_aba_ltmc(root: ET.Element, nome_aba: str, registros: list[dict]) -> None:
    """Popula uma aba específica do template XML do LTMC."""
    worksheet = next(
        (
            ws
            for ws in root.findall("ss:Worksheet", ns)
            if ws.attrib.get(f"{{{NS_SS}}}Name") == nome_aba
        ),
        None,
    )
    if worksheet is None:
        raise ValueError(f"Aba '{nome_aba}' não encontrada no template LTMC")

    table = worksheet.find("ss:Table", ns)
    if table is None:
        raise ValueError(f"Tabela não encontrada na aba '{nome_aba}'")

    rows = table.findall("ss:Row", ns)

    # Cabeçalho está na 5ª linha (índice 4)
    header_row = next(islice(rows, 4, 5), None)
    if header_row is None:
        raise ValueError(f"Linha de cabeçalho não encontrada na aba '{nome_aba}'")

    col_map: dict[str, int] = {}
    col_idx = 1
    for cell in header_row.findall("ss:Cell", ns):
        idx_attr = cell.attrib.get(f"{{{NS_SS}}}Index")
        if idx_attr:
            col_idx = int(idx_attr)
        data = cell.find("ss:Data", ns)
        if data is not None and data.text:
            col_map[data.text.strip()] = col_idx
        col_idx += 1

    # Remove dados antigos (a partir da 9ª linha)
    for old_row in islice(rows, 8, None):
        table.remove(old_row)

    for reg in registros:
        row_elem = ET.SubElement(table, f"{{{NS_SS}}}Row")
        row_elem.set(f"{{{NS_SS}}}AutoFitHeight", "0")

        celulas = sorted(
            (col_map[k], str(v).strip())
            for k, v in reg.items()
            if str(v).strip() and k in col_map
        )

        ultimo_idx = 0
        for col_pos, valor in celulas:
            cell_elem = ET.SubElement(row_elem, f"{{{NS_SS}}}Cell")
            if col_pos != ultimo_idx + 1:
                cell_elem.set(f"{{{NS_SS}}}Index", str(col_pos))
            data_elem = ET.SubElement(cell_elem, f"{{{NS_SS}}}Data")
            data_elem.set(f"{{{NS_SS}}}Type", "String")
            data_elem.text = valor
            ultimo_idx = col_pos

    table.set(f"{{{NS_SS}}}ExpandedRowCount", str(8 + len(registros)))
    logger.info(f"[OK] Aba '{nome_aba}': {len(registros)} linhas")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    logger.info("Iniciando geração do XML LTMC – Supplier BP")

    df = load_and_prepare_data(PATH_CLEAN)
    logger.info(
        f"Registros carregados: {len(df)} | "
        f"de '{df.iloc[0]['PARTNER (Codigo BP)']}' até '{df.iloc[-1]['PARTNER (Codigo BP)']}'"
    )

    if not PATH_TEMPLATE.exists():
        raise FileNotFoundError(f"Template não encontrado: {PATH_TEMPLATE}")

    tree = ET.parse(PATH_TEMPLATE)
    root = tree.getroot()

    # Namespaces necessários para o formato Excel Spreadsheet
    for prefix, uri in [
        ("", NS_SS),
        ("o", "urn:schemas-microsoft-com:office:office"),
        ("x", "urn:schemas-microsoft-com:office:excel"),
        ("dt", "uuid:C2F41010-65B3-11d1-A29F-00AA00C14882"),
        ("ss", NS_SS),
        ("html", "http://www.w3.org/TR/REC-html40"),
    ]:
        ET.register_namespace(prefix, uri)

    col_region = "REGION (UF)" if "REGION (UF)" in df.columns else "REGION"
    records = df.to_dict("records")

    # --- Dados gerais ---
    popular_aba_ltmc(
        root,
        "Dados gerais",
        [
            {
                "LIFNR": r["PARTNER (Codigo BP)"],
                "BU_GROUP": r["AGRUPAMENTO"],
                "NAME_FIRST": r["NAME_ORG1 (Razao Social)"][: MAX_LEN["NAME_FIRST"]],
                "SORTL": r["BU_SORT1 (Termo Pesquisa)"][: MAX_LEN["SORTL"]],
                "STREET": r["STREET (Rua)"][: MAX_LEN["STREET"]],
                "HOUSE_NUM1": r["HOUSE_NUM1 (Numero)"][: MAX_LEN["HOUSE_NUM1"]],
                "POST_CODE1": r["POST_CODE1 (CEP SP)"],
                "CITY1": r["CITY1 (Cidade SP)"][: MAX_LEN["CITY1"]],
                "COUNTRY": r["COUNTRY"],
                "REGION": r[col_region],
                "LANGU_CORR": DEFAULT_LANGUAGE,
            }
            for r in records
        ],
    )

    # --- Funções do parceiro ---
    funcoes = []
    for r in records:
        for func_col in ("FUNCAO_1", "FUNCAO_2"):
            if r.get(func_col):
                funcoes.append(
                    {"LIFNR": r["PARTNER (Codigo BP)"], "BP_ROLE": r[func_col]}
                )
    popular_aba_ltmc(root, "Funções do parceiro de negócio", funcoes)

    # --- Dados da empresa ---
    popular_aba_ltmc(
        root,
        "Dados da empresa",
        [
            {
                "LIFNR": r["PARTNER (Codigo BP)"],
                "BUKRS": r["BUKRS (Empresa)"],
                "AKONT": r["AKONT (Conta Conciliacao)"],
                "ZTERM1": r["ZTERM (Cond Pagamento)"],
            }
            for r in records
        ],
    )

    # --- Organização de compra ---
    popular_aba_ltmc(
        root,
        "Dados de organização de compra",
        [
            {
                "LIFNR": r["PARTNER (Codigo BP)"],
                "EKORG": r["EKORG (Org Compras)"],
                "WAERS": r["WAERS (Moeda)"],
                "ZTERM": r["ZTERM (Cond Pagamento)"],
            }
            for r in records
        ],
    )

    # --- Identificação fiscal ---
    popular_aba_ltmc(
        root,
        "Números de identificação fiscal",
        [
            {
                "LIFNR": r["PARTNER (Codigo BP)"],
                "TAXTYPE": r["TAXTYPE"],
                "TAXNUM": r["TAXNUM (CNPJ)"],
            }
            for r in records
        ],
    )

    # --- Detalhes bancários (somente se preenchido) ---
    popular_aba_ltmc(
        root,
        "Detalhes bancários",
        [
            {
                "LIFNR": r["PARTNER (Codigo BP)"],
                "BANKS": r["COUNTRY"],
                "BANKL": r["BANKL (Chave Banco)"],
                "BANKN": r["BANKN (Conta Bancaria)"],
                "BKONT": r["BKONT (Chave Controle)"][: MAX_LEN["BKONT"]],
                "KOINH": r["NAME_ORG1 (Razao Social)"][: MAX_LEN["KOINH"]],
            }
            for r in records
            if r["BANKL (Chave Banco)"] and r["BANKN (Conta Bancaria)"]
        ],
    )

    # Garante que a pasta de saída existe
    PATH_OUTPUT.parent.mkdir(parents=True, exist_ok=True)

    # Escrita do XML
    tree.write(PATH_OUTPUT, encoding="utf-8", xml_declaration=True)

    # Ajuste da declaração para o formato esperado pelo Excel/LTMC
    content = PATH_OUTPUT.read_text(encoding="utf-8")
    content = content.replace(
        "<?xml version='1.0' encoding='utf-8'?>",
        '<?xml version="1.0"?>\n<?mso-application progid="Excel.Sheet"?>',
        1,
    )
    PATH_OUTPUT.write_text(content, encoding="utf-8")

    logger.info(f"Arquivo gerado com sucesso: {PATH_OUTPUT}")
    logger.info("Pronto para upload no LTMC / Migrate Your Data")


if __name__ == "__main__":
    main()
