# 🚚 LTMC Supplier BP Migration – SAP S/4HANA (Brasil)

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python&logoColor=white)](https://www.python.org/)
[![Pandas](https://img.shields.io/badge/Pandas-2.0%2B-150458?logo=pandas&logoColor=white)](https://pandas.pydata.org/)
[![SAP](https://img.shields.io/badge/SAP-S%2F4HANA%20LTMC-0FAAFF?logo=sap&logoColor=white)](https://help.sap.com/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Status](https://img.shields.io/badge/Status-Production%20Ready-success)]()

> Projeto real e bem-sucedido de migração de **Business Partners Fornecedores** para o SAP S/4HANA utilizando o **Legacy Transfer Migration Cockpit (LTMC)**.

---

## 📌 Visão Geral

Este repositório contém a solução completa de **Data Cleansing + geração automática do arquivo XML** no formato exigido pelo LTMC para o objeto de migração de **Supplier (Fornecedor)**.

**Problema de negócio resolvido:**
- Dados legados de fornecedores sujos (CNPJ inválido/duplicado, endereço inconsistente, formatação incorreta de CEP, etc.).
- Necessidade de gerar o arquivo XML multi-sheet do LTMC de forma automatizada e confiável, respeitando os tamanhos de campo e regras brasileiras (CNPJ, UF, CEP, roles FLVN00/FLVN01).

**Resultado da limpeza de dados:**

| Etapa                  | Quantidade | Observação                                   |
|------------------------|------------|----------------------------------------------|
| Base suja (raw)        | **50**     | Dados fictícios originais                    |
| Base limpa (processed) | **40**     | Após remoção de duplicatas de CNPJ + limpeza |
| Registros removidos    | **10**     | Principalmente por CNPJ duplicado            |

**Como a solução funciona:**
1. Planilha suja (**50** registros) → Notebook de limpeza (Jupyter)
2. Planilha limpa (**40** registros) → Script Python que popula o template XML oficial do LTMC
3. XML final pronto para upload na transação **LTMC** / app **Migrate Your Data**

---

## 🏗 Arquitetura e Fluxo de Dados

```
data/raw/
└── base_fornecedores_ficticia-v3.xlsx     ← 50 registros (dados legados sujos)

          │
          ▼
notebooks/
└── Limpeza.ipynb                          ← Cleansing (CNPJ válido, dedup, endereço...)
                                             → remove 10 duplicados de CNPJ
          │
          ▼
data/processed/
└── Base_Limpa_BP_Fornecedor.xlsx          ← 40 registros (base limpa)

          │
          ▼
src/gerar_xml_ltmc.py  +  templates/PT_Fornecedor.xml

          │
          ▼
output/
└── PT_Fornecedor_LTMC_V3_FINAL.xml        ← Arquivo pronto para LTMC (40 BPs)
```

**Abas populadas automaticamente:**

| Aba LTMC                              | Campos principais                                      |
|---------------------------------------|--------------------------------------------------------|
| Dados gerais                          | LIFNR, BU_GROUP, NAME_FIRST, SORTL, STREET, HOUSE_NUM1, POST_CODE1, CITY1, COUNTRY, REGION, LANGU_CORR |
| Funções do parceiro de negócio        | LIFNR + BP_ROLE (FLVN00 / FLVN01)                      |
| Dados da empresa                      | LIFNR, BUKRS, AKONT, ZTERM1                            |
| Dados de organização de compra        | LIFNR, EKORG, WAERS, ZTERM                             |
| Números de identificação fiscal       | LIFNR, TAXTYPE (BR1), TAXNUM (CNPJ)                    |
| Detalhes bancários                    | LIFNR, BANKS, BANKL, BANKN, BKONT, KOINH               |

---

## 🛠 Tecnologias e Dependências

| Tecnologia              | Uso                                      |
|-------------------------|------------------------------------------|
| Python 3.10+            | Script principal                         |
| Pandas ≥ 2.0            | Manipulação de dados                     |
| openpyxl ≥ 3.0          | Leitura/escrita Excel                    |
| xml.etree.ElementTree   | Manipulação do XML Spreadsheet do LTMC   |
| Jupyter Notebook        | Data Cleansing                           |
| SAP S/4HANA LTMC        | Destino final (Migrate Your Data)        |

**Transações / Apps SAP:**
- `LTMC` / Fiori App **Migrate Your Data**
- Objeto de migração: **Supplier** (Business Partner – Vendor)

---

## 📁 Estrutura do Repositório

```
ltmc-supplier-bp-migration/
├── .gitignore
├── README.md
├── requirements.txt
├── src/
│   ├── gerar_xml_ltmc.py          ← Versão 2.0 (refatorada e modular)
│   └── gerar_xml_ltmc_v1.py       ← Versão 1.0 (protótipo funcional)
├── data/
│   ├── raw/                           ← Dados sujos
│   └── processed/                     ← Dados limpos
├── templates/
│   └── PT_Fornecedor.xml              ← Template oficial LTMC
├── output/                            ← XMLs gerados (não versionar)
├── notebooks/
│   └── Limpeza.ipynb                  ← Notebook de data cleansing
└── docs/
```

---

## 🚀 Guia de Instalação e Uso

### 1. Abrir no VS Code e criar ambiente virtual

```bash
# No terminal do VS Code (na raiz do projeto)
python -m venv .venv

# Windows
.venv\Scripts\activate

# Linux / macOS
source .venv/bin/activate

pip install -r requirements.txt
```

### 2. (Opcional) Re-executar a limpeza de dados

Abra o notebook `notebooks/Limpeza.ipynb` e execute todas as células.  
A planilha limpa será salva em `data/processed/Base_Limpa_BP_Fornecedor.xlsx`.

### 3. Gerar o XML do LTMC

```bash
python src/gerar_xml_ltmc.py
```

O arquivo será gerado em:
```
output/PT_Fornecedor_LTMC_V3_FINAL.xml
```

### 4. Upload no SAP

1. Acesse o app **Migrate Your Data** (ou transação `LTMC`)
2. Crie ou abra o projeto de migração do objeto **Supplier**
3. Faça o upload do arquivo XML gerado
4. Execute a simulação e a migração

---

## 📋 Exemplo de Input / Output (fictício)

**Input (trecho da base limpa):**

| PARTNER | NAME_ORG1                  | TAXNUM (CNPJ)   | CITY1      | REGION |
|---------|----------------------------|-----------------|------------|--------|
| 001     | Comércio Atlântico Ltda.   | 00662766000131  | Piracicaba | SP     |
| 002     | Plásticos Horizonte Ltda   | 03292694000176  | São Paulo  | SP     |

**Output:** XML multi-sheet pronto, com as 6 abas preenchidas e declaração correta do Excel Spreadsheet (`<?mso-application progid="Excel.Sheet"?>`).

---

## ⚠️ Aviso de Segurança e LGPD

> **Importante:** Todos os dados presentes neste repositório são **fictícios e anonimizados**.  
> Nenhum CNPJ, razão social, conta bancária ou informação real de clientes/fornecedores foi versionado.  
> Em projetos reais, **nunca** faça commit de bases com dados pessoais ou corporativos sensíveis.  
> Utilize sempre dados mascarados ou sintéticos em repositórios públicos.

---

## 📄 Licença

MIT License – sinta-se livre para usar, adaptar e evoluir.

---

**Desenvolvido com foco em qualidade de dados e automação de migração SAP S/4HANA.**
