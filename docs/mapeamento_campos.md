# Mapeamento de Campos LTMC - Supplier BP

| Coluna Excel | Campo LTMC | Aba | Observação |
|--------------|------------|-----|------------|
| PARTNER (Codigo BP) | LIFNR | Todas | Código do BP (gerado sequencialmente) |
| AGRUPAMENTO | BU_GROUP | Dados gerais | Padrão 0001 |
| NAME_ORG1 (Razao Social) | NAME_FIRST | Dados gerais | Truncado em 40 caracteres |
| BU_SORT1 (Termo Pesquisa) | SORTL | Dados gerais | Truncado em 20 |
| STREET (Rua) | STREET | Dados gerais | Truncado em 60 |
| HOUSE_NUM1 (Numero) | HOUSE_NUM1 | Dados gerais | Truncado em 10 |
| POST_CODE1 (CEP SP) | POST_CODE1 | Dados gerais | Formato 99999-999 |
| CITY1 (Cidade SP) | CITY1 | Dados gerais | Truncado em 40 |
| REGION (UF) | REGION | Dados gerais | UF |
| COUNTRY | COUNTRY | Dados gerais / Bancários | BR |
| FUNCAO_1 / FUNCAO_2 | BP_ROLE | Funções | FLVN00 / FLVN01 |
| BUKRS (Empresa) | BUKRS | Dados da empresa | |
| AKONT (Conta Conciliacao) | AKONT | Dados da empresa | |
| ZTERM (Cond Pagamento) | ZTERM1 / ZTERM | Empresa / Org. Compras | |
| EKORG (Org Compras) | EKORG | Org. de compra | |
| WAERS (Moeda) | WAERS | Org. de compra | BRL |
| TAXTYPE | TAXTYPE | ID Fiscal | BR1 |
| TAXNUM (CNPJ) | TAXNUM | ID Fiscal | 14 dígitos |
| BANKL (Chave Banco) | BANKL | Detalhes bancários | |
| BANKN (Conta Bancaria) | BANKN | Detalhes bancários | |
| BKONT (Chave Controle) | BKONT | Detalhes bancários | Truncado em 2 |

