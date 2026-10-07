---
name: "spec-002"
title: "VBA Engeneering Cost Calculation Engine"
description: "O agente de IA deverá implementar uma rotina de fechamento de custos e renderização de relatório executivo"
version: "1.0.0"
skills:
  - "vba-high-performance"
  - "vba-coding-standards"
  - "code-spec-validator"
  - "end-user-doc-generator"
environment:
  runtime: "Excel VBA 7.1 (64-bit)"
  host_file: "Nano_Cores_CoreEngine.xlsm"
---

## 1. CONTEXT & ACCEPTANCE CRITERIA
### 1.1 Bussiness Context
Desenvolver 1 rotina de fechamento de custos que se especializa em calcular custos por sku, a partir do relatório de explosão de engenharia do produto. Com base em uma lista de custos das matérias primas, e uma lista de associação de perda de matéria prima por sku, a rotina deverá calcular o custo total do sku remontando sua engenharia completa e salvar o resultado adicionando uma nova aba na planilha de banco de dados, situada em `G:\Ícaro\Cerulean_DB\sheets\database\Cerulean_DB_Custos.xlsx`.

### 1.2 Functional Requirements
- **RF-01 (Taxonomia e Classificação de SKUs):** Categorizar os itens da explosão de engenharia em Produto Acabado (PA), Semiacabado (SA) e Matéria-Prima (MP), filtrando agrupadores intermediários para evitar dupla contagem.
- **RF-02 (Cálculo de Custo com Perdas):** Consolidar o custo unitário do PA através da soma dos componentes folha (MP), aplicando os fatores de perda cadastrados.

### 1.3 Acceptance Criteria (Definition of Done)
- [ ] **Critério 1 (Taxonomia Estrita de SKUs):** O classificador deve mapear os itens exclusivamente pelas seguintes regras de código/prefixo:
  | Categoria | Papel no Pipeline | Códigos / Prefixos Válidos |
  | :--- | :--- | :--- |
  | **PA** (Produto Acabado) | SKU Pai receptor do custo total | `40000`, `44000`, `45000` |
  | **SA** (Semiacabado) | Nó intermediário agrupador (Excluído do somatório) | `30000`, `33000` |
  | **MP** (Matéria-Prima) | Componente produtivo (Entra no cálculo) | `10000` |
  | **MP** (Embalagem) | Componente de embalagem (Entra no cálculo) | `20000`, `21000` |

- [ ] **Critério 2 (Isolamento de SA e Supressão de Dupla Contagem):** SKUs categorizados como `SA` não devem ter seus valores adicionados ao custo do `PA`. O motor deve perfurar o nível do SA e computar unicamente os itens filhos do tipo `MP` (`10000`) subordinados a ele.

- [ ] **Critério 3 (Equação Determinística de Custo):** O custo final de cada SKU PA deve coincidir exatamente com a fórmula matemática:
  $$\text{Custo\_PA} = \sum_{i \in \text{MPs}} \left[ \text{Qtd}_i \times \text{Custo\_Unit}_i \times (1 + \text{Taxa\_Perda}_i) \right]$$
  *(onde $\text{Taxa\_Perda}_i$ é o percentual em decimal referente ao SKU filho $i$)*.

- [ ] **Critério 4 (Resiliência de Perdas Ausentes):** Caso um SKU MP não conste na lista de perdas ou apresente valor em branco, o motor deve atribuir $\text{Taxa\_Perda} = 0.00$ e prosseguir o cálculo sem interrupção.

## 2. DATA CONTRACTS & SOURCES
### Input Contract 01:
* **Origin Path:** `G:\Ícaro\Cerulean_DB\sheets\database\Cerulean_DB_Parametros_Associacao_Produtos.xlsx`
* **Origin Entity:** Worksheet `Assoc_PerdasEngenharia`
* **Ingestion Method:** `ADODB.Connection` (Provider: `Microsoft.ACE.OLEDB.12.0; Extended Properties="Excel 12.0 Xml;HDR=YES;IMEX=1;"`)
* **SQL Query Definition:**
  ```sql
  SELECT 
    [Cod_PA] & '-' & [Cod_MP] AS [Chave_Combinada],
    [Cod_PA], 
    [Cod_MP], 
    [Perda_Percentual] 
  FROM [Assoc_PerdasEngenharia$]
  WHERE [Cod_PA] IS NOT NULL 
  AND [Cod_MP] IS NOT NULL 
  AND [Perda_Percentual] IS NOT NULL;
  ```
* **Schema Specification:**
  | Coluna SQL | Tipo VBA | Papel no Contrato | Regra / Validação |
  | :--- | :--- | :--- | :--- |
  | `Cod_PA` | `String` | Chave de Agrupamento | Obrigatório; Converter para string limpa (`Trim$`) |
  | `Cod_MP` | `String` | SKU Matéria Prima | Obrigatório; Converter para string limpa (`Trim$`) |
  | `Perda_Percentual` | `Double` | Taxa de Perda | Obrigatório; rejeitar valores nulos ou negativos |

* **Post-Ingestion Lifecycle & In-Memory Target:**
  * **Destinos em memória:** Criar dicionário `dictPerdas` com chave composta `Cod_PA & "-" & Cod_MP` e valor `Perda_Percentual`:
  - **Papel:** Armazenar as taxas de perda associadas a cada combinação de SKU PA e MP.

* **Fallback & Fault Tolerance:**
  * Se houver `Cod_PA` duplicado em `Assoc_PerdasEngenharia`: sobrescrever com o último registro e logar no painel de debug.

---

### Output Contract 01:

---

## 3. ALGORITHMIC PIPELINE (TAXONOMY RESOLUTION ENGINE)
### 3.1 

## 4. ARCHITECTURAL BLUEPRINT (MODULE DECOMPOSITION)
### Module 1: 