---
name: "spec-002"
title: "VBA Engeneering Cost Calculation Engine"
description: "O agente de IA deverá implementar uma rotina de fechamento de custos e renderização de relatório executivo"
version: "1.0.0"
skills:
  - "vba-high-performance"
  - "vba-coding-standards"
  - "code-spec-validator"
environment:
  runtime: "Excel VBA 7.1 (64-bit)"
  host_file: "Nano_Cores_CoreEngine.xlsm"
---

## CONTEXT & ACCEPTANCE CRITERIA
Desenvolver 1 rotina de fechamento de custos.
A rotina se especializa em formar o custo teórico de um sku, com base no custo combinado de todas as matérias primas necessárias para fabricá-lo. Os dados da rotina 2 são obtidos a partir de um arquivo de engenharia de produção em formato '.xlsx', extraído do ERP.

### FUNCTIONAL REQUIREMENTS
- **Cálculo Engenharia do Produto:** Apuração de custo total com perda % estimada, associada por produto acabado &  matéria-prima.