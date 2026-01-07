# Sistema de Auditoria e Ajuste Patrimonial

> **Projeto de Otimização de Processos Administrativos utilizando Excel Avançado, Power Query e VBA.**

![Ferramentas](https://img.shields.io/badge/Stack-Excel%20|%20Power%20Query%20|%20VBA%20|%20Power%20BI-blue)

## Visão Geral

Este projeto foi desenvolvido para resolver um problema crítico de inconsistência de dados no inventário patrimonial de uma unidade corporativa. O processo manual anterior era moroso, propenso a erros humanos e não refletia a realidade física dos ativos.

A solução consiste num **pipeline de dados automatizado** que ingere dados brutos de leituras físicas (QR Code), cruza com a base legado do sistema (ERP) e gera relatórios de auditoria prontos para tomada de decisão.

## ⚠️ Disclaimer (LGPD)
**Nota:** Os dados apresentados neste repositório (ficheiros CSV/Excel) são **fictícios (dummy data)** gerados apenas para demonstrar a lógica e o funcionamento da ferramenta, respeitando integralmente as políticas de privacidade e proteção de dados da empresa.

---

## Problema

* **Cenário:** Inventário físico realizado manualmente, sala a sala.
* **Dores:**
    * Base de dados do sistema desatualizada (Localizações incorretas).
    * Cadastros inconsistentes (Códigos misturados com nomes de responsáveis).
    * Dificuldade em identificar ativos transferidos de outras unidades (Curitiba, Maringá, etc.).
    * Tempo excessivo para consolidar "O que está na sala" vs "O que deveria estar".

## Solução

Desenvolvi uma solução "Full Microsoft Stack" para sanear e auditar os dados:

1.  **Coleta de Dados (Input):** Utilização de scanner mobile para gerar CSVs brutos da realidade física.
2.  **ETL (Extract, Transform, Load) com Power Query:**
    * Limpeza de strings (remoção de caracteres especiais).
    * Padronização de nomes de salas (De/Para de códigos de localização).
    * Separação de colunas (Responsável / Descrição / Património).
3.  **Reconciliação de Dados:** Algoritmo lógico para cruzar a "Base Real" com a "Base Sistema".
4.  **Automação Visual (VBA):** Script para formatar o relatório final, destacando automaticamente divergências (Faltas/Sobras) com formatação condicional.
5.  **Dashboard Gerencial:** Visualização dos KPIs de acuracidade no Power BI (ou Excel).

---

## Resultados

*  **Acuracidade:** Elevação da precisão do inventário da unidade de **~40% para 100%** nas salas auditadas.
*  **Recuperação:** Identificação e regularização de **36 ativos** que não constavam na base local (sobras/transferências não oficializadas).
*  **Eficiência:** Redução estimada de **80% no tempo** de tratamento de dados comparado ao método manual anterior.
*  **Padronização:** Implementação de uma tabela oficial de "De/Para" para códigos de salas, eliminando ambiguidades.

---

##  Ferramentas

* **Microsoft Excel:** Interface principal e armazenamento de dados.
* **Power Query (Linguagem M):** Para limpeza, transformação e modelagem dos dados brutos.
* **VBA (Visual Basic for Applications):** Para automação da formatação do relatório final e alertas visuais.
* **Fórmulas Avançadas:** `XLOOKUP`, `INDEX/MATCH` e lógica condicional para validação cruzada.

---
