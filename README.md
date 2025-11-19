# 📊 Macro VBA — Importação Automática de Dados RREO

<div align="center">

![VBA](https://img.shields.io/badge/VBA-Macro%2520Excel-yellow?style=for-the-badge&logo=microsoft-excel)
![License](https://img.shields.io/badge/License-P%25C3%25BAblica-blue?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Em%2520Produ%C3%A7%C3%A3o-green?style=for-the-badge)

</div>

---

## 📋 Sobre o Projeto

Este projeto contém uma **macro VBA** desenvolvida para automatizar a importação de dados da planilha de origem `planilha_auditoria.xls` para a base oficial `SICONFI_RREO_XXXX_BASE.xls`.

A ferramenta foi desenhada para preencher **somente as células vazias** nos anexos do RREO (Relatório Resumido da Execução Orçamentária), garantindo a integridade dos dados pré-existentes.

### 🎯 Objetivo Principal
> A automação evita **sobrescritas indevidas**, reduz **erros manuais** e acelera significativamente o processo de conferência e consolidação dos dados contábeis.

---

## ✨ Funcionalidades Principais

| Funcionalidade | Descrição |
| :--- | :--- |
| **🔄 Cópia Segura** | Copia dados apenas para células vazias, evitando sobrescrever valores já preenchidos. |
| **📑 Multi-Anexos** | Compatível com múltiplos anexos RREO (01, 02, 03, 04, 06, 07, 13, 14). |
| **🧠 Código Flexível** | Lógica expansível com intervalos configurados em bloco único. |
| **📌 Modo Invisível** | Funciona via VBScript/Background sem a necessidade de interação visual constante. |
| **⚡ Alta Performance** | Processamento otimizado para grandes volumes de dados. |

---

## 🏗️ Arquitetura da Solução

### 🔄 Fluxo de Execução

```mermaid
graph TD
    A[📂 Início] --> B[Abre Arquivos Origem/Destino]
    B --> C{🔍 Varredura dos Anexos}
    C --> D[Verifica Célula Destino]
    D -- Célula Vazia? --> E[✅ Copia Dado]
    D -- Célula Cheia? --> F[🚫 Pula (Não Sobrescreve)]
    E & F --> G{Mais Células?}
    G -- Sim --> D
    G -- Não --> H[📝 Log e Debug]
    H --> I[💾 Salva e Fecha]
