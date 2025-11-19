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
```

1. **📂 Abertura**: Abre arquivo de origem e destino.
2. **🔍 Varredura**: Percorre cada anexo configurado.
3. **✅ Validação**: Verifica se a célula de destino está vazia.
4. **📤 Cópia**: Transfere dados apenas se a validação for positiva.
5. **📝 Log**: Registra eventuais erros no Debug do VBA.
6. **💾 Finalização**: Salva e fecha o arquivo base.

---

## 🚀 Como Usar

### 📥 Instalação Rápida

1. Abra o Excel.
2. Pressione `ALT + F11` para abrir o Editor VBA.
3. Insira um novo módulo (`Inserir > Módulo`).
4. Cole o código da macro.
5. Execute a subrotina: `Importar_RREO`.

### ⚙️ Configuração

No início do código, certifique-se de ajustar as constantes conforme o nome do seu arquivo:

```vba
' Ajuste o nome do arquivo base se necessário
Const ARQUIVO_BASE As String = "SICONFI_RREO_XXXX_BASE.xls"
```

### 🗂️ Estrutura de Arquivos

Para o funcionamento correto, mantenha a seguinte estrutura de diretórios:

```text
📁 Pasta do Projeto/
├── 📊 planilha_auditoria.xls        <-- Origem dos dados
├── 🎯 SICONFI_RREO_XXXX_BASE.xls    <-- Destino (Oficial)
├── 🛠️ macro_rreo.vba                <-- Código Fonte
└── 📖 README.md
```

---

## 📋 Anexos Suportados

A ferramenta cobre os seguintes demonstrativos do RREO:

| Anexo | Descrição | Status |
| :--- | :--- | :---: |
| **RREO-Anexo 01** | Demonstrações Contábeis | ✅ |
| **RREO-Anexo 02** | Receita Orçamentária | ✅ |
| **RREO-Anexo 03** | Despesa Orçamentária | ✅ |
| **RREO-Anexo 04** | Receitas e Despesas | ✅ |
| **RREO-Anexo 06** | Restos a Pagar | ✅ |
| **RREO-Anexo 07** | Dívida Consolidada | ✅ |
| **RREO-Anexo 13** | Operações de Crédito | ✅ |
| **RREO-Anexo 14** | Garantias | ✅ |

---

## ⚠️ Requisitos e Observações

### 🔧 Pré-requisitos
* ✅ Microsoft Excel (2010 ou superior)
* ✅ Macros habilitadas nas configurações de segurança
* ✅ Arquivos (origem e destino) na mesma pasta
* ✅ Permissões de escrita no diretório

### 📌 Observações Importantes
* **Caminhos:** A macro utiliza `ThisWorkbook.Path`, portanto, não inclua caminhos absolutos (ex: `C:\Users...`). Apenas garanta que os arquivos estejam juntos.
* **Proteção:** A lógica principal é **não destrutiva**. Se houver um valor na célula de destino, ele será preservado.
* **Formatos:** Funciona tanto com `.xls` (Excel 97-2003) quanto `.xlsx`.

---

## 🛠️ Melhorias Futuras (Roadmap)

| Melhoria | Status | Prioridade |
| :--- | :---: | :---: |
| 📝 Registro de logs em arquivo .txt externo | 🟡 Planejado | Alta |
| ⚡ Otimização de array para milhares de células | 🟡 Planejado | Alta |
| 👨‍💻 Mensagens de erro mais amigáveis (MsgBox) | 🟡 Planejado | Média |
| 🎨 Interface gráfica simples com UserForm | 🔴 Futuro | Baixa |
| 🔄 Controle de versões automático | 🔴 Futuro | Baixa |

---

## 📄 Licença e Termos de Uso

**Licença Pública**

Este projeto pode ser **reutilizado livremente** dentro de:
* 🏢 Órgãos públicos federais, estaduais e municipais.
* 👁️ Controladorias e tribunais de contas.
* 🏛️ Secretarias de Fazenda e Planejamento.
* 📊 Departamentos de auditoria interna.

---

## 🤝 Suporte

Para questões sobre implementação, bugs ou customização para novos anexos, entre em contato com a equipe de desenvolvimento responsável.
