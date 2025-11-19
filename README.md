📊 Macro VBA — Importação Automática de Dados RREO
https://img.shields.io/badge/VBA-Macro%2520Excel-yellow?style=for-the-badge&logo=microsoft-excel
https://img.shields.io/badge/License-P%25C3%25BAblica-blue?style=for-the-badge
https://img.shields.io/badge/Status-Em%2520Produ%C3%A7%C3%A3o-green?style=for-the-badge

📋 Sobre o Projeto
Este projeto contém uma macro VBA desenvolvida para automatizar a importação de dados da planilha planilha_auditoria.xls para a base oficial SICONFI_RREO_XXXX_BASE.xls, preenchendo somente as células vazias dos anexos RREO (Relatório Resumido da Execução Orçamentária).

🎯 Objetivo Principal
A automação evita sobrescritas indevidas, reduz erros manuais e acelera significativamente o processo de conferência e consolidação dos dados.

✨ Funcionalidades Principais
Funcionalidade	Descrição
🔄 Cópia Segura	Copia dados apenas para células vazias, evitando sobrescrever valores já preenchidos
📑 Multi-Anexos	Compatível com múltiplos anexos RREO (01, 02, 03, 04, 06, 07, 13, 14)
🧠 Código Flexível	Lógica expansível com intervalos configurados em bloco único
📌 Modo Invisível	Funciona via VBScript sem abrir o Excel visualmente
⚡ Alta Performance	Processamento otimizado para grandes volumes de dados
🏗️ Arquitetura da Solução
📂 Estrutura de Processamento








🔄 Fluxo de Execução
📂 Abertura - Abre arquivo de origem e destino

🔍 Varredura - Percorre cada anexo configurado

✅ Validação - Verifica se célula destino está vazia

📤 Cópia - Transfere dados apenas para células vazias

📝 Log - Registra eventuais erros no Debug

💾 Finalização - Salva e fecha o arquivo base

🚀 Como Usar
📥 Instalação Rápida
vba
' 1. Abra o Excel
' 2. Pressione ALT + F11
' 3. Insira um novo módulo
' 4. Cole o código da macro
' 5. Execute: Importar_RREO
⚙️ Configuração
vba
' Ajuste o nome do arquivo base se necessário
Const ARQUIVO_BASE As String = "SICONFI_RREO_XXXX_BASE.xls"
🗂️ Estrutura de Arquivos
text
📁 Pasta do Projeto/
├── 📊 planilha_auditoria.xls
├── 🎯 SICONFI_RREO_XXXX_BASE.xls
├── 🛠️ macro_rreo.vba
└── 📖 README.md
📋 Anexos Suportados
Anexo	Descrição	Status
RREO-Anexo 01	Demonstrações Contábeis	✅ Suportado
RREO-Anexo 02	Receita Orçamentária	✅ Suportado
RREO-Anexo 03	Despesa Orçamentária	✅ Suportado
RREO-Anexo 04	Receitas e Despesas	✅ Suportado
RREO-Anexo 06	Restos a Pagar	✅ Suportado
RREO-Anexo 07	Dívida Consolidada	✅ Suportado
RREO-Anexo 13	Operações de Crédito	✅ Suportado
RREO-Anexo 14	Garantias	✅ Suportado
⚠️ Requisitos e Observações
🔧 Pré-requisitos
✅ Microsoft Excel (2010 ou superior)

✅ Macros habilitadas

✅ Arquivos na mesma pasta

✅ Permissões de escrita

📌 Observações Importantes
⚠️ Atenção: Não inclua caminhos completos - a macro assume que os arquivos estão na mesma pasta de execução.

🔒 Proteção de Dados: Não sobrescreve células preenchidas

📁 Compatibilidade: Funciona com .xls e .xlsx

🏢 Público-Alvo: Órgãos públicos e controladorias
