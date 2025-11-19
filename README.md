📊 Macro VBA — Importação Automática de Dados para o RREO
Copia apenas células vazias dos anexos RREO a partir da planilha de auditoria

Este projeto contém uma macro VBA desenvolvida para automatizar a importação de dados da planilha planilha_auditoria.xls para a base oficial SICONFI_RREO_XXXX_BASE.xls, preenchendo somente as células vazias dos anexos RREO (Relatório Resumido da Execução Orçamentária).

A automação evita sobrescritas indevidas, reduz erros manuais e acelera significativamente o processo de conferência e consolidação dos dados.

🚀 Funcionalidades Principais
🔄 Copia dados apenas para células vazias
Evita sobrescrever valores já preenchidos no arquivo oficial do RREO.

📑 Compatível com múltiplos anexos RREO
Inclui intervalos específicos de linhas e colunas para:

RREO-Anexo 01
RREO-Anexo 02
RREO-Anexo 03
RREO-Anexo 04
RREO-Anexo 06
RREO-Anexo 07
RREO-Anexo 13
RREO-Anexo 14

🧠 Lógica flexível e expansível
Os intervalos de cada anexo são configurados em um único bloco, facilitando manutenção.

📌 Funciona mesmo quando usado via VBScript (modo invisível)
Pode rodar sem abrir o Excel visualmente.

📂 Estrutura Geral da Macro
A macro:

1.Abre o arquivo de origem (planilha_auditoria.xls)
2.Abre o arquivo de destino (SICONFI_RREO...BASE.xls)
3.Varre cada anexo configurado
4.Copia dados somente se a célula destino estiver vazia
5.Registra erros básicos no Debug
6.Salva e fecha o arquivo base

🛠️ Como usar
1.Abra o Excel
2.Pressione ALT + F11
3.Insira um novo módulo
4.Cole o conteúdo da macro
5.Ajuste o nome do arquivo base caso necessário
6.Execute Importar_RREO

⚠ Não inclua caminhos completos — a macro assume que os arquivos estão na mesma pasta onde ela está sendo executada.

📌 Observações Importantes
A macro não sobrescreve células preenchidas
Necessário habilitar macros no Excel
Arquivos devem estar na mesma pasta que a macro, conforme solicitado
Projetada para arquivos .xls e .xlsx

🧩 Melhorias Futuras (sugestões)
Registro de logs em arquivo .txt
Mensagens amigáveis ao usuário
Interface simples com UserForm
Otimização para milhares de células

📄 Licença
Este projeto pode ser reutilizado livremente dentro de órgãos públicos, controladorias, secretarias municipais, etc.
