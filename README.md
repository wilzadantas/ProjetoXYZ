# 📊 Projeto VB6 com Integração ao SQL Server

Este projeto, desenvolvido em Visual Basic 6 (VB6), permite o cadastro e gerenciamento de transações de cartão de crédito, além de oferecer exportação de dados para Excel.

### 📌 Referências Necessárias
Para o correto funcionamento do sistema, adicione as seguintes referências ao projeto:
- ADO (ActiveX Data Objects) – Para conexão com o SQL Server.
- MSCOMCTL.OCX – Para controles avançados de interface.
- Excel Object Library – Para exportação de dados para Excel.

### 🛠 Componentes Necessários
Inclua os seguintes componentes no projeto:
- Microsoft Common Controls – Para botões, listas e barras de progresso.
- Microsoft DataGrid Control – Para exibição de dados do SQL Server.
- Microsoft FlexGrid Control – Para tabelas interativas.
- Microsoft Common Dialog Control 6.0 – Para diálogos comuns do sistema.

### ⚙️ Configuração do Banco de Dados 
Para conectar o sistema ao SQL Server, siga os passos abaixo:

1. No formulário principal do projeto, edite a função responsável pela conexão, ajustando os seguintes parâmetros conforme necessário:
   - **Nome do servidor**: Insira o nome ou endereço do servidor onde o SQL Server está instalado.
   - **Nome do banco de dados**: Especifique o banco de dados que será utilizado na aplicação.
   - **Usuário e senha de acesso**: Ajuste as credenciais de autenticação para acessar o banco de dados.

2. Antes de rodar a aplicação, **certifique-se de realizar os seguintes passos adicionais**:
   - Localize o arquivo de backup do banco de dados (**`XYZ.bak`**), que está disponível no repositório do projeto.
   - Restaure o arquivo de backup (**attach**) no SQL Server para garantir que o banco de dados está configurado corretamente.

💡 **Observação Importante**: A configuração do usuário e senha é crucial para evitar erros de autenticação. Além disso, certifique-se de que o arquivo de backup foi restaurado corretamente antes de prosseguir com a execução do sistema.
