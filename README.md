# Sistema Corretagem

- Este é um sistema básico, desenvolvido em VB6 para cadastro de clientes e corretores.
- O projeto utiliza como banco de dados o SQL Server e referências ao ADODB para construção dos Recordsets

## Como instalar o projeto
- Primeiramente, você vai precisar ter o serviço do Sql Server instalado em sua máquina
- Será preciso executar os scripts que estão no arquivo ScriptBanco.sql deste repositório, via prompt de comando ou pelo Sql Server Management Studio
  - O script engloba a criação das schemas e inserção de alguns dados "mockeados"
 
## Como configurar o sistema:
- Após configurada a base de dados, execute o SistemaCorretagem.exe
- Na tela de boas vindas, informe o host, usuário e senha do servidor do Sql Server (por default ele traz localhost root root, porém é alterável).
- Clique em acessar.
- Caso o sistema apresente algum erro de conexão, certifique-se de que suas informações coincidam com as do servidor corretamente.

### Como utilizar o sistema:
- O cadastro de clientes serve tanto para pesquisa quanto para editar um cliente já existente. Sendo assim, para utilizar a pesquisa, basta passar os filtros que gostaria de utilizar.
- Caso queira editar um cliente, clique sobre o nome do cliente em questão na tabela, altere os dados que precisar, e por fim clique em salvar.
- Caso queira excluir um cliente, o procedimento é o mesmo. Basta clicar sobre o nome do cliente, e depois em excluir.
