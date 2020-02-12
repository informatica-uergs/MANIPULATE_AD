# AD_MANIPULATE
## Scripts de manipulação do AD da UERGS

# Requisitos para alterar os Scripts

- Biblioteca pyad
- Biblioteca pypiwin32
    
# Requisitos para executar os Scripts

- Possuir um arquivo CSV com os campos:
        "FUNCIONÁRIO,CARGO,EMPRESA,DEPARTAMENTO,RAMAL".
    - Obs: As informações do CSV devem estar na ordem descrita acima e separadas por vírgula.
- Possuir um arquivo de configuração com os campos:
        "NOME DO AD COMPLETO,USUÁRIO,SENHA".
    - Obs: Cada campo deve estar cada um em uma linha do arquivo.
    - Obs2: O nome do arquivo deve ser 'config' e salvo com o formato 'Documentos de Texto(.txt)'(opcional)
- Todos os arquivos devem estar no mesmo diretório que o arquivo executável.

# Funcionamento do Script

1. Executar o arquivo AD_USER_UPDATE.exe.
2. Digitar o nome do arquivo de configuração sem a extensão.
3. Digitar o nome do arquivo CSV sem a extensão.
    1. Obs: O arquivo CSV deve ser gerado antes da execução do script.
4. Esperar a conclusão da modificação no AD.