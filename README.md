# Avaliação de Preços em MarketPlace
## Automação feita com o propósito de avaliar preços em marketplace
Projeto de processo seletivo realizado em python. Este projeto tem como objetivo gerar um excel com nomes e valores de produtos encontrados no marketplace "Americanas", e comparar seus preços gerando média.

A automação recebe uma lista em excel com nomes de produtos e envia por e-mail o relatório resultante da pesquisa.

#### Como instalar as dependências e executar o projeto:

No terminal, navegue até o diretório do projeto.

Instale as dependências necessárias que estão no arquivo “requirements.txt” através da linha de comando: 
```pip install –r requirements.txt```

Também no terminal, execute o arquivo “main.py” através da linha de comando: 
```python3 main.py``` ou ```python main.py```

Um e-mail será enviado ao endereço especificado no arquivo "email.ini" com o assunto “Relatório de Avaliação de Preços” e em anexo a planilha de Excel gerada no projeto. Um log também será gerado e implementado em cada execução do projeto.

Obs: É necessário a configuração do arquivo "email.sample.ini" antes da execução do projeto, e renomeie-o para "email.ini".
    Também é necessário observar que, se for a primeira execução, pode demorar alguns segundos para a abertura do navegador da automação.