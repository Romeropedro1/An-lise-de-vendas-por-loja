# Análise de Vendas por Loja

## Descrição

Este projeto tem como objetivo realizar uma análise de vendas de uma base de dados em formato Excel, sumarizando informações como faturamento, quantidade de produtos vendidos e o ticket médio por loja. Além disso, o sistema envia automaticamente um e-mail com um relatório detalhado sobre as vendas para um destinatário especificado.

O código utiliza as bibliotecas **Pandas** para manipulação de dados, **win32com** para interação com o Outlook e **Excel** para carregar a base de dados.

## Tecnologias Utilizadas

- **Python**: Linguagem utilizada para análise dos dados e envio do e-mail.
- **Pandas**: Biblioteca para análise e manipulação de dados.
- **win32com**: Biblioteca para automação de envio de e-mails via Outlook.
- **Excel**: Formato de entrada de dados (arquivo `Vendas.xlsx`).

## Funcionalidades

- **Leitura e visualização da base de dados**: Importação de dados de vendas a partir de um arquivo Excel (`Vendas.xlsx`).
- **Faturamento por loja**: Cálculo do faturamento total de cada loja.
- **Quantidade de produtos vendidos**: Cálculo da quantidade de produtos vendidos por loja.
- **Cálculo do ticket médio**: Cálculo do ticket médio dos produtos vendidos em cada loja.
- **Envio de relatório por e-mail**: Envio automático do relatório de vendas por e-mail com o Outlook.

## Como Rodar o Projeto

### Pré-requisitos

Antes de executar o script, você precisará:

- **Python**: Certifique-se de ter o Python instalado em seu sistema. Caso contrário, baixe e instale a versão mais recente de [python.org](https://www.python.org/downloads/).
- **Bibliotecas necessárias**: O projeto utiliza as bibliotecas **pandas** e **win32com**. Você pode instalar essas dependências com o seguinte comando:

   ```bash
   pip install pandas pywin32

