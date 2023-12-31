# Notas Fiscais XML to Excel Converter

Este é um projeto em Python para converter informações de arquivos XML de notas fiscais em um arquivo Excel formatado. O script percorre os arquivos XML no diretório especificado, extrai as informações desejadas e gera um arquivo Excel com as informações organizadas em colunas.

## Funcionalidades

- Extrai informações como número da nota fiscal, data de emissão, descrição, valor da nota fiscal e forma de pagamento dos arquivos XML de notas fiscais.
- Gera um arquivo Excel formatado com as informações extraídas.
- Verifica se já existe um arquivo Excel para a data atual e adiciona as informações ao arquivo existente, se aplicável.

## Requisitos

- Python 3.x
- Bibliotecas: `xmltodict`, `openpyxl`

Você pode instalar as bibliotecas necessárias usando o seguinte comando:

```sh
pip install xmltodict openpyxl
