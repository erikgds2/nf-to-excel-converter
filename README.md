Uso
Clone ou baixe este repositório para o seu sistema.

Coloque seus arquivos XML de notas fiscais na pasta XML dentro do diretório do projeto.

Execute o script main.py para processar os arquivos XML e gerar um arquivo Excel com as informações.

O script verifica se já existe um arquivo Excel para a data atual. Se existir, ele adicionará as informações ao arquivo existente. Caso contrário, criará um novo arquivo Excel.

Configuração
No início do arquivo main.py, você pode ajustar as seguintes configurações:

diretorio_xml: Diretório onde os arquivos XML de notas fiscais estão localizados.
diretorio_excel: Diretório onde serão salvos os arquivos Excel gerados.
Contribuição
Se você quiser contribuir com melhorias ou correções para este projeto, sinta-se à vontade para abrir um pull request ou reportar problemas na seção de Issues.


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
#   n f - t o - e x c e l - c o n v e r t e r 
 
 
