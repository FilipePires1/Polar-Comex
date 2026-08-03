# POLAR Comex — Inteligência de Importação

O POLAR Comex é um mini sistema de Business Intelligence voltado à análise de operações de importação. A ferramenta permite importar planilhas Excel com dados de comércio exterior e transformar essas informações em indicadores, rankings e comparativos de preço. EM FASE DE TESTES, PRIMEIRA VERSÃO!!

## Funcionalidades

- Importação de arquivos `.xlsx` e `.xls`
- Leitura de múltiplas planilhas e produtos/NCMs
- Armazenamento local dos dados no navegador
- Filtros por ano, período, fornecedor, NCM e país de origem
- Resumo executivo das compras da empresa analisada
- Comparação entre o preço pago e o mercado
- Gráfico de evolução mensal de preços em US$/kg
- Ranking de fornecedores conforme a competitividade de preço
- Cálculo de economia potencial ao comprar pelo melhor preço identificado
- Detalhamento de cada compra, incluindo volume, FOB, CIF, frete e seguro
- Anonimização de concorrentes, com opção para exibir nomes reais

## Como funciona

1. Importe uma ou mais planilhas Excel com os dados de importação.
2. Informe o nome da empresa que será analisada.
3. Use os filtros para selecionar o período, fornecedor, produto ou país desejado.
4. Analise os indicadores e compare os preços da empresa com os praticados pelo mercado.

## Formato esperado da planilha

Cada aba representa um produto ou NCM. O sistema espera:

- Linha 1: metadados
- Linha 2: cabeçalhos das colunas
- Linha 3 em diante: dados das operações

Colunas principais esperadas:

- `ANO/MÊS`
- `PAIS DE ORIGEM`
- `NCM`
- `PROVÁVEL EXPORTADOR`
- `PROVÁVEL IMPORTADOR`
- `Valor FOB Estimado Unitário`

## Tecnologias utilizadas

- HTML5
- CSS3
- JavaScript
- IndexedDB para armazenamento local
- SheetJS (XLSX) para leitura de planilhas
- Chart.js para gráficos

## Privacidade

Os dados importados são armazenados apenas no navegador do usuário, usando IndexedDB. Nenhuma informação é enviada a servidores externos.

## Objetivo

Apoiar decisões de compras internacionais ao tornar a comparação de preços de importação mais simples, visual e estratégica.
