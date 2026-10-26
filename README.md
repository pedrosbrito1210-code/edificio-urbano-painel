# Painel de prestação de contas de condomínio em obra

Painel web que mostra, para os proprietários de um edifício em construção, quanto cada um aportou, quanto deve pela sua fração das despesas, o extrato completo da conta, o gasto por fornecedor e as despesas projetadas.

A fonte dos dados é a planilha de prestação de contas que já era mantida em Excel. Um script em Python lê as tabelas da planilha, refaz todos os cálculos e gera uma página HTML estática, que é publicada no Netlify. Os proprietários acessam sempre o mesmo link e veem a versão mais recente.

Este repositório roda com uma planilha de exemplo, com nomes e valores fictícios. Os dados reais ficam fora do Git.

## Como funciona

```
Planilha Excel  ->  gerar_painel.py  ->  site/index.html  ->  Netlify
 (tabelas)          (lê e calcula)       (template + dados)    (link fixo)
```

- `template.html`: layout, abas, gráficos (Chart.js) e toda a lógica de exibição. Os dados entram em marcadores `/*DADOS:...*/`.
- `gerar_painel.py`: lê as tabelas nomeadas da aba "Extrato" (proprietários, extrato bancário, caixinha, rateio especial e projeções), recalcula saldo, parte de cada proprietário e valores pagos, e injeta tudo no template.
- `atualizar.py`: usado no computador do escritório. Baixa a versão mais nova do código, gera o painel com a planilha real e publica no Netlify pela API, só quando algo mudou.
- `exemplo/`: planilha fictícia e o script que a cria.

## Cuidados com os números

- As colunas são localizadas pelo nome do cabeçalho, então mudar a ordem não quebra a leitura.
- Se uma tabela ou coluna sumir, se um proprietário aparecer no extrato com nome diferente do cadastro, ou se os percentuais não somarem 100%, o script para e não gera nada.
- Quando a planilha foi salva pelo Excel, o script compara o que calculou com os valores das fórmulas da própria planilha. Qualquer diferença acima de cinco centavos interrompe a publicação.

## Rodando localmente

```bash
pip install -r requirements.txt
python gerar_painel.py --excel exemplo/dados_exemplo.xlsx
# abre site/index.html no navegador
```

## Publicação do painel real

1. Copie `config_local.exemplo.json` para `config_local.json` e preencha o caminho da planilha, o Site ID e um token do Netlify.
2. Depois de atualizar a planilha, rode `atualizar_painel.bat` (ou `python atualizar.py`).

`config_local.json`, a pasta `site/` e qualquer arquivo Excel além do de exemplo estão no `.gitignore`.

## Tecnologias

Python (openpyxl), HTML, CSS, JavaScript, Chart.js, Netlify.
