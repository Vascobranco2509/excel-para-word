---
name: excel-para-word
description: Transforma um ficheiro Excel (.xlsx) num relatório Word (.docx) com gráficos e uma descrição factual de cada um. Usar quando alguém pede um relatório, gráficos ou uma análise a partir de uma folha de cálculo, ou fala em Excel, xlsx, folha de cálculo, relatório em Word ou docx.
---

# Excel para Word

Recebe um `.xlsx` e instruções em linguagem natural; devolve um `.docx` com um
gráfico por página, cada um com título, descrição com números reais e,
opcionalmente, a tabela de dados.

## Regras que não se negoceiam

1. **Nunca alterar o Excel de origem.** Só se lê.
2. **Nunca inventar dados.** Nenhum número vai para o relatório sem sair de uma
   conta feita sobre o ficheiro. Não estimar valores a olho a partir do gráfico.
3. **Nunca gerar o relatório por cima de avisos sem o utilizador dizer que sim.**
   O script bloqueia sozinho; não contornar com `--avisos-aceites` por iniciativa
   própria.
4. **Nunca perguntar o que está dentro do ficheiro.** Colunas, folhas e número de
   linhas leem-se, não se perguntam.
5. **Nunca escrever recomendações, causas, previsões ou juízos de valor** no
   relatório. Só a descrição do que o gráfico mostra.
6. **Nunca dar o trabalho por terminado sem ter verificado o `.docx` gerado.**

## Fluxo de trabalho

### 1. Inspecionar

Ler o ficheiro antes de falar sobre ele:

```bash
python -c "import pandas as pd; x=pd.ExcelFile('DADOS.xlsx'); print(x.sheet_names); [print(f, x.parse(f).dtypes, sep='\n') for f in x.sheet_names]"
```

### 2. Confirmar o pedido

Poucas perguntas de cada vez, e só as que mudam o resultado: que gráficos, de que
colunas, agrupados como. Nunca perguntar sobre cores, tipos de letra ou tamanhos.

### 3. Escrever o `plano.json`

O utilizador não escreve o plano — o Claude escreve-o a partir da conversa.
A referência dos campos está no [README](README.md#campos-do-planojson).
Não existem outros campos além dos que lá estão; qualquer campo a mais é recusado
pelo script.

O campo `nota` é opcional e **não pode ter números**: serve para enquadrar
(«período da campanha de verão»). Os números do relatório são sempre calculados
a partir dos dados.

### 4. Verificar — passo obrigatório

```bash
python scripts/gerar_relatorio.py --dados DADOS.xlsx --plano plano.json --verificar
```

Se houver avisos, **traduzi-los para linguagem simples e mostrá-los ao
utilizador**. Depois **esperar pela resposta dele**. Exemplos de tradução:

| O que o script diz | O que dizer ao utilizador |
|---|---|
| linhas ignoradas por células vazias | «Há N linhas sem valor preenchido. Vou deixá-las de fora — pode ser?» |
| coluna guardada como texto | «A coluna X está guardada como texto no Excel. Consegui ler os números, mas convém confirmar.» |
| colunas sem cabeçalho | «A folha tem células soltas ao lado da tabela. Vou ignorá-las.» |
| circular com muitas fatias | «Este gráfico circular ficaria com N fatias e não se leria. Sugiro barras.» |

### 5. Gerar

Só depois do «sim»:

```bash
python scripts/gerar_relatorio.py --dados DADOS.xlsx --plano plano.json --saida relatorio.docx
```

Se houver avisos já aceites pelo utilizador, acrescentar `--avisos-aceites`.

### 6. Verificar o resultado

Confirmar que o documento saiu como devia, antes de o entregar:

```bash
soffice --headless --convert-to pdf relatorio.docx
pdftoppm -jpeg -r 100 relatorio.pdf pagina
```

E olhar para as imagens das páginas. Sem LibreOffice instalado, extrair as
imagens de dentro do `.docx` (que é um ZIP, com os gráficos em `word/media/`) e
olhar para elas.

## Erros mais comuns e o que significam

| Mensagem | Causa |
|---|---|
| «Só são suportados ficheiros .xlsx» | O ficheiro é `.xls`, `.xlsm` ou `.csv`. Gravar como `.xlsx`. |
| «está aberto noutro programa» | O Excel tem o ficheiro aberto. Fechar. |
| «não tem números» | A coluna escolhida para `eixo_y` é de texto. |
| «tem campos que não existem» | O plano usa campos de uma versão antiga. Ver o README. |
| «há valores negativos» | Um gráfico circular não pode ter fatias negativas. Usar barras. |
