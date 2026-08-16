---
name: excel-para-word
description: Transforma um ficheiro Excel ou CSV (.xlsx, .xlsm, .csv) num relatório Word (.docx) com gráficos e uma análise estatística de cada um — dispersão, concentração, valores atípicos, tendência, sazonalidade, comparação ano a ano e previsão dos períodos seguintes. Usar quando alguém pede um relatório, gráficos, uma análise, uma evolução, uma previsão ou uma comparação entre anos a partir de uma folha de cálculo, ou fala em Excel, xlsx, folha de cálculo, relatório em Word ou docx.
---

# Excel para Word

Recebe um `.xlsx`, `.xlsm` ou `.csv` e instruções em linguagem natural; devolve um
`.docx` com um gráfico por página, cada um com título, descrição com números
reais, um bloco de análise estatística e, opcionalmente, a tabela de dados.

O bloco **Análise** traz âmbito, amplitude, dispersão, concentração, valores
atípicos e — quando o eixo é uma linha do tempo — evolução, tendência por
regressão, crescimento médio e índice sazonal. Se os dados abrangerem vários
anos, traz também os totais ano a ano e a variação interanual.

## Regras que não se negoceiam

1. **Nunca alterar o Excel de origem.** Só se lê.
2. **Nunca inventar dados.** Nenhum número vai para o relatório sem sair de uma
   conta feita sobre o ficheiro. Não estimar valores a olho a partir do gráfico.
3. **Nunca gerar o relatório por cima de avisos sem o utilizador dizer que sim.**
   O script bloqueia sozinho; não contornar com `--avisos-aceites` por iniciativa
   própria.
4. **Nunca perguntar o que está dentro do ficheiro.** Colunas, folhas e número de
   linhas leem-se, não se perguntam.
5. **Nunca escrever recomendações nem apontar causas** no relatório. O ficheiro
   tem números, não tem o mundo: nunca se saberá se as vendas subiram pela
   campanha ou porque o concorrente fechou.
6. **Nunca prever por iniciativa própria.** A previsão só entra quando o
   utilizador a pede; sem o campo `previsao` não há previsão. E **nunca
   apresentar o valor central sem o intervalo** — um número solto lê-se como
   promessa.
7. **Nunca avaliar desempenho sem uma meta dada pelo utilizador.** Nada de
   «bom», «mau», «fraco» ou «preocupante» — essas palavras exigem uma referência
   que a folha de cálculo não contém. **Se o utilizador der uma meta**, usa-se o
   campo `meta` e a comparação passa a ser aritmética: quanto foi, contra quanto
   se queria, e a diferença. **Nunca inventar uma meta** nem sugerir uma «meta
   razoável»: o número tem de vir dele.
8. **Nenhuma palavra qualitativa sem uma regra calculada ao lado.** O relatório
   escreve «tendência crescente», «dispersão elevada», «ajuste moderado» — sempre
   com o número **e** o critério à frente. Quem lê pode discordar do critério;
   não há opinião com que discordar. Os limiares estão no
   [README](README.md#os-termos-qualitativos-e-a-regra-de-cada-um) e em constantes
   no topo do script. Não inventar limiares novos no meio do código.
9. **Nunca aplicar um método estatístico sem pontos que cheguem.** O script já
   trava sozinho e escreve no relatório que não o aplicou, e porquê. Não contornar.
10. **Nunca dar o trabalho por terminado sem ter verificado o `.docx` gerado.**

## Fluxo de trabalho

### 1. Inspecionar

Ler o ficheiro antes de falar sobre ele. Para Excel (`.xlsx`, `.xlsm`):

```bash
python -c "import pandas as pd; x=pd.ExcelFile('DADOS.xlsx'); print(x.sheet_names); [print(f, x.parse(f).dtypes, sep='\n') for f in x.sheet_names]"
```

Para CSV, dois passos. Primeiro as linhas em bruto, para ver o **separador** e as
linhas de metadados por cima da tabela:

```bash
python -c "print(open('DADOS.csv', encoding='utf-8', errors='replace').read()[:400])"
```

Depois a **forma do ficheiro**. Um CSV real tem dezenas de colunas e os 400
caracteres acima nem chegam ao fim do cabeçalho:

```bash
python -c "
l=open('DADOS.csv',encoding='utf-8',errors='replace').read().splitlines()
s=';' if l[0].count(';')>l[0].count(',') else ','
print('linhas:',len(l)-1,'colunas:',len(l[0].split(s)))
print('colunas:',l[0].split(s))
print('exemplo:',l[1].split(s))"
```

### 1b. Ver a forma dos dados antes de prometer o que quer que seja

Antes de escolher colunas, decidir **em que formato estão os dados**, porque isso
muda tudo:

| Forma | Como se reconhece | O que fazer |
|---|---|---|
| **Longo** | Uma coluna com o nome da categoria (`Canal`) e outra com os valores (`Valor`) | `serie: "Canal"` |
| **Largo** | Uma coluna por categoria (`vendas_norte`, `vendas_sul`, `vendas_centro`) | `serie: ["vendas_norte", "vendas_sul", ...]`, **sem `eixo_y`** |

O formato largo é muito comum em exportações e em folhas feitas à mão. **Não
tentar reorganizar o ficheiro**: a skill lê os dois, e alterar o original está
proibido pela regra 1.

### 2. Confirmar o pedido — e oferecer o que o ficheiro pede

Poucas perguntas de cada vez, e só as que mudam o resultado: que gráficos, de que
colunas, agrupados como. Nunca perguntar sobre cores, tipos de letra ou tamanhos.

**Depois de inspecionar, reparar no que dá escolha e perguntar.** O utilizador
não sabe o que a ferramenta consegue fazer; quem leu o ficheiro foi o Claude.
Estas quatro são para oferecer, e **nunca para decidir sozinho**:

| O que se vê no ficheiro | O que perguntar |
|---|---|
| Mais períodos do que os que ele mencionou (três anos, e ele falou de um) | «Tem dados de 2023, 2024 e 2025. Queres o relatório dos três anos, ou só de um?» → `filtro` |
| Categorias com nomes compridos (mais de ~20 letras) | «Os nomes das campanhas são compridos e não cabem por baixo das barras. Queres barras deitadas?» → `barras_horizontais` |
| Duas colunas de números que façam sentido juntas (investimento e conversões, cliques e vendas) | «Queres um gráfico a cruzar as duas, para ver se andam juntas?» → `dispersao` |
| Uma linha do tempo com várias séries | «Queres ver o peso de cada uma no total ao longo do tempo?» → `area` |

**Nunca filtrar por iniciativa própria.** Filtrar é esconder dados: só entra se
ele pedir. Quando entra, o relatório escreve sozinho o que ficou de fora.

**Nunca prometer o que a dispersão não dá.** Ela mostra se duas colunas andam
juntas — não mostra que uma causa a outra. Se ele perguntar «então investir mais
traz mais vendas?», a resposta honesta é que os dados mostram a relação e não a
explicam.

### 3. Escrever o `plano.json`

O utilizador não escreve o plano — o Claude escreve-o a partir da conversa.
A referência dos campos está no [README](README.md#campos-do-planojson).
Não existem outros campos além dos que lá estão; qualquer campo a mais é recusado
pelo script.

O campo `nota` é opcional e **não pode ter números**: serve para enquadrar
(«período da campanha de verão»). Os números do relatório são sempre calculados
a partir dos dados.

Três campos decidem a profundidade da análise, e vale a pena pensar neles antes
de escrever o plano:

- **`coluna_periodo`** — a coluna com as datas ou os anos. Sem ela **não há
  quebra ano a ano**. Se o ficheiro abranger mais do que um ano, indicá-la.
- **`eixo_temporal`** — só é preciso quando os rótulos são invulgares
  (`Semana 1`, `P1`). Meses em português, trimestres, anos e datas são
  reconhecidos sozinhos.
- **`analise`**, ao nível do plano — `completa` por omissão. `curta` só se o
  utilizador disser que quer apenas o gráfico. `detalhada` repete a análise
  inteira para **cada** série: com três canais, o relatório fica cerca de duas
  vezes e meia mais longo, por isso só a pedido.
- **`serie`** — sempre que o utilizador falar em **comparar** canais, campanhas,
  regiões ou lojas, é isto que ele quer: um gráfico com várias linhas, não vários
  gráficos. Três gráficos com escalas diferentes não se comparam a olho.
  Aceita duas formas, conforme a forma dos dados (ver o passo 1b):
  `"serie": "Canal"` quando há uma coluna com o nome da categoria;
  `"serie": ["vendas_norte", "vendas_sul"]` quando cada categoria tem a sua
  própria coluna — e nesse caso **não se dá `eixo_y`**, porque cada coluna já é
  os valores de uma série.
- **`meta`** — `{"valor": N, "ambito": "total"}`, `"categoria"` ou `"serie"`. Se o utilizador
  falar em objetivo, meta ou target, perguntar-lhe **qual dos dois âmbitos**:
  «700.000 no trimestre» é `total`; «cada canal tem de trazer 50.000» é
  `categoria`. Adivinhar aqui dá um relatório errado com ar de certo.
- **`previsao`** — número de períodos a prever. **Só quando o utilizador pedir.**
  Se ele pedir uma previsão, vale a pena avisá-lo do que ela pressupõe: que a
  tendência e a sazonalidade se mantêm, e que nada do que está fora do ficheiro
  entra na conta.

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
| parece ser linha de totais | «A folha tem uma linha TOTAL no fim. Se ficar, o gráfico conta os valores duas vezes. Quer que apague essa linha?» |
| não começa na primeira linha | «A tabela começa na linha N, por causa do título por cima. Usei essa — está certo?» |
| números gravados como texto | «Os valores estão como texto no Excel, com o símbolo de euro. Consegui lê-los, mas convém confirmar dois ou três.» |
| parece um valor acumulado | «Esta coluna nunca desce — parece um acumulado. Se for, somá-la dá um número sem sentido: o total daria X quando o valor real é Y. Quer que use antes o último período?» |
| séries no mesmo gráfico | «Seriam N linhas no mesmo gráfico e iam confundir-se. Quer que fique só com as principais?» |
| categorias … ilegíveis | «O gráfico ficaria com N barras, finas demais para se lerem. Quer agrupar, ou usar outra coluna?» |
| a tabela … ficariam de fora | «A tabela só mostra as primeiras 30 de N categorias. Quer assim, ou sem tabela?» |
| não é possível saber (datas) | «As datas estão em texto e nenhum dia passa de 12, portanto não dá para saber se é dia/mês ou mês/dia. Vou tratá-las como texto — sem tendência nem sazonalidade.» |
| Previsão não calculada | «Não dá para prever com confiança: [a razão que o script deu]. Prefiro dizer isso a inventar um número.» |

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

**O `.docx` de saída não pode estar aberto no Word** enquanto se gera. Se estiver,
o script diz isso e não escreve nada.

Ao ler a análise, confirmar que **nenhum número foi inventado** e que cada termo
qualitativo aparece com o critério ao lado. Se algum método disser que não correu,
está certo — é a guarda a funcionar, não uma falha.

## Erros mais comuns e o que significam

| Mensagem | Causa |
|---|---|
| «Não sei ler ficheiros …» | Suportados: `.xlsx`, `.xlsm`, `.csv`. O `.xls` antigo tem de ser gravado como `.xlsx`. |
| «não diz de que folha ler» | O ficheiro tem várias folhas e o plano não indica `folha`. |
| «mistura formatos de data incompatíveis» | Umas datas parecem dia/mês e outras mês/dia. |
| «Não consigo ler … está aberto noutro programa» | O Excel de origem está aberto. Fechar. |
| «não tem números» | A coluna escolhida para `eixo_y` é de texto. |
| «tem campos que não existem» | O plano usa campos de uma versão antiga. Ver o README. |
| «há valores negativos» | Um gráfico circular não pode ter fatias negativas. Usar barras. |
| «Não consigo escrever … está aberto noutro programa» | O `.docx` de saída está aberto no Word. Fechar. |
| «formato ambíguo» | Números como texto em `1.250`, sem nada a provar se é mil duzentos e cinquenta ou um vírgula vinte e cinco. Formatar a coluna como número no Excel. |
| «parece(m) ser linha(s) de totais» | A folha tem uma linha `TOTAL` no fim, que contaria os valores duas vezes. |
| «não começa na primeira linha» | A folha tem título por cima da tabela; confirmar a linha escolhida ou indicar `linha_cabecalho`. |
