# excel-para-word

Transforma um ficheiro Excel num relatório Word com gráficos e uma descrição
factual de cada um. É uma [skill do Claude Code](https://docs.claude.com/en/docs/claude-code):
descreves em português o que queres, e o Claude lê a folha de cálculo, escreve o
plano, verifica os dados e gera o documento.

> **In English —** A Claude Code skill that turns an Excel file into a Word report.
> You describe the charts you want in plain language; it reads the spreadsheet,
> checks the data, draws the charts and writes a factual description of each one
> using numbers computed from the file itself. It refuses to generate the report
> when it finds problems in the data, unless you explicitly accept them.

![Primeira página do relatório de exemplo](exemplos/preview.png)

## O que faz, e o que não faz

**Faz:**

- Lê **`.xlsx`, `.xlsm` e CSV**. No CSV deteta sozinho o separador (`;` ou `,`) e a
  codificação. Nos ficheiros Excel, aceita várias folhas.
- Desenha gráficos de **barras**, **linhas** e **circular**.
- Põe **várias séries no mesmo gráfico** — canais, campanhas ou regiões lado a lado, com
  legenda, para se compararem de relance em vez de em três gráficos com escalas
  diferentes.
- Agrega por **soma**, **média** ou **contagem**.
- Escreve, por baixo de cada gráfico, uma **análise estatística completa**: dispersão,
  concentração, valores atípicos, tendência, crescimento médio, sazonalidade e quebra
  **ano a ano** com variação interanual. Tudo calculado a partir da mesma conta que
  desenhou o gráfico.
- **Prevê os períodos seguintes, se pedires** — sempre com intervalo de confiança, e
  só quando a série dá garantias para isso.
- Junta a tabela de dados, se pedires.
- **Recusa gerar o relatório** quando encontra problemas nos dados, até tu dizeres
  que os aceitas.

**Não faz, de propósito:**

- Não lê `.xls`, o formato antigo do Excel.
- Não escreve recomendações nem aponta causas.
  [A razão está aqui abaixo](#o-que-continua-de-fora-e-porquê).
- Não diz se um resultado é «bom» ou «fraco». Avalia desempenho **só contra uma meta
  que tu dês** — [a razão está aqui abaixo](#avaliação-face-a-uma-meta).
- Não inventa dados nem preenche células vazias em silêncio.
- Não altera o ficheiro Excel de origem.
- Não aplica um método estatístico quando não há pontos que cheguem — diz que não
  o aplicou, e porquê.

## Instalação

Precisas de saber abrir um terminal. Se nunca abriste, esta ferramenta ainda não
é para ti — o relatório de exemplo em `exemplos/` mostra o resultado sem instalar
nada.

```bash
git clone https://github.com/Vascobranco2509/excel-para-word.git
cd excel-para-word
pip install -r requirements.txt
```

Precisas de Python 3.10 ou mais recente.

Para usar como skill do Claude Code, copia a pasta para `~/.claude/skills/`.

## Exemplo

O repositório traz um exemplo completo, com três anos de dados mensais. Para o repetir:

```bash
python scripts/gerar_relatorio.py \
  --dados exemplos/vendas_2023_2025.xlsx \
  --plano exemplos/plano.json \
  --saida relatorio.docx
```

> **Os dados do exemplo são inventados.** `vendas_2023_2025.xlsx` foi gerado por uma
> fórmula, número a número, para a análise ter alguma coisa que dizer. Não representa
> vendas de ninguém: serve para mostrar o que a ferramenta faz, não para tirar conclusões.

## Comandos

```bash
# ver os problemas dos dados sem gerar nada
python scripts/gerar_relatorio.py --dados X.xlsx --plano plano.json --verificar

# gerar o relatório
python scripts/gerar_relatorio.py --dados X.xlsx --plano plano.json --saida r.docx

# gerar mesmo havendo avisos, depois de os teres lido
python scripts/gerar_relatorio.py --dados X.xlsx --plano plano.json --saida r.docx --avisos-aceites
```

Códigos de saída: `0` correu bem, `1` erro nos dados, `2` bloqueado por avisos.

## Campos do `plano.json`

Esta é a referência única. O `SKILL.md` aponta para aqui e não a repete.

```json
{
  "titulo_relatorio": "Vendas 2025",
  "graficos": [
    {
      "tipo": "barras",
      "folha": "Vendas",
      "eixo_x": "Mes",
      "eixo_y": "Valor",
      "agregacao": "soma",
      "titulo": "Valor vendido por mês",
      "nota": "campanha de fim de ano",
      "tabela_dados": true
    }
  ]
}
```

| Campo | Obrigatório | O que é |
|---|---|---|
| `titulo_relatorio` | sim | Título na primeira página do documento. |
| `graficos` | sim | Lista de gráficos. Um gráfico por página. |
| `tipo` | sim | `barras`, `linhas` ou `circular`. |
| `folha` | não | Nome da folha do Excel. Dispensável se o ficheiro tiver uma só folha, ou se for CSV. |
| `eixo_x` | sim | Coluna que dá as categorias (meses, canais, regiões…). |
| `eixo_y` | sim, exceto em `contagem` | Coluna com os números. |
| `agregacao` | sim | `soma`, `media` ou `contagem`. |
| `titulo` | sim | Título do gráfico, no Word. Não entra na imagem. |
| `nota` | não | Frase de enquadramento, **sem números**. Fica em itálico por baixo da descrição. |
| `tabela_dados` | não | `true` acrescenta a tabela por baixo do gráfico. Acima de 30 categorias mostra as primeiras 30 e diz quantas ficaram de fora. |
| `linha_cabecalho` | não | Número da linha do Excel onde estão os nomes das colunas. Só é preciso quando a deteção automática escolher a linha errada. |
| `coluna_periodo` | não | Coluna com as datas ou os anos, para a quebra ano a ano. Só é preciso quando o `eixo_x` é outra coisa (ex.: gráfico por canal, com a data noutra coluna). |
| `eixo_temporal` | não | `true`/`false` para forçar ou travar a deteção de linha do tempo. Só é preciso com rótulos invulgares (`Semana 1`, `P1`). |
| `previsao` | não | Número de períodos a prever. Sem este campo não há previsão. [Ver as guardas](#previsões). |
| `meta` | não | Objeto `{"valor": N, "ambito": "total"\|"categoria"\|"serie"}`. [Ver como funciona](#avaliação-face-a-uma-meta). |
| `serie` | não | Coluna cujos valores dão as séries do gráfico. [Ver como funciona](#várias-séries-no-mesmo-gráfico). |

E, ao nível do plano (fora da lista `graficos`):

| Campo | Obrigatório | O que é |
|---|---|---|
| `analise` | não | `completa` (por omissão), `curta` (só o gráfico e uma frase) ou `detalhada` (repete a análise inteira para cada série). |

Qualquer campo que não esteja nesta tabela é recusado com erro. É de propósito:
apanha planos escritos contra uma versão antiga.

### Duas decisões que valem a pena explicar

**A tabela não tem de começar na primeira linha.** Se a folha tiver o título do
relatório e a data de exportação por cima — como quase todas as folhas que saem
de um sistema — o cabeçalho é encontrado sozinho. Mas nunca em silêncio: fica
registado um aviso a dizer que linha foi usada, e o aviso bloqueia a geração até
tu confirmares. Se a escolha estiver errada, indicas a linha certa com
`linha_cabecalho`.

**Números guardados como texto são convertidos, mas só quando é seguro.**
`1.250,00 €` e `$1,250.00` são lidos corretamente, e a convenção é deduzida dos
valores da própria coluna: se lá estiver um `980,50`, fica provado que a vírgula
é decimal. Quando nada o prova — uma coluna só com `1.250` e `2.500` — para com
erro em vez de adivinhar, porque `1.250` tanto pode ser mil duzentos e cinquenta
como um vírgula vinte e cinco, e um palpite errado dá um relatório mil vezes ao
lado com ar perfeitamente normal.

**A ordem das categorias é a ordem do ficheiro, não a alfabética.** Se a tua
coluna tem os meses por ordem, o gráfico sai por ordem. Ordenar por alfabeto
poria Abril antes de Janeiro e daria uma série temporal errada.

**O campo `nota` não aceita números.** Os números do relatório saem sempre dos
dados. Se a nota pudesse ter números escritos à mão, bastava mudar o Excel e não
refazer o plano para o texto passar a desmentir o gráfico ao lado.

## Formatos de entrada

| Formato | Notas |
|---|---|
| `.xlsx` | O caso normal. Várias folhas; indica-se qual com `folha`. |
| `.xlsm` | Lido como o `.xlsx`. **As macros nunca são executadas** — só se leem os dados. |
| `.csv`, `.tsv`, `.txt` | Separador detetado sozinho (`;`, `,`, tabulação). Codificação tentada por UTF-8, depois cp1252, depois latin-1 — a que funcionou fica registada nas notas. |
| `.xls` | **Não suportado.** É o formato antigo do Excel e obrigaria a mais uma dependência. Abre-o e grava como `.xlsx`. |

Num CSV não há folhas: o campo `folha` é ignorado se estiver no plano. Num Excel com uma
só folha, também não é preciso indicá-la.

**Exportações de plataformas** (Meta Ads, Google Ads, Analytics) costumam trazer duas ou
três linhas de metadados antes da tabela. São detetadas e ignoradas, com aviso a dizer que
linha foi usada como cabeçalho.

### Datas em texto

Uma coluna de datas em texto só é tratada como linha do tempo quando a convenção estiver
**provada**:

| O que lá está | O que acontece |
|---|---|
| `2026-01-31` (ISO) | Inequívoco. Aceite sempre. |
| Algum dia acima de 12 (`31/01/2026`) | Prova que é dia/mês. Aplicado à coluna toda. |
| Algum mês acima de 12 (`01/31/2026`) | Prova que é mês/dia. Aplicado à coluna toda. |
| As duas coisas | Erro: a coluna mistura convenções. |
| Nada prova (só dias ≤ 12) | Fica como categoria, **sem análise temporal**, e o relatório diz porquê. |

O último caso é deliberado: falhar a análise é seguro, trocar Janeiro por Fevereiro não é.
Um índice sazonal com os meses trocados estaria errado com ar de certo.

## Várias séries no mesmo gráfico

O campo `serie` aceita duas formas, conforme a forma dos dados.

**Formato longo** — uma coluna com o nome da categoria e outra com os valores:

```json
{ "tipo": "linhas", "eixo_x": "Periodo", "eixo_y": "Valor",
  "serie": "Canal", "agregacao": "soma", "titulo": "Valor por canal" }
```

**Formato largo** — cada categoria na sua própria coluna, como em quase todas as
exportações e nas folhas feitas à mão. Aqui `serie` é a **lista das colunas**, e
**não se dá `eixo_y`**, porque cada coluna já é os valores de uma série:

```json
{ "tipo": "linhas", "eixo_x": "data",
  "serie": ["vendas_norte", "vendas_centro", "vendas_sul"],
  "agregacao": "soma", "titulo": "Vendas por região" }
```

As séries ficam com **cores, traços e marcas diferentes** — o gráfico continua a ler-se
impresso a preto e branco e por quem não distingue vermelho de verde.

**Uma série que não cubra todos os períodos** — um canal que só existiu a partir de certa
altura — fica com um buraco no gráfico, nunca com um zero. Zero seria um valor inventado.

**O conjunto não é a soma das séries.** A análise geral é feita reagregando os dados em
bruto, ignorando a coluna das séries. Com `agregacao: media`, somar as médias de cada
série daria um número errado.

**Guardas:** um gráfico circular com `serie` dá erro (um circular mostra a repartição de
um total, não a evolução de várias séries); acima de 6 séries há aviso, porque as linhas
começam a confundir-se; e usar a mesma coluna como eixo x e como série dá erro, porque
cada série ficaria com um ponto só.

## A análise

Por baixo de cada gráfico entra um bloco **Análise**, com secções etiquetadas:

| Secção | O que traz |
|---|---|
| Âmbito | Nº de categorias e de linhas, total, média, mediana |
| Amplitude | Máximo, mínimo, diferença e rácio entre eles |
| Dispersão | Desvio-padrão e coeficiente de variação |
| Concentração | Peso da maior categoria e das três maiores |
| Valores atípicos | Método de Tukey (1,5×IQR) |
| Evolução | Variação primeiro→último, quantas subidas e descidas, maior subida e descida |
| Tendência | Regressão linear: declive e R² |
| Crescimento médio | Taxa geométrica por período |
| Sazonalidade | Índice sazonal: média de cada mês a dividir pela média geral |
| Ano a ano | Totais por ano, variação interanual, ano mais alto e mais baixo |

### Nenhum método corre sem pontos que cheguem

Uma regressão sobre três pontos faz um relatório **parecer** sério sendo falso. Cada
método tem um mínimo, e quando não é cumprido o relatório **diz que não correu e porquê**,
em vez de calcular na mesma:

| Método | Mínimo |
|---|---|
| Regressão linear | 4 períodos, e um eixo temporal |
| Crescimento médio geométrico | 3 períodos, todos maiores que zero |
| Valores atípicos (Tukey) | 5 categorias |
| Índice sazonal | 2 ciclos anuais completos de dados mensais |
| Variação interanual | 2 anos |

**Métodos temporais nunca correm num eixo categórico.** Uma regressão sobre «Canal» não
significa nada: a ordem das categorias é arbitrária, e o declive mudaria só por trocares
duas colunas no Excel. Nesse caso o relatório diz isso mesmo.

### Os termos qualitativos, e a regra de cada um

Nenhuma palavra qualitativa entra sem uma regra calculada, e a regra aparece ao lado dela
no documento:

| Termo | Regra |
|---|---|
| dispersão baixa / moderada / elevada | coeficiente de variação < 15% / 15–35% / > 35% |
| ajuste fraco / moderado / forte | R² < 0,3 / 0,3–0,7 / > 0,7 |
| tendência crescente / decrescente | declive positivo/negativo **e** R² ≥ 0,3 |
| tendência indefinida | R² < 0,3 |
| série monótona | todas as variações com o mesmo sinal |
| concentração baixa / moderada / elevada | maior categoria < 25% / 25–50% / > 50% do total |

### Previsões

Pedes com o campo `previsao` (número de períodos). **Sem esse campo não há previsão** —
tem de ser pedida, nunca aparece por iniciativa da ferramenta.

**Método.** Decomposição clássica, por esta ordem: dessazonalizar a série (dividir cada
valor pelo seu índice sazonal), ajustar a reta de tendência sobre a série já sem ciclo,
extrapolar, e reaplicar o fator sazonal a cada período previsto. Ajustar a reta depois de
tirar a sazonalidade não é um detalhe: no exemplo deste repositório, o R² sobe de **0,47**
na série bruta para **0,89** na dessazonalizada. Sem sazonalidade identificável, o método
é só a reta sobre a série observada — e o relatório diz qual dos dois usou.

**Nunca um número solto.** Cada período sai como intervalo a 95%, calculado com o
erro-padrão dos resíduos:

```
ŷ ± t(0,975; n−2) · s · √(1 + 1/n + (x₀ − x̄)² / Sxx)
```

**Quando é que recusa**, e cada recusa fica escrita no relatório com a razão:

| Condição | Mínimo |
|---|---|
| Eixo temporal | obrigatório — prever o «período seguinte» de «Canal» não significa nada |
| Períodos observados | 8 |
| Qualidade do ajuste | R² ≥ 0,3 |
| Horizonte | até 1/3 da série, e nunca mais de 12 períodos |

Também avisa quando a previsão desce abaixo de zero numa série sempre positiva (sinal de
que a reta deixou de servir) e quando o intervalo é mais largo do que o próprio valor
previsto.

**A previsão não vai para o gráfico**, de propósito. O gráfico mostra dados; a previsão
vive no texto. Uma linha tracejada num gráfico é fácil de confundir com facto quando se
tira uma captura de ecrã.

**O que a previsão não sabe.** Pressupõe que a tendência e a sazonalidade se mantêm. Não
sabe de campanhas que aí vêm, de concorrentes, de mudanças de preço nem de nada que não
esteja no ficheiro. É uma extrapolação, não uma bola de cristal — e é assim que está
escrita no relatório.

### O que continua de fora, e porquê

Duas coisas nunca vão entrar, e não é por falta de vontade:

**Causas.** O ficheiro tem números, não tem o mundo. A ferramenta nunca saberá se as
vendas subiram por causa da campanha, porque o concorrente fechou ou porque choveu.
Correlação não é causa, e qualquer frase sobre o porquê seria inventada.

**Recomendações.** Precisam de orçamento, capacidade, calendário e estratégia — nada
disso está numa folha de cálculo. Uma recomendação sem esse contexto é um palpite com ar
de conselho.

### Avaliação face a uma meta

O relatório nunca escreve «bom», «fraco» ou «preocupante» por sua conta. Não é timidez:
essas palavras exigem uma **referência que a folha de cálculo não contém**. Dizer que um
canal teve «fraco desempenho» sem saber contra que objetivo é uma afirmação que não se
consegue defender.

A solução não foi proibir a avaliação — foi **exigir a referência**. Dás a meta no plano,
e a avaliação deixa de ser opinião e passa a ser aritmética:

```json
"meta": { "valor": 2000000, "ambito": "total" }
```

| `ambito` | O que compara |
|---|---|
| `total` | O total do período (ou a média, se a agregação for `media`) contra a meta. |
| `categoria` | Cada categoria contra a meta: quantas atingiram, quais ficaram mais longe, e o défice somado. |

Sai assim, do exemplo deste repositório:

> **Meta.** Meta definida: 600.000 por categoria. 2 de 3 categorias atingiram-na (66,7%).
> Mais distantes da meta: «Loja fisica» com 296.476 (303.524 abaixo). Somando o que faltou
> a cada uma, o défice total é 303.524. A média por categoria é 686.227, 114,4% da meta.

Repara no que **não** está lá: nenhuma palavra sobre o resultado ser bom ou mau. Diz-se
quanto foi, contra quanto se queria, e a diferença. O juízo é de quem lê.

Sem `meta`, esta secção não aparece — e o resto do relatório continua a usar só
vocabulário técnico com significado fixo: `tendência crescente`, `dispersão elevada`,
`ajuste moderado`, `valor atípico`. Cada um com o número e o critério à frente, para quem
lê poder discordar do critério em vez de discordar da opinião.

## Quando é que se recusa a gerar

O script analisa sempre os dados antes de gerar. Se encontrar algum destes casos,
**não gera nada** e explica porquê:

- a tabela não começar na primeira linha da folha;
- um gráfico de barras com mais de 40 categorias — as barras ficam finas como cabelos;
- uma tabela de dados com mais de 30 categorias, que seria truncada;
- uma coluna que **nunca desce ao longo de 24 ou mais períodos** — parece um valor
  acumulado, e somar um acumulado dá um número sem significado (num caso real de teste,
  1.531.289.090 em vez dos 15.537.056 verdadeiros);
- uma categoria que parece ser a **linha de totais** da folha (`TOTAL`, `Soma`,
  `Total Geral`, ou a última categoria a valer exatamente a soma das outras) —
  se passasse, o gráfico contaria os mesmos valores duas vezes;
- linhas com células vazias nas colunas usadas;
- coluna numérica guardada como texto no Excel;
- colunas sem cabeçalho na folha (células soltas ao lado da tabela);
- gráfico circular com mais de 12 fatias;
- tabela de dados com mais de 30 categorias.

Lês os avisos e, se estiverem entendidos, repetes com `--avisos-aceites`. Todos
eles ficam registados na secção «Notas sobre os dados», no fim do documento.

Casos mais graves param mesmo, sem hipótese de contornar: folha ou coluna que não
existe, coluna sem números nenhuns, gráfico circular com valores negativos,
ficheiro que não é `.xlsx`, ficheiro aberto no Excel.

## Testes

```bash
python -m pytest testes/ -v
```

Os testes usam ficheiros propositadamente sujos: texto onde deviam estar números,
células vazias, meses fora de ordem alfabética, folhas que não existem, caminhos
com espaços e acentos.

## Desempenho, medido

Com um ficheiro de **250 000 linhas** e 3 séries, num portátil comum:

| Formato | Tempo até ao `.docx` |
|---|---|
| CSV | **2,7 s** |
| `.xlsx` | **16 s** |

O Excel é mais lento de propósito: além de o ler, o programa vai buscar as **células em
bruto** para detetar números guardados como texto. É esse trabalho extra que impede o
`1.250` de ser lido como 1,25. Trocar esses segundos por risco de números errados não
compensa.

## Versões testadas

pandas 3.0.5 · openpyxl 3.1.5 · matplotlib 3.11.1 · python-docx 1.2.0 ·
pytest 9.1.1 · Python 3.14.6 em Windows 11.

## Licença

MIT — ver [LICENSE](LICENSE).
