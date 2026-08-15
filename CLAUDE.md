# CLAUDE.md

> Regras deste projeto. Ler antes de mexer em código ou criar ficheiros.
> O que está em «Decisões fechadas» não se volta a discutir sem pedido explícito.

## 1. O que é o projeto

Skill pública do Claude Code que recebe um ficheiro Excel e instruções em
linguagem natural, e devolve um relatório Word com os gráficos pedidos e a
descrição factual de cada um.

## 2. Ferramentas e porquê

- **pandas + openpyxl** — ler `.xlsx`, agrupar. Padrão da área.
- **matplotlib** — desenhar os gráficos em PNG. Sem dependências pesadas.
- **python-docx** — montar o `.docx` e inserir os PNG.
- **JSON (`plano.json`)** — separa o que se quer do como se gera; torna o
  resultado reprodutível sem repetir a conversa.
- **LibreOffice (`soffice`) + `pdftoppm`** — só para verificar o resultado.
  Opcionais para quem usa a skill; necessários para quem a desenvolve.
- **Python 3.10+**.

## 3. Comandos

```bash
# instalar
pip install -r requirements.txt

# verificar os dados sem gerar nada
python scripts/gerar_relatorio.py --dados X.xlsx --plano plano.json --verificar

# gerar
python scripts/gerar_relatorio.py --dados X.xlsx --plano plano.json --saida r.docx

# testes
python -m pytest testes/ -v

# verificar o resultado, olhando para as páginas
soffice --headless --convert-to pdf r.docx
pdftoppm -jpeg -r 100 r.pdf pagina
```

## 4. Estrutura

```
SKILL.md  README.md  CLAUDE.md  LICENSE  requirements.txt  .gitignore
scripts/gerar_relatorio.py
exemplos/vendas_2023_2025.xlsx  plano.json  relatorio_exemplo.docx  preview.png
testes/test_gerar_relatorio.py
```

## 5. Convenções

- **Um único script**: `scripts/gerar_relatorio.py`. Não partir em módulos sem
  necessidade real e demonstrada.
- Nomes de ficheiros, funções, variáveis e chaves do JSON em **português sem
  acentos** (`gerar_relatorio`, `eixo_x`, `tabela_dados`).
- Texto visível ao utilizador (README, SKILL.md, mensagens, relatório) em
  **português de Portugal, com acentos**.
- Comentários e docstrings do código **sem acentos**, para evitar problemas de
  codificação.
- **Fonte única de documentação**: a referência dos campos do `plano.json` vive
  só no `README.md`. O `SKILL.md` aponta para lá. Nunca a mesma tabela em dois
  ficheiros.
- Se o formato do `plano.json` mudar, atualizar `exemplos/plano.json`, regenerar
  `exemplos/relatorio_exemplo.docx` e `exemplos/preview.png`, e acrescentar um
  teste.
- Mensagens de commit curtas e em português.

## 6. Nunca fazer

- Nunca alterar nem gravar por cima do Excel original.
- Nunca inventar dados, preencher células vazias em silêncio, nem estimar
  números a olho a partir do gráfico.
- Nunca escrever no relatório recomendações, causas ou previsões.
- Nunca escrever avaliações de desempenho («bom», «fraco»). Ver a decisão em 7.
- Nunca perguntar ao utilizador algo que está dentro do ficheiro (colunas,
  folhas, número de linhas) — ler primeiro.
- Nunca despejar perguntas em série. Nunca perguntar sobre cores ou tipos de letra.
- Nunca acrescentar dependências novas sem justificação (seaborn, plotly,
  streamlit e afins ficam de fora).
- Nunca deixar um problema nos dados passar em silêncio: ou erro claro, ou aviso
  registado.
- Nunca acrescentar funcionalidades não pedidas (PDF, PowerPoint, dashboards).
- Nunca dar o trabalho por terminado sem ter verificado o `.docx` gerado.
- Nunca escrever código quando o pedido é para planear.

## 7. Decisões fechadas

- Nome do repositório e da skill: `excel-para-word`. Público, licença MIT.
- Relatório sempre em **português**. README em português com resumo em inglês.
- **A análise é estatística e completa** (decisão revista em 15/08/2026; antes era
  «só descritiva, sem estatística avançada»). Entra dispersão, concentração,
  atípicos, tendência por regressão, crescimento geométrico, índice sazonal e
  quebra ano a ano. Continua sem causas e sem recomendações.
- **Previsões só a pedido** (decisão de 15/08/2026; antes era «sem previsões»).
  Campo `previsao` no gráfico; sem ele não há previsão. Método: dessazonalizar,
  ajustar a reta, extrapolar, reaplicar a sazonalidade. Sempre com intervalo a
  95% — nunca um valor central sozinho. Guardas: 8 períodos, R² ≥ 0,3, horizonte
  até 1/3 da série e no máximo 12. **A previsão não entra no gráfico**: num
  gráfico, uma linha tracejada confunde-se com facto numa captura de ecrã.
- **Avaliação de desempenho só contra uma meta dada pelo utilizador** (decisão de
  15/08/2026; antes era «nunca avaliar desempenho»). A limitação nunca foi de
  pudor: faltava a referência. Campo `meta` com `valor` e `ambito`
  (`total` ou `categoria`). Sem meta, a secção não aparece. **Nunca inventar uma
  meta nem sugerir uma «razoável»** — o número vem de fora ou não há avaliação.
  Mesmo com meta, o relatório diz quanto foi contra quanto se queria, e nunca se
  o resultado é bom.
- **Várias séries no mesmo gráfico** com o campo `serie` (decisão de 15/08/2026).
  Internamente, `preparar_dados` devolve `{nome: pd.Series}`; sem o campo é um
  dicionário com uma entrada de nome vazio. **Nenhuma função de análise mudou** —
  continuam a receber uma `pd.Series` e a ser chamadas uma vez por série.
- **O conjunto não é a soma das séries**: é reagregar os dados em bruto ignorando
  a coluna das séries. Somar médias daria um número errado.
- **Todas as séries são alinhadas à grelha do conjunto**, nunca a uma grelha
  própria: uma série que comece a meio ficaria desenhada por cima dos rótulos
  errados, em silêncio. Falta entra como buraco, nunca como zero.
- **Séries distinguem-se por cor, traço e marca** — o gráfico tem de ler-se
  impresso a preto e branco e por quem não distingue vermelho de verde.
- **O campo `serie` aceita uma lista de colunas**, para dados em formato largo —
  uma coluna por região, por mês ou por canal. Descoberto a usar a skill sobre
  dados públicos reais: sem isto, «comparar por região» era impossível num
  ficheiro que traz as regiões em colunas separadas, que é o caso comum.
- **Séries acumuladas são detetadas e avisadas.** Uma coluna que nunca desce ao
  longo de 24 ou mais períodos parece um acumulado, e somá-la dá um número sem
  significado. O limite é 24 de propósito: com 12, uma série legítima a crescer
  todos os meses era acusada, e os próprios testes apanharam o falso positivo.
- **A linha de cabeçalho é a mais preenchida, não a que tem a seguinte igual.**
  A regra antiga rejeitava o cabeçalho verdadeiro assim que a primeira linha de
  dados tivesse uma célula vazia — banal em dados reais.
- **Causas e recomendações ficam de fora para sempre.** O ficheiro tem números,
  não tem o mundo; e recomendar exige orçamento, capacidade e estratégia que não
  estão numa folha de cálculo. Não voltar a propor.
- **A linha que não se atravessa**: nenhuma palavra qualitativa entra sem uma
  regra calculada, e a regra vai escrita ao lado. Nunca «bom», «mau», «fraco» ou
  «preocupante» — essas exigem uma meta que o ficheiro Excel não tem, e uma
  afirmação indefensável é pior do que um relatório curto. Sim a «tendência
  crescente», «dispersão elevada», «ajuste moderado», sempre com número e critério.
- **Nenhum método corre sem pontos que cheguem.** Regressão: 4 períodos.
  Crescimento geométrico: 3 períodos, todos positivos. Tukey: 5 categorias.
  Índice sazonal: 2 ciclos anuais. Quando falta, o relatório **diz que não correu
  e porquê** — uma regressão sobre 3 pontos faz um documento parecer sério sendo
  falso.
- **Métodos temporais nunca correm num eixo categórico.** O declive de «Canal»
  mudaria só por trocar duas colunas no Excel.
- Os limiares dos termos qualitativos vivem em constantes no topo do script, não
  espalhados pelo código, e estão publicados no README.
- As instruções são dadas em **linguagem natural**; o Claude escreve o
  `plano.json`. Não há formulário para o utilizador preencher.
- Formatos de entrada: **`.xlsx`, `.xlsm` e CSV** (decisão revista em 15/08/2026;
  antes era só `.xlsx`). As exportações de plataformas são quase todas CSV, e é
  esse o caso de uso real. O `.xls` antigo continua de fora: obrigaria ao `xlrd`.
  Macros de `.xlsm` nunca são executadas.
- **Num CSV é tudo texto.** Por isso lê-se com `dtype=str` e a conversão passa
  toda pelo `converter_texto_formatado`, que só aceita o inequívoco. E por isso
  «coluna guardada como texto» é **nota**, não aviso: um aviso que dispara em
  todas as colunas de todos os ficheiros deixa de significar alguma coisa e
  transforma o bloqueio em ruído.
- **Datas em texto só viram eixo temporal com a convenção provada** (um dia
  acima de 12 algures na coluna). Sem prova, ficam categóricas e o relatório
  explica: falhar a análise é seguro, trocar Janeiro por Fevereiro não é.
- **O índice sazonal é uma razão** e só corre com todos os valores positivos.
  Descoberto com dados reais de anomalias de temperatura, que oscilam à volta de
  zero e faziam o índice explodir.
- **Formatar números com casas suficientes para dizerem algo** — um declive de
  0,0005 arredondado a duas casas aparecia como «+0».
- Tipos de gráfico: **`barras`, `linhas`, `circular`**. Não acrescentar mais sem
  pedido explícito.
- Agregações: `soma`, `media`, `contagem`.
- Um gráfico por página; título, descrição e tabela opcional por gráfico.
- O título vai no cabeçalho do Word, não dentro da imagem.
- A ordem das categorias é a **ordem de aparição no ficheiro**, nunca a
  alfabética. Ordenar por alfabeto estraga séries temporais.
- O campo `nota` **não aceita números**. Os números saem sempre dos dados.
- Campos desconhecidos no `plano.json` dão erro, para apanhar planos antigos.
- **A aprovação é mecânica, não uma promessa em prosa**: erros param a execução;
  avisos bloqueiam a geração até `--avisos-aceites`.
- Códigos de saída: `0` correu bem, `1` erro nos dados, `2` bloqueado por avisos.
- **Prioridade: robustez acima de cosmética.** A ferramenta tem de aguentar
  ficheiros feios.
- Público-alvo do repositório: recrutadores e utilizadores de Excel. O README
  assume que se sabe abrir um terminal, e diz isso em vez de fingir o contrário.

- **A tabela pode não começar na primeira linha.** O cabeçalho é detetado, mas
  a deteção gera sempre um aviso — que bloqueia — a dizer que linha foi usada.
  O campo `linha_cabecalho` fixa a escolha. Adivinhar em silêncio nunca.
- **Linhas de totais** deixadas por sistemas de exportação são detetadas e
  bloqueiam, porque contariam os valores duas vezes.
- **Números gravados como texto** (`1.250,00 €`) são convertidos a partir da
  célula em bruto, com a convenção deduzida da própria coluna. Quando nada a
  prova, para com erro: o pandas lê o texto `1.250` como 1,25, e um palpite
  errado dá um relatório mil vezes ao lado sem dar erro nenhum.
- O `plano.json` é lido em `utf-8-sig`, para aceitar o BOM que o Bloco de Notas
  e o PowerShell põem.

## 8. Fora de âmbito até alguém sentir a falta

`barras_horizontais`, `area`, `dispersao`, células combinadas, filtros por
coluna, numeração de páginas, saída em PDF ou PowerPoint.
