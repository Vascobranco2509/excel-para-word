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
exemplos/vendas_2025.xlsx  plano.json  relatorio_exemplo.docx  preview.png
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
- Nunca escrever no relatório recomendações, causas, previsões ou juízos de valor.
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
- A análise no Word é **só descritiva**: o que o gráfico mostra, com números
  concretos. Sem insights, sem recomendações, sem estatística avançada.
- As instruções são dadas em **linguagem natural**; o Claude escreve o
  `plano.json`. Não há formulário para o utilizador preencher.
- Formatos de entrada: **só `.xlsx`**. CSV, `.xls` e `.xlsm` ficam de fora.
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

## 8. Fora de âmbito até alguém sentir a falta

`barras_horizontais`, `area`, `dispersao`, deteção automática de cabeçalho,
células combinadas, filtros por coluna, numeração de páginas, saída em PDF ou
PowerPoint.
