"""Testes com ficheiros propositadamente sujos.

O que se testa aqui nao e o caso bonito: e o Excel real, com texto onde
deviam estar numeros, celulas vazias, folhas que nao existem e meses
fora de ordem alfabetica.
"""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import pandas as pd
import pytest

RAIZ = Path(__file__).resolve().parent.parent
SCRIPT = RAIZ / "scripts" / "gerar_relatorio.py"
sys.path.insert(0, str(SCRIPT.parent))

import gerar_relatorio as gr  # noqa: E402

MESES = ["Janeiro", "Fevereiro", "Marco", "Abril", "Maio", "Junho",
         "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]


# ------------------------------------------------------------------ auxiliares


def escrever_excel(caminho: Path, dados: dict, folha: str = "Vendas") -> Path:
    caminho.parent.mkdir(parents=True, exist_ok=True)
    pd.DataFrame(dados).to_excel(caminho, sheet_name=folha, index=False)
    return caminho


def escrever_plano(caminho: Path, graficos: list[dict],
                   titulo: str = "Relatório de teste") -> Path:
    caminho.parent.mkdir(parents=True, exist_ok=True)
    caminho.write_text(
        json.dumps({"titulo_relatorio": titulo, "graficos": graficos}, ensure_ascii=False),
        encoding="utf-8",
    )
    return caminho


def grafico(**alteracoes) -> dict:
    base = {
        "tipo": "barras",
        "folha": "Vendas",
        "eixo_x": "Mes",
        "eixo_y": "Valor",
        "agregacao": "soma",
        "titulo": "Teste",
    }
    base.update(alteracoes)
    return base


def preparar(caminho: Path, config: dict):
    """Corre a preparacao de dados e devolve (serie, n_linhas, avisos, notas)."""
    avisos: list[str] = []
    notas: list[str] = []
    excel = gr.abrir_excel(caminho)
    serie, n_linhas = gr.preparar_dados(excel, config, 1, avisos, notas)
    return serie, n_linhas, avisos, notas


# --------------------------------------------------------- ordem das categorias


def test_ordem_dos_meses_segue_o_ficheiro_e_nao_o_alfabeto(tmp_path):
    """O erro classico: agrupar por ordem alfabetica poe Abril antes de Janeiro."""
    ficheiro = escrever_excel(
        tmp_path / "meses.xlsx",
        {"Mes": MESES, "Valor": list(range(1, 13))},
    )
    serie, _, avisos, _ = preparar(ficheiro, grafico())

    assert list(serie.index) == MESES
    assert list(serie.index) != sorted(MESES)
    assert avisos == []


# ------------------------------------------------------------- colunas de texto


def test_coluna_de_texto_convertivel_avisa_e_nao_soma_em_silencio(tmp_path):
    ficheiro = escrever_excel(
        tmp_path / "texto.xlsx",
        {"Mes": ["Janeiro", "Fevereiro", "Marco"], "Valor": ["10", "20", "trinta"]},
    )
    serie, n_linhas, avisos, _ = preparar(ficheiro, grafico())

    assert serie.sum() == 30  # "trinta" nao entrou na conta
    assert n_linhas == 2
    assert any("texto" in aviso for aviso in avisos)
    assert any("trinta" in aviso for aviso in avisos)


def test_coluna_sem_numeros_nenhuns_e_erro(tmp_path):
    ficheiro = escrever_excel(
        tmp_path / "so_texto.xlsx",
        {"Mes": ["Janeiro", "Fevereiro"], "Valor": ["muito", "pouco"]},
    )
    with pytest.raises(gr.ErroDados, match="não tem números"):
        preparar(ficheiro, grafico())


def test_coluna_de_datas_como_valor_e_erro(tmp_path):
    ficheiro = escrever_excel(
        tmp_path / "datas.xlsx",
        {"Mes": ["Janeiro", "Fevereiro"],
         "Valor": pd.to_datetime(["2025-01-01", "2025-02-01"])},
    )
    with pytest.raises(gr.ErroDados, match="datas"):
        preparar(ficheiro, grafico())


# ------------------------------------------------------- numeros com formato
# Encontrados no teste com um ficheiro a serio: uma exportacao portuguesa
# grava o dinheiro como texto, «1.250,00 €», que nao e numero nenhum para
# o pandas.


def test_moeda_a_portuguesa_e_convertida(tmp_path):
    ficheiro = escrever_excel(
        tmp_path / "euros.xlsx",
        {"Mes": ["Janeiro", "Fevereiro", "Marco"],
         "Valor": ["1.250,00 €", "980,50 €", "0,00 €"]},
    )
    serie, _, avisos, _ = preparar(ficheiro, grafico())

    assert serie["Janeiro"] == pytest.approx(1250.0)
    assert serie["Fevereiro"] == pytest.approx(980.5)
    assert serie["Marco"] == pytest.approx(0.0)
    assert any("moeda" in aviso for aviso in avisos)


def test_moeda_a_inglesa_e_convertida(tmp_path):
    ficheiro = escrever_excel(
        tmp_path / "dolares.xlsx",
        {"Mes": ["Janeiro", "Fevereiro"], "Valor": ["$1,250.00", "$980.50"]},
    )
    serie, _, _, _ = preparar(ficheiro, grafico())
    assert serie["Janeiro"] == pytest.approx(1250.0)
    assert serie["Fevereiro"] == pytest.approx(980.5)


def escrever_excel_com_texto(caminho: Path, cabecalhos: list[str],
                             linhas: list[list]) -> Path:
    """Escreve celulas tal e qual, sem o pandas pelo meio a interpretar."""
    from openpyxl import Workbook

    caminho.parent.mkdir(parents=True, exist_ok=True)
    livro = Workbook()
    folha = livro.active
    folha.title = "Vendas"
    folha.append(cabecalhos)
    for linha in linhas:
        folha.append(linha)
    livro.save(caminho)
    return caminho


def test_formato_ambiguo_para_com_erro_em_vez_de_adivinhar(tmp_path):
    """«1.250» tanto pode ser 1250 como 1,25. Adivinhar da um erro invisivel."""
    ficheiro = escrever_excel_com_texto(
        tmp_path / "ambiguo.xlsx",
        ["Mes", "Valor"],
        [["Janeiro", "1.250"], ["Fevereiro", "2.500"]],
    )
    with pytest.raises(gr.ErroDados, match="ambíguo"):
        preparar(ficheiro, grafico())


def test_texto_ambiguo_nao_e_lido_como_decimal_pelo_pandas(tmp_path):
    """O pandas le o texto «1.250» como 1,25. Num Excel portugues isso e mil
    vezes menos. A celula em bruto tem de mandar."""
    ficheiro = escrever_excel_com_texto(
        tmp_path / "mil_vezes.xlsx",
        ["Mes", "Valor"],
        [["Janeiro", "1.250"], ["Fevereiro", "980,50"]],
    )
    serie, _, avisos, _ = preparar(ficheiro, grafico())

    assert serie["Janeiro"] == pytest.approx(1250.0), "leu 1,25 em vez de 1.250"
    assert serie["Fevereiro"] == pytest.approx(980.5)
    assert any("texto" in aviso for aviso in avisos)


def test_numeros_a_serio_nao_sao_mexidos(tmp_path):
    """Uma coluna de numeros verdadeiros nao pode passar pelo caminho do texto."""
    ficheiro = escrever_excel(
        tmp_path / "numeros.xlsx",
        {"Mes": ["Janeiro", "Fevereiro", "Marco"], "Valor": [1.234, 2.5, 1250.0]},
    )
    serie, _, avisos, _ = preparar(ficheiro, grafico())
    assert serie["Janeiro"] == pytest.approx(1.234)
    assert serie["Marco"] == pytest.approx(1250.0)
    assert avisos == []


def test_valor_ambiguo_resolve_se_a_coluna_o_provar(tmp_path):
    """«980,50» prova que a virgula e decimal; logo «1.250» e mil duzentos e cinquenta."""
    ficheiro = escrever_excel(
        tmp_path / "resolvido.xlsx",
        {"Mes": ["Janeiro", "Fevereiro"], "Valor": ["1.250", "980,50"]},
    )
    serie, _, _, _ = preparar(ficheiro, grafico())
    assert serie["Janeiro"] == pytest.approx(1250.0)
    assert serie["Fevereiro"] == pytest.approx(980.5)


def test_formatos_misturados_param_com_erro(tmp_path):
    ficheiro = escrever_excel(
        tmp_path / "misturado.xlsx",
        {"Mes": ["Janeiro", "Fevereiro"], "Valor": ["1.250,00", "2,500.00"]},
    )
    with pytest.raises(gr.ErroDados, match="incompatíveis"):
        preparar(ficheiro, grafico())


def test_espaco_como_separador_de_milhares(tmp_path):
    ficheiro = escrever_excel(
        tmp_path / "espacos.xlsx",
        {"Mes": ["Janeiro", "Fevereiro"], "Valor": ["1 250,00", "980,50"]},
    )
    serie, _, _, _ = preparar(ficheiro, grafico())
    assert serie["Janeiro"] == pytest.approx(1250.0)


# --------------------------------------------------------------- celulas vazias


def test_celulas_vazias_sao_ignoradas_com_aviso(tmp_path):
    ficheiro = escrever_excel(
        tmp_path / "vazias.xlsx",
        {"Mes": ["Janeiro", "Fevereiro", None, "Abril"], "Valor": [10, None, 30, 40]},
    )
    serie, n_linhas, avisos, _ = preparar(ficheiro, grafico())

    assert n_linhas == 2  # so Janeiro e Abril tem os dois campos preenchidos
    assert serie.sum() == 50
    assert any("ignoradas" in aviso for aviso in avisos)


def test_ficheiro_todo_vazio_depois_da_limpeza_e_erro(tmp_path):
    ficheiro = escrever_excel(
        tmp_path / "tudo_vazio.xlsx",
        {"Mes": ["Janeiro", "Fevereiro"], "Valor": [None, None]},
    )
    with pytest.raises(gr.ErroDados, match="ficou sem dados"):
        preparar(ficheiro, grafico())


# ------------------------------------------------------------ folhas e colunas


def test_folha_inexistente_lista_as_que_existem(tmp_path):
    ficheiro = escrever_excel(tmp_path / "f.xlsx", {"Mes": ["Janeiro"], "Valor": [1]})
    with pytest.raises(gr.ErroDados) as erro:
        preparar(ficheiro, grafico(folha="Faturacao"))
    assert "Vendas" in str(erro.value)


def test_coluna_inexistente_lista_as_que_existem(tmp_path):
    ficheiro = escrever_excel(tmp_path / "c.xlsx", {"Mes": ["Janeiro"], "Valor": [1]})
    with pytest.raises(gr.ErroDados) as erro:
        preparar(ficheiro, grafico(eixo_y="Receita"))
    assert "Valor" in str(erro.value)


def test_colunas_fantasma_geram_aviso(tmp_path):
    caminho = tmp_path / "fantasma.xlsx"
    df = pd.DataFrame({"Mes": ["Janeiro", "Fevereiro"], "Valor": [10, 20]})
    df["Unnamed: 5"] = [None, "nota solta"]
    df.to_excel(caminho, sheet_name="Vendas", index=False)

    _, _, avisos, _ = preparar(caminho, grafico())
    assert any("sem cabeçalho" in aviso for aviso in avisos)


# -------------------------------------------------------- linhas de totais
# Encontrados num teste com um ficheiro de exportacao a serio: a linha TOTAL
# do fim da folha virava uma categoria e o relatorio contava tudo duas vezes,
# sem avisar.


def test_linha_total_no_fim_da_folha_e_apanhada(tmp_path):
    ficheiro = escrever_excel(
        tmp_path / "com_total.xlsx",
        {"Campanha": ["Inverno", "Primavera", "Verao", "TOTAL"],
         "Valor": [100, 200, 300, 600]},
    )
    _, _, avisos, _ = preparar(ficheiro, grafico(eixo_x="Campanha"))
    assert any("totais" in aviso for aviso in avisos)


def test_total_com_acentos_e_maiusculas_tambem_e_apanhado(tmp_path):
    ficheiro = escrever_excel(
        tmp_path / "totais.xlsx",
        {"Campanha": ["Inverno", "Primavera", "Total Geral"],
         "Valor": [100, 200, 300]},
    )
    _, _, avisos, _ = preparar(ficheiro, grafico(eixo_x="Campanha"))
    assert any("totais" in aviso for aviso in avisos)


def test_dados_legitimos_nao_sao_confundidos_com_totais(tmp_path):
    """10, 20 e 30 sao dados a serio: o 30 e a soma dos outros por acaso."""
    ficheiro = escrever_excel(
        tmp_path / "legitimo.xlsx",
        {"Mes": ["Janeiro", "Fevereiro", "Marco"], "Valor": [10, 20, 30]},
    )
    _, _, avisos, _ = preparar(ficheiro, grafico())
    assert avisos == []


def test_aviso_repetido_so_aparece_uma_vez(tmp_path):
    """Dois graficos sobre a mesma folha suja davam o mesmo aviso duas vezes."""
    ficheiro = escrever_excel(
        tmp_path / "suja.xlsx",
        {"Mes": ["Janeiro", "Fevereiro", "Marco"],
         "Valor": [10, None, 30],
         "Cliques": [1, 2, 3]},
    )
    avisos: list[str] = []
    notas: list[str] = []
    excel = gr.abrir_excel(ficheiro)
    gr.preparar_dados(excel, grafico(), 1, avisos, notas)
    gr.preparar_dados(excel, grafico(titulo="Outro"), 2, avisos, notas)

    assert len(avisos) == len(set(avisos)), f"avisos repetidos: {avisos}"


# --------------------------------------------------- cabecalho fora da linha 1
# Folhas reais costumam ter o titulo do relatorio e a data de exportacao por
# cima da tabela. Sem isto, o titulo virava cabecalho e as colunas ficavam
# todas «Unnamed».


def folha_com_titulo_por_cima(caminho: Path) -> Path:
    from openpyxl import Workbook

    caminho.parent.mkdir(parents=True, exist_ok=True)
    livro = Workbook()
    folha = livro.active
    folha.title = "Vendas"
    folha["A1"] = "Relatório de Campanhas — 1.º trimestre"
    folha["A2"] = "Exportado em 31/03/2026"
    folha.append([])
    folha.append(["Mes", "Valor"])
    for mes, valor in zip(MESES[:3], [10, 20, 35]):
        folha.append([mes, valor])
    livro.save(caminho)
    return caminho


def test_cabecalho_fora_da_primeira_linha_e_encontrado(tmp_path):
    ficheiro = folha_com_titulo_por_cima(tmp_path / "titulo.xlsx")
    serie, n_linhas, avisos, _ = preparar(ficheiro, grafico())

    assert list(serie.index) == MESES[:3]
    assert serie.sum() == 65
    assert n_linhas == 3
    assert any("não começa na primeira linha" in aviso for aviso in avisos)


def test_linha_cabecalho_explicita_nao_gera_aviso(tmp_path):
    ficheiro = folha_com_titulo_por_cima(tmp_path / "titulo.xlsx")
    serie, _, avisos, _ = preparar(ficheiro, grafico(linha_cabecalho=4))

    assert serie.sum() == 65
    assert avisos == []


def test_folha_normal_continua_a_usar_a_primeira_linha(tmp_path):
    ficheiro = escrever_excel(
        tmp_path / "normal.xlsx",
        {"Mes": MESES[:3], "Valor": [10, 20, 35]},
    )
    serie, _, avisos, _ = preparar(ficheiro, grafico())
    assert serie.sum() == 65
    assert avisos == []


def test_linha_cabecalho_invalida_e_recusada():
    with pytest.raises(gr.ErroDados, match="linha_cabecalho"):
        gr.validar_grafico(grafico(linha_cabecalho=0), 1)


# ----------------------------------------------------------------- circular


def test_circular_com_negativos_e_erro(tmp_path):
    ficheiro = escrever_excel(
        tmp_path / "neg.xlsx",
        {"Mes": ["Janeiro", "Fevereiro"], "Valor": [10, -5]},
    )
    with pytest.raises(gr.ErroDados, match="negativos"):
        preparar(ficheiro, grafico(tipo="circular"))


def test_circular_com_muitas_fatias_avisa(tmp_path):
    categorias = [f"Categoria {i}" for i in range(20)]
    ficheiro = escrever_excel(
        tmp_path / "fatias.xlsx",
        {"Mes": categorias, "Valor": list(range(1, 21))},
    )
    _, _, avisos, _ = preparar(ficheiro, grafico(tipo="circular"))
    assert any("fatias" in aviso for aviso in avisos)


# ------------------------------------------------------------ ficheiro e plano


def test_extensao_errada_da_mensagem_clara(tmp_path):
    falso = tmp_path / "antigo.xls"
    falso.write_bytes(b"nao interessa")
    with pytest.raises(gr.ErroDados, match=r"\.xlsx"):
        gr.abrir_excel(falso)


def test_ficheiro_inexistente_da_mensagem_clara(tmp_path):
    with pytest.raises(gr.ErroDados, match="não foi encontrado"):
        gr.abrir_excel(tmp_path / "nao_existe.xlsx")


def test_caminho_com_espacos_e_acentos_funciona(tmp_path):
    pasta = tmp_path / "relatórios de gestão" / "ano corrente"
    ficheiro = escrever_excel(pasta / "vendas anuais.xlsx",
                              {"Mes": MESES[:3], "Valor": [1, 2, 3]})
    serie, _, _, _ = preparar(ficheiro, grafico())
    assert serie.sum() == 6


def test_nota_com_numeros_e_recusada():
    with pytest.raises(gr.ErroDados, match="números"):
        gr.validar_grafico(grafico(nota="cresceu 20% no verão"), 1)


def test_nota_sem_numeros_passa():
    gr.validar_grafico(grafico(nota="período da campanha de verão"), 1)


def test_campo_que_nao_existe_e_recusado():
    """Apanha planos escritos contra a especificacao antiga."""
    with pytest.raises(gr.ErroDados, match="ordem_categorias"):
        gr.validar_grafico(grafico(ordem_categorias=["Janeiro"]), 1)


def test_tipo_desconhecido_lista_os_disponiveis():
    with pytest.raises(gr.ErroDados) as erro:
        gr.validar_grafico(grafico(tipo="dispersao"), 1)
    assert "barras" in str(erro.value)


def test_media_sem_eixo_y_e_erro():
    config = grafico(agregacao="media")
    del config["eixo_y"]
    with pytest.raises(gr.ErroDados, match="eixo_y"):
        gr.validar_grafico(config, 1)


def test_contagem_sem_eixo_y_conta_linhas(tmp_path):
    ficheiro = escrever_excel(
        tmp_path / "contagem.xlsx",
        {"Mes": ["Janeiro", "Janeiro", "Fevereiro"], "Valor": [1, 2, 3]},
    )
    config = grafico(agregacao="contagem")
    del config["eixo_y"]
    serie, _, _, _ = preparar(ficheiro, config)
    assert serie["Janeiro"] == 2
    assert serie["Fevereiro"] == 1


def test_plano_com_bom_e_aceite(tmp_path):
    """O Bloco de Notas e o PowerShell gravam UTF-8 com BOM. Nao e motivo para recusar."""
    plano = tmp_path / "com_bom.json"
    plano.write_text(
        json.dumps({"titulo_relatorio": "x", "graficos": [grafico()]}, ensure_ascii=False),
        encoding="utf-8-sig",
    )
    assert gr.ler_plano(plano)["titulo_relatorio"] == "x"


def test_plano_com_json_invalido_diz_a_linha(tmp_path):
    plano = tmp_path / "mau.json"
    plano.write_text('{"titulo_relatorio": "x",}', encoding="utf-8")
    with pytest.raises(gr.ErroDados, match="linha"):
        gr.ler_plano(plano)


# --------------------------------------------------------- bloqueio por avisos


def correr_cli(*argumentos: str) -> subprocess.CompletedProcess:
    return subprocess.run(
        [sys.executable, str(SCRIPT), *argumentos],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
    )


def test_avisos_bloqueiam_a_geracao_e_nao_criam_docx(tmp_path):
    ficheiro = escrever_excel(
        tmp_path / "sujo.xlsx",
        {"Mes": ["Janeiro", "Fevereiro", "Marco"], "Valor": [10, None, 30]},
    )
    plano = escrever_plano(tmp_path / "plano.json", [grafico()])
    saida = tmp_path / "relatorio.docx"

    resultado = correr_cli("--dados", str(ficheiro), "--plano", str(plano),
                           "--saida", str(saida))

    assert resultado.returncode == 2
    assert not saida.exists(), "gerou o .docx apesar dos avisos"
    assert "--avisos-aceites" in resultado.stdout


def test_avisos_aceites_desbloqueiam_a_geracao(tmp_path):
    ficheiro = escrever_excel(
        tmp_path / "sujo.xlsx",
        {"Mes": ["Janeiro", "Fevereiro", "Marco"], "Valor": [10, None, 30]},
    )
    plano = escrever_plano(tmp_path / "plano.json", [grafico()])
    saida = tmp_path / "relatorio.docx"

    resultado = correr_cli("--dados", str(ficheiro), "--plano", str(plano),
                           "--saida", str(saida), "--avisos-aceites")

    assert resultado.returncode == 0
    assert saida.exists()


def test_verificar_nunca_gera_nada(tmp_path):
    ficheiro = escrever_excel(tmp_path / "limpo.xlsx",
                              {"Mes": MESES[:3], "Valor": [1, 2, 3]})
    plano = escrever_plano(tmp_path / "plano.json", [grafico()])
    saida = tmp_path / "relatorio.docx"

    resultado = correr_cli("--dados", str(ficheiro), "--plano", str(plano),
                           "--saida", str(saida), "--verificar")

    assert resultado.returncode == 0
    assert not saida.exists()


def test_erro_de_dados_nao_mostra_traceback(tmp_path):
    ficheiro = escrever_excel(tmp_path / "limpo.xlsx",
                              {"Mes": MESES[:3], "Valor": [1, 2, 3]})
    plano = escrever_plano(tmp_path / "plano.json", [grafico(folha="Inexistente")])

    resultado = correr_cli("--dados", str(ficheiro), "--plano", str(plano), "--verificar")

    assert resultado.returncode == 1
    assert "Traceback" not in resultado.stderr
    assert "Erro:" in resultado.stderr


# ------------------------------------------------------- limpeza e integridade


def test_o_excel_de_origem_nao_e_alterado(tmp_path):
    ficheiro = escrever_excel(tmp_path / "origem.xlsx",
                              {"Mes": MESES[:3], "Valor": [1, 2, 3]})
    antes = ficheiro.read_bytes()

    plano = escrever_plano(tmp_path / "plano.json", [grafico()])
    correr_cli("--dados", str(ficheiro), "--plano", str(plano),
               "--saida", str(tmp_path / "r.docx"))

    assert ficheiro.read_bytes() == antes


def test_nao_ficam_pastas_temporarias_para_tras(tmp_path):
    import tempfile

    ficheiro = escrever_excel(tmp_path / "limpo.xlsx",
                              {"Mes": MESES[:3], "Valor": [1, 2, 3]})
    plano = escrever_plano(tmp_path / "plano.json", [grafico()])
    base = Path(tempfile.gettempdir())
    antes = set(base.glob("excel-para-word-*"))

    correr_cli("--dados", str(ficheiro), "--plano", str(plano),
               "--saida", str(tmp_path / "r.docx"))

    assert set(base.glob("excel-para-word-*")) == antes


def test_documento_gerado_tem_o_conteudo_esperado(tmp_path):
    from docx import Document

    ficheiro = escrever_excel(tmp_path / "limpo.xlsx",
                              {"Mes": MESES[:3], "Valor": [10, 20, 30]})
    plano = escrever_plano(
        tmp_path / "plano.json",
        [grafico(titulo="Vendas por mês", tabela_dados=True)],
        titulo="Relatório anual",
    )
    saida = tmp_path / "r.docx"
    correr_cli("--dados", str(ficheiro), "--plano", str(plano), "--saida", str(saida))

    documento = Document(str(saida))
    texto = "\n".join(p.text for p in documento.paragraphs)

    assert "Relatório anual" in texto
    assert "Vendas por mês" in texto
    assert "Notas sobre os dados" in texto
    assert "O total é 60." in texto
    assert len(documento.inline_shapes) == 1, "faltou a imagem do gráfico"
    assert len(documento.tables) == 1
    assert documento.tables[0].rows[1].cells[0].text == "Janeiro"
