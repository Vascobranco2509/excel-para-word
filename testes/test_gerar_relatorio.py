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
