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
    excel = gr.abrir_dados(caminho)
    serie, n_linhas, _, _ = gr.preparar_dados(excel, config, 1, avisos, notas)
    return serie, n_linhas, avisos, notas


def preparar_completo(caminho: Path, config: dict):
    """Como preparar(), mas devolve tambem o eixo temporal e a serie anual."""
    avisos: list[str] = []
    notas: list[str] = []
    excel = gr.abrir_dados(caminho)
    serie, n_linhas, temporal, anual = gr.preparar_dados(excel, config, 1, avisos, notas)
    return serie, n_linhas, temporal, anual, avisos


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
    excel = gr.abrir_dados(ficheiro)
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


# ------------------------------------------------------------------- CSV


def escrever_csv(caminho: Path, linhas: list[str], separador: str = ";",
                 codificacao: str = "utf-8") -> Path:
    caminho.parent.mkdir(parents=True, exist_ok=True)
    caminho.write_text("\n".join(linhas), encoding=codificacao)
    return caminho


def test_csv_com_ponto_e_virgula(tmp_path):
    """O Excel portugues exporta com «;»."""
    ficheiro = escrever_csv(tmp_path / "v.csv", [
        "Mes;Valor", "Janeiro;10", "Fevereiro;20", "Marco;30",
    ])
    serie, _, _, _ = preparar(ficheiro, grafico())
    assert list(serie.index) == ["Janeiro", "Fevereiro", "Marco"]
    assert serie.sum() == 60


def test_csv_com_virgula(tmp_path):
    ficheiro = escrever_csv(tmp_path / "v.csv", [
        "Mes,Valor", "Janeiro,10", "Fevereiro,20",
    ])
    serie, _, _, _ = preparar(ficheiro, grafico())
    assert serie.sum() == 30


def test_csv_em_cp1252_com_acentos(tmp_path):
    """Ficheiros exportados por Windows antigos nao sao UTF-8."""
    ficheiro = escrever_csv(tmp_path / "v.csv", [
        "Regiao;Valor", "Lisboa;10", "Évora;20", "Bragança;30",
    ], codificacao="cp1252")
    serie, _, _, _ = preparar(ficheiro, grafico(eixo_x="Regiao"))
    assert "Évora" in list(serie.index)
    assert "Bragança" in list(serie.index)


def test_csv_com_metadados_por_cima(tmp_path):
    """Exportacoes do Google Ads trazem linhas de relatorio antes da tabela."""
    ficheiro = escrever_csv(tmp_path / "v.csv", [
        "Relatorio de campanhas", "Intervalo: ultimo mes", "",
        "Mes;Valor", "Janeiro;10", "Fevereiro;20",
    ])
    serie, _, avisos, _ = preparar(ficheiro, grafico())
    assert serie.sum() == 30
    assert any("não começa na primeira linha" in a for a in avisos)


def test_csv_com_moeda_a_portuguesa(tmp_path):
    ficheiro = escrever_csv(tmp_path / "v.csv", [
        "Mes;Valor", "Janeiro;1.250,00 €", "Fevereiro;980,50 €",
    ])
    serie, _, _, _ = preparar(ficheiro, grafico())
    assert serie["Janeiro"] == pytest.approx(1250.0)
    assert serie["Fevereiro"] == pytest.approx(980.5)


def test_csv_ignora_a_folha_indicada_no_plano(tmp_path):
    ficheiro = escrever_csv(tmp_path / "v.csv", ["Mes;Valor", "Janeiro;10", "Fevereiro;20"])
    avisos: list[str] = []
    notas: list[str] = []
    fonte = gr.abrir_dados(ficheiro)
    gr.preparar_dados(fonte, grafico(folha="Vendas"), 1, avisos, notas)
    assert any("não tem folhas" in n for n in notas)


def test_csv_nao_transforma_texto_em_aviso_bloqueante(tmp_path):
    """Num CSV e tudo texto: se isso fosse aviso, bloqueava sempre e virava ruido."""
    ficheiro = escrever_csv(tmp_path / "v.csv", ["Mes;Valor", "Janeiro;10", "Fevereiro;20"])
    _, _, avisos, notas = preparar(ficheiro, grafico())
    assert avisos == []
    assert any("guardada como texto" in n for n in notas)


def test_csv_com_celulas_perdidas_continua_a_avisar(tmp_path):
    ficheiro = escrever_csv(tmp_path / "v.csv", [
        "Mes;Valor", "Janeiro;10", "Fevereiro;vinte", "Marco;30",
    ])
    _, _, avisos, _ = preparar(ficheiro, grafico())
    assert any("vinte" in a for a in avisos)


def test_xlsm_e_aceite(tmp_path):
    caminho = tmp_path / "com_macros.xlsm"
    pd.DataFrame({"Mes": ["Janeiro", "Fevereiro"], "Valor": [10, 20]}).to_excel(
        caminho, sheet_name="Vendas", index=False)
    serie, _, _, _ = preparar(caminho, grafico())
    assert serie.sum() == 30


def test_xls_antigo_da_mensagem_util(tmp_path):
    falso = tmp_path / "antigo.xls"
    falso.write_bytes(b"nao interessa")
    with pytest.raises(gr.ErroDados, match="formato antigo"):
        gr.abrir_dados(falso)


def test_folha_em_falta_com_varias_folhas_e_erro(tmp_path):
    caminho = tmp_path / "duas.xlsx"
    with pd.ExcelWriter(caminho) as escritor:
        pd.DataFrame({"Mes": ["Janeiro"], "Valor": [1]}).to_excel(
            escritor, sheet_name="Vendas", index=False)
        pd.DataFrame({"Mes": ["Janeiro"], "Valor": [2]}).to_excel(
            escritor, sheet_name="Custos", index=False)
    config = grafico()
    config.pop("folha", None)
    with pytest.raises(gr.ErroDados, match="não diz de que folha"):
        preparar(caminho, config)


def test_folha_unica_dispensa_o_campo(tmp_path):
    ficheiro = escrever_excel(tmp_path / "uma.xlsx", {"Mes": ["Janeiro"], "Valor": [7]})
    config = grafico()
    config.pop("folha", None)
    serie, _, _, _ = preparar(ficheiro, config)
    assert serie.sum() == 7


# ---------------------------------------------------------- datas em texto


def test_datas_em_texto_com_convencao_provada(tmp_path):
    """Um dia acima de 12 prova que e dia/mes."""
    ficheiro = escrever_csv(tmp_path / "d.csv", [
        "Data;Valor",
        "31/01/2026;10", "15/02/2026;20", "03/03/2026;30", "20/04/2026;40",
    ])
    serie, _, _, notas = preparar(ficheiro, grafico(eixo_x="Data"))
    assert any("reconhecida como datas" in n for n in notas)
    assert [d.month for d in serie.index] == [1, 2, 3, 4]


def test_datas_ambiguas_ficam_categoricas(tmp_path):
    """Nenhum dia acima de 12: nao da para saber se e dia/mes ou mes/dia."""
    ficheiro = escrever_csv(tmp_path / "d.csv", [
        "Data;Valor", "01/02/2026;10", "03/04/2026;20", "05/06/2026;30",
    ])
    _, _, _, notas = preparar(ficheiro, grafico(eixo_x="Data"))
    assert any("não é possível saber" in n for n in notas)


def test_datas_iso_sao_sempre_aceites(tmp_path):
    ficheiro = escrever_csv(tmp_path / "d.csv", [
        "Data;Valor", "2026-01-05;10", "2026-02-06;20", "2026-03-07;30",
    ])
    serie, _, _, _ = preparar(ficheiro, grafico(eixo_x="Data"))
    assert [d.month for d in serie.index] == [1, 2, 3]


def test_datas_com_convencoes_misturadas_param_com_erro(tmp_path):
    ficheiro = escrever_csv(tmp_path / "d.csv", [
        "Data;Valor", "31/01/2026;10", "01/31/2026;20",
    ])
    with pytest.raises(gr.ErroDados, match="incompatíveis"):
        preparar(ficheiro, grafico(eixo_x="Data"))


def test_separador_detetado_em_varios_formatos():
    assert gr.detetar_separador("a;b;c\n1;2;3") == ";"
    assert gr.detetar_separador("a,b,c\n1,2,3") == ","
    assert gr.detetar_separador("a\tb\tc\n1\t2\t3") == "\t"


# ------------------------------------------------------------ ficheiro e plano


def test_extensao_errada_da_mensagem_clara(tmp_path):
    falso = tmp_path / "antigo.xls"
    falso.write_bytes(b"nao interessa")
    with pytest.raises(gr.ErroDados, match=r"\.xlsx"):
        gr.abrir_dados(falso)


def test_ficheiro_inexistente_da_mensagem_clara(tmp_path):
    with pytest.raises(gr.ErroDados, match="não foi encontrado"):
        gr.abrir_dados(tmp_path / "nao_existe.xlsx")


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


# ------------------------------------------------------------- estatistica
# Valores conferidos a parte. Um R2 errado nao se ve a olho como se ve um
# grafico partido: tem de ser batido contra outra conta.


def test_regressao_com_reta_perfeita():
    reta = gr.regressao_linear([0.0, 2.0, 4.0, 6.0, 8.0])
    assert reta["declive"] == pytest.approx(2.0)
    assert reta["intercecao"] == pytest.approx(0.0)
    assert reta["r2"] == pytest.approx(1.0)
    assert reta["erro_padrao"] == pytest.approx(0.0)


def test_regressao_com_serie_plana_da_declive_zero():
    reta = gr.regressao_linear([5.0, 5.0, 5.0, 5.0])
    assert reta["declive"] == pytest.approx(0.0)
    assert reta["r2"] == pytest.approx(1.0)


def test_regressao_nao_corre_com_poucos_pontos():
    """Uma reta sobre tres pontos faz um relatorio parecer serio sendo falso."""
    assert gr.regressao_linear([1.0, 2.0, 3.0]) is None


def test_crescimento_medio_duplica_por_periodo():
    # 100 -> 200 -> 400: 100% por periodo
    assert gr.crescimento_medio([100.0, 200.0, 400.0]) == pytest.approx(100.0)


def test_crescimento_medio_recusa_zeros_e_negativos():
    assert gr.crescimento_medio([100.0, 0.0, 400.0]) is None
    assert gr.crescimento_medio([100.0, -50.0, 400.0]) is None


def test_crescimento_medio_recusa_poucos_pontos():
    assert gr.crescimento_medio([100.0, 200.0]) is None


def test_atipicos_encontra_o_valor_disparatado():
    serie = pd.Series([10, 11, 12, 11, 10, 900])
    encontrados = gr.detetar_atipicos(serie)
    assert [v for _, v in encontrados] == [900]


def test_atipicos_nao_correm_com_poucas_categorias():
    assert gr.detetar_atipicos(pd.Series([1, 2, 3, 4])) is None


def test_classificacao_usa_os_limiares_documentados():
    termos = ("baixa", "moderada", "elevada")
    assert gr.classificar(14.9, 15, 35, termos) == "baixa"
    assert gr.classificar(20.0, 15, 35, termos) == "moderada"
    assert gr.classificar(35.1, 15, 35, termos) == "elevada"


# ------------------------------------------- achados em dados publicos reais
# Cada um destes apareceu ao correr a skill sobre dados abertos de terceiros
# (PIB do Banco Mundial, anomalias de temperatura global). Ficam aqui presos
# com ficheiros sinteticos minimos, sem meter dados de ninguem no repositorio.


def test_indice_sazonal_recusa_series_com_negativos():
    """Anomalias de temperatura oscilam a volta de zero: o indice e uma razao
    e explodia, saindo «índice -0,34 — 134% abaixo», um disparate com ar de rigor."""
    rotulos, valores = [], []
    for ano in (2024, 2025):
        for mes in range(1, 13):
            rotulos.append(f"{ano}-{mes:02d}")
            valores.append(-0.5 if mes % 2 else 0.4)
    resultado = gr.indice_sazonal(pd.Series(valores, index=rotulos))
    assert resultado["impossivel"] == "negativos"


def test_seccao_sazonalidade_explica_a_recusa_por_negativos():
    rotulos, valores = [], []
    for ano in (2024, 2025):
        for mes in range(1, 13):
            rotulos.append(f"{ano}-{mes:02d}")
            valores.append(-0.5 if mes % 2 else 0.4)
    blocos = gr.montar_analise(pd.Series(valores, index=rotulos), grafico(), 24, True, None)
    assert "positivos" in texto_de(blocos, "Sazonalidade")


def test_declive_pequeno_nao_e_arredondado_para_zero():
    """Um declive de 0,0005 aparecia como «+0», o que nao diz nada."""
    valores = [i * 0.0005 for i in range(12)]
    serie = pd.Series(valores, index=[f"P{i}" for i in range(12)])
    blocos = gr.montar_analise(serie, grafico(), 12, True, None)
    assert "+0 por período" not in texto_de(blocos, "Tendência")
    assert "0,0005" in texto_de(blocos, "Tendência")


def test_previsao_de_valores_pequenos_nao_colapsa_em_inteiros():
    """A previsao arredondada a inteiros dava «entre 0 e 1» em dados pequenos."""
    serie = pd.Series([i * 0.01 for i in range(12)], index=[f"P{i}" for i in range(12)])
    _, texto = gr.secao_previsao(serie, 3)
    assert "entre 0 e 1" not in texto


def test_anos_anteriores_a_1900_sao_reconhecidos():
    """A serie de temperatura comeca em 1850 e ficava como categoria."""
    assert gr.e_ano(1850) is True
    assert gr.extrair_ano("1850-01") == 1850
    assert gr.parece_temporal(pd.Index(["1850-01", "1850-02", "1850-03"])) is True


def test_formato_ano_mes_e_reconhecido_sem_ajuda():
    """As exportacoes de plataformas trazem «2023-01» e nenhuma dica no plano."""
    assert gr.parece_temporal(pd.Index(["2023-01", "2023-02", "2023-03"])) is True


def test_barras_com_categorias_a_mais_avisam(tmp_path):
    """262 paises num grafico de barras sao ilegiveis, e nada avisava."""
    categorias = [f"Pais {i}" for i in range(60)]
    ficheiro = escrever_excel(tmp_path / "muitos.xlsx",
                              {"Mes": categorias, "Valor": list(range(60))})
    _, _, avisos, _ = preparar(ficheiro, grafico())
    assert any("categorias" in a and "ilegíveis" in a for a in avisos)


def test_linhas_com_muitos_pontos_nao_avisam(tmp_path):
    """Numa linha, 60 pontos leem-se bem; o aviso das barras seria falso positivo."""
    categorias = [f"P{i}" for i in range(60)]
    ficheiro = escrever_excel(tmp_path / "linha.xlsx",
                              {"Mes": categorias, "Valor": list(range(60))})
    _, _, avisos, _ = preparar(ficheiro, grafico(tipo="linhas"))
    assert not any("ilegíveis" in a for a in avisos)


def test_tabela_truncada_avisa_antes_de_gerar(tmp_path):
    """O README dizia que bloqueava; o codigo so o registava depois de gerar."""
    categorias = [f"C{i}" for i in range(50)]
    ficheiro = escrever_excel(tmp_path / "tabela.xlsx",
                              {"Mes": categorias, "Valor": list(range(50))})
    _, _, avisos, _ = preparar(ficheiro, grafico(tipo="linhas", tabela_dados=True))
    assert any("ficariam de fora" in a for a in avisos)


# ---------------------------------------------------------------- previsao


def serie_reta(n: int, declive: float = 2.0) -> pd.Series:
    """Serie perfeitamente linear, com rotulos que nao ativam a sazonalidade."""
    return pd.Series([declive * i for i in range(n)],
                     index=[f"P{i}" for i in range(n)])


def test_previsao_acerta_em_cheio_numa_reta_sem_ruido():
    """Sanidade matematica: se falhar aqui, esta errada em todo o lado."""
    resultado = gr.prever(serie_reta(10), 3)
    previstos = [p["centro"] for p in resultado["previsoes"]]

    assert previstos == pytest.approx([20.0, 22.0, 24.0])
    for p in resultado["previsoes"]:
        assert p["superior"] - p["inferior"] == pytest.approx(0.0, abs=1e-6)


def test_previsao_traz_sempre_intervalo():
    serie = pd.Series([10, 13, 11, 18, 21, 19, 26, 30, 28, 35],
                      index=[f"P{i}" for i in range(10)])
    resultado = gr.prever(serie, 2)
    for p in resultado["previsoes"]:
        assert p["inferior"] < p["centro"] < p["superior"]


def test_previsao_recusa_serie_curta():
    assert gr.prever(serie_reta(5), 2) is None


def test_previsao_recusa_ajuste_fraco():
    """Uma serie sem forma nao se extrapola."""
    serie = pd.Series([10, 90, 20, 85, 15, 95, 25, 80, 12, 88],
                      index=[f"P{i}" for i in range(10)])
    resultado = gr.prever(serie, 2)
    assert resultado["recusa"] == "ajuste"
    assert resultado["r2"] < gr.R2_MINIMO_PREVISAO


def test_horizonte_limitado_a_um_terco_da_serie():
    assert gr.horizonte_permitido(9) == 3
    assert gr.horizonte_permitido(36) == 12
    assert gr.horizonte_permitido(90) == 12  # travado pelo maximo absoluto


def test_seccao_previsao_diz_quando_corta_o_horizonte():
    _, texto = gr.secao_previsao(serie_reta(12), 99)
    assert "máximo defensável" in texto
    assert "ficção" in texto


def test_seccao_previsao_explica_a_recusa_por_ajuste():
    serie = pd.Series([10, 90, 20, 85, 15, 95, 25, 80, 12, 88],
                      index=[f"P{i}" for i in range(10)])
    _, texto = gr.secao_previsao(serie, 3)
    assert "Não calculada" in texto
    assert "ajuste" in texto


def test_seccao_previsao_explica_a_recusa_por_serie_curta():
    _, texto = gr.secao_previsao(serie_reta(5), 2)
    assert "pelo menos 8" in texto


def test_previsao_avisa_quando_desce_abaixo_de_zero():
    serie = serie_reta(12, declive=-2.0) + 22  # comeca em 22 e desce ate zero
    _, texto = gr.secao_previsao(serie, 4)
    assert "abaixo de zero" in texto


def test_eixo_categorico_recusa_previsao():
    serie = pd.Series([10, 20, 30], index=["Meta", "Google", "TikTok"])
    blocos = gr.montar_analise(serie, grafico(eixo_x="Canal", previsao=3), 3, False, None)
    assert "período seguinte" in texto_de(blocos, "Previsão")


def test_sem_campo_previsao_nao_ha_seccao_previsao():
    serie = serie_reta(12)
    blocos = gr.montar_analise(serie, grafico(), 12, True, None)
    assert "Previsão" not in etiquetas(blocos)


def test_previsao_e_recusada_se_nao_for_inteiro_positivo():
    with pytest.raises(gr.ErroDados, match="previsao"):
        gr.validar_grafico(grafico(previsao=0), 1)
    with pytest.raises(gr.ErroDados, match="previsao"):
        gr.validar_grafico(grafico(previsao="seis"), 1)


# ------------------------------------------------------- rotulos futuros


def test_rotulos_futuros_viram_o_ano():
    rotulos = gr.proximos_rotulos(pd.Index(["2025-10", "2025-11", "2025-12"]), 3)
    assert rotulos == ["2026-01", "2026-02", "2026-03"]


def test_rotulos_futuros_continuam_os_meses():
    rotulos = gr.proximos_rotulos(pd.Index(["Outubro", "Novembro", "Dezembro"]), 2)
    assert rotulos == ["Janeiro", "Fevereiro"]


def test_rotulos_futuros_continuam_os_anos():
    assert gr.proximos_rotulos(pd.Index([2023, 2024, 2025]), 2) == ["2026", "2027"]


def test_rotulos_futuros_desistem_em_vez_de_inventar():
    rotulos = gr.proximos_rotulos(pd.Index(["Fase A", "Fase B"]), 2)
    assert rotulos == ["período +1", "período +2"]


# ------------------------------------------------------------ eixo temporal


def test_meses_em_portugues_sao_linha_do_tempo():
    assert gr.parece_temporal(pd.Index(MESES)) is True


def test_anos_sao_linha_do_tempo():
    assert gr.parece_temporal(pd.Index([2023, 2024, 2025])) is True


def test_canais_nao_sao_linha_do_tempo():
    """Uma regressao sobre «Canal» mudaria de declive so por trocar colunas."""
    assert gr.parece_temporal(pd.Index(["Meta", "Google Ads", "TikTok"])) is False


def test_extrair_ano_de_varios_formatos():
    assert gr.extrair_ano("2023-01") == 2023
    assert gr.extrair_ano(2024) == 2024
    assert gr.extrair_ano(pd.Timestamp("2025-06-30")) == 2025
    assert gr.extrair_ano("Janeiro") is None


def test_indice_sazonal_precisa_de_dois_ciclos():
    um_ano = pd.Series(range(1, 13), index=[f"2025-{m:02d}" for m in range(1, 13)])
    assert gr.indice_sazonal(um_ano) is None


def test_indice_sazonal_encontra_o_mes_forte():
    rotulos, valores = [], []
    for ano in (2024, 2025):
        for mes in range(1, 13):
            rotulos.append(f"{ano}-{mes:02d}")
            valores.append(300 if mes == 12 else 100)
    resultado = gr.indice_sazonal(pd.Series(valores, index=rotulos))

    assert resultado["ciclos"] == 2
    assert max(resultado["indices"], key=resultado["indices"].get) == 12


# -------------------------------------------------------------- bloco analise


def etiquetas(blocos):
    return [etiqueta for etiqueta, _ in blocos]


def texto_de(blocos, etiqueta):
    return next(t for e, t in blocos if e == etiqueta)


def test_eixo_categorico_nao_leva_tendencia():
    serie = pd.Series([10, 20, 30], index=["Meta", "Google", "TikTok"])
    blocos = gr.montar_analise(serie, grafico(eixo_x="Canal"), 3, False, None)

    assert "Tendência" not in etiquetas(blocos)
    assert "categórico" in texto_de(blocos, "Evolução")


def test_serie_curta_diz_que_a_regressao_nao_correu():
    serie = pd.Series([10, 20, 30], index=["Janeiro", "Fevereiro", "Marco"])
    blocos = gr.montar_analise(serie, grafico(), 3, True, None)

    assert "não calculada" in texto_de(blocos, "Tendência")
    assert "pelo menos 4" in texto_de(blocos, "Tendência")


def test_analise_ano_a_ano_com_dois_anos():
    anual = pd.Series([100.0, 150.0], index=[2024, 2025])
    serie = pd.Series([10, 20, 30, 40], index=MESES[:4])
    blocos = gr.montar_analise(serie, grafico(), 4, True, anual)

    comparacao = texto_de(blocos, "Comparação entre anos")
    assert "+50" in comparacao
    assert "+50%" in comparacao
    assert "2025" in texto_de(blocos, "Ano a ano")


def test_sem_juizos_de_valor_no_texto():
    """A linha que nao se atravessa: nada de «bom», «mau» ou «fraco desempenho»."""
    serie = pd.Series([10, 20, 15, 40, 35, 60], index=[f"2025-{m:02d}" for m in range(1, 7)])
    blocos = gr.montar_analise(serie, grafico(), 6, True, None)
    tudo = " ".join(t for _, t in blocos).lower()

    for proibida in ("bom ", "mau ", "excelente", "preocupante", "fraco desempenho"):
        assert proibida not in tudo, f"apareceu um juízo de valor: {proibida}"


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


def test_analise_curta_nao_traz_o_bloco_de_analise(tmp_path):
    from docx import Document

    ficheiro = escrever_excel(tmp_path / "limpo.xlsx",
                              {"Mes": MESES, "Valor": list(range(1, 13))})
    plano = tmp_path / "plano.json"
    plano.write_text(json.dumps(
        {"titulo_relatorio": "T", "analise": "curta", "graficos": [grafico()]},
        ensure_ascii=False), encoding="utf-8")
    saida = tmp_path / "r.docx"
    correr_cli("--dados", str(ficheiro), "--plano", str(plano), "--saida", str(saida))

    texto = "\n".join(p.text for p in Document(str(saida)).paragraphs)
    assert "Análise" not in texto
    assert "O total é" in texto


def test_analise_completa_e_o_comportamento_por_omissao(tmp_path):
    from docx import Document

    ficheiro = escrever_excel(tmp_path / "limpo.xlsx",
                              {"Mes": MESES, "Valor": list(range(1, 13))})
    plano = escrever_plano(tmp_path / "plano.json", [grafico()])
    saida = tmp_path / "r.docx"
    correr_cli("--dados", str(ficheiro), "--plano", str(plano), "--saida", str(saida))

    texto = "\n".join(p.text for p in Document(str(saida)).paragraphs)
    assert "Análise" in texto
    assert "Dispersão" in texto


def test_campo_desconhecido_no_plano_e_recusado(tmp_path):
    plano = tmp_path / "plano.json"
    plano.write_text(json.dumps(
        {"titulo_relatorio": "T", "graficos": [grafico()], "autor": "eu"},
        ensure_ascii=False), encoding="utf-8")
    with pytest.raises(gr.ErroDados, match="autor"):
        gr.ler_plano(plano)


def test_analise_com_valor_invalido_e_recusada(tmp_path):
    plano = tmp_path / "plano.json"
    plano.write_text(json.dumps(
        {"titulo_relatorio": "T", "analise": "profunda", "graficos": [grafico()]},
        ensure_ascii=False), encoding="utf-8")
    with pytest.raises(gr.ErroDados, match="profunda"):
        gr.ler_plano(plano)


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
