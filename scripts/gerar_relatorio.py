"""Gera um relatorio Word a partir de um ficheiro Excel e de um plano em JSON.

Fluxo: le o .xlsx -> valida o plano -> verifica os dados -> desenha os graficos
em PNG -> monta o .docx.

A verificacao corre sempre. Se encontrar avisos, a geracao e bloqueada ate o
utilizador passar --avisos-aceites. Isto e mecanico de proposito: uma promessa
em prosa nao trava nada.

O ficheiro Excel de origem nunca e escrito nem alterado.
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
import tempfile
import unicodedata
from pathlib import Path

import matplotlib

matplotlib.use("Agg")  # sem janela; obrigatorio para correr sem ecra

import matplotlib.ticker
import matplotlib.pyplot as plt
import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt

TIPOS_GRAFICO = ("barras", "linhas", "circular")
AGREGACOES = ("soma", "media", "contagem")

CHAVES_OBRIGATORIAS = ("tipo", "folha", "eixo_x", "agregacao", "titulo")
CHAVES_OPCIONAIS = ("eixo_y", "nota", "tabela_dados")

MAX_LINHAS_TABELA = 30
MAX_FATIAS_CIRCULAR = 12

# rotulos que denunciam uma linha de totais deixada por um sistema de exportacao
ROTULOS_DE_TOTAL = frozenset({"total", "totais", "total geral", "subtotal", "soma", "sum"})

# simbolos a ignorar quando os numeros vem gravados como texto
SIMBOLOS_MOEDA = ("€", "$", "£", "¥", "R$", "EUR", "USD")
LARGURA_IMAGEM = Inches(6.0)

# tamanhos pensados para a figura ja encolhida dentro da pagina do Word
plt.rcParams.update({
    "font.size": 11,
    "axes.labelsize": 11,
    "xtick.labelsize": 10,
    "ytick.labelsize": 10,
})


class ErroDados(Exception):
    """Problema que impede continuar. Mostrado ao utilizador sem traceback."""


# ---------------------------------------------------------------- utilitarios


def formatar_numero(valor) -> str:
    """Formata um numero a portuguesa: 1.234 e 1.234,56."""
    if valor is None:
        return "-"
    try:
        if pd.isna(valor):
            return "-"
    except (TypeError, ValueError):
        pass
    try:
        numero = float(valor)
    except (TypeError, ValueError):
        return str(valor)
    if abs(numero - round(numero)) < 1e-9:
        return f"{int(round(numero)):,}".replace(",", ".")
    texto = f"{numero:,.2f}"
    # troca os separadores anglo-saxonicos pelos portugueses
    return texto.replace(",", "\x00").replace(".", ",").replace("\x00", ".")


def formatar_categoria(valor) -> str:
    """Formata o rotulo de uma categoria; datas ficam em dd/mm/aaaa."""
    if isinstance(valor, pd.Timestamp):
        return valor.strftime("%d/%m/%Y")
    if valor is None:
        return "(vazio)"
    try:
        if pd.isna(valor):
            return "(vazio)"
    except (TypeError, ValueError):
        pass
    return str(valor)


def listar(itens) -> str:
    return ", ".join(f"«{item}»" for item in itens)


def acrescentar(avisos: list[str], mensagem: str) -> None:
    """Junta um aviso so uma vez.

    Varios graficos leem a mesma folha; sem isto o mesmo aviso sairia
    repetido uma vez por grafico.
    """
    if mensagem not in avisos:
        avisos.append(mensagem)


def normalizar(texto) -> str:
    """Minusculas e sem acentos, para comparar rotulos."""
    sem_acentos = unicodedata.normalize("NFKD", str(texto))
    return "".join(c for c in sem_acentos if not unicodedata.combining(c)).strip().lower()


def limpar_numero(texto) -> str:
    """Tira simbolos de moeda e espacos, deixando so digitos e separadores."""
    limpo = str(texto).strip()
    for simbolo in SIMBOLOS_MOEDA:
        limpo = limpo.replace(simbolo, "")
    return limpo.replace(" ", "").replace(" ", "").strip()


def classificar_formato(limpo: str) -> str | None:
    """Diz que convencao de separadores um valor usa.

    Devolve 'pt' (1.250,00), 'en' (1,250.00), 'simples' (sem separadores),
    'ambiguo' (1.250 -- tanto pode ser mil duzentos e cinquenta como um
    virgula vinte e cinco) ou None se nem sequer parece um numero.
    """
    if not re.fullmatch(r"-?\d[\d.,]*", limpo):
        return None

    virgulas = limpo.count(",")
    pontos = limpo.count(".")

    if virgulas and pontos:
        # o ultimo separador a aparecer e o decimal
        return "pt" if limpo.rindex(",") > limpo.rindex(".") else "en"
    if virgulas:
        if virgulas > 1:
            return "en"  # 1,250,000 -- so pode ser separador de milhares
        return "ambiguo" if len(limpo.split(",")[-1]) == 3 else "pt"
    if pontos:
        if pontos > 1:
            return "pt"  # 1.250.000
        return "ambiguo" if len(limpo.split(".")[-1]) == 3 else "en"
    return "simples"


def aplicar_formato(limpo: str, convencao: str) -> float | None:
    if convencao == "pt":
        limpo = limpo.replace(".", "").replace(",", ".")
    elif convencao == "en":
        limpo = limpo.replace(",", "")
    try:
        return float(limpo)
    except ValueError:
        return None


def ler_coluna_bruta(excel: pd.ExcelFile, folha: str, nome: str,
                     n_linhas: int) -> list | None:
    """Devolve os valores da coluna tal como estao gravados na celula.

    E preciso porque o pandas decide sozinho que o texto «1.250» vale 1,25.
    Num Excel portugues isso quer dizer mil vezes menos, sem erro nenhum e
    sem aviso nenhum. A celula em bruto ainda sabe que aquilo era texto.

    Devolve None se nao conseguir alinhar com o que o pandas leu, para o
    resto do programa seguir pelo caminho normal em vez de arriscar.
    """
    try:
        folha_bruta = excel.book[folha]
        cabecalhos = [c.value for c in next(folha_bruta.iter_rows(max_row=1))]
        indice = cabecalhos.index(nome)
    except (AttributeError, KeyError, StopIteration, ValueError):
        return None

    valores = [
        linha[indice].value if indice < len(linha) else None
        for linha in folha_bruta.iter_rows(min_row=2)
    ]
    # o pandas ignora linhas totalmente vazias no fim; so seguimos se bater certo
    while valores and valores[-1] is None and len(valores) > n_linhas:
        valores.pop()
    return valores if len(valores) == n_linhas else None


def converter_texto_formatado(coluna: pd.Series, onde: str,
                              nome: str) -> tuple[pd.Series, bool]:
    """Converte texto com separadores de milhares e moeda em numeros.

    A convencao e deduzida dos valores inequivocos da propria coluna. Se um
    «980,50 €» prova que a virgula e decimal, entao o «1.250» ao lado deixa
    de ser ambiguo. Se nada a prova, para com erro em vez de adivinhar: um
    palpite errado aqui nao da erro nenhum, da um relatorio mil vezes ao lado.
    """
    limpos = {}
    formatos = {}
    for indice, valor in coluna.items():
        if pd.isna(valor):
            continue
        limpo = limpar_numero(valor)
        limpos[indice] = limpo
        formatos[indice] = classificar_formato(limpo)

    convencoes = {f for f in formatos.values() if f in ("pt", "en")}
    ambiguos = [limpos[i] for i, f in formatos.items() if f == "ambiguo"]

    if len(convencoes) > 1:
        raise ErroDados(
            f"O {onde} usa «{nome}» como valor, mas a coluna mistura formatos de "
            "número incompatíveis (uns à portuguesa, outros à inglesa). "
            "Formata a coluna toda como número no Excel."
        )
    if ambiguos and not convencoes:
        exemplos = list(dict.fromkeys(ambiguos))[:3]
        raise ErroDados(
            f"O {onde} usa «{nome}» como valor, mas os números estão guardados "
            f"como texto num formato ambíguo: {listar(exemplos)}. «1.250» tanto "
            "pode ser mil duzentos e cinquenta como um vírgula vinte e cinco, e "
            "não vou adivinhar. Formata a coluna como número no Excel."
        )

    convencao = convencoes.pop() if convencoes else "simples"
    houve_formatacao = bool(convencoes) or convencao != "simples"

    valores = {}
    for indice, limpo in limpos.items():
        if formatos[indice] is None:
            valores[indice] = None
        else:
            valores[indice] = aplicar_formato(limpo, convencao)

    convertida = pd.Series(valores, dtype="float64").reindex(coluna.index)
    return convertida, houve_formatacao


def detetar_linha_de_total(serie: pd.Series, grafico: dict,
                           avisos: list[str]) -> None:
    """Avisa quando uma categoria parece ser a linha de totais da folha.

    Exportacoes de sistemas costumam deixar uma linha TOTAL no fim. Sem
    isto, essa linha vira uma categoria e o relatorio conta tudo duas
    vezes -- em silencio, que e o pior dos casos.

    Nao se apaga nada: avisa-se, e o aviso bloqueia a geracao ate haver
    uma decisao humana.
    """
    if len(serie) < 2:
        return

    suspeitas = [c for c in serie.index if normalizar(c) in ROTULOS_DE_TOTAL]

    # Sinal independente do nome: a ultima categoria vale a soma das outras.
    # So a partir de 5 categorias, e so na ultima. Com poucas categorias isto
    # acontece por acaso -- 10, 20 e 30 sao dados legitimos, e o 30 e a soma
    # dos outros dois. Linhas de totais ficam no fim da folha, nao no meio.
    if len(serie) >= 5:
        ultima = serie.index[-1]
        valor = float(serie.iloc[-1])
        resto = float(serie.sum()) - valor
        if resto > 0 and abs(valor - resto) <= abs(resto) * 1e-6:
            if ultima not in suspeitas:
                suspeitas.append(ultima)

    if suspeitas:
        acrescentar(avisos, (
            f"{grafico['titulo']}: {listar(suspeitas)} parece(m) ser linha(s) de "
            "totais da folha, e não categorias. Se ficarem, o gráfico conta os "
            "mesmos valores duas vezes. Apaga essa linha do Excel, ou confirma "
            "que é mesmo uma categoria."
        ))


# ------------------------------------------------------------------- leitura


def ler_plano(caminho: Path) -> dict:
    if not caminho.exists():
        raise ErroDados(f"O plano «{caminho}» não foi encontrado.")
    try:
        texto = caminho.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        raise ErroDados(
            f"O plano «{caminho}» não está em UTF-8. Grava-o outra vez em UTF-8."
        ) from None
    try:
        plano = json.loads(texto)
    except json.JSONDecodeError as erro:
        raise ErroDados(
            f"O plano «{caminho}» não é um JSON válido: {erro.msg} "
            f"(linha {erro.lineno}, coluna {erro.colno})."
        ) from None

    if not isinstance(plano, dict):
        raise ErroDados("O plano tem de ser um objeto JSON com chaves e valores.")
    titulo = plano.get("titulo_relatorio")
    if not isinstance(titulo, str) or not titulo.strip():
        raise ErroDados("Falta «titulo_relatorio» no plano, ou está vazio.")
    graficos = plano.get("graficos")
    if not isinstance(graficos, list) or not graficos:
        raise ErroDados("Falta «graficos» no plano, ou a lista está vazia.")

    for indice, grafico in enumerate(graficos, start=1):
        validar_grafico(grafico, indice)
    return plano


def validar_grafico(grafico, indice: int) -> None:
    onde = f"gráfico {indice}"
    if not isinstance(grafico, dict):
        raise ErroDados(f"O {onde} não é um objeto JSON.")

    for chave in CHAVES_OBRIGATORIAS:
        if chave not in grafico:
            raise ErroDados(f"Falta «{chave}» no {onde}.")

    conhecidas = set(CHAVES_OBRIGATORIAS) | set(CHAVES_OPCIONAIS)
    desconhecidas = sorted(set(grafico) - conhecidas)
    if desconhecidas:
        raise ErroDados(
            f"O {onde} tem campos que não existem: {listar(desconhecidas)}. "
            f"Campos aceites: {listar(sorted(conhecidas))}."
        )

    if grafico["tipo"] not in TIPOS_GRAFICO:
        raise ErroDados(
            f"O {onde} pede o tipo «{grafico['tipo']}», que não existe. "
            f"Tipos disponíveis: {listar(TIPOS_GRAFICO)}."
        )
    if grafico["agregacao"] not in AGREGACOES:
        raise ErroDados(
            f"O {onde} pede a agregação «{grafico['agregacao']}», que não existe. "
            f"Agregações disponíveis: {listar(AGREGACOES)}."
        )
    if grafico["agregacao"] != "contagem" and not grafico.get("eixo_y"):
        raise ErroDados(
            f"O {onde} usa a agregação «{grafico['agregacao']}» e por isso "
            "precisa de «eixo_y» (a coluna com os números)."
        )

    nota = grafico.get("nota")
    if nota is not None:
        if not isinstance(nota, str):
            raise ErroDados(f"A «nota» do {onde} tem de ser texto.")
        if any(caractere.isdigit() for caractere in nota):
            raise ErroDados(
                f"A «nota» do {onde} tem números. A nota serve para enquadrar "
                "(ex.: «período da campanha de verão»); os números do relatório "
                "são calculados a partir dos dados, para nunca desmentirem o gráfico."
            )

    if "tabela_dados" in grafico and not isinstance(grafico["tabela_dados"], bool):
        raise ErroDados(f"O campo «tabela_dados» do {onde} tem de ser true ou false.")


def abrir_excel(caminho: Path) -> pd.ExcelFile:
    if not caminho.exists():
        raise ErroDados(f"O ficheiro «{caminho}» não foi encontrado.")
    if caminho.suffix.lower() != ".xlsx":
        extensao = caminho.suffix or "sem extensão"
        raise ErroDados(
            f"Só são suportados ficheiros .xlsx. Recebi «{extensao}». "
            "Abre o ficheiro no Excel e grava como .xlsx."
        )
    try:
        return pd.ExcelFile(caminho, engine="openpyxl")
    except PermissionError:
        raise ErroDados(
            f"Não consigo ler «{caminho}»: o ficheiro está aberto noutro programa "
            "(provavelmente o Excel). Fecha-o e tenta outra vez."
        ) from None
    except Exception as erro:  # ficheiro corrompido, zip invalido, etc.
        raise ErroDados(f"Não consigo abrir «{caminho}» como Excel: {erro}") from None


# --------------------------------------------------------------- preparacao


def preparar_dados(excel: pd.ExcelFile, grafico: dict, indice: int,
                   avisos: list[str], notas: list[str]) -> tuple[pd.Series, int]:
    """Le a folha, valida colunas e devolve a serie agregada e o nº de linhas usadas.

    A ordem das categorias e a ordem de aparicao no ficheiro (sort=False).
    Ordenar por ordem alfabetica estragaria series temporais: Abril viria
    antes de Janeiro.
    """
    onde = f"gráfico {indice} («{grafico['titulo']}»)"
    folha = grafico["folha"]
    if folha not in excel.sheet_names:
        raise ErroDados(
            f"O {onde} pede a folha «{folha}», que não existe no ficheiro. "
            f"Folhas disponíveis: {listar(excel.sheet_names)}."
        )

    df = excel.parse(folha)
    if df.empty:
        raise ErroDados(f"A folha «{folha}» não tem linhas nenhumas.")

    fantasma = [c for c in df.columns if str(c).startswith("Unnamed:")]
    if fantasma:
        acrescentar(avisos, (
            f"A folha «{folha}» tem {len(fantasma)} coluna(s) sem cabeçalho "
            f"({listar(fantasma)}). Costuma ser sinal de células soltas ao lado "
            "da tabela. Foram ignoradas."
        ))

    eixo_x = grafico["eixo_x"]
    if eixo_x not in df.columns:
        raise ErroDados(
            f"O {onde} pede a coluna «{eixo_x}», que não existe na folha «{folha}». "
            f"Colunas disponíveis: {listar(df.columns)}."
        )

    agregacao = grafico["agregacao"]
    eixo_y = grafico.get("eixo_y")

    if eixo_y is not None:
        if eixo_y not in df.columns:
            raise ErroDados(
                f"O {onde} pede a coluna «{eixo_y}», que não existe na folha «{folha}». "
                f"Colunas disponíveis: {listar(df.columns)}."
            )
        df = proteger_de_texto_reinterpretado(df, excel, folha, eixo_y, onde, avisos)
        df = converter_para_numero(df, eixo_y, folha, onde, avisos)
        colunas_necessarias = [eixo_x, eixo_y]
    else:
        colunas_necessarias = [eixo_x]

    antes = len(df)
    df = df.dropna(subset=colunas_necessarias)
    ignoradas = antes - len(df)
    if ignoradas:
        acrescentar(avisos, (
            f"{ignoradas} linha(s) da folha «{folha}» foram ignoradas por terem "
            f"células vazias em {listar(colunas_necessarias)}."
        ))
    if df.empty:
        raise ErroDados(
            f"Depois de ignorar as linhas com células vazias, o {onde} ficou sem dados."
        )

    if eixo_y is None:
        serie = df.groupby(eixo_x, sort=False).size()
    else:
        grupo = df.groupby(eixo_x, sort=False)[eixo_y]
        serie = {"soma": grupo.sum, "media": grupo.mean, "contagem": grupo.count}[agregacao]()

    if serie.empty:
        raise ErroDados(f"O {onde} não produziu nenhuma categoria depois de agrupar.")

    detetar_linha_de_total(serie, grafico, avisos)

    if len(df) > len(serie):
        notas.append(
            f"{grafico['titulo']}: {len(df)} linhas agrupadas em {len(serie)} "
            f"categorias de «{eixo_x}» ({agregacao})."
        )

    if grafico["tipo"] == "circular":
        if (serie < 0).any():
            negativas = [formatar_categoria(c) for c in serie[serie < 0].index]
            raise ErroDados(
                f"O {onde} é circular, mas há valores negativos em {listar(negativas)}. "
                "Uma fatia não pode ser negativa. Usa barras, ou filtra esses valores."
            )
        if (serie == 0).all():
            raise ErroDados(f"O {onde} é circular e todos os valores são zero.")
        if len(serie) > MAX_FATIAS_CIRCULAR:
            acrescentar(avisos, (
                f"{grafico['titulo']}: o gráfico circular ficaria com {len(serie)} "
                f"fatias. Acima de {MAX_FATIAS_CIRCULAR} deixa de se ler. "
                "Considera barras, ou agrupar as categorias mais pequenas."
            ))

    return serie, len(df)


def proteger_de_texto_reinterpretado(df: pd.DataFrame, excel: pd.ExcelFile,
                                     folha: str, coluna: str, onde: str,
                                     avisos: list[str]) -> pd.DataFrame:
    """Reconverte a coluna a partir das celulas em bruto quando la havia texto.

    So faz alguma coisa se a coluna tiver mesmo celulas de texto com
    separadores ou simbolos de moeda. Numeros a serio nao passam por aqui.
    """
    brutos = ler_coluna_bruta(excel, folha, coluna, len(df))
    if brutos is None:
        return df

    formatados = [
        v for v in brutos
        if isinstance(v, str) and classificar_formato(limpar_numero(v)) in
        ("pt", "en", "ambiguo")
    ]
    if not formatados:
        return df

    coluna_bruta = pd.Series(brutos, index=df.index, dtype="object")
    convertida, _ = converter_texto_formatado(coluna_bruta, onde, coluna)

    mensagem = (
        f"A coluna «{coluna}» da folha «{folha}» tem números gravados como texto "
        "(com separadores de milhares ou símbolos de moeda). Foram convertidos a "
        "partir da célula original."
    )
    if pd.api.types.is_numeric_dtype(df[coluna]):
        antes = pd.to_numeric(df[coluna], errors="coerce")
        diferentes = int((~antes.sub(convertida).abs().le(1e-9)).sum())
        if diferentes:
            mensagem += (
                f" Atenção: {diferentes} valor(es) tinham sido lidos com o ponto "
                "como separador decimal, o que dava mil vezes menos. Confirma-os "
                "no relatório."
            )
    acrescentar(avisos, mensagem)

    df = df.copy()
    df[coluna] = convertida
    return df


def converter_para_numero(df: pd.DataFrame, coluna: str, folha: str,
                          onde: str, avisos: list[str]) -> pd.DataFrame:
    """Garante que a coluna e numerica. Nunca soma texto em silencio."""
    if pd.api.types.is_numeric_dtype(df[coluna]):
        return df

    if pd.api.types.is_datetime64_any_dtype(df[coluna]):
        raise ErroDados(
            f"O {onde} usa «{coluna}» como valor, mas essa coluna tem datas, "
            "não números."
        )

    convertida, com_formatacao = converter_texto_formatado(df[coluna], onde, coluna)
    validos = int(convertida.notna().sum())
    if validos == 0:
        exemplos = [str(v) for v in df[coluna].dropna().unique()[:3]]
        raise ErroDados(
            f"O {onde} usa «{coluna}» como valor, mas essa coluna não tem números. "
            f"Exemplos do que lá está: {listar(exemplos)}."
        )

    ja_vazias = int(df[coluna].isna().sum())
    perdidas = int(convertida.isna().sum()) - ja_vazias
    mensagem = (
        f"A coluna «{coluna}» da folha «{folha}» estava guardada como texto e "
        "foi convertida para número"
        + (" (separadores de milhares e símbolos de moeda incluídos)." if com_formatacao
           else ".")
    )
    if perdidas > 0:
        nao_numericas = df.loc[convertida.isna() & df[coluna].notna(), coluna]
        exemplos = [str(v) for v in nao_numericas.unique()[:3]]
        mensagem += (
            f" {perdidas} célula(s) não eram números e foram ignoradas "
            f"(exemplos: {listar(exemplos)})."
        )
    acrescentar(avisos, mensagem)

    df = df.copy()
    df[coluna] = convertida
    return df


# ----------------------------------------------------------------- desenho


def desenhar(serie: pd.Series, grafico: dict, destino: Path) -> None:
    """Desenha o grafico em PNG. O titulo nao entra na imagem: vai no Word."""
    rotulos = [formatar_categoria(c) for c in serie.index]
    valores = list(serie.values)
    tipo = grafico["tipo"]

    # a figura entra no Word com LARGURA_IMAGEM; desenhar muito maior do que
    # isso encolhe o texto ate ficar ilegivel na pagina impressa
    figura, eixos = plt.subplots(figsize=(7, 4), dpi=200)

    if tipo == "barras":
        eixos.bar(rotulos, valores, color="#3b6ea5")
        eixos.grid(axis="y", linestyle=":", alpha=0.6)
    elif tipo == "linhas":
        eixos.plot(rotulos, valores, marker="o", color="#3b6ea5", linewidth=2)
        eixos.grid(linestyle=":", alpha=0.6)
    else:  # circular
        eixos.pie(
            valores,
            labels=rotulos,
            # percentagem a portuguesa, com virgula, como o resto do relatorio
            autopct=lambda p: f"{formatar_numero(round(p, 1))}%",
            startangle=90,
            counterclock=False,
        )
        eixos.axis("equal")

    if tipo in ("barras", "linhas"):
        eixos.set_axisbelow(True)
        eixos.set_xlabel(str(grafico["eixo_x"]))
        eixos.set_ylabel(str(grafico.get("eixo_y") or "Registos"))
        eixos.yaxis.set_major_formatter(
            matplotlib.ticker.FuncFormatter(lambda v, _: formatar_numero(v))
        )
        rodar_rotulos(eixos, rotulos)

    figura.tight_layout()
    figura.savefig(destino, bbox_inches="tight")
    plt.close(figura)


def rodar_rotulos(eixos, rotulos: list[str]) -> None:
    """Roda os rotulos do eixo x quando sao muitos ou compridos."""
    mais_comprido = max((len(r) for r in rotulos), default=0)
    if len(rotulos) > 8 or mais_comprido > 6:
        for rotulo in eixos.get_xticklabels():
            rotulo.set_rotation(45)
            rotulo.set_horizontalalignment("right")


# ------------------------------------------------------------------- texto


def frase_descritiva(serie: pd.Series, grafico: dict, n_linhas: int) -> str:
    """Frase descritiva, com numeros vindos da mesma conta que desenhou o grafico."""
    agregacao = grafico["agregacao"]
    eixo_x = grafico["eixo_x"]
    eixo_y = grafico.get("eixo_y")

    if agregacao == "soma":
        o_que = f"o total de «{eixo_y}»"
    elif agregacao == "media":
        o_que = f"a média de «{eixo_y}»"
    elif eixo_y is None:
        o_que = "o número de registos"
    else:
        o_que = f"a contagem de «{eixo_y}»"

    maior_categoria = formatar_categoria(serie.idxmax())
    menor_categoria = formatar_categoria(serie.idxmin())

    partes = [
        f"O gráfico mostra {o_que} por «{eixo_x}», a partir de "
        f"{formatar_numero(n_linhas)} linhas agrupadas em "
        f"{formatar_numero(len(serie))} categorias."
    ]

    if agregacao == "soma":
        partes.append(f"O total é {formatar_numero(serie.sum())}.")
    elif agregacao == "media":
        partes.append(f"A média das categorias é {formatar_numero(serie.mean())}.")
    else:
        partes.append(f"O total de registos é {formatar_numero(serie.sum())}.")

    partes.append(
        f"O valor mais alto é {formatar_numero(serie.max())}, em «{maior_categoria}», "
        f"e o mais baixo é {formatar_numero(serie.min())}, em «{menor_categoria}»."
    )

    if grafico["tipo"] == "circular" and serie.sum() > 0:
        fatia = round(float(serie.max()) / float(serie.sum()) * 100, 1)
        partes.append(
            f"A maior fatia, «{maior_categoria}», representa "
            f"{formatar_numero(fatia)}% do total."
        )

    return " ".join(partes)


# ---------------------------------------------------------------- documento


def montar_documento(plano: dict, resultados: list[dict], avisos: list[str],
                     notas: list[str], saida: Path) -> None:
    documento = Document()
    documento.add_heading(plano["titulo_relatorio"], level=0)

    for posicao, resultado in enumerate(resultados):
        if posicao > 0:
            documento.add_page_break()

        grafico = resultado["grafico"]
        serie = resultado["serie"]

        documento.add_heading(grafico["titulo"], level=1)
        documento.add_picture(str(resultado["imagem"]), width=LARGURA_IMAGEM)
        documento.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

        documento.add_paragraph(frase_descritiva(serie, grafico, resultado["n_linhas"]))

        nota = grafico.get("nota")
        if nota:
            paragrafo = documento.add_paragraph()
            paragrafo.add_run(nota).italic = True

        if grafico.get("tabela_dados"):
            acrescentar_tabela(documento, serie, grafico, notas)

    documento.add_page_break()
    documento.add_heading("Notas sobre os dados", level=1)
    if avisos:
        documento.add_paragraph(
            f"Foram encontrados {len(avisos)} problema(s) nos dados de origem:"
        )
        for linha in avisos:
            documento.add_paragraph(linha, style="List Bullet")
    else:
        documento.add_paragraph("Não foram encontrados problemas nos dados de origem.")

    if notas:
        documento.add_paragraph("Como os dados foram agrupados:")
        for linha in notas:
            documento.add_paragraph(linha, style="List Bullet")

    saida.parent.mkdir(parents=True, exist_ok=True)
    documento.save(saida)


def acrescentar_tabela(documento, serie: pd.Series, grafico: dict,
                       notas: list[str]) -> None:
    mostradas = serie.head(MAX_LINHAS_TABELA)
    escondidas = len(serie) - len(mostradas)

    tabela = documento.add_table(rows=1, cols=2)
    tabela.style = "Light Grid Accent 1"
    cabecalho = tabela.rows[0].cells
    cabecalho[0].text = str(grafico["eixo_x"])
    cabecalho[1].text = str(grafico.get("eixo_y") or "Registos")

    for categoria, valor in mostradas.items():
        celulas = tabela.add_row().cells
        celulas[0].text = formatar_categoria(categoria)
        celulas[1].text = formatar_numero(valor)

    if escondidas > 0:
        aviso = (
            f"{grafico['titulo']}: a tabela mostra as primeiras "
            f"{MAX_LINHAS_TABELA} de {len(serie)} categorias; "
            f"{escondidas} ficaram de fora."
        )
        notas.append(aviso)
        paragrafo = documento.add_paragraph()
        corrida = paragrafo.add_run(aviso)
        corrida.italic = True
        corrida.font.size = Pt(9)


# --------------------------------------------------------------------- main


def executar(args) -> int:
    plano = ler_plano(Path(args.plano))
    excel = abrir_excel(Path(args.dados))

    avisos: list[str] = []
    notas: list[str] = []
    preparados = []

    for indice, grafico in enumerate(plano["graficos"], start=1):
        serie, n_linhas = preparar_dados(excel, grafico, indice, avisos, notas)
        preparados.append({"grafico": grafico, "serie": serie, "n_linhas": n_linhas})

    print(f"Plano lido: {len(preparados)} gráfico(s).")
    for nota in notas:
        print(f"  · {nota}")

    if avisos:
        print(f"\n{len(avisos)} aviso(s) sobre os dados:")
        for aviso in avisos:
            print(f"  ! {aviso}")
    else:
        print("\nSem avisos: os dados estão limpos.")

    if args.verificar:
        print("\n--verificar: nada foi gerado.")
        return 0

    if avisos and not args.avisos_aceites:
        print(
            "\nO relatório NÃO foi gerado por causa dos avisos acima.\n"
            "Se estiverem todos entendidos e quiseres avançar mesmo assim, "
            "repete o comando com --avisos-aceites."
        )
        return 2

    saida = Path(args.saida)
    pasta_temporaria = Path(tempfile.mkdtemp(prefix="excel-para-word-"))
    try:
        for posicao, resultado in enumerate(preparados, start=1):
            imagem = pasta_temporaria / f"grafico_{posicao}.png"
            desenhar(resultado["serie"], resultado["grafico"], imagem)
            resultado["imagem"] = imagem
        montar_documento(plano, preparados, avisos, notas, saida)
    finally:
        shutil.rmtree(pasta_temporaria, ignore_errors=True)

    print(f"\nRelatório gerado: {saida}")
    return 0


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        # a consola do Windows nao usa UTF-8 por omissao e engasga-se nos acentos
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

    analisador = argparse.ArgumentParser(
        description="Gera um relatório Word com gráficos a partir de um Excel."
    )
    analisador.add_argument("--dados", required=True, help="ficheiro .xlsx de origem")
    analisador.add_argument("--plano", required=True, help="ficheiro plano.json")
    analisador.add_argument("--saida", help="ficheiro .docx a criar")
    analisador.add_argument(
        "--verificar", action="store_true",
        help="analisa os dados e mostra os problemas, sem gerar nada",
    )
    analisador.add_argument(
        "--avisos-aceites", dest="avisos_aceites", action="store_true",
        help="gera o relatório mesmo havendo avisos",
    )
    args = analisador.parse_args()

    if not args.verificar and not args.saida:
        analisador.error("é preciso --saida (ou usa --verificar para só analisar).")

    try:
        return executar(args)
    except ErroDados as erro:
        print(f"\nErro: {erro}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
