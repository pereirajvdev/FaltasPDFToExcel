import re
import sys
import math
from pathlib import Path
from openpyxl.styles import Font, PatternFill, Alignment

import pdfplumber
from openpyxl import load_workbook


# ============================================================
# CONFIGURAÇÃO
# ============================================================

ARQUIVO_TEMPLATE = Path(
    r"C:\Users\Joao Castro\Documents\AutoFaltas\data\Faltas.xlsx"
)

NOME_ABA = "Faltas"

CELULA_INICIAL = "A1"

CARGA_HORARIA_PADRAO = 8


PADRAO_ARQUIVO = re.compile(
    r"^FALTAS\s+\d+\s+(\d{4})\.pdf$",
    re.IGNORECASE
)


PADRAO_CODIGO_VERBA = re.compile(
    r"Lançamento::\s*(\d+)\s*-\s*(.+)",
    re.IGNORECASE
)


PADRAO_LINHA_619 = re.compile(
    r"^"
    r"(?P<mes_folha>\d{2}-\d{2})\s+"
    r"(?P<codigo>\d{6})\s+"
    r"(?P<funcionario>.+?)\s+"
    r"(?P<admissao>\d{2}/\d{2}/\d{4})\s+"
    r"(?P<quantidade>[\d.,]+)\s+"
    r"(?P<valor>[\d.,]+)"
    r"$"
)


PADRAO_LINHA_682 = re.compile(
    r"^"
    r"(?P<mes_folha>\d{2}-\d{2})\s+"
    r"(?P<codigo>\d{6})\s+"
    r"(?P<funcionario>.+?)\s+"
    r"(?P<admissao>\d{2}/\d{2}/\d{4})\s+"
    r"(?P<base>[\d.,]+)\s+"
    r"(?P<quantidade>[\d.,]+)\s+"
    r"(?P<valor>[\d.,]+)"
    r"$"
)


NOMES_MESES = {
    1: "jan",
    2: "fev",
    3: "mar",
    4: "abr",
    5: "mai",
    6: "jun",
    7: "jul",
    8: "ago",
    9: "set",
    10: "out",
    11: "nov",
    12: "dez",
}


# ============================================================
# UTILITÁRIOS
# ============================================================

def normalizar_numero(valor):
    """
    Converte número brasileiro para float.

    1.234,56 -> 1234.56
    146,60   -> 146.60
    """

    valor = valor.strip()

    valor = valor.replace(".", "")
    valor = valor.replace(",", ".")

    try:
        return float(valor)

    except ValueError:
        return None


def extrair_ano_do_arquivo(caminho):

    match = PADRAO_ARQUIVO.match(caminho.name)

    if not match:
        return None

    return int(match.group(1))


def converter_mes_619(mes_folha, ano_relatorio):
    """
    Para a verba 00619:

    01-11 -> dezembro do ano anterior
    02-11 -> janeiro do ano
    03-11 -> fevereiro
    04-11 -> março
    ...
    12-11 -> novembro
    """

    mes_folha = int(mes_folha)

    if mes_folha == 1:
        return 12, ano_relatorio - 1

    return mes_folha - 1, ano_relatorio


# ============================================================
# PDF
# ============================================================

def extrair_texto_pdf(caminho):

    textos = []

    with pdfplumber.open(caminho) as pdf:

        print(
            f"  Páginas: {len(pdf.pages)}"
        )

        for numero, pagina in enumerate(
            pdf.pages,
            start=1
        ):

            texto = pagina.extract_text()

            if texto:

                textos.append(texto)

                print(
                    f"  Página {numero}: "
                    f"{len(texto)} caracteres"
                )

    return "\n".join(textos)


# ============================================================
# VERBA
# ============================================================

def identificar_verba(texto):

    for linha in texto.splitlines():

        linha = linha.strip()

        match = PADRAO_CODIGO_VERBA.search(linha)

        if match:

            codigo = match.group(1).strip()
            descricao = match.group(2).strip()

            return codigo, descricao

    return None, None


# ============================================================
# PARSER 00619
# ============================================================

def extrair_horas_619(texto, ano_relatorio):

    resultados = []

    for linha in texto.splitlines():

        linha = linha.strip()

        if not linha:
            continue

        match = PADRAO_LINHA_619.match(linha)

        if not match:
            continue

        mes_folha = match.group(
            "mes_folha"
        )

        funcionario = match.group(
            "funcionario"
        ).strip()

        quantidade_texto = match.group(
            "quantidade"
        )

        quantidade = normalizar_numero(
            quantidade_texto
        )

        if quantidade is None:
            continue

        mes, ano = converter_mes_619(
            mes_folha[:2],
            ano_relatorio
        )

        resultados.append({
            "funcionario": funcionario,
            "ano": ano,
            "mes": mes,
            "horas": quantidade,
            "mes_folha": mes_folha,
        })

    return resultados


# ============================================================
# PARSER 00682
# ============================================================

def extrair_dias_682(texto, ano_relatorio):

    resultados = []

    for linha in texto.splitlines():

        linha = linha.strip()

        if not linha:
            continue

        match = PADRAO_LINHA_682.match(linha)

        if not match:
            continue

        funcionario = match.group(
            "funcionario"
        ).strip()

        mes_folha = match.group(
            "mes_folha"
        )

        quantidade = normalizar_numero(
            match.group("quantidade")
        )

        valor = normalizar_numero(
            match.group("valor")
        )

        resultados.append({
            "funcionario": funcionario,
            "mes_folha": mes_folha,
            "quantidade": quantidade,
            "valor": valor,
            "ano_relatorio": ano_relatorio,
        })

    return resultados


# ============================================================
# PROCESSAR PDF
# ============================================================

def processar_pdf(caminho, dados_horas):

    print()
    print("=" * 80)
    print(
        f"ARQUIVO: {caminho.name}"
    )
    print("=" * 80)

    ano_relatorio = extrair_ano_do_arquivo(
        caminho
    )

    if ano_relatorio is None:

        print(
            "[IGNORADO] Nome fora do padrão: "
            f"{caminho.name}"
        )

        return

    print(
        f"Ano do relatório: {ano_relatorio}"
    )

    print()
    print("Extraindo texto...")

    try:

        texto = extrair_texto_pdf(caminho)

    except Exception as erro:

        print(
            f"[ERRO] Falha ao ler PDF: {erro}"
        )

        return

    if not texto.strip():

        print(
            "[ERRO] Nenhum texto encontrado."
        )

        return

    codigo_verba, descricao = identificar_verba(
        texto
    )

    print(
        f"Verba: "
        f"{codigo_verba or 'não identificada'}"
    )

    print(
        f"Descrição: "
        f"{descricao or 'não identificada'}"
    )

    # --------------------------------------------------------
    # 00619 - HORAS / FALTAS
    # --------------------------------------------------------

    if codigo_verba == "00619":

        resultados = extrair_horas_619(
            texto,
            ano_relatorio
        )

        print(
            "Lançamentos de horas encontrados: "
            f"{len(resultados)}"
        )

        for item in resultados:

            print(
                f"  {item['funcionario']} | "
                f"{item['mes']:02d}/{item['ano']} | "
                f"{item['horas']:.2f} horas | "
                f"{item['mes_folha']}"
            )

            funcionario = item["funcionario"]
            ano = item["ano"]
            mes = item["mes"]
            horas = item["horas"]

            if funcionario not in dados_horas:

                dados_horas[funcionario] = {}

            if ano not in dados_horas[funcionario]:

                dados_horas[
                    funcionario
                ][ano] = {}

            if mes not in dados_horas[
                funcionario
            ][ano]:

                dados_horas[
                    funcionario
                ][ano][mes] = 0

            dados_horas[
                funcionario
            ][ano][mes] += horas

    # --------------------------------------------------------
    # 00682 - DIAS / FALTAS
    # --------------------------------------------------------

    elif codigo_verba == "00682":

        resultados = extrair_dias_682(
            texto,
            ano_relatorio
        )

        print(
            "Lançamentos de dias encontrados: "
            f"{len(resultados)}"
        )

        for item in resultados:

            print(
                f"  {item['funcionario']} | "
                f"Mês-Folha: {item['mes_folha']} | "
                f"Quantidade: "
                f"{item['quantidade']:.2f} | "
                f"Valor: "
                f"{item['valor']:.2f}"
            )

        print(
            "  [INFO] Verba 00682 não será "
            "lançada na matriz de HORAS."
        )

    else:

        print(
            "[AVISO] Verba desconhecida. "
            "Nenhum lançamento feito."
        )


# ============================================================
# EXCEL
# ============================================================

def criar_excel(dados_horas, caminho_saida):

    if not ARQUIVO_TEMPLATE.exists():

        print()
        print("[ERRO] Template não encontrado:")
        print(f"       {ARQUIVO_TEMPLATE}")

        return False

    print()
    print("Abrindo template...")

    try:

        wb = load_workbook(ARQUIVO_TEMPLATE)

    except Exception as erro:

        print(
            f"[ERRO] Não foi possível abrir "
            f"o template: {erro}"
        )

        return False

    # --------------------------------------------------------
    # Aba
    # --------------------------------------------------------

    if NOME_ABA in wb.sheetnames:

        ws = wb[NOME_ABA]

    else:

        ws = wb.active

        print(
            f"[INFO] Aba '{NOME_ABA}' não encontrada. "
            f"Usando '{ws.title}'."
        )

    # --------------------------------------------------------
    # Formatação
    # --------------------------------------------------------

    preenchimento = PatternFill(
        fill_type="solid",
        fgColor="EB7B71"
    )

    fonte_branca_negrito = Font(
        color="FFFFFF",
        bold=True
    )

    alinhamento_centro = Alignment(
        horizontal="center",
        vertical="center"
    )

    # --------------------------------------------------------
    # Configuração das tabelas
    # --------------------------------------------------------

    COLUNA_INICIAL_BRUTA = 1      # A
    COLUNA_INICIAL_CALCULADA = 17 # Q

    # --------------------------------------------------------
    # Posição inicial
    # --------------------------------------------------------

    linha = 1

    # --------------------------------------------------------
    # Funcionários
    # --------------------------------------------------------

    for funcionario in sorted(dados_horas):

        dados_funcionario = dados_horas[funcionario]

        anos = sorted(dados_funcionario)

        # ====================================================
        # TABELA 1 - HORAS BRUTAS
        # ====================================================

        coluna_inicial = COLUNA_INICIAL_BRUTA

        linha_cabecalho_bruto = linha

        # Funcionário
        ws.cell(
            row=linha,
            column=coluna_inicial,
            value=funcionario
        )

        # Carga horária
        ws.cell(
            row=linha,
            column=coluna_inicial + 1,
            value="Carga horária:"
        )

        ws.cell(
            row=linha,
            column=coluna_inicial + 2,
            value=CARGA_HORARIA_PADRAO
        )

        # Célula da carga horária
        linha_carga_horaria = linha
        coluna_carga_horaria = coluna_inicial + 2

        # Meses
        for mes in range(1, 13):

            coluna = coluna_inicial + 3 + (mes - 1)

            celula = ws.cell(
                row=linha,
                column=coluna,
                value=NOMES_MESES[mes]
            )

            celula.fill = preenchimento
            celula.font = fonte_branca_negrito
            celula.alignment = alinhamento_centro

        linha += 1

        # Primeira linha dos anos
        linha_inicial_bruta = linha

        # Anos + horas
        for ano in anos:

            celula_ano = ws.cell(
                row=linha,
                column=coluna_inicial + 2,
                value=ano
            )

            celula_ano.fill = preenchimento
            celula_ano.font = fonte_branca_negrito
            celula_ano.alignment = alinhamento_centro

            for mes in range(1, 13):

                horas = dados_funcionario.get(
                    ano,
                    {}
                ).get(mes)

                if horas is None:
                    continue

                coluna = (
                    coluna_inicial
                    + 3
                    + (mes - 1)
                )

                ws.cell(
                    row=linha,
                    column=coluna,
                    value=horas
                )

            linha += 1

        # ====================================================
        # TABELA 2 - HORAS / CARGA HORÁRIA
        # ====================================================

        coluna_inicial = COLUNA_INICIAL_CALCULADA

        # Mesmo funcionário
        ws.cell(
            row=linha_cabecalho_bruto,
            column=coluna_inicial,
            value=f"{funcionario} - Horas / Carga horária"
        )

        # Meses
        for mes in range(1, 13):

            coluna = coluna_inicial + 3 + (mes - 1)

            celula = ws.cell(
                row=linha_cabecalho_bruto,
                column=coluna,
                value=NOMES_MESES[mes]
            )

            celula.fill = preenchimento
            celula.font = fonte_branca_negrito
            celula.alignment = alinhamento_centro

        # Primeira linha dos anos calculados
        linha_calculada = linha_inicial_bruta

        for indice, ano in enumerate(anos):

            # Ano
            celula_ano = ws.cell(
                row=linha_calculada,
                column=coluna_inicial + 2,
                value=ano
            )

            celula_ano.fill = preenchimento
            celula_ano.font = fonte_branca_negrito
            celula_ano.alignment = alinhamento_centro

            # Linha correspondente na tabela bruta
            linha_bruta = (
                linha_inicial_bruta + indice
            )

            for mes in range(1, 13):

                coluna_calculada = (
                    coluna_inicial
                    + 3
                    + (mes - 1)
                )

                coluna_bruta = (
                    COLUNA_INICIAL_BRUTA
                    + 3
                    + (mes - 1)
                )

                coluna_bruta_letra = ws.cell(
                    row=1,
                    column=coluna_bruta
                ).column_letter

                coluna_carga_letra = ws.cell(
                    row=1,
                    column=COLUNA_INICIAL_BRUTA + 2
                ).column_letter

                # Exemplo:
                # =SE(D3="";"";TRUNCAR(D3/$C$1))
                #
                # openpyxl recebe a fórmula em inglês:
                # =IF(D3="","",TRUNC(D3/$C$1))

                formula = (
                    f'=IF('
                    f'{coluna_bruta_letra}{linha_bruta}="",'
                    f'"",'
                    f'TRUNC('
                    f'{coluna_bruta_letra}{linha_bruta}/'
                    f'${coluna_carga_letra}$'
                    f'{linha_carga_horaria}'
                    f')'
                    f')'
                )

                ws.cell(
                    row=linha_calculada,
                    column=coluna_calculada,
                    value=formula
                )

            linha_calculada += 1

        # ----------------------------------------------------
        # Espaço entre funcionários
        # ----------------------------------------------------

        linha += 2

    # --------------------------------------------------------
    # Salvar
    # --------------------------------------------------------

    try:

        wb.save(caminho_saida)

    except Exception as erro:

        print(
            f"[ERRO] Não foi possível salvar "
            f"o Excel: {erro}"
        )

        return False

    return True


# ============================================================
# MAIN
# ============================================================

def main():

    if len(sys.argv) < 2:

        print(
            'Uso: python src/main.py "folder"'
        )

        sys.exit(1)

    pasta = Path(sys.argv[1])

    if not pasta.exists():

        print(
            f"Pasta não encontrada: {pasta}"
        )

        sys.exit(1)

    if not pasta.is_dir():

        print(
            f"O caminho não é uma pasta: {pasta}"
        )

        sys.exit(1)

    arquivos_pdf = sorted(
        pasta.glob("*.pdf")
    )

    if not arquivos_pdf:

        print(
            "Nenhum PDF encontrado."
        )

        sys.exit(0)

    print()
    print("=" * 80)
    print("PROCESSAMENTO DE FALTAS")
    print("=" * 80)

    print(
        f"Pasta: {pasta}"
    )

    print(
        f"PDFs encontrados: "
        f"{len(arquivos_pdf)}"
    )

    dados_horas = {}

    for caminho_pdf in arquivos_pdf:

        processar_pdf(
            caminho_pdf,
            dados_horas
        )

    # ========================================================
    # RESUMO
    # ========================================================

    print()
    print("=" * 80)
    print("RESUMO DAS HORAS")
    print("=" * 80)

    total = 0

    for funcionario in sorted(
        dados_horas
    ):

        print()
        print(
            f"FUNCIONÁRIO: {funcionario}"
        )

        for ano in sorted(
            dados_horas[
                funcionario
            ],
            reverse=True
        ):

            for mes in sorted(
                dados_horas[
                    funcionario
                ][ano]
            ):

                horas = dados_horas[
                    funcionario
                ][ano][mes]

                print(
                    f"  {mes:02d}/{ano}: "
                    f"{horas:.2f} horas"
                )

                total += 1

    print()
    print(
        f"Total de lançamentos de horas: "
        f"{total}"
    )

    # ========================================================
    # EXCEL
    # ========================================================

    if dados_horas:

        caminho_saida = (
            pasta / "Faltas.xlsx"
        )

        sucesso = criar_excel(
            dados_horas,
            caminho_saida
        )

        if sucesso:

            print()
            print("=" * 80)
            print(
                f"Excel criado: "
                f"{caminho_saida}"
            )
            print("=" * 80)

    else:

        print()
        print(
            "Nenhum lançamento de HORAS encontrado."
        )


if __name__ == "__main__":
    main()