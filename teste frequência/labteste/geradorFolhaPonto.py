import fitz  # PyMuPDF
import os
import pandas as pd
from babel.dates import format_date
from datetime import datetime

# ===============================
# Constantes globais
# ===============================
ANO_REFERENCIA = 2025
MES_REFERENCIA = 10


def inicializar_programa():
    """Verifica arquivo e parâmetros iniciais e retorna lista de abas válidas."""
    nome_arquivo_base = f"Base-folhaPonto-{ANO_REFERENCIA}-{MES_REFERENCIA}.xlsx"

    if not os.path.exists(nome_arquivo_base):
        print(f"Arquivo {nome_arquivo_base} não encontrado.")
        return None, None

    # Ler aba parametros
    df_parametros = pd.read_excel(nome_arquivo_base, sheet_name='parametros', engine='openpyxl')

    ano_base = int(df_parametros.iloc[3, 1])
    mes_base = int(df_parametros.iloc[4, 1])

    if ano_base != ANO_REFERENCIA or mes_base != MES_REFERENCIA:
        print(f"Erro: Ano/Mês base do arquivo ({ano_base}/{mes_base}) "
              f"não corresponde ao esperado ({ANO_REFERENCIA}/{MES_REFERENCIA}).")
        return None, None

    # Extrair lista de abas válidas
    if "Nome da Aba" not in df_parametros.columns:
        print("Erro: coluna 'Nome da Aba' não encontrada na aba 'parametros'.")
        return None, None

    abas_validas = df_parametros["Nome da Aba"].dropna().tolist()

    # Mensagem de início
    data = datetime(ANO_REFERENCIA, MES_REFERENCIA, 1)
    nome_mes = format_date(data, "MMMM", locale="pt_BR")
    print(f"Atenção, estamos a processar a folha do mês/ano: {nome_mes}/{ANO_REFERENCIA}")
    print(f"Abas a processar: {abas_validas}")

    return nome_arquivo_base, abas_validas


def excel_to_idx(cell):
    """Converte coordenada Excel (ex: B19) para índice (linha, coluna)."""
    col = cell[0]
    row = int(cell[1:])
    col_idx = ord(col.upper()) - ord('A')
    return row - 1, col_idx


def montar_grade(df_aba):
    """
    Lê as matrizes da grade horária (manhã, tarde, noite) da aba do professor.
    Cada matriz é 6x6 e será armazenada como tupla de tuplas.
    Se não houver dados suficientes, preenche com vazio.
    """
    coords = {
        "GradeManha": ("B19", "G24"),
        "GradeTarde": ("H19", "M24"),
        "GradeNoite": ("N19", "Q24")
    }

    max_rows, max_cols = df_aba.shape
    dados_grade = {}

    for turno, (inicio, fim) in coords.items():
        r1, c1 = excel_to_idx(inicio)
        r2, c2 = excel_to_idx(fim)

        matriz = []
        for r in range(r1, r2 + 1):
            linha = []
            for c in range(c1, c2 + 1):
                if r < max_rows and c < max_cols:
                    val = df_aba.iloc[r, c]
                    linha.append(str(val).strip() if pd.notna(val) else "")
                else:
                    linha.append("")  # célula inexistente -> vazio
            matriz.append(tuple(linha))
        dados_grade[turno] = tuple(matriz)

    return dados_grade


def preencher_pdf(dados, modelo_path, saida_path):
    """Preenche o modelo PDF com os dados de um dicionário."""
    doc = fitz.open(modelo_path)
    page = doc[0]

    # Cabeçalho
    page.insert_text((80, 135), f"{dados['Nome']}", fontsize=9)
    page.insert_text((360, 135), f"{dados['Matricula']}", fontsize=8)
    page.insert_text((460, 135), f"{dados['Regime']}", fontsize=8)
    page.insert_text((545, 135), f"{dados['Categoria']}", fontsize=8)

    # Disciplinas
    disciplinas_texto = ", ".join(dados["Disciplinas"].values())
    page.insert_text((95, 142), disciplinas_texto, fontsize=5)
    page.insert_text((545, 144), f"{dados['CHS']}", fontsize=10)

    # Observações
    page.insert_text((55, 273), f"{dados['Observacao1_Grade']}", fontsize=8)
    page.insert_text((245, 273), f"{dados['Observacao2_Grade']}", fontsize=8)
    page.insert_text((435, 273), f"{dados['Observacao3_Grade']}", fontsize=8)

    doc.save(saida_path)
    doc.close()


def processamento_central():
    nome_arquivo_base, abas_validas = inicializar_programa()
    if nome_arquivo_base is None or not abas_validas:
        return

    xls = pd.ExcelFile(nome_arquivo_base)

    for aba in abas_validas:
        if aba not in xls.sheet_names:
            print(f"Aba '{aba}' não encontrada no arquivo. Pulando...")
            continue

        df_aba = pd.read_excel(nome_arquivo_base, sheet_name=aba, engine='openpyxl')

        # Montar grade horária
        dados_grade = montar_grade(df_aba)

        # Exibir resultado no console
        print(f"\nProfessor: {aba}")
        for turno, matriz in dados_grade.items():
            print(f"\n{turno}:")
            for linha in matriz:
                print(linha)

        # Preparar dados para PDF (grade compactada)
        dados_pdf = {
            "Nome": aba,
            "Matricula": "123456",
            "Regime": "CLT",
            "Categoria": "Docente",
            "Disciplinas": {"1": "Estrutura de Dados"},
            "CHS": "20",
            "HoraAtividade": "5",
            "HAE-O": "2",
            "HAE-C": "1",
            "Observacao1_Grade": "",
            "Observacao2_Grade": "",
            "Observacao3_Grade": ""
        }

        # Adicionar grade como texto (cada linha concatenada)
        for turno, matriz in dados_grade.items():
            dados_pdf[turno] = {f"Linha{i+1}": " | ".join(linha) for i, linha in enumerate(matriz)}

        # Gerar PDF
        pdf_modelo = "_ model.pdf"
        output_dir = "formularios_preenchidos"
        os.makedirs(output_dir, exist_ok=True)
        saida_pdf = os.path.join(output_dir, f"{aba.replace(' ', '_')}_formulario.pdf")
        preencher_pdf(dados_pdf, pdf_modelo, saida_pdf)
        print(f"PDF gerado: {saida_pdf}")


def main():
    print("=== Iniciando processamento da folha de ponto ===")
    processamento_central()
    print("=== Fim do processamento ===")


if __name__ == "__main__":
    main()