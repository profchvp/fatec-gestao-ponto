import fitz  # PyMuPDF
import os
import pandas as pd
from babel.dates import format_date
from datetime import datetime

# ===============================
# Constantes globais de referência
# ===============================
ANO_REFERENCIA = 2025
MES_REFERENCIA = 10
#Linha onde inicia a lista de professores 
INICIO_DADOS_PROFESSOR=10


def inicializar_programa():
    """
    Executa os procedimentos iniciais:
    - Usa as constantes ANO_REFERENCIA e MES_REFERENCIA
    - Verifica existência do arquivo Excel
    - Lê e valida a aba 'parametros'
    Retorna (ANO_REFERENCIA, MES_REFERENCIA, df_parametros, dicionario_dados) se tudo estiver OK, caso contrário retorna None.
    """
    nome_arquivo_base = f"Base-folhaPonto-{ANO_REFERENCIA}-{MES_REFERENCIA}.xlsx"

    if not os.path.exists(nome_arquivo_base):
        print(f"Arquivo {nome_arquivo_base} não encontrado.")
        return None

    df_parametros = pd.read_excel(nome_arquivo_base, sheet_name='parametros', engine='openpyxl')
    #............................print(df_parametros)
    #............................print(df_parametros.shape)
    # Obter valores base da planilha
    ano_base = int(df_parametros.iloc[3, 1])
    mes_base = int(df_parametros.iloc[4, 1])

    if ano_base != ANO_REFERENCIA or mes_base != MES_REFERENCIA:
        print(f"Erro: Ano/Mês base do arquivo ({ano_base}/{mes_base}) "
              f"não corresponde ao esperado ({ANO_REFERENCIA}/{MES_REFERENCIA}).")
        return None

    # Mensagem de início
    data = datetime(ANO_REFERENCIA, MES_REFERENCIA, 1)
    nome_mes = format_date(data, "MMMM", locale="pt_BR")
    print(f"Atenção, estamos a processar a folha do mês/ano: {nome_mes}/{ANO_REFERENCIA}")

    # Extrair dados a partir da linha 12 (índice 11) das colunas A, B, C e D
    df_dados = df_parametros.iloc[INICIO_DADOS_PROFESSOR:, [0, 1, 2, 3]]
    df_dados.columns = ['Sequencia', 'Matricula', 'Nome', 'Nome da Pasta']
    df_dados = df_dados.dropna(how='all')  # Remove linhas completamente vazias

    # Criar dicionário com os dados
    dicionario_dados = df_dados.to_dict(orient='records')

    return ANO_REFERENCIA, MES_REFERENCIA, df_parametros, dicionario_dados


def preencher_pdf(dados, modelo_path, saida_path):
    """Preenche o modelo PDF com os dados de um dicionário."""
    doc = fitz.open(modelo_path)
    page = doc[0]

    # Cabeçalho
    page.insert_text((80, 135), f"{dados['Nome']}", fontsize=9)
    page.insert_text((360, 135), f"{dados['Matricula']}", fontsize=8)
    page.insert_text((460, 135), f"{dados['Regime']}", fontsize=8)
    page.insert_text((545, 135), f"{dados['Categoria']}", fontsize=8)

    # Disciplinas e carga horária
    disciplinas_texto = ", ".join(dados["Disciplinas"].values())
    page.insert_text((95, 142), disciplinas_texto, fontsize=5)
    page.insert_text((545, 144), f"{dados['CHS']}", fontsize=10)

    # Hora Atividade / HAE
    page.insert_text((98, 154), f"{dados['HoraAtividade']}", fontsize=9)
    page.insert_text((295, 154), f"{dados['HAE-O']}", fontsize=9)
    page.insert_text((473, 154), f"{dados['HAE-C']}", fontsize=9)

    # Grade - manhã / tarde / noite
    for i, dia in enumerate(["SegundaManha", "TercaManha", "QuartaManha", "QuintaManha", "SextaManha", "SabadoManha"]):
        y = 221 + i * 9
        page.insert_text((95, y), dados["GradeManha"].get(dia, ""), fontsize=9)

    for i, dia in enumerate(["SegundaTarde", "TercaTarde", "QuartaTarde", "QuintaTarde", "SextaTarde", "SabadoTarde"]):
        y = 221 + i * 9
        page.insert_text((294, y), dados["GradeTarde"].get(dia, ""), fontsize=9)

    for i, dia in enumerate(["SegundaNoite", "TercaNoite", "QuartaNoite", "QuintaNoite", "SextaNoite", "SabadoNoite"]):
        y = 221 + i * 9
        page.insert_text((473, y), dados["GradeNoite"].get(dia, ""), fontsize=9)

    # Observações
    page.insert_text((55, 273), f"{dados['Observacao1_Grade']}", fontsize=8)
    page.insert_text((245, 273), f"{dados['Observacao2_Grade']}", fontsize=8)
    page.insert_text((435, 273), f"{dados['Observacao3_Grade']}", fontsize=8)

    # Salvar resultado
    doc.save(saida_path)
    doc.close()


def processamento_central():
    """Executa o processamento principal do programa."""
    resultado = inicializar_programa()
    if resultado is None:
        return

    ano, mes, df_parametros, dicionario_dados = resultado

    # Exibir o conteúdo do dicionário extraído da planilha
    print("Conteúdo extraído da planilha:")
    for item in dicionario_dados:
        print(item)

    pdf_modelo = "_ model.pdf"
    output_dir = "formularios_preenchidos"
    os.makedirs(output_dir, exist_ok=True)

    # Exemplo de dados para preenchimento
    dados_exemplo = {
        "Nome": "Carlos Henrique",
        "Matricula": "123456",
        "Regime": "CLT",
        "Categoria": "Docente",
        "Disciplinas": {
            "1": "Estrutura de Dados",
            "2": "Banco de Dados"
        },
        "CHS": "20",
        "HoraAtividade": "5",
        "HAE-O": "2",
        "HAE-C": "1",
        "GradeManha": {
            "SegundaManha": "X      X       X       X      X       X",
            "TercaManha":   "X      X       X       X      X       X",
            "QuartaManha":  "X      X       X       X      X       X",
            "QuintaManha":  "X      X       X       X      X       X",
            "SextaManha":   "X      X       X       X      X       X",
            "SabadoManha":  "X      X       X       X      X       X"
        },
        "GradeTarde": {
            "SegundaTarde": "X      X      X      X     X      X",
            "TercaTarde":   "X      X      X      X     X      X",
            "QuartaTarde":  "X      X      X      X     X      X",
            "QuintaTarde":  "X      X      X      X     X      X",
            "SextaTarde":   "X      X      X      X     X      X",
            "SabadoTarde":  "X      X      X      X     X      X"
        },
        "GradeNoite": {
            "SegundaNoite": "X      X       X       X",
            "TercaNoite":   "X      X       X       X",
            "QuartaNoite":  "X      X       X       X",
            "QuintaNoite":  "X      X       X       X",
            "SextaNoite":   "X      X       X       X",
            "SabadoNoite":  "X      X       X       X"
        },
        "Observacao1_Grade": "esta é a observacao #1 da grade",
        "Observacao2_Grade": "esta é a observacao #2 da grade",
        "Observacao3_Grade": "esta é a observacao #3 da gradex",
    }

    saida_pdf = os.path.join(output_dir, f"{dados_exemplo['Nome'].replace(' ', '_')}_formulario.pdf")
    preencher_pdf(dados_exemplo, pdf_modelo, saida_pdf)
    print(f"Formulário preenchido salvo em: {saida_pdf}")


def main():
    """Procedimento principal."""
    print("=== Iniciando processamento da folha de ponto ===")
    processamento_central()
    print("=== Fim do processamento ===")


# Execução direta
if __name__ == "__main__":
    main()