import fitz  # PyMuPDF
import os
import pandas as pd
from babel.dates import format_date
from datetime import datetime

# Definir ano e mês de processamento
ano = 2025
mes = 10

# Compor nome do arquivo base
nome_arquivo_base = f"Base-folhaPonto-{ano}-{mes}.xlsx"

# Verificar se o arquivo existe
if not os.path.exists(nome_arquivo_base):
    print(f"Arquivo {nome_arquivo_base} não encontrado.")
else:
    # Ler a aba 'parametros' do arquivo Excel
    df_parametros = pd.read_excel(nome_arquivo_base, sheet_name='parametros', engine='openpyxl')
    print(df_parametros) 
    print(df_parametros.shape)  
    # Obter os valores de AnoBase e MesBase
    ano_base = int(df_parametros.iloc[3,1])
    mes_base = int(df_parametros.iloc[4,1])

    # Verificar se os valores correspondem
    if ano_base != ano or mes_base != mes:
        print(f"Erro: Ano/Mês base do arquivo ({ano_base}/{mes_base}) não corresponde ao esperado ({ano}/{mes}).")
    else:
        # Mensagem de início
        data = datetime(ano, mes, 1)
        nome_mes = format_date(data, "MMMM", locale="pt_BR")
        print(f"Atenção, estamos a processar a folha do mês/ano: {nome_mes}/{ano}")

        # Caminho do modelo PDF
        pdf_modelo = "_ model.pdf"

        # Pasta de saída para os formulários preenchidos
        output_dir = "formularios_preenchidos"
        os.makedirs(output_dir, exist_ok=True)

        # Função para preencher o PDF com dados de uma linha da planilha
        def preencher_pdf(dados, modelo_path, saida_path):
            doc = fitz.open(modelo_path)
            page = doc[0]

            # Linha par Nome do professor
            page.insert_text((80, 135), f"{dados['Nome']}", fontsize=9)
            page.insert_text((360, 135), f"{dados['Matricula']}", fontsize=8)
            page.insert_text((460, 135), f"{dados['Regime']}", fontsize=8)
            page.insert_text((545, 135), f"{dados['Categoria']}", fontsize=8)

            # linha para lista de disciplinas
            disciplinas_texto = ", ".join(dados["Disciplinas"].values())
            page.insert_text((95, 142), disciplinas_texto, fontsize=5)
            page.insert_text((545, 144), f"{dados['CHS']}", fontsize=10)

            # linha hora Atividade - HAE-O - HAE-C 
            page.insert_text((98, 154), f"{dados['HoraAtividade']}", fontsize=9)
            page.insert_text((295, 154), f"{dados['HAE-O']}", fontsize=9)
            page.insert_text((473, 154), f"{dados['HAE-C']}", fontsize=9)

            # Grade Manhã
            for i, dia in enumerate(["SegundaManha", "TercaManha", "QuartaManha", "QuintaManha", "SextaManha", "SabadoManha"]):
                y = 221 + i * 9
                texto = dados["GradeManha"].get(dia, "")
                page.insert_text((95, y), texto, fontsize=9)

            # Grade Tarde
            for i, dia in enumerate(["SegundaTarde", "TercaTarde", "QuartaTarde", "QuintaTarde", "SextaTarde", "SabadoTarde"]):
                y = 221 + i * 9
                texto = dados["GradeTarde"].get(dia, "")
                page.insert_text((294, y), texto, fontsize=9)

            # Grade Noite
            for i, dia in enumerate(["SegundaNoite", "TercaNoite", "QuartaNoite", "QuintaNoite", "SextaNoite", "SabadoNoite"]):
                y = 221 + i * 9
                texto = dados["GradeNoite"].get(dia, "")
                page.insert_text((473, y), texto, fontsize=9)
            
            # Observações da grade
            page.insert_text((55, 273), f"{dados['Observacao1_Grade']}", fontsize=8)
            page.insert_text((245, 273), f"{dados['Observacao2_Grade']}", fontsize=8)
            page.insert_text((435, 273), f"{dados['Observacao3_Grade']}", fontsize=8)

            # Salvar o novo PDF
            doc.save(saida_path)
            doc.close()

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

        # Caminho de saída para o PDF preenchido
        saida_pdf = os.path.join(output_dir, f"{dados_exemplo['Nome'].replace(' ', '_')}_formulario.pdf")

        # Executar a função para preencher o PDF
        preencher_pdf(dados_exemplo, pdf_modelo, saida_pdf)

        # Mensagem de sucesso
        print(f"Formulário preenchido salvo em: {saida_pdf}")