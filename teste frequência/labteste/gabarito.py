import fitz  # PyMuPDF

# Abrir o PDF modelo
doc = fitz.open("_ model.pdf")
page = doc[0]

# Tamanho da página
width, height = page.rect.width, page.rect.height

# Espaçamento entre os pontos
spacing = 20

# Inserir pontos em grade
for y in range(0, int(height), spacing):
    for x in range(0, int(width), spacing):
        page.insert_text((x, y), "*", fontsize=5, color=(0, 0, 0))

# Salvar o novo PDF com a máscara
doc.save("mascara_grade.pdf")
doc.close()