def calcular_coordenada(indice_caractere, caracteres_por_linha=30, espacamento=20):
    indice = indice_caractere - 1
    linha = indice // caracteres_por_linha
    coluna = indice % caracteres_por_linha
    x = coluna * espacamento
    y = linha * espacamento
    return (x, y)
coordenada = calcular_coordenada(244)
print(coordenada)  # Saída: (60, 140)