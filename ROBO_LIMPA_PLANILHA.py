#21/10/2023
#@PLima
#ROBO - LIMPA PLANILHA
import openpyxl
import string
import unicodedata

caracteres_a_substituir = {
    "Á": "A",
    "Â": "A",
    "À": "A",
    "Ã": "A",    
    "É": "E",
    "Ê": "E",
    "Í": "I",
    "Î": "I",
    "Ó": "O",
    "Ô": "O",
    "Õ": "O",
    "Ú": "U",
    "Û": "U",
    "Ü": "U",
    "Ç": "C",
    "#": "",
    "*": "",
    "!": "",
    "@": "",
    "#": "",
    "$": "",
    "%": "",
    "&": ""
} 

# Retira os acentos das vogais
def remover_acentos(texto):
    for caractere_original, caractere_substituido in caracteres_a_substituir.items():
        texto = texto.strip()
        texto = texto.upper()
        texto = texto.replace(caractere_original, caractere_substituido)
    return texto

print("================================= INICIALIZADO ======================")

# Abre o arquivo Excel
print("Abre o arquivo Excel;")
wb = openpyxl.load_workbook("planilha.xlsx")

# Seleciona a planilha
print("Seleciona a planilha")
sheet = wb["aba"]

# Cria uma lista com os nomes da coluna A
print("Cria uma lista com os nomes da coluna;")
coluna_1 = []
coluna_2 = []
for i in range(1, sheet.max_row + 1):
    #print(sheet["A" + str(i)].value)
    coluna_1.append(remover_acentos(sheet["A" + str(i)].value))  
    coluna_2.append(remover_acentos(sheet["D" + str(i)].value))  

# Salva os nomes sem caracteres especiais na planilha
print("Salva os nomes sem caracteres especiais na planilha;\n")
for i in range(1, sheet.max_row + 1):
    #print( nomes_limpos[i - 1])
    sheet["A" + str(i)].value = coluna_1[i - 1]
    sheet["D" + str(i)].value = coluna_2[i - 1]
    #print(coluna_1[i - 1] + " - " + coluna_2[i - 1])

# Salva o arquivo Excel
wb.save("planilha_limpa_.xlsx")
print("wb.save(planilha_limpa_.xlsx);")
print("================================= FINALIZADO ======================")