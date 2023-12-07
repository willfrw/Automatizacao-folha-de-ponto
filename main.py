from functions import gerador_folhas_col, gerador_pdf
import os

print("Gerando")
gerador_folhas_col("modelo-dados-teste.xlsx", "modelo-tabela-folha-ponto-teste.xlsx")

lista_caminho = []

dirlist = os.listdir(".")

for file in dirlist:
    filename = os.path.abspath(file)
    lista_caminho.append(filename)

for arquivo in lista_caminho:
    # Se o arquivo for .xlsx imprime o caminho do arquivo
    if arquivo[-5:] == ".xlsx" and arquivo[-10:] != "teste.xlsx":
        gerador_pdf(arquivo)
        os.remove(arquivo)

print("Finalizado")
# Falta colocar mudança de diretório para salvamento correto dos arquivos

"""
1. Colocar o projeto
"""

"""print("Gerador de folha de ponto")

folha_de_ponto = input("Digite o caminho do modelo de folha de ponto: ")

dados_func = input("Digite o caminho da planilha com os dados dos funcionários: ")
"""