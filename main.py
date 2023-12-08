from functions import gerador_folhas_col, gerador_pdf
from openpyxl import load_workbook
import os

tupla_pastas = ("EXERCÍCIO 2024//1 - Janeiro", "EXERCÍCIO 2024//2 - Fevereiro", "EXERCÍCIO 2024//3 - Março", "EXERCÍCIO 2024//4 - Abril", "EXERCÍCIO 2024//5 - Maio", "EXERCÍCIO 2024//6 - Junho", "EXERCÍCIO 2024//7 - Julho", "EXERCÍCIO 2024//8 - Agosto", "EXERCÍCIO 2024//9 - Setembro", "EXERCÍCIO 2024//10 - Outubro", "EXERCÍCIO 2024//11 - Novembro", "EXERCÍCIO 2024//12 - Dezembro")
dic_setor = {}
lista_caminho = []
lista_nome_arquivo = []

planilha_setores = load_workbook("modelo-dados-teste.xlsx")
aba_ativa_setores = planilha_setores.active

for celula in aba_ativa_setores["D"]:
    linha_setores = celula.row
    if linha_setores > 1:
        cod_setor = aba_ativa_setores[f'D{linha_setores}'].value
        setor = aba_ativa_setores[f'E{linha_setores}'].value
        if cod_setor != None:
            dic_setor[cod_setor] = setor

while True:
    pasta_salvar = int(input("Digite o número correspondente a pasta que quer salvar o arquivo gerado:\n1 - Janeiro\n2 - Fevereiro\n3 - Março\n4 - Abril\n5 - Maio\n6 - Junho\n7 - Julho\n8 - Agosto\n9 - Setembro\n10 - Outubro\n11 - Novembro\n12 - Dezembro\n"))
    if pasta_salvar > 0 and pasta_salvar <= 12:
        break
    else:
        print("Valor inválido! Tente novamente.")

pasta_salvar -= 1


print("Gerando os arquivos xlsx\n")
gerador_folhas_col("modelo-dados-teste.xlsx", "modelo-tabela-folha-ponto-teste.xlsx")

dirlist = os.listdir(".")

for file in dirlist:
    filename = os.path.abspath(file)
    lista_caminho.append(filename)
    if 'xlsx' in file and 'teste' not in file:
        newname = file.replace('xlsx', 'pdf')
        lista_nome_arquivo.append(newname)

print("Convertendo os arquivos para pdf e excluindo os arquivos xlsx\n")

for arquivo_path in lista_caminho:
    # Se o arquivo for .xlsx imprime o caminho do arquivo
    if ".xlsx" in arquivo_path and "teste" not in arquivo_path:
        gerador_pdf(arquivo_path)
        os.remove(arquivo_path)

print("Criando pastas de acordo com o setor")


print("Movendo arquivos para pastas de acordo com o solicitado")

pasta_mes = tupla_pastas[pasta_salvar]

for arquivo_name in lista_nome_arquivo:
    # Se o arquivo for .pdf move o arquivo
    if "pdf" in arquivo_name:
        pasta_setor = dic_setor[arquivo_name[:3]]
        if os.path.isdir(f"{pasta_mes}//{pasta_setor}")==False:
            os.mkdir(f"{pasta_mes}//{pasta_setor}")

        os.rename(arquivo_name, f"{pasta_mes}//{pasta_setor}//{arquivo_name}")


print("Finalizado")

"""
1. Colocar o projeto
print("Gerador de folha de ponto")

folha_de_ponto = input("Digite o caminho do modelo de folha de ponto: ")

dados_func = input("Digite o caminho da planilha com os dados dos funcionários: ")
"""