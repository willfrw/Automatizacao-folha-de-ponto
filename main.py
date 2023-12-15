import os
from functions import gerador_folhas_col, gerador_pdf, gerador_folha_mensal
from openpyxl import load_workbook
from tkinter.filedialog import askopenfilename, askdirectory


tupla_pastas = ("1 - Janeiro", "2 - Fevereiro", "3 - Março", "4 - Abril", "5 - Maio", "6 - Junho", "7 - Julho", "8 - Agosto", "9 - Setembro", "10 - Outubro", "11 - Novembro", "12 - Dezembro")
dic_setor = {}
lista_caminho = []
lista_nome_arquivo = []
nome_folha_gerada = []
dados = askopenfilename(title='Planilha de Dados')
modelo_folha = askopenfilename(title='Modelo Folha de Ponto')
pasta_destino = askdirectory(title='Pasta de Destino dos Arquivos - Ano')


planilha_setores = load_workbook(dados)
aba_ativa_setores = planilha_setores.active

for celula in aba_ativa_setores["D"]:
    linha_setores = celula.row
    if linha_setores > 1:
        cod_setor = aba_ativa_setores[f'D{linha_setores}'].value
        setor = aba_ativa_setores[f'E{linha_setores}'].value
        if cod_setor != None and cod_setor not in dic_setor:
            dic_setor[cod_setor] = setor
            print(dic_setor[cod_setor])

while True:
    mes = int(input('Digite o mês de referência: '))
    ano = int(input('Digite o ano de referência: '))
    if mes > 0 and mes <= 12:
        break
    else:
        print("Valor inválido! Tente novamente.")


print("Gerando os arquivos xlsx\n")

gerador_folha_mensal(mes, ano, modelo_folha, nome_folha_gerada)
gerador_folhas_col(dados, nome_folha_gerada)

dirlist = os.listdir(".")

for file in dirlist:
    filename = os.path.abspath(file)
    lista_caminho.append(filename)
    if 'xlsx' in file and 'dne' not in file:
        newname = file.replace('xlsx', 'pdf')
        lista_nome_arquivo.append(newname)

print("Convertendo os arquivos para pdf e excluindo os arquivos xlsx\n")

for arquivo_path in lista_caminho:
    # Se o arquivo for .xlsx imprime o caminho do arquivo
    if ".xlsx" in arquivo_path and "dne" not in arquivo_path:
        gerador_pdf(arquivo_path)
        os.remove(arquivo_path)

print("Criando pastas de acordo com o setor")


print("Movendo arquivos para pastas de acordo com o solicitado")

pasta_mes = tupla_pastas[mes-1]

for arquivo_name in lista_nome_arquivo:
    # Se o arquivo for .pdf move o arquivo
    if ".pdf" in arquivo_name:
        pasta_setor = dic_setor[arquivo_name[:3]]
        if os.path.isdir(f"{pasta_destino}//{pasta_mes}//{pasta_setor}")==False:
            os.mkdir(f"{pasta_destino}//{pasta_mes}//{pasta_setor}")

        os.rename(arquivo_name, f"{pasta_destino}//{pasta_mes}//{pasta_setor}//{arquivo_name}") # Tem que arrumar, está criando duas pastas


print("Finalizado")

"""
1. Colocar o projeto
print("Gerador de folha de ponto")

folha_de_ponto = input("Digite o caminho do modelo de folha de ponto: ")

dados_func = input("Digite o caminho da planilha com os dados dos funcionários: ")
"""