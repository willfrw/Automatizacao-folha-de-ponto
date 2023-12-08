import os
from openpyxl import Workbook, load_workbook
import win32com.client
from datetime import datetime, timedelta

def gerador_folhas_col(dados, modelo_folha_ponto):
    planilha_nomes = load_workbook(dados)

    aba_ativa_nomes = planilha_nomes.active
    for celula in aba_ativa_nomes["A"]:
        planilha_folha_ponto = load_workbook(modelo_folha_ponto)
        aba_ativa_folha = planilha_folha_ponto.active

        linha_nome = celula.row
        if linha_nome > 1:
            setor = aba_ativa_nomes[f"D{linha_nome}"].value
            nome = aba_ativa_nomes[f"B{linha_nome}"].value
            cargo = aba_ativa_nomes[f"C{linha_nome}"].value
            mat = aba_ativa_nomes[f"A{linha_nome}"].value

            aba_ativa_folha["C4"] = "NOME: " + str(nome)
            aba_ativa_folha["C5"] = "CARGO: " + str(cargo)
            aba_ativa_folha["G5"] = "MATRÍCULA: " + str(mat)

            nome_planilha = str(f"{setor}_{nome}.xlsx")
            planilha_folha_ponto.save(nome_planilha)
    return 0

# A função irá criar um arquivo em pdf com base no excel
# Criar um loop for com o tamanho da lista com os arquivo para criar conforme percorsse o diretório
def gerador_pdf(path):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    workbook = excel.Workbooks.Open(path)
    path_pdf = path.replace(".xlsx", ".pdf")

    try:
        workbook.ActiveSheet.ExportAsFixedFormat(0, path_pdf)
    except Exception as e:
        print(f"Erro ao exportar {path} para PDF: {str(e)}")
    finally:
        workbook.Close(False)
        excel.Quit()

    return 0



"""def gerador_folha_mensal(modelo_folha_ponto):
    print("Teste")

meses = {"jan":("01/01/2024")}"""




