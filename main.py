from tkinter import *
import xlsxwriter
from openpyxl import load_workbook
import openpyxl
import pandas as pd
import os

class InfoPlanilha:

def arquivo(nome,oqfeito,data):

    file_path = 'planilha.xlsx'

    if os.path.exists(file_path):
        workbook = load_workbook(file_path)
        worksheet = workbook.active
    else:
        workbook = xlsxwriter.Workbook(file_path)
        worksheet = workbook.add_worksheet()
        worksheet.write('A1', 'Nome')
        worksheet.write('B1', 'Serviço Prestado')
        worksheet.write('C1', 'Data')
        workbook.close()
        workbook = load_workbook(file_path)
        worksheet = workbook.active

    existing_file = 'planilha.xlsx'

    #Adicionar as informações na planilha
    new_data = [[nome, oqfeito, data]]
    wb = load_workbook(existing_file)
    ws = wb.active
    for row in new_data:
        ws.append(row)
    wb.save(existing_file)

    workbook.close()

def mostrarplanilha():

    df = pd.read_excel("planilha.xlsx")

    print(df)

print("O que deseja fazer?")
print("1 - Ler planilha \n2 - Inserir dados \n3 - Modificar dados "
      "\n 4 - Criar planilha")
opcao = int(input())
if opcao == 1:
    mostrarplanilha()
elif opcao == 2:
    print("Insira o nome: ")
    nome = str(input())
    print("Insira o serviço prestado: ")
    oqfeito = str(input())
    print("Insira a data que foi feito: ")
    data = str(input())
    arquivo(nome, oqfeito, data)
elif opcao == 3:
    print("Deseja modiciar uma linha (1) ou coluna (2)?")
    opcao = int(input())
    if opcao == 1:
        print("TTT")
    elif opcao == 1:
        print("TTT")
elif opcao == 4:
    print("Quantas colunas deseja criar? ")