try:
    import xlsxwriter
    from openpyxl import load_workbook
    import pandas as pd
    import os
    import tkinter as tk
    from tkinter import filedialog
except Exception as error:
    print(error)
    print("Houve um erro na hora de importar a dependências. Por favor, os instale com o comando a seguir no seu CMD/Terminal: "
    "\n pip install -r requirements.txt \n\nSe isto não funcionar, tente executar o arquivo normalmente com duplo click invés do terminal.")
    input()
else:
    
    def ler_planilha(filepath):

        df = pd.read_excel(filepath)
        print(f'\n {df}')
        input("\n\nPressione qualquer tecla para continuar\n")

    def carregarplanilha(filepath):

        try:
            if os.path.exists(filepath):
                book = load_workbook(filepath)
                sheet = book.active
                return True
            else:
                return False
        except Exception as error:
            print(error)
            print("Aparentemente ocorreu um erro na hora de carregar a planilha. Se certifique que o arquivo foi criado de forma correta \n"
            "Neste caso se certifique que o arquivo foi criado por algum programa ou software de maneira correta. (Criar planilha -> Salvar como: ex_nome.xlsx)")
            input()

    def analisar_planilha(filepath, tipo):

        arquivo = pd.read_excel(filepath)
        rowcol = None
        while True:
            print("Deseja selecionar: Linha (1) | Coluna (2) | Documento inteiro (3)")
            while True:
                try:
                    opc = int(input())
                except ValueError:
                    print("Valor inválido, tente novamente:")
                else:
                    break
            if opc == 1:
                print("Insira a linha: (0, 1, 2, ...)")
                rowcol = int(input())
                break
            elif opc == 2:
                print("Insira o nome ou letra associada com a coluna: ")
                rowcol = str(input())
                break
            elif opc == 3:
                break
            else:
                print("Inválido, tente novamente:")

        if tipo == 1:
            print("Tipo de expressão matemática: ")
            tipo_res = str(input())
            if opc == 1:
                tipo_res = tipo_res + f" na linha {rowcol}"
            if opc == 2:
                tipo_res = tipo_res + f" na coluna {rowcol}"
        if tipo == 2:
            print("Qual o tipo de análise que tem que ser feita?")
            tipo_res = "A seguinte análise: "+str(input())
            if opc == 1:
                tipo_res = tipo_res + f" na linha {rowcol}"
            if opc == 2:
                tipo_res = tipo_res + f" na coluna {rowcol}"
        if tipo == 3:
            print("Qual descrepância deve ser achada no documento? ")
            tipo_res = "o seguinte, ache a seguinte descrepância: "+str(input())
            if opc == 1:
                tipo_res = tipo_res + f" na linha {rowcol}"
            if opc == 2:
                tipo_res = tipo_res + f" na coluna {rowcol}"
        
        print("Input para o GPT: \n \n"
            "Dado a seguinte planilha: \n \n"
            f"{arquivo} \n \n"
            f"Faça {tipo_res}")

        input()

#======================#
#[ Começo do programa ]#
#======================#

    def passo_escolher(filepath):
        while True:
            print("O que deseja fazer?")
            print("1 - Ler planilha \n"
                "2 - Expressão Matemática \n"
                "3 - Análise de dados \n"
                "4 - Descrepâncias \n"
                "5 - Sair \n")

            while True:
                try:
                    opcao = int(input("Insira a opção: "))
                except ValueError:
                    print("Valor inválido, tente novamente:")
                else:
                    break
                
            if opcao == 1:
                ler_planilha(filepath)
            elif opcao == 2:
                analisar_planilha(filepath, 1)
            elif opcao == 3:
                analisar_planilha(filepath, 2)
            elif opcao == 4:
                analisar_planilha(filepath, 3)
            else:
                break

    def GPTxlsx_main():
        def selecionar_arquivo():
            root = tk.Tk()
            root.withdraw()

            caminho_arquivo = filedialog.askopenfilename(title="Selecione a planilha")
            
            return caminho_arquivo

        filepath = selecionar_arquivo()

        if filepath:
            print(f"Arquivo selecionado: {filepath}")
        else:
            print("Nenhum arquivo selecionado.")

        while True:
            if filepath == '0':
                break
            else:
                if carregarplanilha(filepath):
                    passo_escolher(filepath)
                else:
                    print("Planilha inexistente.")
                    input()