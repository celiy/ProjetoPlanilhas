import openpyxl
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
import random
import string

def gerar_string_aleatoria(tamanho=8):
    return ''.join(random.choices(string.ascii_uppercase + string.digits, k=tamanho))

def gerar_numero_aleatorio(min_val=1, max_val=1000):
    return random.randint(min_val, max_val)

def gerar_data_aleatoria():
    start_date = datetime(2020, 1, 1)
    end_date = datetime(2024, 12, 31)
    delta = end_date - start_date
    random_days = random.randint(0, delta.days)
    random_date = start_date + timedelta(days=random_days)
    return random_date

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Dados Aleatórios"

headers = ['ID', 'Nome', 'Idade', 'Data de Cadastro', 'Saldo', 'Produto', 'Quantidade', 'Categoria']
ws.append(headers)

for i in range(1, 1001):
    id_ = i
    nome = gerar_string_aleatoria()
    idade = gerar_numero_aleatorio(18, 90)
    data_cadastro = gerar_data_aleatoria().strftime("%d/%m/%Y")
    saldo = round(random.uniform(10, 1000), 2)
    produto = gerar_string_aleatoria(6)
    quantidade = gerar_numero_aleatorio(1, 50)
    categoria = random.choice(['Eletrônicos', 'Roupas', 'Alimentos', 'Móveis', 'Beleza'])

    ws.append([id_, nome, idade, data_cadastro, saldo, produto, quantidade, categoria])

wb.save('planilha_aleatoria.xlsx')

print("Planilha 'planilha_aleatoria.xlsx' criada com sucesso!")
