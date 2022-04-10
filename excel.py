from openpyxl import load_workbook

"""
Colunas solicitante => a h i
Colunas solicitado  => b k l
"""
def abrir_excel():
    wb = load_workbook(filename=r"C:\Users\jbrag\PycharmProjects\python_word\pasta_troca_excel\trocas_10-04-2022.xlsx")
    planilha = wb["Report"]
    lista = list()
    for i in range(2, 20):
        lista_troca = list()
        if planilha[f"A{i}"].value is None:
            break
        lista_troca.append(planilha[f"A{i}"].value)
        lista_troca.append(tratamento_data(planilha[f"H{i}"].value))
        lista_troca.append(planilha[f"I{i}"].value)
        lista_troca.append(planilha[f"B{i}"].value)
        lista_troca.append(tratamento_data(planilha[f"K{i}"].value))
        lista_troca.append(planilha[f"L{i}"].value)
        lista.append(lista_troca)
    return lista


def tratamento_data(data):
    data_lista = data.split('-')
    return f"{data_lista[2]}/{data_lista[1]}/{data_lista[0]}"
