from avaliaprecos import AvaliaPrecos
import pandas


def get_valores_excel(excel_sheet):
    data_frame = (pandas.read_excel(excel_sheet)).values.tolist()
    lista_produtos = []
    for x in range(len(data_frame)):
        lista_produtos.append(data_frame[x][0])
    return lista_produtos


excel = "GoLiveTech - Processo Seletivo - Desafio Técnico - Exemplo.xlsx"
produtos = get_valores_excel(excel)
avaliaprecos = AvaliaPrecos()
avaliaprecos.gerador_relatorio_excel(produtos)
