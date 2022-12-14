import sys
from datetime import datetime
from time import sleep
from openpyxl.styles import Font
from openpyxl.workbook import Workbook
from selenium import webdriver
from selenium.webdriver.common.by import By


class AvaliaPrecos:
    def extrai_preco_como_float(self, preco_string):
        return float(preco_string.replace('R$ ', '').replace('.', '').replace(',', '.'))

    def media_entre_valores(self, skus):
        total_precos = 0

        for x in range(len(skus)):
            total_precos += self.extrai_preco_como_float(skus[x][1])

        media_precos = total_precos / 5

        for x in range(len(skus)):
            skus[x].append(media_precos)

    def acessa_loja(self, produtos):
        driver = webdriver.Chrome()
        try:
            driver.get("https://www.americanas.com.br/")
            log = open("log.txt", 'a')
            log.write(datetime.now().strftime('%d/%m/%y %H:%M:%S') +
                      " - Conexao com o site feita com Sucesso!\n")
        except Exception:
            log = open("log.txt", 'a')
            log.write(datetime.now().strftime('%d/%m/%y %H:%M:%S') +
                      " - Erro {} ao tentar acessar o site!\n".format(sys.exc_info()[0]))
            return

        lista_todos_produtos = []
        for x in range(len(produtos)):
            elemento = driver.find_element(By.XPATH,
                                           """//*[@id="rsyswpsdk"]/div/header/div[1]/div[1]/div/div[1]/form/input""")
            pesquisa = driver.find_element(By.XPATH,
                                           """//*[@id="rsyswpsdk"]/div/header/div[1]/div[1]/div/div[1]/form/button""")
            elemento.clear()
            elemento.send_keys(produtos[x])
            driver.execute_script("arguments[0].click()", pesquisa)
            sleep(5)

            skus = []
            i = 1
            while i <= 5:
                sku = []
                nome_produto = driver.find_element(By.XPATH,
                                                   """//*[@id="rsyswpsdk"]/div/main/div/div[3]/div[2]/div[{}]/div/div/a/div[2]/div[2]/h3""".format(i))
                valor_produto = driver.find_element(By.XPATH,
                                                    """//*[@id="rsyswpsdk"]/div/main/div/div[3]/div[2]/div[{}]/div/div/a/div[3]/span[1]""".format(i))
                sku.append(nome_produto.text)
                sku.append(valor_produto.text)
                skus.append(sku)
                i += 1

            self.media_entre_valores(skus)
            lista_todos_produtos.append(skus)
        return lista_todos_produtos

    def gerador_relatorio_excel(self, produtos):
        workbook = Workbook()
        sheet = workbook.active
        sheet['A1'] = "Produtos"
        sheet['B1'] = "SKUs"
        sheet['C1'] = "Pre??os"
        sheet['D1'] = "M??dia dos Pre??os"
        sheet['A1'].font = Font(bold=True)
        sheet['B1'].font = Font(bold=True)
        sheet['C1'].font = Font(bold=True)
        sheet['D1'].font = Font(bold=True)

        query_results = self.acessa_loja(produtos)

        row = 2
        for x in range(len(query_results)):
            produto = query_results[x]
            for y in range(len(produto)):
                sheet["A{}".format(row)] = produtos[x]
                sheet["B{}".format(row)] = produto[y][0]
                sheet["C{}".format(row)] = produto[y][1]
                sheet["D{}".format(row)] = produto[y][2]
                row += 1

        try:
            workbook.save("relatorio.xlsx")
            log = open("log.txt", 'a')
            log.write(datetime.now().strftime('%d/%m/%y %H:%M:%S') +
                      " - Criado relatorio final com sucesso!\n")
        except TypeError():
            log = open("log.txt", 'a')
            log.write(datetime.now().strftime('%d/%m/%y %H:%M:%S') +
                      " - Erro {} ao salvar relat??rio!\n".format(sys.exc_info()[0]))
            return
