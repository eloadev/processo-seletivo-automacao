from datetime import datetime
from time import sleep
from openpyxl.styles import Font
from openpyxl.workbook import Workbook
from selenium import webdriver
from selenium.webdriver.common.by import By


class AvaliaPrecos:
    @staticmethod
    def media_entre_valores(skus):
        media_precos = 0

        for x in range(len(skus)):
            refatorado = skus[x][1].replace('R$ ', '')
            refatorado = refatorado.replace('.', '')
            refatorado = refatorado.replace(',', '.')
            media_precos += float(refatorado)

        for x in range(len(skus)):
            skus[x].append(media_precos)

    def acessa_loja(self, produto):
        driver = webdriver.Chrome()
        driver.get("https://www.americanas.com.br/")
        elemento = driver.find_element(By.XPATH,
                                       """//*[@id="rsyswpsdk"]/div/header/div[1]/div[1]/div/div[1]/form/input""")
        pesquisa = driver.find_element(By.XPATH,
                                       """//*[@id="rsyswpsdk"]/div/header/div[1]/div[1]/div/div[1]/form/button""")
        elemento.send_keys(produto)
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

        log = open("log.txt", 'a')
        log.write(datetime.now().strftime('%d/%m/%y %H:%M:%S') + " - Pesquisado SKUs do produto {}\n".format(produto))

        self.media_entre_valores(skus)
        return skus

    def gerador_relatorio_excel(self, produtos):
        workbook = Workbook()
        sheet = workbook.active
        sheet['A1'] = "Produtos"
        sheet['B1'] = "SKUs"
        sheet['C1'] = "Preços"
        sheet['D1'] = "Média dos Preços"
        sheet['A1'].font = Font(bold=True)
        sheet['B1'].font = Font(bold=True)
        sheet['C1'].font = Font(bold=True)
        sheet['D1'].font = Font(bold=True)

        row = 2
        for x in range(len(produtos)):
            resultado = self.acessa_loja(produtos[x])
            for y in range(len(resultado)):
                sheet["A{}".format(row)] = produtos[x]
                sheet["B{}".format(row)] = resultado[y][0]
                sheet["C{}".format(row)] = resultado[y][1]
                sheet["D{}".format(row)] = resultado[y][2]
                row += 1

        workbook.save("Relatório Produtos.xlsx")

        log = open("log.txt", 'a')
        log.write(datetime.now().strftime('%d/%m/%y %H:%M:%S') + " - Criado relatorio final\n")
