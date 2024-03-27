import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import messagebox
from selenium.common.exceptions import NoSuchElementException

meses_pt_br = {
    "January": "janeiro",
    "February": "fevereiro",
    "March": "março",
    "April": "abril",
    "May": "maio",
    "June": "junho",
    "July": "julho",
    "August": "agosto",
    "September": "setembro",
    "October": "outubro",
    "November": "novembro",
    "December": "dezembro"
}
 








print("Verificando se o arquivo 'dados.csv' existe. Se não existir, será criado.")
def verificar_csv():
    if os.path.exists('dados.csv'):
        # Verificar se o arquivo está vazio
        if os.stat('dados.csv').st_size == 0:
            # Se estiver vazio, exibir mensagem de alerta
            messagebox.showwarning("Aviso", "O arquivo CSV existe, mas está vazio. Por favor, preencha-o com os dados necessários.")
    else:
        # Se não existir, criar o arquivo e exibir mensagem de alerta
        with open('dados.csv', 'w') as f:
            f.write("CNPJ,Senha,CNPJ_Tomador,CEP_Tomador,Local_Prestacao,Descricao_Servico,Pedido_Compra,Valor_Servico\n")
        messagebox.showinfo("Aviso", "Um novo arquivo CSV foi criado. Por favor, preencha-o com os dados necessários.")


print ("inicio ")
def preencher_formulario():
    print("Ler os dados do CSV")
    dados = pd.read_csv('dados.csv')
    for index, row in dados.iterrows():
        # Iniciar o driver Selenium
        driver = webdriver.Chrome()
        driver.maximize_window()

        # URL da página
        
        url = "https://www.nfse.gov.br/EmissorNacional/Login?ReturnUrl=%2fEmissorNacional%2fDPS%2fPessoas"

        print("open navegador")
        driver.get(url)
        # login
        print("login")        
        
        sleep(2)
        print("preenchendo CNPJ")
        driver.find_element(By.XPATH, '//*[@id="Inscricao"]').send_keys(row['CNPJ'])
        sleep(2)
        print("preenchendo SENHA")
        driver.find_element(By.XPATH, '//*[@id="Senha"]').send_keys(row['Senha'])
        sleep(2)
        print("click em entra")
        driver.find_element(By.XPATH, '/html/body/section/div/div/div[2]/div[2]/div[1]/div/form/div[3]/button').click()
        sleep(2)
        print("Fim login")
        
  
    try:
        elemento =   driver.find_element(By.XPATH, '/html/body/div[1]/div[1]/div/ul/li[1]/a[1]/span') 
        # Verifique se o elemento está visível
        if elemento.is_displayed():           
        
        # preencher pessoas
            sleep(2)
            print("tratando data para emissao, usando data atual ")
            now = datetime.now()
            data_formatada = now.strftime("%d%m%Y")
            print("preenchendo data serviço prestado")
            driver.find_element(By.XPATH, '//*[@id="DataCompetencia"]').send_keys(data_formatada)
            driver.find_element(By.XPATH, '//*[@id="btn_DataCompetencia"]').click()
            
            print("preenchedo cnpj do  tomador")
            sleep(2)
            driver.find_element(By.XPATH, '//*[@id="pnlTomador"]/div[1]/div/div/div[2]/label/span').click()
            driver.find_element(By.XPATH, '//*[@id="Tomador_Inscricao"]').send_keys(row['CNPJ_Tomador'])
            driver.find_element(By.XPATH, '//*[@id="btn_Tomador_Inscricao_pesquisar"]').click()
            print("preenchedo cep do  tomador")
            driver.find_element(By.XPATH, '//*[@id="Tomador_EnderecoNacional_CEP"]').send_keys(row['CEP_Tomador'])
            driver.find_element(By.XPATH, '//*[@id="btn_Tomador_EnderecoNacional_CEP"]').click()
            print("proximo formulario")

            driver.find_element(By.XPATH, '//*[@id="btnAvancar"]').click()

            # preenche serviços
            sleep(2)
            driver.find_element(By.XPATH, '//*[@id="pnlLocalPrestacao"]/div/div/div[2]/div/span[1]/span[1]/span/span[2]').click()
            sleep(2)
            campo_input = driver.find_element(By.CSS_SELECTOR, 'body > span > span > span.select2-search.select2-search--dropdown > input')
            campo_input.send_keys(row['Local_Prestacao'])
            sleep(5)
            driver.find_element(By.CSS_SELECTOR, '#select2-LocalPrestacao_CodigoMunicipioPrestacao-results > li').click()
            driver.find_element(By.XPATH, '//*[@id="pnlServicoPrestado"]/div/div[1]/div/div/span[1]/span[1]/span/span[2]/b').click()
            driver.find_element(By.XPATH, '/html/body/span/span/span[1]/input').send_keys(row['Descricao_Servico'])
            sleep(5)
            driver.find_element(By.CSS_SELECTOR, '#select2-ServicoPrestado_CodigoTributacaoNacional-results > li').click()
            driver.find_element(By.CSS_SELECTOR, '#pnlServicoPrestado > div > div.form-group.form-group-lg > div > div:nth-child(1) > label > span').click()

            # Obtendo o mês anterior
            data_atual = datetime.now()
            data_mes_anterior = data_atual.replace(day=1) - timedelta(days=1)
            mes_anterior_en = data_mes_anterior.strftime('%B')  # Nome do mês em inglês
            mes_anterior_pt_br = meses_pt_br[mes_anterior_en]  # Tradução para português
            ano_mes_anterior = data_mes_anterior.year

            # Formatar a descrição do serviço
            descricao_servico = f"PRESTAÇÃO DE SERVIÇOS REFERENTE AO MÊS DE {mes_anterior_pt_br.upper()}/{ano_mes_anterior} Nº DO PEDIDO DE COMPRA: {row['Pedido_Compra']}"

            # Preencher a descrição do serviço
            driver.find_element(By.XPATH, '//*[@id="ServicoPrestado_Descricao"]').send_keys(descricao_servico)
            sleep(2)

            driver.find_element(By.XPATH, '/html/body/div[1]/form/div[7]/button').click()
            sleep(2)
            driver.find_element(By.XPATH, '//*[@id="Valores_ValorServico"]').send_keys(row['Valor_Servico'])
            driver.find_element(By.XPATH, '//*[@id="pnlOpcaoParaMEI"]/div/div/label/span').click()
            driver.find_element(By.XPATH, '/html/body/div[1]/form/div[7]/button').click()
            sleep(2)
                # Mostrar caixa de diálogo para confirmação
            if messagebox.askyesno("Confirmação", "Confirme os dados para continuar e baixar o arquivo. Deseja continuar?"):
                print("Continuando o processo...")
                sleep(2)
                driver.find_element(By.XPATH, '//*[@id="btnProsseguir"]/span').click()  
                sleep(2)
                
                driver.find_element(By.XPATH, '//*[@id="btnDownloadDANFSE"]/span').click()  
                driver.quit()
            else:
                print("Processo cancelado pelo usuário.")
                # Encerrar o script
                driver.quit()
        
        else:
            print("falha login.")
            # Encerre o script se o elemento não estiver visível
            driver.quit()
        
        
    except NoSuchElementException: 
        driver.quit()
        print("Login nao realizado. Encerrando o script.")
        
        
        
        
     

# Verificar e criar o arquivo CSV se necessário
verificar_csv()

# Preencher o formulário com os dados do CSV
preencher_formulario()

sleep(2)

