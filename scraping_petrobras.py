#=========================================================== 
# Request By: eng.Raphael Sales
# Created By: João Holanda
# Supports By: Jumario, João Pedro.
# Enterprise: Normatel Engenharia
# Purpose: Portal Petrobras Scrape Automatation
#=========================================================== 

 


import os
from _send_email import sendEmail

def PegarContratos(): 
    
    #linha sempre comeca no zero!
    Linha = 0
    Verificador = 1


    #iniciando o looping para ele executar linha por linha !
    while(Verificador != 0):

        
        import pandas as pd
        from time import sleep
        from selenium import webdriver
        from selenium.webdriver.common.by import By
        
    
        #localizando o valor do contrato posto pelo cliente
        tabela = pd.read_excel('Contratos.xlsx')
        
       
        #acessando o link 
        driver = webdriver.Chrome()
        driver.get('https://spot.petrobras.com.br/portal/ContratoConsultarPortal.aspx')


        #Fazendo condicoes para o uso ou do icj ou do contrato!
        while True:
            try:
                labelc = int(tabela.loc[Linha,"Contrato"])
                inputcontrato =  driver.find_element(by=By.ID, value='txtNumeroContrato')
                inputcontrato.send_keys(labelc)
                break
            except ValueError:
                labelicj = tabela.loc[Linha,"ICJ"]
                inputICJ = driver.find_element(by=By.ID, value='txtNumeroICJ')
                inputICJ.send_keys(labelicj) 
                break 

        sleep(1)
        #achando o botao consulta e atribuindo variavel
        csv_button = driver.find_element(by=By.ID, value='lbtnConsultar')

        sleep(1)
        #clicando no butao consultar
        csv_button.click()
        sleep(1)
       
        #achando valores que o Cliente precisa!

        Fornec =  driver.find_element(by=By.XPATH,value='//*[@id="ContentPlaceHolder1_ccContratoConsultar_grvCon"]/tbody/tr[2]/td/div[1]/table/tbody/tr/td[3]')
        ContratoN =  driver.find_element(by=By.XPATH,value='//*[@id="ContentPlaceHolder1_ccContratoConsultar_grvCon"]/tbody/tr[2]/td/div[1]/table/tbody/tr/td[2]')
        Icj = driver.find_element(by=By.XPATH,value='//span[@id="ContentPlaceHolder1_ccContratoConsultar_grvCon_subNumeroICJ_0"]')
        dateI = driver.find_element(by=By.XPATH,value='//*[@id="ContentPlaceHolder1_ccContratoConsultar_grvCon"]/tbody/tr[2]/td/div[1]/table/tbody/tr/td[4]')
        dateF = driver.find_element(by=By.XPATH,value='//*[@id="ContentPlaceHolder1_ccContratoConsultar_grvCon"]/tbody/tr[2]/td/div[1]/table/tbody/tr/td[5]')
        valorPrice = driver.find_element(by=By.XPATH,value='//*[@id="ContentPlaceHolder1_ccContratoConsultar_grvCon_subValorContrato_0"]')
        obj =  driver.find_element(by=By.XPATH,value='//*[@id="ContentPlaceHolder1_ccContratoConsultar_grvCon_lblObjeto_0"]')

        sleep(1)
        #resgatando os valores necessario

        fornecedor = Fornec.get_attribute("innerHTML").splitlines()[0]
        valorContrato = ContratoN.get_attribute("innerHTML").splitlines()[0]
        valoricj = Icj.get_attribute("innerHTML").splitlines()[0]
        dataInicial = dateI.get_attribute("innerHTML").splitlines()[0]
        dataFinal = dateF.get_attribute("innerHTML").splitlines()[0]
        valor_do_Contrato = valorPrice.get_attribute("innerHTML").splitlines()[0]
        ObjetoContrato = obj.get_attribute("innerHTML").splitlines()[0]

    #prevencao de erros pt.2
            


        #escrevendo os valores  no excel!
    
        #Condicoes se o valor do Contrato for preenchido
        while True:
            try:
                labelc = int(tabela.loc[Linha,"Contrato"]) 
                tabela.loc[tabela["Contrato"] == labelc, "Fornecedor"] = fornecedor
                tabela.loc[tabela["Contrato"] == labelc, "Número do Contrato"] = valorContrato
                tabela.loc[tabela["Contrato"] == labelc, "Número de ICJ"] = valoricj
                tabela.loc[tabela["Contrato"] == labelc, "INICIO"] = dataInicial
                tabela.loc[tabela["Contrato"] == labelc, "Fim do Contrato"] = dataFinal
                tabela.loc[tabela["Contrato"] == labelc, "Valor do Contrato"] = valor_do_Contrato
                tabela.loc[tabela["Contrato"] == labelc, "OBJETO"] = ObjetoContrato
                break
                
            except ValueError:
                labelicj = tabela.loc[Linha,"ICJ"]
                tabela.loc[tabela["ICJ"] == labelicj, "Fornecedor"] = fornecedor
                tabela.loc[tabela["ICJ"] == labelicj, "Número do Contrato"] = valorContrato
                tabela.loc[tabela["ICJ"] == labelicj, "Número de ICJ"] = valoricj
                tabela.loc[tabela["ICJ"] == labelicj, "INICIO"] = dataInicial
                tabela.loc[tabela["ICJ"] == labelicj, "Fim do Contrato"] = dataFinal
                tabela.loc[tabela["ICJ"] == labelicj, "Valor do Contrato"] = valor_do_Contrato
                tabela.loc[tabela["ICJ"] == labelicj, "OBJETO"] = ObjetoContrato
                break
                
        
        
        Linha = Linha + 1
        tabela.to_excel("Valores.xlsx",index=False)
        driver.close()
        #mudando o verificar para 0, para travar e acabar o While! 
        if(tabela.loc[Linha,"Contrato"] == None and tabela.loc[Linha,"ICJ"] == None ):
            Verificador = 0
            #sendEmail()
            sleep(1)
            
        

    else:
        
        quit()
        


if(os.path.exists('Contratos.xlsx')):
    PegarContratos()
else:
    quit()
         




