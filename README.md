# RECEBIMENTO-MTR
Desenvolvido para receber de forma simples o Manifesto de transporte de resíduos, e gerar uma planilha com seus dados.

from ast import Break
from operator import index
from sqlite3 import Row
import pandas as pd
from tkinter import Button
from playwright.sync_api import sync_playwright

arquivo = "PASTA"
bd = pd.read_excel(arquivo)
linha_final = len(bd) - 1
linha = 0


with sync_playwright() as p :
    browser = p.chromium.launch(headless=False)
    page = browser.new_page()

    #Login
    page.goto("http://200.20.53.11/")
    page.fill("input[name='txtCnpj']",'CNPJ')
    page.wait_for_timeout(1000)
    page.fill("input[id='txtSenha']",'SENHA')
    page.wait_for_timeout(1000)
    page.fill("input[id='txtUnidadeCodigo']",'UNID COD')
    page.wait_for_timeout(1000)
    page.click("input[id='txtSenha']")
    page.wait_for_timeout(1000)
    page.fill("input[name='txtCpfUsuario']",'CPF')
    page.wait_for_timeout(1000)
    page.click("[id='btEntrar']")

    #Navegação até meus mtr
    page.wait_for_timeout(1000)
    page.click('#dv_principal_menu > ul > li:nth-child(2) > a')
    page.wait_for_timeout(1000)
    page.click('#dv_principal_menu > ul > li:nth-child(2) > ul > li:nth-child(5) > a')

    while True:
        if linha == linha_final: #Final do loop
            bd.to_excel('mtr2.xlsx', index = False)
            break
        else:
            print(linha)
            CodMTR = bd.iloc[linha][5]
            dtReceb = bd.iloc[linha][1]
            mot = bd.iloc[linha][2]
            placa = bd.iloc[linha][3]
            QNTr = bd.iloc[linha][4]

            #Recebimento do MTR
            page.fill("input[id='txtCodigo']", str(CodMTR))
            page.wait_for_timeout(1000)
            page.keyboard.press('Enter')
            page.wait_for_timeout(1000)

            #Se MTR ja for recebido ou não existir
            if page.is_visible("input[id='txtDataRecebimento']") == True :

                #Se o Campo do motorista estiver em branco
                if page.get_attribute('[id="txtTransportadorNomeMotorista"]','value') == "":
                    page.fill("input[id='txtTransportadorNomeMotorista']", mot)
                    page.fill("input[id='txtTransportadorPlacaVeiculo']", placa)
                    bd.at[linha, 'CAMPO MOT'] = "em branco"
                    page.wait_for_timeout(500)
                else:
                    bd.at[linha, 'CAMPO MOT'] = "Ok"
                    page.wait_for_timeout(500)

                #Calculo diferença de quantidade
                raw_txt = page.locator('#tbRecebeMTR > tbody > tr > td:nth-child(4)').text_content()
                nounit_txt = raw_txt.replace("(t)","")
                QNTi = nounit_txt.replace(",",".")
                dif_qr = float(QNTr) - float(QNTi)
                dif_qi = float(QNTi) - float(QNTr)
                dif_qr = ((dif_qr ** 2)**(1/2))
                dif_qi = ((dif_qi ** 2)**(1/2))
                mtr_num = page.get_attribute('input[id="txtCodigoMtr"]', 'value')
                bd.at[linha, 'Num MTR'] = mtr_num

                #Filtro diferença de quantidade
                if dif_qr < 2 or dif_qi < 2 :
                    page.locator('[value="0,00000"]').fill(QNTr)
                    page.locator('[title="Divergência do resíduo."]').click()
                    page.wait_for_timeout(1000)
                    page.fill('textarea[id="txtJust"]', "Divergencia de balança.")
                    page.keyboard.press('Tab')
                    page.keyboard.press('Enter')
                    bd.at[linha, 'STATUS'] = "Ok"

                elif 2 < dif_qr < 5 or 2 < dif_qr < 5:
                    page.locator('[value="0,00000"]').fill(QNTr)
                    page.locator('[title="Divergência do resíduo."]').click()
                    page.wait_for_timeout(1000)
                    page.fill('textarea[id="txtJust"]', "Quantidade indicada estimada/Quantidade indicada não verificada com balança.")
                    page.keyboard.press('Tab')
                    page.keyboard.press('Enter')
                    bd.at[linha, 'STATUS'] = "Ok c/ dif"
                else:
                    # mtr_erro = page.get_attribute('input[id="txtCodigoMtr"]', 'value') se codigo ficar lento, deixar apenas para esse e remover "status"
                    bd.at[linha, 'STATUS'] = "RECU"
                    bd.at[linha, 'QNT-I'] = QNTi
                    page.locator('text=Recebimento do MTRclose >> button[role="button"]').click() #se tiver mais de um button/input procurar no terminal o respectivo apos o erro

                #Preencher restante dos campos
                page.wait_for_timeout(1000)
                page.eval_on_selector("input[id='txtDataRecebimento']", "el => el.removeAttribute('disabled')")
                page.fill("input[id='txtDataRecebimento']", str(dtReceb) )
                page.click('#formRespRecebimento > fieldset:nth-child(6) > table > tbody > tr > td:nth-child(1) > a')
                page.click('#tabelaResRecebimento > tbody > tr:nth-child(2) > td:nth-child(3)')
                page.locator('[type="submit"]').click()
                page.wait_for_timeout(1000)
                page.keyboard.press('Tab')
                page.keyboard.press('Enter')

            else :
                page.wait_for_timeout(1000)
                bd.at[linha, 'STATUS'] = "RECEB." 
                page.keyboard.press('Tab')
                page.keyboard.press('Enter')

            linha = linha + 1       
