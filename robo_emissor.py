import os
import time
import json
import glob
import re
import pandas as pd
from datetime import datetime

# Import para a janelinha de pop-up
import tkinter as tk
from tkinter import messagebox

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager

# ===== SOLUÇÃO PARA O ERRO DO DDDDOCR (PILLOW 10+ / Python 3.14) =====
from PIL import Image
from dotenv import load_dotenv
load_dotenv()  # lê o .env local (não vai para o GitHub)
if not hasattr(Image, 'ANTIALIAS'):
    Image.ANTIALIAS = Image.Resampling.LANCZOS if hasattr(Image, "Resampling") else Image.LANCZOS
# =====================================================================

import ddddocr 

# ================= CONFIGURAÇÕES =================
CAMINHO_PLANILHA = os.getenv("CAMINHO_PLANILHA_FILA")
CAMINHO_SENHAS = os.getenv("CAMINHO_SENHAS")
PASTA_DRIVE_NOTAS = os.getenv("PASTA_DRIVE_NOTAS")

def registrar_log(mensagem):
    data_hora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    texto_formatado = f"[{data_hora}] {mensagem}"
    
    if not os.path.exists(PASTA_DRIVE_NOTAS):
        os.makedirs(PASTA_DRIVE_NOTAS)
        
    try:
        with open(os.path.join(PASTA_DRIVE_NOTAS, "historico_do_robo.txt"), "a", encoding="utf-8") as f:
            f.write(texto_formatado + "\n")
    except: 
        pass

ocr = ddddocr.DdddOcr()

def configurar_navegador(pasta_download):
    chrome_options = Options()
    settings = {"recentDestinations": [{"id": "Save as PDF", "origin": "local", "account": ""}], "selectedDestinationId": "Save as PDF", "version": 2}
    prefs = {"printing.print_preview_sticky_settings.appState": json.dumps(settings), "savefile.default_directory": pasta_download, "download.default_directory": pasta_download, "download.prompt_for_download": False, "safebrowsing.enabled": True}
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument('--kiosk-printing') 
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    driver.maximize_window()
    return driver

# Função original intacta
def resolver_captcha(driver):
    try:
        wait = WebDriverWait(driver, 1)
        try:
            captcha_img = driver.find_element(By.CSS_SELECTOR, "img[src*='aptcha'], img[src*='APTCHA']")
        except:
            captcha_img = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@id='tbCaptcha']/preceding::img[1]")))
        captcha_img.screenshot("captcha.png")
        with open("captcha.png", 'rb') as f:
            texto_captcha = ocr.classification(f.read())
        return ''.join(e for e in texto_captcha if e.isalnum()).upper()
    except Exception as e:
        registrar_log(f"Erro ao resolver captcha: {e}")
        return ""

def formatar_cnpj(valor):
    if pd.isna(valor): return ""
    digitos = str(valor).replace('.', '').replace('/', '').replace('-', '').strip()
    if len(digitos) > 11:
        return digitos.zfill(14)
    return digitos

def preencher_campo(driver, seletor, valor):
    if pd.notna(valor) and str(valor).strip() != "":
        try:
            campo = driver.find_element(By.CSS_SELECTOR, seletor)
            campo.clear()
            val_str = f"{float(valor):.2f}".replace('.', ',') if isinstance(valor, (int, float)) else str(valor).replace('.', ',')
            campo.send_keys(val_str)
            campo.send_keys(Keys.TAB)
        except Exception as e:
            msg = f"Aviso: Falha ao preencher campo {seletor} com valor {valor}. Detalhe: {e}"
            print(msg)
            registrar_log(msg)

def escrever_erro(id_nota, mensagem):
    if not os.path.exists(PASTA_DRIVE_NOTAS): os.makedirs(PASTA_DRIVE_NOTAS)
    with open(os.path.join(PASTA_DRIVE_NOTAS, f"erro_{id_nota}.txt"), "w", encoding="utf-8") as f:
        f.write(mensagem)

def main():
    print("Iniciando o robô...")
    registrar_log("=== ROBÔ INICIADO PELO AGENDADOR DE TAREFAS ===")
    
    if not os.path.exists(CAMINHO_SENHAS):
        print(f"ERRO: Planilha de senhas não encontrada: {CAMINHO_SENHAS}")
        return

    dicionario_senhas = {}
    try:
        df_senhas = pd.read_excel(CAMINHO_SENHAS)
        
        # O sublinhado da variável anônima está aqui: _
        for _, row in df_senhas.iterrows():
            if len(row) >= 7:  
                cnpj_cru = row.iloc[5]   
                senha_cru = row.iloc[6]  
                if pd.notna(cnpj_cru) and pd.notna(senha_cru):
                    cnpj_fmt = formatar_cnpj(cnpj_cru)
                    dicionario_senhas[cnpj_fmt] = str(senha_cru).strip()
    except Exception as e:
        print(f"ERRO LENDO SENHAS: {e}")
        registrar_log(f"ERRO LENDO SENHAS: {e}")
        return

    if not os.path.exists(CAMINHO_PLANILHA):
        print(f"ERRO: Planilha de fila não encontrada: {CAMINHO_PLANILHA}")
        return

    try:
        df = pd.read_excel(CAMINHO_PLANILHA)
        df = df.astype(object)
        df.columns = [str(c).lower().strip() for c in df.columns]
        
        notas_pendentes = []
        for _, row in df.iterrows():
            status_val = str(row.get('status', '')).strip()
            if (status_val == "" or status_val == "nan") and len(row) >= 22:
                status_val = str(row.iloc[21]).strip()
            
            if status_val.upper() == 'EMITIR':
                notas_pendentes.append(row)

        if not notas_pendentes:
            print("AVISO: Nenhuma nota para 'EMITIR'.")
            return
            
    except Exception as e:
        print(f"ERRO LENDO FILA: {e}")
        registrar_log(f"ERRO LENDO FILA: {e}")
        return

    driver = None
    cnpj_logado = ""

    for row in notas_pendentes:
        id_nota = str(row.get('id', '')).strip()
        cnpj_p = formatar_cnpj(row.get('prestador_cnpj', ''))
        tomador = formatar_cnpj(row.get('tomador_cnpj', row.get('cnpj_tomador', '')))
        senha = dicionario_senhas.get(cnpj_p, "")
        nome_empresa_planilha = str(row.get('prestador_name', row.get('socio_solicitante', cnpj_p))).strip()

        print(f"\n-> Emitindo Nota ID {id_nota} (Tomador: {tomador})")

        if not senha:
            erro_senha = f"ERRO: Senha não encontrada para CNPJ {cnpj_p}"
            escrever_erro(id_nota, erro_senha)
            continue

        if cnpj_p != cnpj_logado:
            print(f"Realizando login para o CNPJ: {cnpj_p}")
            if driver: driver.quit()
            driver = configurar_navegador(PASTA_DRIVE_NOTAS)
            driver.get("https://nfse.salvador.ba.gov.br/")
            
            sucesso_login = False
            
            for tentativa in range(1, 51):
                texto = resolver_captcha(driver)
                if len(texto) == 5:
                    print(f"  -> Tentativa {tentativa}: Testando captcha '{texto}'...")
                    
                    try:
                        campo_cnpj = driver.find_element(By.CSS_SELECTOR, "input#txtLogin")
                        if not campo_cnpj.get_attribute('value'):
                            campo_cnpj.send_keys(cnpj_p)
                            
                        campo_senha = driver.find_element(By.CSS_SELECTOR, "input#txtSenha")
                        if not campo_senha.get_attribute('value'):
                            campo_senha.send_keys(senha)
                    except Exception as e:
                        registrar_log(f"Erro ao inserir credenciais: {e}")
                    
                    try:
                        driver.find_element(By.CSS_SELECTOR, "input#tbCaptcha").clear()
                        driver.find_element(By.CSS_SELECTOR, "input#tbCaptcha").send_keys(texto)
                        driver.find_element(By.CSS_SELECTOR, "input#cmdLogin").click()
                        time.sleep(7)
                        
                        if "Sair" in driver.page_source or "InformacaoDebito.aspx" in driver.current_url:
                            sucesso_login = True
                            cnpj_logado = cnpj_p
                            print("  -> Login efetuado com sucesso!")
                            break
                    except Exception as e:
                        registrar_log(f"Erro ao interagir com captcha/login: {e}")
                    
                try: 
                    driver.find_element(By.XPATH, "//*[contains(text(), 'Recarregar')]").click()
                    time.sleep(1.5)
                except:
                    pass
            
            if not sucesso_login:
                erro_msg = "ERRO: Falha Login (Captcha ou Senha Incorreta após 50 tentativas)"
                print(f"   [FALHA] Nota {id_nota} - {erro_msg}")
                escrever_erro(id_nota, erro_msg)
                continue

        try:
            print(f"Preenchendo dados da nota para o tomador {tomador}...")
            driver.get("https://nfse.salvador.ba.gov.br/site/contribuinte/nota/nota.aspx")
            wait = WebDriverWait(driver, 15)
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input#tbCPFCNPJTomador"))).send_keys(tomador)
            driver.find_element(By.XPATH, "//*[@id='btAvancar']").click()
            time.sleep(3) 

            # ================= CORREÇÃO DO CNAE (MÉTODO JQUERY) =================
            try:
                cnae_completo = str(row.get('cnae', ''))
                cnae_parte = cnae_completo.split('(')[0]
                cnae_numeros = ''.join(filter(str.isdigit, cnae_parte))
                
                if cnae_numeros:
                    driver.execute_script(f"$('#ddlCNAE').val('{cnae_numeros}').trigger('change');")
                    time.sleep(4)
            except Exception as erro_cnae:
                msg_cnae = f"AVISO: Não foi possível preencher o CNAE automaticamente: {erro_cnae}"
                print(msg_cnae)
                registrar_log(msg_cnae)
            # ====================================================================

            preencher_campo(driver, "input#tbAliquota", row.get('aliquota', ''))
            driver.find_element(By.CSS_SELECTOR, "textarea#tbDiscriminacao").send_keys(str(row.get('discriminacao', '')))
            preencher_campo(driver, "input#tbValor", row.get('valor_total', ''))
            preencher_campo(driver, "input#tbCodTribNac", row.get('codigo_tributario_nacional', ''))
            preencher_campo(driver, "input#tbCTISS", row.get('ctiss', ''))
            preencher_campo(driver, "input#tbNBS", row.get('nbs', ''))

            if "SIM" in str(row.get('iss_retido', '')).upper():
                driver.find_element(By.XPATH, "//table[@id='rblISSRetido']//label[contains(text(), 'Sim')]/preceding-sibling::input").click()
            else:
                driver.find_element(By.XPATH, "//table[@id='rblISSRetido']//label[contains(text(), 'Não')]/preceding-sibling::input").click()

            preencher_campo(driver, "input#tbRPSEmissao", row.get('data_competencia', ''))
            preencher_campo(driver, "input#tbINSS", row.get('inss', ''))
            preencher_campo(driver, "input#tbIRPJ", row.get('irrf', ''))
            preencher_campo(driver, "input#tbCSLL", row.get('csll', ''))
            preencher_campo(driver, "input#tbCOFINS", row.get('cofins', ''))
            preencher_campo(driver, "input#tbPisPasep", row.get('pis', ''))
            preencher_campo(driver, "input#tbOutrasRetencoes", row.get('outras_retencoes', ''))

            # ==== PAUSA VISUAL (POP-UP) PARA VALIDAÇÃO MANUAL ====
            print("\nAguardando validação manual através da janela de pop-up...")
            
            root = tk.Tk()
            root.withdraw() 
            root.attributes("-topmost", True) 
            
            messagebox.showinfo(
                "Ação Necessária - Robô Emissor", 
                "1. O robô terminou de preencher os dados!\n"
                "2. Vá no navegador e confira se está tudo certo.\n"
                "► SÓ DEPOIS DISSO CLIQUE EM 'OK' AQUI.\n\n"
                "Após clicar em OK, o robô clicará em Emitir, confirmará o aviso e baixará o PDF."
            )
            
            root.destroy()
            
            # ================= CLICA EM EMITIR, ACEITA ALERTA E CAPTURA NOTA =================
            print("Clicando no botão Emitir...")
            try:
                botao_emitir = wait.until(EC.element_to_be_clickable((By.ID, "btEmitir")))
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao_emitir)
                time.sleep(0.5)
                botao_emitir.click()
                
                try:
                    wait_alerta = WebDriverWait(driver, 5)
                    wait_alerta.until(EC.alert_is_present())
                    alerta = driver.switch_to.alert
                    print(f"Aceitando alerta: '{alerta.text}'")
                    alerta.accept()
                except Exception as e:
                    registrar_log(f"Aviso: Alerta de confirmação não processado. Detalhe: {e}")
            except Exception as e:
                msg_emitir = f"Aviso: Falha ao clicar em Emitir automaticamente: {e}"
                print(msg_emitir)
                registrar_log(msg_emitir)

            # ================= NOVA EXTRAÇÃO INTELIGENTE DO NÚMERO DA NOTA =================
            numero_nota = ""
            try:
                WebDriverWait(driver, 20).until(
                    lambda d: re.search(
                        r'N[uú]mero\s*(?:da)?\s*(?:Nota|NFS-?e)?\s*:?\s*\d+', 
                        d.find_element(By.TAG_NAME, "body").text, 
                        re.IGNORECASE
                    )
                )

                texto_pagina = driver.find_element(By.TAG_NAME, "body").text

                match_nota = re.search(
                    r'N[uú]mero\s*(?:da)?\s*(?:Nota|NFS-?e)?\s*:?\s*(\d+)',
                    texto_pagina,
                    re.IGNORECASE
                )

                if match_nota:
                    numero_nota = match_nota.group(1).strip()
                else:
                    caminho_debug = os.path.join(PASTA_DRIVE_NOTAS, f"debug_pagina_{id_nota}.txt")
                    with open(caminho_debug, "w", encoding="utf-8") as f:
                        f.write(texto_pagina)
                    msg_debug = f"AVISO: número da nota não encontrado para ID {id_nota}. Texto da tela salvo em {caminho_debug} para análise."
                    print(msg_debug)
                    registrar_log(msg_debug)

            except Exception as e:
                registrar_log(f"Aviso: Falha ao extrair número da nota da tela: {e}")
            
            if not numero_nota:
                numero_nota = f"ID_{id_nota}_GERADA_{datetime.now().strftime('%H%M%S')}"
                print("AVISO: Não foi possível capturar o 'Número da Nota'. Usando ID alternativo.")
            else:
                print(f"Número da Nota capturado com sucesso: {numero_nota}")
            # ================================================================================
            
            print("Comandando impressão (PDF)...")
            driver.execute_script("window.print();")
            time.sleep(5)
            
            if not os.path.exists(PASTA_DRIVE_NOTAS): os.makedirs(PASTA_DRIVE_NOTAS)
            list_of_files = [f for f in glob.glob(os.path.join(PASTA_DRIVE_NOTAS, '*.pdf')) if not os.path.basename(f).startswith("[OK]")]

            if list_of_files:
                ultimo_pdf = max(list_of_files, key=os.path.getctime)
                nome_limpo = re.sub(r'[\\/*?:"<>|]', "", nome_empresa_planilha) 
                
                novo_nome = os.path.join(PASTA_DRIVE_NOTAS, f"{datetime.now().strftime('%d%m%Y')}-{nome_limpo} - NF {numero_nota}.pdf")
                
                if os.path.exists(novo_nome): os.remove(novo_nome)
                os.rename(ultimo_pdf, novo_nome)
                print(f"   [SUCESSO] Salvo: {novo_nome}")
            else:
                erro_pdf = "ERRO: O PDF não foi gerado/baixado."
                print(f"   [FALHA] Nota {id_nota} - {erro_pdf}")
                escrever_erro(id_nota, erro_pdf)

        except Exception as e:
            msg_curta = f"ERRO SISTEMA: {e}"
            print(f"   [FALHA] {msg_curta}")
            registrar_log(f"Erro geral no processamento da nota {id_nota}: {e}")
            escrever_erro(id_nota, msg_curta)

    if driver: driver.quit()

# Os underlines duplos de __name__ e __main__ estão aplicados nesta linha:
if __name__ == "__main__":
    main()