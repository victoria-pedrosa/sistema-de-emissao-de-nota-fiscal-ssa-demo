import os
import time
import json
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
import ddddocr 
from dotenv import load_dotenv
load_dotenv()  # lê o .env local (não vai para o GitHub)

# =============================================================================
# CONFIGURAÇÕES DA PLANILHA LOCAL (Sincronizada do arquivo Notas Exemplo)
# =============================================================================
CAMINHO_PLANILHA = os.getenv("CAMINHO_PLANILHA_EMISSAO")
NOME_DA_ABA = "Emissao"
BASE_DIR_PDF = os.getenv("BASE_DIR_PDF")

MMAAAA_ATUAL = datetime.now().strftime("%m%Y")
ocr = ddddocr.DdddOcr(show_ad=False)

def configurar_navegador(pasta_download):
    chrome_options = Options()
    settings = {
        "recentDestinations": [{"id": "Save as PDF", "origin": "local", "account": ""}],
        "selectedDestinationId": "Save as PDF",
        "version": 2
    }
    prefs = {
        "printing.print_preview_sticky_settings.appState": json.dumps(settings),
        "savefile.default_directory": pasta_download,
        "download.default_directory": pasta_download,
        "download.prompt_for_download": False,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument('--kiosk-printing') 
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    driver.maximize_window()
    return driver

def resolver_captcha(driver):
    try:
        wait = WebDriverWait(driver, 1)
        try:
            captcha_img = driver.find_element(By.CSS_SELECTOR, "img[src*='aptcha'], img[src*='APTCHA']")
        except:
            xpath_imagem = "//input[@id='tbCaptcha']/preceding::img[1]"
            captcha_img = wait.until(EC.presence_of_element_located((By.XPATH, xpath_imagem)))
        caminho_imagem = "captcha.png"
        captcha_img.screenshot(caminho_imagem)
        with open(caminho_imagem, 'rb') as f:
            img_bytes = f.read()
        texto_captcha = ocr.classification(img_bytes)
        return ''.join(e for e in texto_captcha if e.isalnum()).upper()
    except: return ""

def preencher_campo(driver, seletor, valor):
    if pd.notna(valor) and str(valor).strip() != "":
        try:
            campo = driver.find_element(By.CSS_SELECTOR, seletor)
            campo.clear()
            val_str = f"{float(valor):.2f}".replace('.', ',') if isinstance(valor, (int, float)) else str(valor).replace('.', ',')
            campo.send_keys(val_str)
            campo.send_keys(Keys.TAB)
        except: pass

def main():
    print(f"Lendo a planilha: {CAMINHO_PLANILHA}")
    try:
        df = pd.read_excel(CAMINHO_PLANILHA, sheet_name=NOME_DA_ABA)
        df = df.astype(object) 
    except Exception as e:
        print(f"[ERRO] Falha ao ler arquivo: {e}")
        return

    driver = None
    cnpj_logado = ""

    for index, row in df.iterrows():
        empresa = str(row.get('Empresa', '')).strip()
        cnpj_p = str(row.get('CNPJ Prestador', '')).strip()
        senha = str(row.get('Senha', '')).strip()
        tomador = str(row.get('CNPJ Tomador', '')).strip()
        revisao_necessaria = str(row.get('Revisão', 'EMITIR DIRETO')).strip().upper()
        
        if pd.isna(empresa) or empresa == 'nan' or empresa == '' or "Emitido" in str(row.get('Status', '')):
            continue

        print(f"\nProcessando Fila: {empresa}")
        pasta_final = os.path.join(BASE_DIR_PDF, MMAAAA_ATUAL, empresa)
        os.makedirs(pasta_final, exist_ok=True)

        if cnpj_p != cnpj_logado:
            if driver: driver.quit()
            driver = configurar_navegador(pasta_final)
            wait = WebDriverWait(driver, 20)
            driver.get("https://nfse.salvador.ba.gov.br/")
            
            sucesso_login = False
            for tentativa in range(1, 31):
                texto = resolver_captcha(driver)
                if len(texto) == 5:
                    try:
                        campo_cnpj = driver.find_element(By.CSS_SELECTOR, "input#txtLogin")
                        if not campo_cnpj.get_attribute('value'): campo_cnpj.send_keys(cnpj_p)
                        campo_senha = driver.find_element(By.CSS_SELECTOR, "input#txtSenha")
                        if not campo_senha.get_attribute('value'): campo_senha.send_keys(senha)
                    except: pass

                    driver.find_element(By.CSS_SELECTOR, "input#tbCaptcha").clear()
                    driver.find_element(By.CSS_SELECTOR, "input#tbCaptcha").send_keys(texto)
                    driver.find_element(By.CSS_SELECTOR, "input#cmdLogin").click()
                    time.sleep(7)
                    if "Sair" in driver.page_source or "InformacaoDebito.aspx" in driver.current_url:
                        sucesso_login = True
                        cnpj_logado = cnpj_p
                        break
                try: 
                    driver.find_element(By.XPATH, "//*[contains(text(), 'Recarregar')]").click()
                    time.sleep(1.5)
                except: pass
            
            if not sucesso_login:
                df.at[index, 'Status'] = "Erro Login"
                df.to_excel(CAMINHO_PLANILHA, index=False)
                continue

        try:
            driver.get("https://nfse.salvador.ba.gov.br/site/contribuinte/nota/nota.aspx")
            wait = WebDriverWait(driver, 15)
            campo_t = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input#tbCPFCNPJTomador")))
            campo_t.clear()
            campo_t.send_keys(tomador)
            driver.find_element(By.XPATH, "//*[@id='btAvancar']").click()
            time.sleep(3) 

            # CNAE
            chosen_elements = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a.chosen-single")))
            if chosen_elements:
                cnae_dropdown = chosen_elements[0]
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", cnae_dropdown)
                time.sleep(0.5)
                cnae_dropdown.click()
                busca = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div.chosen-drop input[type='text']")))
                busca.clear()
                busca.send_keys(str(row['CNAE']).strip())
                time.sleep(1.5)
                busca.send_keys(Keys.RETURN)
                time.sleep(0.5)

            preencher_campo(driver, "input#tbAliquota", row['Alíquota'])
            driver.find_element(By.CSS_SELECTOR, "textarea#tbDiscriminacao").send_keys(str(row['Discriminação']))
            preencher_campo(driver, "input#tbValor", row['Valor Total'])

            iss = str(row['ISS Retido']).upper()
            if "SIM" in iss:
                driver.find_element(By.XPATH, "//table[@id='rblISSRetido']//label[contains(text(), 'Sim')]/preceding-sibling::input").click()
            else:
                driver.find_element(By.XPATH, "//table[@id='rblISSRetido']//label[contains(text(), 'Não')]/preceding-sibling::input").click()

            if pd.notna(row['Data Competência']): preencher_campo(driver, "input#tbRPSEmissao", row['Data Competência'])
            preencher_campo(driver, "input#tbINSS", row['INSS'])
            preencher_campo(driver, "input#tbIRPJ", row['IRRF'])
            preencher_campo(driver, "input#tbCSLL", row['CSLL'])
            preencher_campo(driver, "input#tbCOFINS", row['COFINS'])
            preencher_campo(driver, "input#tbPisPasep", row['PIS'])
            preencher_campo(driver, "input#tbOutrasRetencoes", row['Outras Retenções'])

            time.sleep(2)
            df.at[index, 'Status'] = "Emitido com Sucesso"
            df.to_excel(CAMINHO_PLANILHA, index=False)
        except Exception as e:
            df.at[index, 'Status'] = f"Erro: {str(e)[:40]}"
            df.to_excel(CAMINHO_PLANILHA, index=False)

    if driver: driver.quit()

if __name__ == "__main__":
    main()