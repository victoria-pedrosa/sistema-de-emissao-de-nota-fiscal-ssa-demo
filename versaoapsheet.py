import os
import io
import json
import time
import base64
from datetime import datetime

import pandas as pd
import gspread
import requests
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager

import ddddocr

# =============================================================================
# CONFIGURAÇÕES (agora via variáveis de ambiente, em vez de caminho local fixo)
# =============================================================================
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "ID_EXEMPLO")
NOME_DA_ABA = os.environ.get("NOME_DA_ABA", "Emissao")
DRIVE_FOLDER_ID_PDFS = os.environ.get("DRIVE_FOLDER_ID_PDFS")  # pasta raiz no Drive p/ salvar PDFs
CHROMEDRIVER_PATH = os.environ.get("CHROMEDRIVER_PATH")  # opcional, se já tiver o driver instalado
CHROME_BIN = os.environ.get("CHROME_BIN")  # opcional, caminho do binário do Chrome/Chromium

MMAAAA_ATUAL = datetime.now().strftime("%m%Y")

# Posição das colunas na aba "Emissao" (A=1 ... Z=26), igual ao mapeamento usado no Apps Script
COL_STATUS = 21
COL_NUMERO_NOTA = 22

ocr = ddddocr.DdddOcr(show_ad=False)


# =============================================================================
# CONEXÃO COM GOOGLE SHEETS E GOOGLE DRIVE (substitui a leitura/escrita do Excel local)
# =============================================================================
def conectar_google():
    escopos = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    credenciais_json = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]
    info = json.loads(credenciais_json)
    credenciais = Credentials.from_service_account_info(info, scopes=escopos)

    cliente_sheets = gspread.authorize(credenciais)
    planilha = cliente_sheets.open_by_key(SPREADSHEET_ID)
    aba = planilha.worksheet(NOME_DA_ABA)

    return aba, credenciais


def carregar_dados(aba):
    registros = aba.get_all_records()
    df = pd.DataFrame(registros)
    df = df.astype(object)
    return df


def atualizar_status_sheet(aba, linha_sheet, status, numero_nota=None):
    try:
        aba.update_cell(linha_sheet, COL_STATUS, status)
        if numero_nota:
            aba.update_cell(linha_sheet, COL_NUMERO_NOTA, numero_nota)
    except Exception as e:
        print(f"  -> Falha ao atualizar status na planilha: {e}")


# =============================================================================
# SALVAMENTO DO PDF NO GOOGLE DRIVE (substitui a gravação em pasta local)
# =============================================================================
def obter_ou_criar_pasta(servico_drive, nome, pasta_pai_id):
    query = (
        f"name = '{nome}' and mimeType = 'application/vnd.google-apps.folder' "
        f"and '{pasta_pai_id}' in parents and trashed = false"
    )
    resultado = servico_drive.files().list(q=query, fields="files(id, name)").execute()
    arquivos = resultado.get("files", [])
    if arquivos:
        return arquivos[0]["id"]

    metadata = {
        "name": nome,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [pasta_pai_id],
    }
    pasta = servico_drive.files().create(body=metadata, fields="id").execute()
    return pasta["id"]


def salvar_pdf_no_drive(credenciais, pdf_bytes, nome_arquivo, empresa):
    if not DRIVE_FOLDER_ID_PDFS:
        print("  -> [AVISO] DRIVE_FOLDER_ID_PDFS não configurado; PDF não será salvo.")
        return

    servico_drive = build("drive", "v3", credentials=credenciais)
    pasta_mes_id = obter_ou_criar_pasta(servico_drive, MMAAAA_ATUAL, DRIVE_FOLDER_ID_PDFS)
    pasta_empresa_id = obter_ou_criar_pasta(servico_drive, empresa, pasta_mes_id)

    media = MediaIoBaseUpload(io.BytesIO(pdf_bytes), mimetype="application/pdf")
    metadata = {"name": nome_arquivo, "parents": [pasta_empresa_id]}
    servico_drive.files().create(body=metadata, media_body=media, fields="id").execute()
    print(f"  -> PDF salvo no Drive: {nome_arquivo}")


def gerar_pdf_da_pagina(driver):
    resultado = driver.execute_cdp_cmd("Page.printToPDF", {
        "printBackground": True,
        "preferCSSPageSize": True,
    })
    return base64.b64decode(resultado["data"])


# =============================================================================
# NAVEGADOR (agora headless, preparado para rodar em servidor)
# =============================================================================
def configurar_navegador():
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")

    if CHROME_BIN:
        chrome_options.binary_location = CHROME_BIN

    service = Service(CHROMEDRIVER_PATH) if CHROMEDRIVER_PATH else Service(ChromeDriverManager().install())

    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.set_window_size(1920, 1080)
    return driver


# =============================================================================
# CAPTCHA (sem alterações)
# =============================================================================
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
    except:
        return ""


def preencher_campo(driver, seletor, valor):
    if pd.notna(valor) and str(valor).strip() != "":
        try:
            campo = driver.find_element(By.CSS_SELECTOR, seletor)
            campo.clear()
            val_str = f"{float(valor):.2f}".replace('.', ',') if isinstance(valor, (int, float)) else str(valor).replace('.', ',')
            campo.send_keys(val_str)
            campo.send_keys(Keys.TAB)
        except:
            pass


# =============================================================================
# FLUXO PRINCIPAL
# =============================================================================
def main():
    print("Conectando à planilha do Google...")
    try:
        aba, credenciais = conectar_google()
        df = carregar_dados(aba)
        print("Planilha carregada com sucesso!")
    except Exception as e:
        print("\n[ERRO CRÍTICO] Falha ao conectar/ler a planilha.")
        print(f"Detalhe do erro: {e}")
        return

    driver = None
    cnpj_logado = ""

    for index, row in df.iterrows():
        linha_sheet = index + 2  # +1 pelo cabeçalho, +1 porque index começa em 0

        empresa = str(row.get('Empresa', '')).strip()
        cnpj_p = str(row.get('CNPJ Prestador', '')).strip()
        senha = str(row.get('Senha', '')).strip()
        tomador = str(row.get('CNPJ Tomador', '')).strip()

        revisao_necessaria = str(row.get('Revisão', 'EMITIR DIRETO')).strip().upper()

        if pd.isna(empresa) or empresa == 'nan' or empresa == '' or "Emitido" in str(row.get('Status', '')):
            continue

        print(f"\nIniciando: {empresa} -> Tomador: {tomador}")

        if cnpj_p != cnpj_logado:
            if driver:
                driver.quit()
            driver = configurar_navegador()
            wait = WebDriverWait(driver, 20)

            driver.get("https://nfse.salvador.ba.gov.br/")

            sucesso_login = False
            for tentativa in range(1, 31):
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
                    except:
                        pass
                    driver.find_element(By.CSS_SELECTOR, "input#tbCaptcha").clear()
                    driver.find_element(By.CSS_SELECTOR, "input#tbCaptcha").send_keys(texto)
                    driver.find_element(By.CSS_SELECTOR, "input#cmdLogin").click()
                    time.sleep(7)

                    if "Sair" in driver.page_source or "InformacaoDebito.aspx" in driver.current_url:
                        sucesso_login = True
                        cnpj_logado = cnpj_p
                        print("  -> Login efetuado com sucesso!")
                        break

                try:
                    driver.find_element(By.XPATH, "//*[contains(text(), 'Recarregar')] | //img[contains(@src, 'reload') or contains(@title, 'Recarregar')]").click()
                    time.sleep(1.5)
                except:
                    pass

            if not sucesso_login:
                print("  -> Falha definitiva no Login.")
                atualizar_status_sheet(aba, linha_sheet, "Erro Login")
                continue

        # Fluxo de Emissão
        try:
            driver.get("https://nfse.salvador.ba.gov.br/site/contribuinte/nota/nota.aspx")
            wait = WebDriverWait(driver, 15)

            campo_t = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input#tbCPFCNPJTomador")))
            campo_t.clear()
            campo_t.send_keys(tomador)
            driver.find_element(By.XPATH, "//*[@id='btAvancar'] | //input[@value='Avançar']").click()
            time.sleep(3)

            print("  -> Preenchendo CNAE...")
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

                try:
                    primeira_opcao = driver.find_element(By.CSS_SELECTOR, "ul.chosen-results li.active-result")
                    if primeira_opcao.is_displayed():
                        primeira_opcao.click()
                except:
                    pass

            preencher_campo(driver, "input#tbAliquota", row['Alíquota'])
            driver.find_element(By.CSS_SELECTOR, "textarea#tbDiscriminacao").send_keys(str(row['Discriminação']))
            preencher_campo(driver, "input#tbValor", row['Valor Total'])

            iss = str(row['ISS Retido']).upper()
            if "SIM" in iss:
                driver.find_element(By.XPATH, "//table[@id='rblISSRetido']//label[contains(text(), 'Sim')]/preceding-sibling::input").click()
            else:
                driver.find_element(By.XPATH, "//table[@id='rblISSRetido']//label[contains(text(), 'Não')]/preceding-sibling::input").click()

            if pd.notna(row['Data Competência']):
                preencher_campo(driver, "input#tbRPSEmissao", row['Data Competência'])

            preencher_campo(driver, "input#tbINSS", row['INSS'])
            preencher_campo(driver, "input#tbIRPJ", row['IRRF'])
            preencher_campo(driver, "input#tbCSLL", row['CSLL'])
            preencher_campo(driver, "input#tbCOFINS", row['COFINS'])
            preencher_campo(driver, "input#tbPisPasep", row['PIS'])
            preencher_campo(driver, "input#tbOutrasRetencoes", row['Outras Retenções'])

            time.sleep(2)

            vl_planilha = float(str(row['Valor Líquido Planilha']).replace(',', '.'))
            try:
                txt_site = driver.find_element(By.CSS_SELECTOR, "input#tbValorLiquido").get_attribute('value')
            except:
                txt_site = driver.find_element(By.CSS_SELECTOR, "span#lblValorLiquido").text

            vl_site = float(txt_site.replace('R$', '').replace('.', '').replace(',', '.').strip())

            emitir_agora = False
            if abs(vl_planilha - vl_site) <= 0.01:
                if revisao_necessaria == 'VERIFICAR':
                    # Sem terminal interativo no servidor: não emite sozinho, só sinaliza para revisão humana
                    print("  -> Valores batem, mas a planilha exige revisão manual. Aguardando revisão.")
                    atualizar_status_sheet(aba, linha_sheet, "Aguardando Revisão Manual")
                else:
                    print("  -> Valores validados automaticamente (Emitir Direto).")
                    emitir_agora = True
            else:
                print(f"  -> [ATENÇÃO] Divergência de valores: Planilha R$ {vl_planilha:.2f} x Site R$ {vl_site:.2f}")
                atualizar_status_sheet(aba, linha_sheet, f"Erro: Divergência de valor (Site: {vl_site:.2f})")

            if emitir_agora:
                print("  -> Emitindo nota...")
                driver.find_element(By.CSS_SELECTOR, "input#btEmitir").click()

                num_nota = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "span#lblNumeroNota"))).text
                print(f"  -> Nota Gerada: {num_nota}")

                nome_pdf = f"NF {num_nota} - {empresa}_{MMAAAA_ATUAL}.pdf"
                pdf_bytes = gerar_pdf_da_pagina(driver)
                salvar_pdf_no_drive(credenciais, pdf_bytes, nome_pdf, empresa)

                atualizar_status_sheet(aba, linha_sheet, "Emitido com Sucesso", numero_nota=num_nota)

        except Exception as e:
            print(f"  -> Erro durante a emissão: {str(e)[:50]}")
            atualizar_status_sheet(aba, linha_sheet, f"Erro: {str(e)[:50]}")

    if driver:
        driver.quit()

    try:
        if os.path.exists("captcha.png"):
            os.remove("captcha.png")
    except:
        pass

    print("\nProcesso concluído.")


if __name__ == "__main__":
    main()