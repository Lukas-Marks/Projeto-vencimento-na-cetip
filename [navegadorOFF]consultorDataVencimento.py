import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException

#URL
URL = "https://www.b3.com.br/pt_br/market-data-e-indices/servicos-de-dados/market-data/consultas/mercado-a-vista/codigo-isin/pesquisa/"


def acessar_pagina(driver, wait):
    driver.get(URL)
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "bvmf_iframe")))
    wait.until(EC.presence_of_element_located((By.ID, "isinCode")))


def voltar_para_pesquisa(driver, wait):
    try:
        botao_voltar = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(.,'Voltar')]"))
        )
        botao_voltar.click()
        wait.until(EC.presence_of_element_located((By.ID, "isinCode")))
    except:
        acessar_pagina(driver, wait)


def main():

    # 1️⃣ Ler Excel
    df = pd.read_excel("isins.xlsx")
    lista_isin = df["ISIN"].dropna().astype(str).tolist()

    resultados = []

    # 2️⃣ Configurar navegador HEADLESS
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")  # 🔥 roda invisível
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )

    wait = WebDriverWait(driver, 3)  # ⬅ 5 segundos é mais seguro que 0.6

    acessar_pagina(driver, wait)

    # 4️⃣ Loop pelos ISINs
    for isin in lista_isin:

        print(f"Consultando {isin}...")

        try:
            driver.switch_to.default_content()
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "bvmf_iframe")))

            campo = wait.until(
                EC.presence_of_element_located((By.ID, "isinCode"))
            )

            campo.clear()

            driver.execute_script("""
                arguments[0].value = arguments[1];
                arguments[0].dispatchEvent(new Event('input'));
            """, campo, isin)

            botao_buscar = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(.,'Buscar')]"))
            )
            botao_buscar.click()

            try:
                wait.until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//table[contains(@class,'table')]//tbody/tr")
                    )
                )

                data_expiracao = driver.find_element(
                    By.XPATH,
                    "//table[contains(@class,'table')]//tbody/tr/td[6]"
                ).text.strip()

                if data_expiracao == "":
                    data_expiracao = "data vazia"

            except TimeoutException:
                data_expiracao = "ISIN não encontrado"

            voltar_para_pesquisa(driver, wait)

        except Exception as e:
            print(f"Erro ao consultar {isin}: {e}")
            data_expiracao = "Erro na consulta"
            acessar_pagina(driver, wait)

        resultados.append({
            "ISIN": isin,
            "Data Expiração": data_expiracao
        })

    driver.quit()

    df_final = pd.DataFrame(resultados)
    df_final.to_excel("resultado.xlsx", index=False)

    print("Processo finalizado com sucesso!")


if __name__ == "__main__":
    main()