import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException


URL = "https://www.b3.com.br/pt_br/market-data-e-indices/servicos-de-dados/market-data/consultas/mercado-a-vista/codigo-isin/pesquisa/"


def acessar_pagina(driver, wait):
    """Acessa a página e entra no iframe."""
    driver.get(URL)
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "bvmf_iframe")))
    wait.until(EC.presence_of_element_located((By.ID, "isinCode")))


def voltar_para_pesquisa(driver, wait):
    """Clica no botão Voltar e aguarda tela de pesquisa."""
    try:
        botao_voltar = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(.,'Voltar')]"))
        )
        botao_voltar.click()

        # Esperar o campo reaparecer
        wait.until(EC.presence_of_element_located((By.ID, "isinCode")))
    except:
        # Se falhar, recarrega a página completamente
        acessar_pagina(driver, wait)


def main():

    # 1️⃣ Ler Excel
    df = pd.read_excel("isins.xlsx")
    lista_isin = df["ISIN"].dropna().astype(str).tolist()

    resultados = []

    # 2️⃣ Configurar navegador
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    driver.maximize_window()
    wait = WebDriverWait(driver, 5)

    # 3️⃣ Acessar página inicial
    acessar_pagina(driver, wait)

    # 4️⃣ Loop pelos ISINs
    for isin in lista_isin:

        print(f"Consultando {isin}...")

        try:
            # Garantir que estamos dentro do iframe
            driver.switch_to.default_content()
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "bvmf_iframe")))

            # Localizar campo
            campo = wait.until(
                EC.presence_of_element_located((By.ID, "isinCode"))
            )

            campo.clear()

            # Inserir valor via JS (Angular)
            driver.execute_script("""
                arguments[0].value = arguments[1];
                arguments[0].dispatchEvent(new Event('input'));
            """, campo, isin)

            # Clicar no botão Buscar
            botao_buscar = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(.,'Buscar')]"))
            )
            botao_buscar.click()

            # Esperar resultado OU mensagem de não encontrado
            try:
                wait.until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//table[contains(@class,'table')]//tbody/tr")
                    )
                )

                # Capturar Data Expiração (6ª coluna)
                data_expiracao = driver.find_element(
                    By.XPATH,
                    "//table[contains(@class,'table')]//tbody/tr/td[6]"
                ).text

            except TimeoutException:
                data_expiracao = "ISIN não encontrado"

            # Voltar para tela de pesquisa
            voltar_para_pesquisa(driver, wait)

        except Exception as e:
            print(f"Erro ao consultar {isin}: {e}")
            data_expiracao = "Erro na consulta"
            acessar_pagina(driver, wait)

        resultados.append({
            "ISIN": isin,
            "Data Expiração": data_expiracao
        })

    # 5️⃣ Fechar navegador
    driver.quit()

    # 6️⃣ Gerar Excel final
    df_final = pd.DataFrame(resultados)
    df_final.to_excel("resultado.xlsx", index=False)

    print("Processo finalizado com sucesso!")


if __name__ == "__main__":
    main()