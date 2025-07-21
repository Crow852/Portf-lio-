from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import re

def buscar_empresas_google_maps(consulta, limite=10, feedback_callback=None):
    resultados = []

    options = Options()
    options.add_argument("--headless=new")  # Modo invisível
    options.add_argument("--disable-gpu")
    options.add_argument("--start-maximized")
    options.add_experimental_option("excludeSwitches", ["enable-logging"])

    driver = webdriver.Chrome(options=options)

    busca_formatada = consulta.replace(" ", "+")
    url = f"https://www.google.com/maps/search/{busca_formatada}"
    print(f"🔍 Acessando: {url}")
    driver.get(url)

    wait = WebDriverWait(driver, 15)
    try:
        wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'Nv2PK')))
    except:
        print("❌ Nenhum resultado encontrado ou página ainda não carregada.")
        driver.quit()
        return []

    # Scroll no painel lateral
    try:
        scrollable_div = driver.find_element(By.XPATH, '//div[@role="feed"]')
        for _ in range(20):
            driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", scrollable_div)
            time.sleep(1)
    except:
        print("⚠️ Não foi possível scrollar resultados.")

    blocos = driver.find_elements(By.CLASS_NAME, 'Nv2PK')
    print(f"🔎 Blocos capturados: {len(blocos)}")

    for bloco in blocos[1:limite+1]:  # pula o primeiro bloco e garante o número total
        try:
            nome = bloco.find_element(By.CLASS_NAME, 'qBF1Pd').text
        except:
            nome = "Nome não encontrado"

        if feedback_callback:
            feedback_callback(f"🔄 Coletando dados de: {nome}")

        try:
            link = bloco.find_element(By.TAG_NAME, 'a').get_attribute('href')
        except:
            link = "Link não encontrado"

        telefone = "Não encontrado"
        email = "Não encontrado"
        endereco = "Não encontrado"

        try:
            if link and "google.com/maps/place" in link:
                driver.execute_script("window.open('');")
                driver.switch_to.window(driver.window_handles[1])
                driver.get(link)
                wait.until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
                time.sleep(5)

                page_text = driver.find_element(By.TAG_NAME, 'body').text

                # Telefone (fixo ou celular)
                tel_match = re.search(r"(?:\(?\d{2}\)?\s?\d{4,5}-\d{4})|(?:0800\s?\d{3}\s?\d{4})", page_text)
                if tel_match:
                    telefone = tel_match.group(0)

                # E-mail
                email_match = re.search(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", page_text)
                if email_match:
                    email = email_match.group(0)

                # Endereço (buscando pelo elemento com data-item-id="address")
                try:
                    address_element = driver.find_element(By.XPATH, '//button[@data-item-id="address"]/div[1]')
                    endereco = re.sub(r"[\n\r\u200b-\u206f\u00a0-\u00bf\u25a0-\u25ff\ufeff]", "", address_element.text).strip()
                except:
                    endereco = "Endereço não encontrado"

                driver.close()
                driver.switch_to.window(driver.window_handles[0])
        except Exception:
            telefone = "Erro ao buscar"
            email = "Erro ao buscar"
            endereco = "Erro ao buscar"

        resultados.append({
            'Nome': nome,
            'Endereço': endereco,
            'Telefone': telefone,
            'E-mail': email,
            'Link': link
        })

    driver.quit()
    return resultados


# Teste local
if __name__ == "__main__":
    def feedback(msg): print(msg)
    dados = buscar_empresas_google_maps("Empresas de topografia em São Paulo", limite=10, feedback_callback=feedback)
    for item in dados:
        print(item)
