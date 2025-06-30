import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Caminho do Excel
caminho_excel = r"pessoas.xlsx"

# Lê os dados
df = pd.read_excel(caminho_excel)

# Inicializar navegador uma vez
driver = webdriver.Chrome()
wait = WebDriverWait(driver, 10)

def preencher_formulario(pessoa):
    try:
        driver.get("https://forms.gle/VyZ37iWCGmKN7ujaA")
        time.sleep(5)  # Aguarda o carregamento do formulário

        # Preenchendo os campos com tratamento de exceção individual
        try:
            driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(str(pessoa['nome']))
        except Exception as e:
            print(f"Erro no campo nome: {e}")
        time.sleep(0)
        try:
            driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(str(pessoa['email']))
        except Exception as e:
            print(f"Erro no campo email: {e}")
        time.sleep(0)

        try:
            driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(str(pessoa['telefone']))
        except Exception as e:
            print(f"Erro no campo telefone: {e}")
        time.sleep(0)

        try:
            data_nascimento = str(pessoa['data_nascimento'])
            driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div/div[2]/div[1]/div/div[1]/input').send_keys(data_nascimento)
        except Exception as e:
            print(f"Erro no campo data de nascimento: {e}")
        time.sleep(0)

        try:
            driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[5]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(str(pessoa['cpf']))
        except Exception as e:
            print(f"Erro no campo cpf: {e}")
        time.sleep(0)

        try:
            driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[6]/div/div/div[2]/div/div[1]/div[2]/textarea').send_keys(str(pessoa['endereco']))
        except Exception as e:
            print(f"Erro no campo endereco: {e}")
        time.sleep(0)

        try:
            driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[7]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(str(pessoa['cep']))
        except Exception as e:
            print(f"Erro no campo cep: {e}")
        time.sleep(0)

        # Gênero
        genero = str(pessoa['genero']).lower()
        try:
            if genero == 'masculino':
                driver.find_element(By.XPATH, '//*[@id="i42"]/div[3]/div').click()
            elif genero == 'feminino':
                driver.find_element(By.XPATH, '//*[@id="i45"]/div[3]/div').click()
            else:
        
                driver.find_element(By.XPATH, '//*[@id="i42"]/div[5]/div').click()
        except Exception as e:
            print(f"Erro ao clicar no gênero: {e}")
        time.sleep(1)
        # Novidades
        aceita = str(pessoa['aceita_novidades']).lower()
        try:
            if aceita == 'sim':
                driver.find_element(By.XPATH, '//*[@id="i59"]/div[3]/div').click()
            else:
                driver.find_element(By.XPATH, '//*[@id="i62"]/div[3]/div').click()
        except Exception as e:
            print(f"Erro ao clicar nas novidades: {e}")

        # Enviar
        try:
            botao_enviar = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[3]/div[1]/div[1]/div/span/span')))
            try:
                botao_enviar.click()
            except:
                driver.execute_script("arguments[0].click();", botao_enviar)
        except Exception as e:
            print(f"Erro ao clicar no botão enviar: {e}")

        try:
            botao_enviar2 = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[3]/div[1]/div[1]/div[2]/span/span')))
            try:
                botao_enviar2.click()
            except:
                driver.execute_script("arguments[0].click();", botao_enviar)
        except Exception as e:
            print(f"Erro ao clicar no botão enviar: {e}")
        # Espera para evitar problemas no próximo envio
        time.sleep(5)

    except Exception as e:
        print(f"Erro geral ao preencher o formulário: {e}")

# Loop para enviar todos os cadastros
for idx, linha in df.iterrows():
    print(f"Enviando cadastro {idx+1} para {linha['nome']}")
    preencher_formulario(linha)

driver.quit()
print("Todos os cadastros foram enviados com sucesso.")
