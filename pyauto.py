import pandas as pd
import pyautogui
import time
import webbrowser

# Caminho do Excel
caminho_excel = r"pessoas.xlsx"
df = pd.read_excel(caminho_excel)

# Função para preencher o formulário com teclado
def preencher_com_teclado(pessoa):
    # Abrir o formulário
    webbrowser.open("https://forms.gle/VyZ37iWCGmKN7ujaA")
    time.sleep(7)  # Espera o navegador carregar

    # Começar a preencher com TABs e digitação
    pyautogui.write(str(pessoa['nome']))
    pyautogui.press('tab')

    pyautogui.write(str(pessoa['email']))
    pyautogui.press('tab')

    pyautogui.write(str(pessoa['telefone']))
    pyautogui.press('tab')

    pyautogui.write(str(pessoa['data_nascimento']))
    pyautogui.press('tab')
    pyautogui.press('tab')

    pyautogui.write(str(pessoa['cpf']))
    pyautogui.press('tab')

    pyautogui.write(str(pessoa['endereco']))
    pyautogui.press('tab')

    pyautogui.write(str(pessoa['cep']))
    pyautogui.press('tab')

    # Gênero (escolhe com espaço, supondo que esteja em uma pergunta do tipo múltipla escolha)
    genero = str(pessoa['genero']).lower()
    if genero == 'masculino':
        pyautogui.press('space')
    elif genero == 'feminino':
        pyautogui.press('down')
        pyautogui.press('space')
    else:
        pyautogui.press('down', presses=2)
        pyautogui.press('space')
    pyautogui.press('tab')

    # Novidades
    aceita = str(pessoa['aceita_novidades']).lower()
    if aceita == 'sim':
        pyautogui.press('space')
    else:
        pyautogui.press('down')
        pyautogui.press('space')
    pyautogui.press('tab')

    # Enviar
    pyautogui.press('enter')

    # Espera carregar a próxima página
    time.sleep(3)

    # Fecha a aba para o próximo envio
    pyautogui.hotkey('ctrl', 'w')
    time.sleep(2)

# Executar o envio
for idx, linha in df.iterrows():
    print(f"Enviando {idx+1}: {linha['nome']}")
    preencher_com_teclado(linha)

print("Todos os cadastros foram enviados.")