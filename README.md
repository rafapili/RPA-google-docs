
# 📋 Preenchimento Automático de Formulários Google Forms

Este projeto automatiza o preenchimento de um formulário do Google Forms com base em dados extraídos de um arquivo Excel (`pessoas.xlsx`). Ele oferece **duas abordagens diferentes** de automação:

- 🧭 **Selenium**: interage diretamente com os elementos do navegador.  
- ⌨️ **PyAutoGUI**: simula ações humanas no teclado e mouse.

---

## 📁 Estrutura do Projeto

```
📦 formulario_automacao
├── pessoas.xlsx           # Arquivo com os dados das pessoas a serem cadastradas
├── selenium_script.py     # Script usando Selenium
├── pyautogui_script.py    # Script usando PyAutoGUI
└── README.md              # Este arquivo
```

---

## ✅ Pré-requisitos

### 1. Python 3.8 ou superior instalado.

### 2. Instale as dependências com:

```bash
pip install -r requirements.txt
```

Ou crie o arquivo `requirements.txt` com o seguinte conteúdo:

```
pandas
selenium
openpyxl
pyautogui
```

### 3. WebDriver do Chrome (para o Selenium)

Acesse: [https://sites.google.com/chromium.org/driver/](https://sites.google.com/chromium.org/driver/)

- Baixe a versão correspondente ao seu navegador Chrome.
- Extraia e coloque o executável no **PATH do sistema** ou na **mesma pasta do script**.

---

## 📊 Estrutura esperada do `pessoas.xlsx`

O arquivo Excel deve conter as colunas a seguir, com os nomes **exatamente iguais**:

```
nome     email     telefone     data_nascimento     cpf     endereco     cep     genero     aceita_novidades
```

- `genero`: **masculino**, **feminino** ou **outro**  
- `aceita_novidades`: **sim** ou **não**

---

## 🚀 Como usar

### 🧭 Versão com Selenium

```bash
python selenium_script.py
```

- O navegador será aberto automaticamente, os dados serão preenchidos e o formulário será enviado.

### ⌨️ Versão com PyAutoGUI

```bash
python pyautogui_script.py
```

> ⚠️ **Atenção**: Não mexa no mouse ou teclado durante a execução, pois o script simula ações humanas e pode ser interrompido.

