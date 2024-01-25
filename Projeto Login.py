#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pandas as pd
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
import time

class FormatLogin:
    def __init__(self):
        self.data = []

    def formatar(self, cpf, nome, cursos, email=None):
        cpf_formatado = re.sub(r'\D', '', cpf)
        if len(cpf_formatado) != 11:
            raise ValueError("CPF deve conter 11 dígitos.")

        partes_nome = nome.split()
        primeiro_nome = partes_nome[0].capitalize()
        sobrenomes = ' '.join(partes_nome[1:]).title()

        email = email or f"{nome.replace(' ', '').lower()}@hotmail.com"

        cursos_formatados = [(curso, 'student') for curso in cursos]

        self.data.append((cpf_formatado, primeiro_nome, sobrenomes, cursos_formatados, email))

    def criar_dataframe(self):
        colunas = ['username', 'password', 'firstname', 'lastname', 'email']
        max_num_cursos = max(len(cursos) for _, _, _, cursos, _ in self.data)

        for i in range(max_num_cursos):
            colunas.extend([f'Course{i + 1}', f'Role{i + 1}'.strip()])

        data_dict = {coluna.lower(): [] for coluna in colunas}

        for cpf, primeiro_nome, sobrenomes, cursos, email in self.data:
            row_data = [cpf, '123', primeiro_nome, sobrenomes, email]
            cursos_preenchidos = cursos + [('','')]*(max_num_cursos - len(cursos))

            for curso, role in cursos_preenchidos:
                row_data.extend([curso, role])

            for coluna, valor in zip(colunas, row_data):
                data_dict[coluna.lower()].append(valor)

        return pd.DataFrame(data_dict)

def obter_cursos_escolhidos(cursos_disponiveis):
    print("Cursos disponíveis:")
    for i, curso in enumerate(cursos_disponiveis, 1):
        print(f"{i}. {curso}")

    escolha = input("Digite os números dos cursos (separados por vírgula): ")
    cursos_escolhidos = []

    for indice in map(str.strip, escolha.split(',')):
        try:
            indice = int(indice)
            if 1 <= indice <= len(cursos_disponiveis):
                cursos_escolhidos.append(cursos_disponiveis[indice - 1])
            else:
                print("Escolha inválida. Ignorando.")
        except ValueError:
            print("Escolha inválida. Ignorando.")

    return cursos_escolhidos

def main():
    formatador = FormatLogin()

    while True:
        cpf = input("CPF (ou pressione Enter para sair): ")
        if not cpf:
            break

        nome = input("Nome: ")
        email = input("E-mail (pressione Enter para gerar automaticamente): ")

        cursos_disponiveis = [
            "NR11TM", "NR10A", "NR26", "SEPI", "NR10I",
            "BI", "CPI", "DD", "EB", "LOTO", "NPSL",
            "NR11OR", "STME", "NR5", "EPI", "NR10",
            "SEP", "NR11OP", "NR11PR", "NR12MEE",
            "NR17E", "NR18SEV", "NR18SS", "NR18IS",
            "NR20SSTIC", "NR33TECS", "NR3316H", "NR3516H", "PSM",
            "LO"
        ]

        cursos_escolhidos = obter_cursos_escolhidos(cursos_disponiveis)

        try:
            formatador.formatar(cpf, nome, cursos_escolhidos, email)
        except ValueError as e:
            print(f"Erro: {e}")

    print("Dados inseridos para confirmação:")
    df_export = formatador.criar_dataframe()
    print(df_export)

    confirmacao = input("Os dados estão corretos? (S/N): ").strip().lower()
    if confirmacao == 's':
        caminho_arquivo_csv = r"C:\Users\cp6at\Documents\Automacao_planilha_csv_ATIC\Planilha_de_automacao_ATIC.csv"
        df_export.to_csv(caminho_arquivo_csv, index=False, sep=';')
        print(f"Dados exportados para {caminho_arquivo_csv}")

if __name__ == "__main__":
    main()

def enviar_arquivo_para_moodle():
    # Configurar o WebDriver (certifique-se de ter o WebDriver correspondente ao seu navegador instalado)
    Login = input("Login: ")
    Senha = input("Senha: ")
    driver = webdriver.Chrome()  # Ou use o webdriver do navegador de sua escolha

    try:
        # Abrir a página do Moodle
        driver.get("https://atic.com.br/metrolms/login/index.php")
        caminho_arquivo_csv = r"C:\Users\cp6at\Documents\Automacao_planilha_csv_ATIC\Planilha_de_automacao_ATIC.csv"

        # Preencher o campo de login
        campo_login = driver.find_element(By.XPATH, '//*[@id="username"]')
        campo_login.send_keys(Login)

        # Preencher o campo de senha
        
        campo_senha = driver.find_element(By.XPATH, '//*[@id="password"]')
        campo_senha.send_keys(Senha)
        campo_senha.submit()

        # Esperar por alguns segundos para garantir que a página seja totalmente carregada
        time.sleep(3)

        # Navegar para a página de cadastro de senhas
        driver.get("https://atic.com.br/metrolms/admin/tool/uploaduser/index.php")

        # Aguardar até que o elemento de botão para fazer upload esteja visível
        time.sleep(3)
        elemento_upload = driver.find_element(By.XPATH, '//*[@id="yui_3_17_2_1_1703093443259_869"]')
        elemento_upload.send_keys(r"C:\Users\cp6at\Documents\Automacao_planilha_csv_ATIC\Planilha_de_automacao_ATIC.csv")

        # Aguardar o upload ser concluído (ajuste o tempo conforme necessário)
        time.sleep(10)

    finally:
        confirmacao = input("Fechar o WebDriver? (s/n): ")
        if confirmacao.lower() == 's':
            driver.quit()

            
# Exemplo de chamada da função
enviar_arquivo_para_moodle()


# In[ ]:




