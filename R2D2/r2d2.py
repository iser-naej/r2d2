"""
Título do projeto: R2D2

Descrição: Acessa o sistema de notas fiscais da prefeitura de lauro de freitas para gerar as notas e posteriormente enviar por e-mail.

Autor: Jean Reis
Data: 13/05/2023

Instruções de uso: **

Dependências: **

Exemplo de uso: **

"""

import os
import configparser

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def main():
    config = configparser.ConfigParser()
    config.read("config.ini")
    login_user = config.get('credenciais', 'username')
    login_pass = config.get('credenciais', 'password')
    

    driver = webdriver.Chrome(executable_path=os.path.abspath('drivers/chromedriver.exe'))
    driver.get('https://lftributos.metropolisweb.com.br/metropolisWEB/?origem=1')
    driver.implicitly_wait(10)

    # fazendo login
    username = driver.find_element(By.NAME, 'login')
    username.click()
    username.send_keys(login_user)
    password = driver.find_element(By.NAME, 'senha')
    password.send_keys(login_pass, Keys.HOME)
    captcha_field = driver.find_element(By.NAME, 'kaptchafield')
    captcha = input('Informe o CAPTCHA para acessar o portal: ')
    captcha_field.send_keys(captcha, Keys.HOME)
    button = driver.find_element(By.XPATH, '//*[@id="botaoEfetuarLogin"]/table/tbody/tr/td[2]/div/input')
    button.click()

    # home page
    try:
        WebDriverWait(driver, 10).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        alert.accept()
    except:
        pass

    driver.get('https://lftributos.metropolisweb.com.br/metropolisWEB/nfa/notaFiscalAvulsaEletronica.do?metodo=executarPrepararIncluir')
    
    driver.find_element(By.NAME, 'tipoTomador').click()

    print()


    driver.quit()


if __name__ == "__main__":
    main()


