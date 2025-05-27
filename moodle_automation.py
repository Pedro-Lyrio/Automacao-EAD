import os
import pandas as pd
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time
from time import sleep
from utils import (
    ler_planilha, enviar_email, atualizar_coluna_nada_para_plataforma,
    conectar_planilha, formatar_username, registrar_status_usuario,
    matricular_usuario_pelo_nome_do_curso, atualizar_dados_certificados
)
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

load_dotenv()

prefixo_para_nome_curso = {
    "CBP": "Curso Básico de Planilhas",
    "CIP": "Curso Intermediário de Planilhas",
    "CAP": "Curso Avançado de Planilhas"
}

urls_certificados = {
    'B': 'https://eadoticsrio.com.br/mod/customcert/view.php?id=208',  # Básico
    'C': 'https://eadoticsrio.com.br/mod/customcert/view.php?id=236',  # Intermediário
    'D': 'https://eadoticsrio.com.br/mod/customcert/view.php?id=42',  # Avançado
}


def gerar_usernames(nome_completo):
    partes = [p.lower() for p in nome_completo.strip().split()]
    if len(partes) < 2:
        return [formatar_username(nome_completo)]

    primeiro = partes[0]
    sobrenomes = partes[1:]

    usernames = []
    for s in reversed(sobrenomes):
        usernames.append(formatar_username(f"{primeiro}_{s}"))
    usernames.append(formatar_username(f"{primeiro}_{'_'.join(sobrenomes)}"))
    usernames.append(formatar_username(f"{primeiro}_1"))  # fallback

    return usernames

aba = conectar_planilha()
linhas_atualizadas = atualizar_coluna_nada_para_plataforma(aba)

if not linhas_atualizadas:
    # Mesmo que não tenha novos alunos, atualiza os certificados.
    chrome_options = Options()
    chrome_options.add_argument("--incognito")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")

    driver = webdriver.Chrome(options=chrome_options)
    driver.get("https://eadoticsrio.com.br/login/index.php")

    try:
        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, "username")))
        driver.find_element(By.ID, "username").send_keys(os.getenv("MOODLE_USERNAME"))
        driver.find_element(By.ID, "password").send_keys(os.getenv("MOODLE_PASSWORD"))
        driver.find_element(By.ID, "loginbtn").click()
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "page-footer")))
        print("Login realizado com sucesso.")
    except Exception as e:
        print("Erro ao tentar fazer login como administrador.", e)
        driver.save_screenshot("erro_login.png")
        driver.quit()
        exit()
if not linhas_atualizadas:
    # Atualiza certificados de todos os cursos
    for coluna, url in urls_certificados.items():
        atualizar_dados_certificados(driver, url, coluna)
    driver.quit()

    enviar_email(
        os.getenv("EMAIL_FROM"),
        os.getenv("EMAIL_PASSWORD"),
        os.getenv("EMAIL_TO"),
        "Automação Moodle - Nenhum novo aluno",
        "Nenhum novo aluno encontrado."
    )
    exit()

df = ler_planilha()
pendentes = df.iloc[[idx - 2 for idx in linhas_atualizadas]].reset_index()

if pendentes.empty:
    # Precaução: deveria ser impossível, mas garante robustez.
    chrome_options = Options()
    chrome_options.add_argument("--incognito")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")

    driver = webdriver.Chrome(options=chrome_options)
    driver.get("https://eadoticsrio.com.br/login/index.php")

    try:
        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, "username")))
        driver.find_element(By.ID, "username").send_keys(os.getenv("MOODLE_USERNAME"))
        driver.find_element(By.ID, "password").send_keys(os.getenv("MOODLE_PASSWORD"))
        driver.find_element(By.ID, "loginbtn").click()
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "page-footer")))
        print("Login realizado com sucesso.")
    except Exception as e:
        print("Erro ao tentar fazer login como administrador.", e)
        driver.save_screenshot("erro_login.png")
        driver.quit()
        exit()

    atualizar_dados_certificados(driver)
    driver.quit()

    enviar_email(
        os.getenv("EMAIL_FROM"),
        os.getenv("EMAIL_PASSWORD"),
        os.getenv("EMAIL_TO"),
        "Automação Moodle - Nenhum novo aluno",
        "Nenhum novo aluno encontrado."
    )
    exit()

chrome_options = Options()
chrome_options.add_argument("--incognito")
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

driver = webdriver.Chrome(options=chrome_options)
driver.get("https://eadoticsrio.com.br/login/index.php")

try:
    WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, "username")))
    driver.find_element(By.ID, "username").send_keys(os.getenv("MOODLE_USERNAME"))
    driver.find_element(By.ID, "password").send_keys(os.getenv("MOODLE_PASSWORD"))
    driver.find_element(By.ID, "loginbtn").click()
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "page-footer")))
    print("Login realizado com sucesso.")
except Exception as e:
    print("Erro ao tentar fazer login como administrador.", e)
    driver.save_screenshot("erro_login.png")
    driver.quit()
    exit()

for _, row in pendentes.iterrows():
    nome_completo = row['Qual o seu nome completo?']
    email = row['Endereço de e-mail']
    numero_ead = row['Numero identificação EAD']

    if pd.isna(nome_completo) or not isinstance(nome_completo, str) or nome_completo.strip() == '':
        print("Nome completo ausente. Pulando registro.")
        continue
    if pd.isna(email) or not isinstance(email, str) or email.strip() == '':
        print("Email ausente. Pulando registro.")
        continue

    usernames_testados = gerar_usernames(nome_completo)
    sucesso = False

    for identificador in usernames_testados:
        try:
            driver.get("https://eadoticsrio.com.br/user/editadvanced.php?id=-1")
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "id_username")))

            driver.find_element(By.ID, "id_username").clear()
            driver.find_element(By.ID, "id_username").send_keys(identificador)
            driver.find_element(By.ID, "id_createpassword").click()
            driver.find_element(By.ID, "id_email").clear()
            driver.find_element(By.ID, "id_email").send_keys(email)
            driver.find_element(By.ID, "id_firstname").clear()
            driver.find_element(By.ID, "id_firstname").send_keys(nome_completo.split()[0])
            driver.find_element(By.ID, "id_lastname").clear()
            driver.find_element(By.ID, "id_lastname").send_keys(" ".join(nome_completo.split()[1:]))

            try:
                botao_opcional = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "collapseElement-4"))
                )
                if botao_opcional.get_attribute("aria-expanded") == "false":
                    botao_opcional.click()
                    WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "id_idnumber")))
            except:
                pass

            if not pd.isna(numero_ead):
                try:
                    campo_idnumber = driver.find_element(By.ID, "id_idnumber")
                    campo_idnumber.clear()
                    campo_idnumber.send_keys(str(numero_ead))
                except:
                    pass

            driver.find_element(By.ID, "id_submitbutton").click()
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            time.sleep(2)

            current_url = driver.current_url
            if "admin/user.php" in current_url:
                sucesso = True
                registrar_status_usuario(nome_completo, email, 'sucesso', 'Usuário criado com sucesso.', username=identificador, acao='criacao')
                break
            else:
                erro_email = driver.find_elements(By.ID, "id_error_email")
                if erro_email and "já foi registrado" in erro_email[0].text.lower():
                    sucesso = True
                    registrar_status_usuario(nome_completo, email, 'atenção', 'E-mail já foi registrado. Prosseguindo com matrícula.', username=identificador, acao='criacao')
                    break
        except Exception as e:
            print(f"Tentativa com {identificador} falhou: {e}")
            continue

    if not sucesso:
        registrar_status_usuario(nome_completo, email, 'erro', 'Falha na criação após várias tentativas.', acao='criacao')
        continue

    # Matrícula
    if numero_ead and isinstance(numero_ead, str) and len(numero_ead) >= 3:
        prefixo = numero_ead[:3].upper()
        nome_curso = prefixo_para_nome_curso.get(prefixo)
        if nome_curso:
            try:
                matricular_usuario_pelo_nome_do_curso(driver, email, nome_curso, prefixo)
                registrar_status_usuario(nome_completo, email, 'sucesso', 'Aluno matriculado com sucesso.', username=identificador, curso=nome_curso, acao='matricula')
            except Exception as e:
                registrar_status_usuario(nome_completo, email, 'erro', f"Erro ao matricular no curso: {e}", username=identificador, curso=nome_curso, acao='matricula')

# Sempre atualiza certificados no final
for coluna, url in urls_certificados.items():
    atualizar_dados_certificados(driver, url, coluna)

driver.quit()

assunto = "Automação Moodle - Relatório de Execução"
corpo = "Resumo da criação e matrícula de usuários no Moodle."

enviar_email(
    os.getenv("EMAIL_FROM"),
    os.getenv("EMAIL_PASSWORD"),
    os.getenv("EMAIL_TO"),
    assunto,
    corpo,
)
