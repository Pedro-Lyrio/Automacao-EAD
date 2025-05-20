import os
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from dotenv import load_dotenv
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from time import sleep
from datetime import datetime
import re
import unicodedata
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys



prefixo_para_url_participantes = {
    "CBP": "https://eadoticsrio.com.br/course/view.php?id=5",
    "CIP": "https://eadoticsrio.com.br/course/view.php?id=4",
    "CAP": "https://eadoticsrio.com.br/course/view.php?id=2"
}

prefixos = {
    "curso básico de planilhas": "CBP",
    "curso intermediário de planilhas": "CIP",
    "curso avançado de planilhas": "CAP"
}

load_dotenv()

# Lista para armazenar status dos usuários
status_usuarios = []

def conectar_planilha():
    escopo = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credenciais = ServiceAccountCredentials.from_json_keyfile_name("credenciais.json", escopo)
    cliente = gspread.authorize(credenciais)
    planilha = cliente.open_by_url(os.getenv("SPREADSHEET_URL"))
    aba = planilha.worksheet("Respostas ao formulário 1")
    return aba

def formatar_username(nome_completo):
    nome_completo = unicodedata.normalize('NFKD', nome_completo).encode('ASCII', 'ignore').decode('utf-8')
    nome_completo = nome_completo.lower().replace(" ", "_")
    nome_completo = re.sub(r'[^a-z0-9_]', '', nome_completo)
    return nome_completo

def atualizar_coluna_nada_para_plataforma(aba):
    valores = aba.col_values(1)
    todos_ids = aba.col_values(3)
    cursos = aba.col_values(6)
    linhas_atualizadas = []

    for idx, valor in enumerate(valores[1:], start=2):  
        if valor.strip().lower() in ("nada", ''):
            linha = aba.row_values(idx)
            if len(linha) >= 6 and linha[3] != '' and linha[4] != '':
                aba.update_cell(idx, 1, "Plataforma")
                aba.update_cell(idx, 2, datetime.now().strftime('%Y-%m-%d %H:%M:%S'))

                if len(linha) < 3 or linha[2].strip() == '':
                    curso_desejado = linha[5].strip().lower()
                    prefixo = next((v for k, v in prefixos.items() if k in curso_desejado), None)
                    if prefixo:
                        todos_ids_atualizados = aba.col_values(3)
                        ids_do_prefixo = [id for id in todos_ids_atualizados if id.startswith(prefixo)]
                        numeros = [int(re.findall(r'\d+', id)[0]) for id in ids_do_prefixo if re.findall(r'\d+', id)]
                        novo_id = f"{prefixo}{max(numeros, default=0) + 1}"
                        aba.update_cell(idx, 3, novo_id)
                        print(f"Linha {idx}: Gerado ID {novo_id} para o curso '{curso_desejado.title()}'")

                print(f"Linha {idx}: 'Nada' -> 'Plataforma' e data de inscrição preenchida.")
                linhas_atualizadas.append(idx)
                sleep(1)
            else:
                print(f"Linha {idx}: Dados incompletos. Pulando este registro.")
                break
    return linhas_atualizadas

def ler_planilha():
    aba = conectar_planilha()
    import pandas as pd
    return pd.DataFrame(aba.get_all_records())


def registrar_status_usuario(nome, email, status, mensagem="", username=None, curso=None, acao="criacao"):
    status_usuarios.append({
        "nome": nome,
        "email": email,
        "status": status,
        "mensagem": mensagem,
        "username": username,
        "curso": curso,
        "acao": acao
    })

def enviar_email(de, senha, para, assunto, corpo_base=""):
    criados = [u['email'] for u in status_usuarios if u['acao'] == 'criacao' and u['status'] == 'sucesso']
    criacao_falha = [f"{u['email']} - erro: {u['mensagem']}" for u in status_usuarios if u['acao'] == 'criacao' and u['status'] == 'erro']
    matriculados = [f"{u['email']} - Curso: {u.get('curso', 'Desconhecido')}" for u in status_usuarios if u['acao'] == 'matricula' and u['status'] == 'sucesso']
    matricula_falha = [f"{u['email']} - erro: {u['mensagem']}" for u in status_usuarios if u['acao'] == 'matricula' and u['status'] == 'erro']

    corpo = (
        "✅ Usuários criados com sucesso:\n" + ("\n".join(criados) or "Nenhum") +
        "\n\n❌ Falha na criação dos seguintes usuários:\n" + ("\n".join(criacao_falha) or "Nenhum") +
        "\n\n✅ Alunos matriculados com sucesso:\n" + ("\n".join(matriculados) or "Nenhum") +
        "\n\n❌ Falha na matrícula dos seguintes usuários:\n" + ("\n".join(matricula_falha) or "Nenhum")
    )

    msg = MIMEMultipart()
    msg["From"] = de
    msg["To"] = para
    msg["Subject"] = assunto
    msg.attach(MIMEText(corpo, "plain"))

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as servidor:
            servidor.login(de, senha)
            servidor.sendmail(de, para, msg.as_string())
        print("✅ E-mail enviado com sucesso.")
    except Exception as e:
        print(f"❌ Erro ao enviar e-mail: {e}")


def matricular_usuario_pelo_nome_do_curso(driver, email, nome_curso, prefixo):
    try:
        print(f"Iniciando matrícula de {email} no curso: {nome_curso}")
        
        # Recupera a URL do curso pelo prefixo
        url_participantes = prefixo_para_url_participantes.get(prefixo)
        if not url_participantes:
            raise Exception(f"URL do curso com prefixo '{prefixo}' não encontrada.")

        # Abre a página do curso
        driver.get(url_participantes)
        time.sleep(2)

        # Clica na aba "Participantes"
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.LINK_TEXT, "Participantes"))
        ).click()

        # Clica no botão "Inscrever usuários"
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@type='submit' and @value='Inscrever usuários']"))
        ).click()

        # Aguarda campo de busca de usuário aparecer
        campo_busca = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[id^='form_autocomplete_input']"))
        )

        # Digita o e-mail do aluno
        campo_busca.send_keys(email)
        time.sleep(3)  # tempo aumentado para garantir carregamento da sugestão
        campo_busca.send_keys(Keys.ENTER)  # seleciona usuário sugerido
        time.sleep(1.5)  # pequeno delay para garantir seleção

        # Clica no botão "Inscrever usuários" final
        botao_final = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-action='save']"))
        )
        botao_final.click()

        print(f"✅ Usuário {email} matriculado com sucesso no curso '{nome_curso}'.")

        # Atualiza planilha
        aba = conectar_planilha()
        registros = aba.get_all_records()
        for idx, registro in enumerate(registros, start=2):
            if registro.get("Endereço de e-mail", "").strip().lower() == email.strip().lower():
                valor_dropdown = "Plataforma e Curso"
                try:
                    aba.update_cell(idx, 1, valor_dropdown)
                    print(f"📝 Linha {idx} atualizada para '{valor_dropdown}'")
                except Exception as e:
                    print(f"⚠️ Erro ao atualizar célula da linha {idx} com valor do dropdown: {e}")

    except Exception as e:
        print(f"❌ Erro ao matricular {email} no curso '{nome_curso}': {e}")
        raise e

