from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
import json
import requests
import logging
from time import sleep
import warnings
import sys
import os
import re
import datetime
import pytz
import base64
from datetime import datetime, time, timedelta
from imap_tools import MailBox, AND, OR, NOT, A
from pathlib import Path
from bs4 import BeautifulSoup
from config.config import Config
config = Config()

warnings.filterwarnings("ignore")

chrome_options = Options()
chrome_options.add_argument('--incognito')
chrome_options.add_experimental_option("detach", True)

browser = None
num_erros = 0
max_err_count = 5
current_filename = config.inter_extract
download_folder = config.path_download
prd_folder = config.path_prd
created_file_at = str(datetime.now().strftime("%d_%m_%Y - %H_%M_%S"))
log_filename = 'Log_Processamento_' + created_file_at + '.txt'

def F_InitEmailConfig():    

    #PRD Configs
    email               = config.email_user
    password            = config.email_pass
    server              = config.email_server

    try:

        mailbox = MailBox(server)
        mailbox.login(email, password, initial_folder='INBOX')  

        F_WriteLog('Configurações de e-mail realizadas com sucesso!')

        return mailbox
        
    except:
        # logging.error('Erro:', exc_info=True)
        return None

def F_GetFileFolderName():

    return str(Path().absolute())

def F_InitBS(filename, folder):

    with open(folder + '\\' + filename) as fp:
        soup = BeautifulSoup(fp, 'html.parser') 

    return soup

def F_WriteFile(filename, data, folder, type_op='w'):

    if not os.path.exists(folder):
        os.makedirs(folder)

    try:
        with open(folder + '\\' + filename, type_op) as arquivo:
            arquivo.write(data)
    except:
        F_WriteLog('Erro ao escrever arquivo!')
        return None

def F_ReadFromFile(filename, folder):

    if not os.path.exists(folder):
        os.makedirs(folder)

    try:
        with open(folder + '\\' + filename, 'r') as arquivo:
            return arquivo.read()
    except:
        F_WriteLog('Erro ao ler arquivo!')
        return None

def F_CheckIfExistsFile(filename, folder):

    if not os.path.exists(folder):
        os.makedirs(folder)

    if Path(folder + '\\' + filename).is_file():
        return True

    return False

def F_DeleteFileFromFolder(filename, folder):

    if not os.path.exists(folder):
        os.makedirs(folder)

    try:  
        os.remove(folder + '\\' + filename)
    except:
        F_WriteLog('Erro ao deletar arquivo')
        return None

def F_DeleteAllExtratosFromFolder(filename, folder):

    if not os.path.exists(folder):
        os.makedirs(folder)

    try:  
        for file_name in os.listdir(folder):
            if file_name.endswith(".pdf") and 'Extrato de Conta Corrente' in filename:
                F_DeleteFileFromFolder(file_name, folder)         
    except:
        F_WriteLog('Erro ao deletar extratos')
        return None

def F_MoveAndRenameFile(old_filename, new_filename, old_folder, new_folder):

    if not os.path.exists(old_folder):
        os.makedirs(old_folder)

    if not os.path.exists(new_folder):
        os.makedirs(new_folder)

    old_file = old_folder + '\\' + old_filename
    new_file = new_folder + '\\' + new_filename

    try:
        os.rename(old_file, new_file)
        return True
    except:
        F_WriteLog('Erro ao renomear/mover arquivo')
        # logging.error('Erro:', exc_info=True)
        return None
    
def F_DeleteOldEmails(mailbox, sent_from):

    mail_filter     = A(from_=sent_from, deleted=False) # E-mail filter
    email_list      = list(mailbox.fetch(mail_filter)) # Fetch e-mail list

    if len(email_list) > 0:
        F_WriteLog(str(len(email_list)) + ' e-mail(s) identificados para deletar')
        for email in email_list:
            F_WriteLog('Deletando e-mail do dia ' + str(email.date) + ' enviado por ' + str(email.from_) +  ' com assunto: ' + str(email.subject))    
            mailbox.delete(email.uid)
        F_WriteLog('Emails antigos deletados com sucesso')
    else:
        F_WriteLog('Nenhum e-mail encontrado para deletar')

def F_ParseEmailBodyToken(email_body = ''):

    token = ''
    filename = 'Token.txt'

    F_WriteFile(filename, email_body, F_GetFileFolderName())

    bs = F_InitBS(filename, F_GetFileFolderName())
    elements = bs.select('span strong')

    for row in elements:
        if str(row.get_text()).isdecimal() and len(str(row.get_text()).strip()) == 6:
            token = str(row.get_text()).strip()
            F_WriteLog('Token extraído com sucesso')
            break
    
    F_DeleteFileFromFolder(filename, F_GetFileFolderName())

    if token == '':
        return None

    return token

def F_GetEmailToken(mailbox, sent_from):

    email_body          = ''
    value_limit_min     = 15
    limit_date          = datetime.now() - timedelta(minutes=value_limit_min) # X minutes ago

    mail_filter     = A(from_=sent_from, deleted=False) # E-mail filter
    email_list      = [] # Fetch e-mail list

    F_WriteLog('Esperando e-mail com o Token chegar')

    while len(email_list) == 0:
        email_list = list(mailbox.fetch(mail_filter))  # Fetch e-mail list
        
    F_WriteLog('Email com o Token chegou. Recuperando informações do e-mail')

    try:      
        for email in email_list:
            email_date = email.date.replace(tzinfo=pytz.UTC)
            limit_date = limit_date.replace(tzinfo=pytz.UTC)
            if email_date > limit_date:
                F_WriteLog('Salvando HTML do e-mail')
                email_body = email.html
                break

        mailbox.logout()

        if email_body == '':
            F_WriteLog('Erro ao processar informações do E-mail')
            return None

        email_token = F_ParseEmailBodyToken(email_body)

        return email_token

    except:
        F_WriteLog('Erro ao processar informações do e-mail')
        # logging.error('Erro:', exc_info=True)
        return None  

def F_GetListAccounts(list_down, is_logged = True):

    if is_logged is False:
        return []

    if list_down is False:
        browser.find_elements_by_css_selector('#HeaderRender span.open-list')[0].click() # Abrindo a lista
        browser.implicitly_wait(1)
        browser.find_elements_by_css_selector('.aumentarBox')[0].click()

    list_contas = browser.find_elements_by_css_selector('.resultadoBusca.scroll ul li')

    return list_contas

def F_Login(mailbox, sent_from, access):  

    try:

        F_WriteLog('Fazendo o acesso. Inserindo o usuário')

        browser.find_element_by_xpath('//*[@id="loginGeral"]/div/div[1]/div[2]/div[5]').click()
        browser.find_element_by_xpath('//*[@id="loginGeral"]/div/div/div[1]/input').send_keys(access) 
        browser.find_element_by_xpath('//*[@id="loginGeral"]/div/div/button').click()

        browser.implicitly_wait(5) # Segundos

        F_WriteLog('Fazendo o acesso. Inserindo a senha')

        browser.find_element_by_xpath('//*[@id="loginGeral"]/div/div/div/div[3]/div[3]/button[4]').click()
        browser.find_element_by_xpath('//*[@id="loginGeral"]/div/div/div/div[3]/div[3]/button[10]').click()
        browser.find_element_by_xpath('//*[@id="loginGeral"]/div/div/div/div[3]/div[4]/button[1]').click()
        browser.find_element_by_xpath('//*[@id="loginGeral"]/div/div/div/div[3]/div[2]/button[2]').click()
        browser.find_element_by_xpath('//*[@id="loginGeral"]/div/div/div/div[3]/div[2]/button[5]').click()
        browser.find_element_by_xpath('//*[@id="loginGeral"]/div/div/div/div[3]/div[2]/button[8]').click()
        browser.find_element_by_xpath('//*[@id="loginGeral"]/div/div/div/div[3]/div[2]/button[5]').click()
        browser.find_element_by_xpath('//*[@id="loginGeral"]/div/div/div/div[2]/button[1]').click()

        F_WriteLog('Token enviado para o e-mail de acesso. Iniciando a validação do Token')
    
        token = F_GetEmailToken(mailbox, sent_from)

        if token is None:
            F_WriteLog('Erro ao recuperar o Token')
            return False

        F_WriteLog('Inserindo o Token')

        browser.find_element_by_xpath('//*[@id="loginGeral"]/div/div/div[2]/div/form/input').send_keys(token)
        browser.implicitly_wait(20) # Segundos

        return True

    except:
        # logging.error('Erro:', exc_info=True)
        return False

def F_RunAccounts(index, list_contas):

    try:

        list_down = False

        browser.implicitly_wait(2)

        conta = list_contas[index]

        conta_title = conta.find_element_by_xpath('.//*[@class="resultadoBusca--left-title"]').text.strip()
        conta_number = conta.find_element_by_xpath('.//*[@class="resultadoBusca--left-conta"]').text.strip()
        conta_number = re.search("\\d+", conta_number).group(0).strip()

        F_WriteLog('Download do extrato da conta: ' + conta_title + ' - ' + conta_number)

        acc_filename = str(index + 1) + '_' + conta_number + '.pdf'

        if F_CheckIfExistsFile(acc_filename, prd_folder) is False:
            conta.find_elements_by_css_selector('span.resultadoBusca--right-expandirMenu')[0].click()
            browser.implicitly_wait(1)
            conta.find_elements_by_css_selector('.resultadoBusca--bottom.ativo button')[0].click()
            browser.implicitly_wait(3)
            menu_conta_digital = browser.find_element_by_xpath('//*[@id="Menu"]/div/div/nav/ul/li[2]')

            actions = ActionChains(browser)
            actions.move_to_element(menu_conta_digital).perform()

            browser.implicitly_wait(1)
            browser.find_element_by_xpath('//*[@id="Menu"]/div/div/nav/ul/li[2]/div[2]/div/div[1]/ul/li/a').click() # Hover
            browser.implicitly_wait(5)
            browser.find_element_by_xpath('//*[@id="frm:periodoExtrato"]/option[4]').click() # 60 Ultimos dias
            browser.implicitly_wait(3)
            browser.find_element_by_xpath('//*[@id="frm"]/div[2]/div/div[7]/input').click() # Consultar
            browser.implicitly_wait(10)
            browser.find_element_by_xpath('//*[@id="j_idt90"]/div[1]/div[2]/a').click() # Download do PDF

            F_WriteLog('Esperando o download do arquivo de PDF...')

            while F_CheckIfExistsFile(current_filename, download_folder) is False:
                sleep(1)

            F_WriteLog('Movendo arquivo para a pasta de processamento')
            
            is_moved = F_MoveAndRenameFile(current_filename, acc_filename, download_folder, prd_folder)

            if is_moved is None:
                F_WriteLog('Erro ao mover arquivo da conta ' + conta_title + ' - ' + conta_number + '\n')
            else:
                F_WriteLog('Extrato da conta baixado e movido com sucesso\n')

            list_down = False

        else:  
            F_WriteLog('Extrato da conta já existe na pasta de processamento.\n')
            browser.implicitly_wait(3)
            list_down = True

        index = index + 1

        browser.implicitly_wait(2)

        if index == len(list_contas):
            F_WriteLog('Fim do download das contas\n')
            return 

        F_RunAccounts(index, F_GetListAccounts(list_down))
    
    except:
        # logging.error('Erro:', exc_info=True)
        F_WriteLog('Erro inesperado ao baixar conta\n')
        global num_erros
        num_erros += 1
        if num_erros == max_err_count:
            return

def F_WriteLog(mensagem):

    print(mensagem)
    F_WriteFile(log_filename, '>>>> ' + str(datetime.now().strftime("%d/%m/%Y %H:%M:%S")) + ' - '  + mensagem + '\n', F_GetFileFolderName(), 'a')
    
if __name__ == "__main__":

    browser = webdriver.Chrome(r"C:\Users\agenterpa2\Desktop\Chromedrive\Chromedriver84.exe",options=chrome_options)

    F_WriteLog('Início do processamento das contas do banco Inter\n')
    F_WriteLog('Fazendo Login no sistema\n')

    browser.get(config.inter_url)

    mailbox = F_InitEmailConfig()

    if mailbox is None:
        F_WriteLog('Erro ao carregar configurações de E-mail')
        F_MoveAndRenameFile(log_filename, log_filename, F_GetFileFolderName(), F_GetFileFolderName() + '\\backup_log')
        exit(1)

    sent_from = config.inter_sent
    access = config.inter_user

    F_DeleteOldEmails(mailbox, sent_from)
    is_logged = F_Login(mailbox, sent_from, access)

    if is_logged is False:
        F_WriteLog('Erro ao Efetuar login. Tente novamente\n')
        F_MoveAndRenameFile(log_filename, log_filename, F_GetFileFolderName(), F_GetFileFolderName() + '\\backup_log')
        exit(1)

    F_WriteLog('Login efetuado com sucesso\n')
    F_WriteLog('Iniciando download dos extratos\n')

    F_DeleteAllExtratosFromFolder(current_filename, download_folder)

    index_contas = int(sys.argv[1])
    F_RunAccounts(index_contas, F_GetListAccounts(False))

    if num_erros == max_err_count:
        F_WriteLog('Erro ao executar o Download das Contas. Tente novamente\n')
        F_MoveAndRenameFile(log_filename, log_filename, F_GetFileFolderName(), F_GetFileFolderName() + '\\backup_log')
        exit(1)

    F_WriteLog('Processamento das contas finalizado')

    F_MoveAndRenameFile(log_filename, log_filename, F_GetFileFolderName(), F_GetFileFolderName() + '\\backup_log')

    browser.close()
