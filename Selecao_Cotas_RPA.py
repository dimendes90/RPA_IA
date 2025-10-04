# Bibliotecas
%pip install undetected-chromedriver

import os                                # Interação com o sistema operacional
import time                              # Controle de tempo e pausas
import pandas as pd                      # Manipulação de dados em tabelas
import pandas_gbq                        # Integra pandas com Google BigQuery
import csv                               # Manipulação de arquivos CSV
import pyperclip                         # Copia e cola texto na área de transferência
import tkinter as tk                     # Cria interfaces gráficas (GUIs) com Python
import json                              # Manipulação de dados no formato JSON
import sys
from PIL import Image, ImageTk                    # Manipulação de imagens (Pillow é uma biblioteca de processamento de imagens)
from datetime import datetime
import re                                # Importe o módulo de expressões regulares no início do seu script
from tkinter import ttk, scrolledtext, messagebox  # Componentes da GUI do tkinter
from tkinter import filedialog            # Diálogo de seleção de arquivos

from selenium import webdriver               # Controla o navegador via Selenium
from selenium.webdriver.common.by import By  # Localiza elementos HTML (por ID, classe, etc.)
from selenium.webdriver.support.ui import Select   # Interage com menus suspensos (<select>)
from selenium.webdriver.common.keys import Keys    # Simula pressionamento de teclas
from selenium.webdriver.common.action_chains import ActionChains # Realiza ações complexas com o mouse e teclado
from selenium.webdriver.chrome.service import Service  # Gerencia o serviço do ChromeDriver
from selenium.webdriver.support.ui import WebDriverWait          # Aguarda elementos na página
from selenium.webdriver.support.wait import WebDriverWait        # (duplicado) Aguarda elementos
from selenium.webdriver.support import expected_conditions as EC # Define condições de espera
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import Select  # Importa a classe Select para interagir com menus suspensos
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

from decimal import Decimal, ROUND_DOWN
from dataclasses import fields  # Permite definir classes com atributos tipados
from dataclasses import dataclass   # Facilita a criação de classes simples para armazenar dados    
from datetime import datetime, timedelta # Manipulação de datas e horas
import logging  # Configuração de logs para depuração e monitoramento


from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time

options = Options()
options.add_argument('--disable-backgrounding-occluded-windows')  # Impede que abas em 2º plano sejam pausadas
options.add_argument('--no-sandbox')
options.add_experimental_option("detach", True)  # Evita que a aba feche com o script

import time
import json
import random
import os
from pathlib import Path

import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import re
import pandas as pd
from narwhals import Time

#Action
global driver, wait, action
action = ActionChains(driver)


# CONFIGURAÇÃO - ajuste conforme seu ambiente
PROFILE_DIR = r"C:/selenium/chrome-profile"   # seu user-data-dir
USER_AGENT = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
              "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/141.0.7390.55 Safari/537.36")
COOKIES_FILE = Path("cookies_saved.json")
#START_URL = "https://example.com"   # página inicial de teste (troque pelo site real)

# utilitário: delays "humanos"
def human_sleep(a=0.1, b=0.9):
    time.sleep(random.uniform(a, b))

# utilitário: digitação com delays entre teclas
def human_type(element, text, delay_min=0.001, delay_max=0.3):
    for ch in text:
        element.send_keys(ch)
        time.sleep(random.uniform(delay_min, delay_max))

# salvar cookies atuais do driver em arquivo json
def save_cookies(driver, path: Path):
    cookies = driver.get_cookies()
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cookies, f, indent=2)
    print(f"Cookies salvos em {path}")

# carregar cookies de arquivo (o driver deve estar na mesma origem/domínio antes)
def load_cookies(driver, path: Path):
    if not path.exists():
        print("Arquivo de cookies não existe:", path)
        return
    with open(path, "r", encoding="utf-8") as f:
        cookies = json.load(f)
    for ck in cookies:
        # remover itens que o selenium pode reclamar (expiry em floats etc)
        ck_copy = {k: v for k, v in ck.items() if k in ("name", "value", "path", "domain", "expiry", "secure", "httpOnly", "sameSite")}
        try:
            driver.add_cookie(ck_copy)
        except Exception as e:
            print("Warning: cookie add failed:", ck_copy.get("name"), e)
    print(f"Cookies carregados de {path}")


#=======================================================================================================================
#                  FUNÇÃO 0 - Iniciar Driver com as configurações iniciais
#=======================================================================================================================
# Função INICIAL - iniciar driver com perfil, user-agent e stealth
def iniciar_driver():
    global driver, wait
    options = uc.ChromeOptions()
    # use profile existente - ajuda a parecer usuário real
    options.add_argument(f"--user-data-dir={PROFILE_DIR}")
    # força user-agent coerente
    options.add_argument(f"--user-agent={USER_AGENT}")

    # outras flags úteis (opcionais)
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--lang=pt-BR")

    # Cria o driver (uc faz download do chromedriver apropriado)
    driver = uc.Chrome(options=options)
    wait = WebDriverWait(driver, 15)

    try:
        # ajuste de tamanho (parece mais "real")
        driver.set_window_size(1200, 900)
        human_sleep(0.1, 1.1)

        # INJETAR stealth JS (adiciona propriedades antes de carregar páginas novas)
        driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            "source": """
                // hide webdriver flag
                Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
                // common properties used by fingerprint scripts
                window.chrome = window.chrome || { runtime: {} };
                Object.defineProperty(navigator, 'plugins', { get: () => [1, 2, 3, 4, 5] });
                Object.defineProperty(navigator, 'languages', { get: () => ['pt-BR','pt','en-US','en'] });
            """
        })

        # opcional: override UA também via CDP para requests futuras
        driver.execute_cdp_cmd("Network.setUserAgentOverride", {
            "userAgent": USER_AGENT
        })

        # driver.get(START_URL)
        # human_sleep(1.0, 2.0)

        # EXEMPLO: esperar e pegar título
        print("Título inicial:", driver.title)

        save_cookies(driver, COOKIES_FILE)

        # fim do fluxo de teste
        human_sleep(1.0, 2.0)
        print("Fluxo finalizado sem exceções aparentes")

    except Exception as e:
        print("Erro durante execução:", e)

#                  FUNÇÃO 0 - Load DataFrame de clientes
def load_df_clientes():
    global df_atual
    # 1 - Ler csv 'base_clientes_fake.csv' e criar um dataframe na pasta do projeto
    df_clientes = pd.read_csv('base_clientes_fake.csv', sep=';')
    #df_clientes.head(2)  # Exibir as primeiras linhas do dataframe para verificação


    # 1 - Filtrar Dataframe (SERA UM LOOP DEPOIS)
    df_atual = df_clientes[df_clientes['cpf'] == 37672119800].copy()
    #pegar primeira linha do dataframe
    df_atual = df_atual.iloc[0]



#=======================================================================================================================
#                  FUNÇÃO 1 - Iniciar inserindo dados do cliente (CPF, data nascimento, tipo do produto)
#=======================================================================================================================
#Inicio - Inserir CPF / data nascimento / tipo do produto e deixar para usuario inserir o reCaptcha
def inserir_dados_cliente():
    global driver, wait, action, df_atual

    
    # Verifica se driver está definido
    if 'driver' not in globals():
        raise RuntimeError("driver não está definido. Execute a célula que inicializa o driver antes de rodar esta função.")

    # Verifica se df_atual está definido
    if 'df_atual' not in globals():
        raise RuntimeError("df_atual não está definido. Execute a célula que define df_atual antes de rodar esta função.")

    # Verifica se df_atual está vazio ou None
    if df_atual is None:
        raise RuntimeError("df_atual está vazio ou None. Execute a célula que define df_atual antes de rodar esta função.")

    action = ActionChains(driver)
    # 0 Incluir CPF
    cpf_atual = str(df_atual['cpf']).strip()
    print(f"CPF do cliente atual: {cpf_atual}")
    input_cpf = driver.find_element(By.CSS_SELECTOR, 'input[formcontrolname="cpfCnpj"]')
    #input_cpf.clear()
    action.move_to_element(input_cpf).pause(random.uniform(0.2, 0.7)).click(input_cpf).perform()
    human_type(input_cpf, cpf_atual)
    human_sleep()

    # 0 Data de nascimento
    data_nascimento_atual = str(df_atual['data_nas']).strip()
    print(f"Data de nascimento do cliente atual: {data_nascimento_atual}")
    input_data_nascimento = driver.find_element(By.CSS_SELECTOR, 'input[formcontrolname="dtaNascimentoFundacao"]')
    #input_data_nascimento.clear()
    action.move_to_element(input_data_nascimento).pause(random.uniform(0.01, 0.3)).click(input_data_nascimento).perform()
    human_type(input_data_nascimento, data_nascimento_atual)
    human_sleep()



    # 0 Tipo de Produto:
    tipo_produto_atual = str(df_atual['tp_produto']).strip().lower()
    print(f"Tipo de produto do cliente atual: {tipo_produto_atual}")
    botao_tipo_produto = driver.find_element(By.ID, 'codigoProduto')
    action.move_to_element(botao_tipo_produto).pause(random.uniform(0.01, 0.3)).click(botao_tipo_produto).perform()
    human_sleep()


    #seleciona o tipo de produto

    #Elemento: <ids-option _ngcontent-cmg-c48="" _nghost-cmg-c11="" id="ids-option-0" role="option" tabindex="0" aria-selected="true" aria-disabled="false" title="imóveis" class="ids-option ids-option--selected ng-star-inserted"><span _ngcontent-cmg-c11="" class="ids-option__text">imóveis </span></ids-option>
    if tipo_produto_atual == 'imoveis':
        #seleciona imoveis
        opcao_imoveis = driver.find_element(By.XPATH, '//ids-option[@title="imóveis"]')
        action.move_to_element(opcao_imoveis).pause(random.uniform(0.01, 0.3)).click(opcao_imoveis).perform()
        human_sleep()

    elif tipo_produto_atual == 'veiculos leves':
        #seleciona veiculos
        opcao_veiculos = driver.find_element(By.XPATH, '//ids-option[@title="veículos leves"]')
        action.move_to_element(opcao_veiculos).pause(random.uniform(0.01, 0.3)).click(opcao_veiculos).perform()
        human_sleep()

    elif tipo_produto_atual == 'motocicletas':
        #seleciona motocicletas
        opcao_motocicletas = driver.find_element(By.XPATH, '//ids-option[@title="motocicletas"]')
        action.move_to_element(opcao_motocicletas).pause(random.uniform(0.01, 0.3)).click(opcao_motocicletas).perform()
        human_sleep()
    elif tipo_produto_atual == 'veiculos pesados':
        #seleciona veiculos pesados
        opcao_veiculos_pesados = driver.find_element(By.XPATH, '//ids-option[@title="veículos pesados"]')
        action.move_to_element(opcao_veiculos_pesados).pause(random.uniform(0.01, 0.3)).click(opcao_veiculos_pesados).perform()
        human_sleep()

    else:
        print("Tipo de produto não reconhecido no dataframe.")


#=======================================================================================================================
#                  FUNÇÃO 2 - PRINCIPAL - Buscar consórcio do cliente e selecionar a melhor opção
#=======================================================================================================================

def buscar_consorcio_cliente():
    global grupo_encontrado
    global driver
    
    list_grupos = ['050127', '50130', '020257', '020269', '020267']  # Exemplo de lista de grupos para ignorar <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    # Variável para controlar se um grupo foi encontrado
    grupo_encontrado = False

    # buscar valor maximo para o cliente
    div_valor_maximo = driver.find_element(By.CLASS_NAME, 'valores-min-max')
    valor_maximo = div_valor_maximo.find_element(By.TAG_NAME, 'h5').text
    match = re.search(r'R\$[\s]*([\d\.,]+)', valor_maximo)
    if match:
        valor_maximo_formatado = match.group(1).replace('.', '').replace(',', '.')
        valor_maximo_float = float(valor_maximo_formatado)
        print(f"Valor máximo extraído: R$ {valor_maximo_float:.2f}")


    ### Selecionar Tabela e interagir com dropdowns - Grupos ### ### ###
    tabela = driver.find_element(By.XPATH, '//*[@aria-describedby="tabelaGrupos"]') # localizar a tabela
    linhas_tabela = tabela.find_elements(By.XPATH, './/tbody/tr') # localizar todas as linhas da tabela, exceto o cabeçalho


    #===========================================
    ##########  1 - Busca de GRUPO #############
    #===========================================

    # Percorrer as linhas da tabela
    for linha in linhas_tabela:
        colunas = linha.find_elements(By.TAG_NAME, 'td')
        botao_grupo = colunas[0].find_element(By.TAG_NAME, 'button')
        numero_grupo = botao_grupo.text.strip()
        print(f"Número do grupo: {numero_grupo}")
        
        #ignorar lista de grupos
        if numero_grupo in list_grupos:
            print(f"Grupo {numero_grupo} está na lista de grupos para ignorar. Pulando...")
            continue
        
        #if numero_grupo == '020257':                            # OLD <---- Aqui entrará a lógica para selecionar o grupo desejado - GRUPO NOVO
        
        action.move_to_element(botao_grupo).pause(random.uniform(0.2, 0.7)).click(botao_grupo).perform()
        human_sleep(1.2, 2.2)


        #### > Clicar em exibir Créditos Disponíveis
        WebDriverWait(driver, 12).until(EC.presence_of_element_located((By.XPATH, '//span[contains(text(), " exibir créditos disponíveis ")]')))
        botao_exibir_creditos = driver.find_element(By.XPATH, '//span[contains(text(), " exibir créditos disponíveis ")]')
        action.move_to_element(botao_exibir_creditos).pause(random.uniform(0.2, 0.7)).click(botao_exibir_creditos).perform()
        human_sleep(1.2, 2.2)


        ### >>> TELA DE CRÉDITOS <<<###


        # Esperar a tabela de créditos ser exibida
        WebDriverWait(driver, 12).until(EC.presence_of_element_located((By.XPATH, "//p[normalize-space()='créditos disponíveis']/following-sibling::div/table")))
        tabela_creditos = driver.find_element(By.XPATH, "//p[normalize-space()='créditos disponíveis']/following-sibling::div/table")
        linhas_creditos = tabela_creditos.find_elements(By.XPATH, './/tbody/tr')

        # Variáveis para armazenar a melhor opção encontrada
        melhor_opcao_encontrada = None  
        maior_credito_encontrado = 0.0  
        codigo_bem_selecionado = None  # Variável para armazenar o código do bem selecionado - PARA CLICAR DEPOIS
        
        print("--- Iniciando análise das linhas de crédito ---")
        print(f"Total de linhas de crédito encontradas: {len(linhas_creditos)}")


        #=====================================================
        ### ### ### Buscar CREDITOS - Melhor Opção ### ### ###
        #=====================================================

        # Loop para analisar cada linha da tabela de créditos
        for linha in linhas_creditos:
            colunas = linha.find_elements(By.TAG_NAME, 'td')
            codigo_bem = colunas[0].text.strip()
            nome_bem = colunas[1].text.strip()
            
            
            #taxa_adm = colunas[2].text.strip()
            valor_credito = colunas[3].text.strip()
            valor_parcela = colunas[4].text.strip()

            # Converter valor_parcela para float antes de comparar
            valor_credito_float = float(valor_credito.replace('.', '').replace(',', '.'))
            valor_parcela_float = float(valor_parcela.replace('.', '').replace(',', '.'))

            print(f"Cód: {codigo_bem}, Nome: {nome_bem}, Vlr Credito: {valor_credito}, Parcela: {valor_parcela}")

            ### 1 - Verifica se o valor da parcela está dentro do valor máximo permitido
            if valor_parcela_float <= valor_maximo_float:
                print("Valor da parcela está dentro do valor máximo permitido.")

                ### 2 - Verifica se o valor do crédito é maior que o maior já encontrado
                if valor_credito_float > maior_credito_encontrado:
                    print(f"Nova melhor opção encontrada: Crédito R$ {valor_credito_float} com Parcela R$ {valor_parcela_float}")
                    maior_credito_encontrado = valor_credito_float
                    codigo_bem_selecionado = codigo_bem
                    print(f"Código do bem selecionado: {codigo_bem_selecionado}")


                    melhor_opcao_encontrada = {
                        'codigo_bem': codigo_bem,
                        'nome_bem': nome_bem,
                        'valor_credito': valor_credito,
                        'valor_parcela': valor_parcela
                    }
            print("--------------------------------------------------")



        print("\n--- Análise Concluída ---")

        if melhor_opcao_encontrada:
            print("✅ A melhor opção de crédito selecionada foi:")
            print(f"Código do bem: {melhor_opcao_encontrada['codigo_bem']}")
            print(f"Nome do bem: {melhor_opcao_encontrada['nome_bem']}")
            print(f"Valor do crédito: {melhor_opcao_encontrada['valor_credito']}")
            print(f"Valor da parcela: {melhor_opcao_encontrada['valor_parcela']} (Dentro do limite de R$ {valor_maximo_float})")
            
            grupo_encontrado = True
            print("==================================================")

            # Loop para encontrar a linha correspondente e clicar
            for linha in linhas_creditos:
                colunas = linha.find_elements(By.TAG_NAME, 'td')
                codigo_bem_na_linha = colunas[0].text.strip()

                # Compara com o código da melhor opção que você já encontrou
                if codigo_bem_na_linha == codigo_bem_selecionado:
                    print(f"Encontrada a linha correspondente ao código {codigo_bem_selecionado}.")
        
                    elemento_clicavel = colunas[0].find_element(By.TAG_NAME, 'u')
                    action.move_to_element(elemento_clicavel).pause(random.uniform(0.2, 0.7)).click().perform()
                    
                    print(f"Elemento do código {codigo_bem_selecionado} clicado com sucesso.")
                    human_sleep(1.5, 2.5)
                    
                    # Retirar Seguro
                    try:
                        time.sleep(1)  # Espera inicial para garantir que a página carregou
                        xpath_seguro = '//input[@formcontrolname="checkSeguro"]'
                        wait = WebDriverWait(driver, 10)
                        botao_seguro = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_seguro)))

                        # Estado atual do botão lendo o atributo 'aria-pressed'
                        estado_atual = botao_seguro.get_attribute('aria-pressed')
                        print(f"🔍 Estado atual do seguro: {estado_atual}")

                        # Se estiver "true" (ativo), clica para desativar
                        if estado_atual == 'true':
                            print("Seguro está ATIVADO. Tentando desativar...")
                            action.move_to_element(botao_seguro).pause(random.uniform(0.3, 0.8)).click().perform()
                            human_sleep(1.0, 1.5) # Dá um tempo para a página processar o clique

                            #  Confirma se o estado mudou
                            botao_seguro_apos_clique = driver.find_element(By.XPATH, xpath_seguro)
                            estado_final = botao_seguro_apos_clique.get_attribute('aria-pressed')
                            
                            if estado_final == 'false':
                                print("✔️ Sucesso! O seguro foi DESATIVADO.")
                            else:
                                print("⚠️ Atenção: O clique foi realizado, mas o estado do seguro não mudou para 'false'.")
                                # Tentar clicar novamente ou registrar para revisão manual
                                time.sleep(3)  # Pequena pausa antes de tentar novamente
                                action.move_to_element(botao_seguro).pause(random.uniform(0.3, 0.8)).click().perform()
                                human_sleep(1.0, 1.5) # Dá um tempo para a página processar o clique
                                # Verifica o estado novamente
                                if botao_seguro_apos_clique.get_attribute('aria-pressed') == 'false':   
                                    print("✔️ Sucesso na segunda tentativa! O seguro foi DESATIVADO.")
                                    human_sleep(1.0, 1.5) # Dá um tempo para a página processar o clique
                                else:
                                    print("❌ Falha: O estado do seguro ainda não é 'false' após duas tentativas. Necessário revisão manual.")
                    
                        # Se o estado já for "false" ou qualquer outra coisa, não faz nada
                        else:
                            print("✅ O seguro já está DESATIVADO. Nenhuma ação foi necessária.")

                    # 5. Tratamento de erros
                    except TimeoutException:
                        print("❌ Erro: Tempo esgotado. O botão de seguro não foi encontrado ou não se tornou clicável em 10 segundos.")
                    except Exception as e:
                        print(f"❌ Ocorreu um erro inesperado ao interagir com o botão de seguro: {e}")



                    ### >>> Clicar em CONTRATAR COTA
                    botao_contratar_cota = driver.find_element(By.XPATH, '//span[contains(text(), " contratar cota ")]')
                    action.move_to_element(botao_contratar_cota).pause(random.uniform(0.2, 0.7)).click(botao_contratar_cota).perform()
                    human_sleep(1, 2)
                    print("Clicado em CONTRATAR COTA, aguardando próxima tela...")


                    #===================================================
                    #===================================================
                    #===================================================
                    # Chamar a função para preencher os dados pessoais
                    #... continuar código Preencher os dados do cliente na próxima tela
                    
                    preencher_dados_pessoais()

                    #===================================================
                    #===================================================
                    #===================================================

                    break
                 
             # Fim do loop de busca por grupos
        else:
            print(f"❌ Nenhuma linha de crédito foi encontrada com parcela menor ou igual a R$ {valor_maximo_float}.")
        

        # Se grupo Não encontrado, clicar em voltar e tentar o próximo grupo
        if not grupo_encontrado:
            print(f"⚠️ Nenhuma opção válida encontrada no grupo {numero_grupo}. Voltando para a lista de grupos...")
            botao_voltar = driver.find_element(By.XPATH, '//p[contains(text(), " voltar para grupos")]')
            action.move_to_element(botao_voltar).pause(random.uniform(0.2, 0.7)).click(botao_voltar).perform()
            human_sleep(0.8, 1.4)

        else:
            print("✅ Grupo e crédito selecionados com sucesso. Saindo do loop de grupos.")
            break  # Sai do loop de grupos se um grupo válido foi encontrado      







     # fim loop de grupos



#=======================================================================================================================
#                                           ##### 2 - Preencher dados pessoais  #####
#=======================================================================================================================
# Finalizar preenchimento dos dados do cliente e finalizar a proposta
def preencher_dados_pessoais():
    global driver, action, df_atual, grupo_encontrado, wait

    print("Iniciando o preenchimento dos dados pessoais do cliente...")
    time.sleep(1)
    # >>> Acessa o Shadow host
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "mf-iparceiros-cadastrocliente")))
    host1 = driver.find_element(By.CSS_SELECTOR, "mf-iparceiros-cadastrocliente")
    shadow_root1 = driver.execute_script("return arguments[0].shadowRoot", host1)
    
    print("Shadow DOM acessado com sucesso.")
    ### 2.1 Genero 
    # Agora pega o ids-select do gênero
    time.sleep(0.5) # webdriverwait não funciona aqui - 
    botao_genero = shadow_root1.find_element(By.CSS_SELECTOR, 'ids-select[formcontrolname="sexo"]')
    combobox_genero = botao_genero.find_element(By.CSS_SELECTOR, "div[role='combobox']")
    action.move_to_element(combobox_genero).pause(random.uniform(0.01, 0.3)).click(combobox_genero).perform()
    human_sleep(0.1, 0.9)

    
    # Seleciona a opção de gênero:
    if str(df_atual['genero']).strip().lower() == 'feminino':
        #seleciona feminino
        opcao_feminino = driver.find_element(By.XPATH, '//span[@class="ids-option__text" and text()="feminino"]')
        action.move_to_element(opcao_feminino).pause(random.uniform(0.01, 0.3)).click(opcao_feminino).perform()
        human_sleep(0.1, 1.1)
    elif str(df_atual['genero']).strip().lower() == 'masculino':
        #seleciona masculino
        opcao_masculino = driver.find_element(By.XPATH, '//span[@class="ids-option__text" and text()="masculino"]')
        action.move_to_element(opcao_masculino).pause(random.uniform(0.01, 0.3)).click(opcao_masculino).perform()
        human_sleep(0.1, 1.1)
    else:
        print("Gênero não reconhecido no dataframe.")

    print("Gênero selecionado com sucesso.")

    ### 2.2 Nacionalidade
    # Agora pega o ids-select da nacionalidade
    botao_nacionalidade = shadow_root1.find_element(By.CSS_SELECTOR, 'ids-select[formcontrolname="nacionalidade"]')
    combobox_nacionalidade = botao_nacionalidade.find_element(By.CSS_SELECTOR, "div[role='combobox']")
    action.move_to_element(combobox_nacionalidade).pause(random.uniform(0.02, 0.3)).click(combobox_nacionalidade).perform()
    human_sleep()


    # lista de nacionalidades
    list_nacionalidades = driver.find_elements(By.XPATH, '//span[@class="ids-option__text"]')
    nacionalidade_cliente_atual = str(df_atual['nacionalidade']).strip().lower()
    print(f"Nacionalidade do cliente atual: {nacionalidade_cliente_atual}")

    for nacionalidade in list_nacionalidades:
        nacionalidade_texto = nacionalidade.text.strip().lower()
        print(nacionalidade_texto)
        
        if nacionalidade_texto == nacionalidade_cliente_atual:
            print(f"Nacionalidade '{nacionalidade_texto}' encontrada. Selecionando...")
            action.move_to_element(nacionalidade).pause(random.uniform(0.02, 0.3)).click(nacionalidade).perform()
            human_sleep()
            break
        
        
    ### 2.3 Estado Civil
    # Agora pega o ids-select do estado civil
    botao_estado_civil = shadow_root1.find_element(By.CSS_SELECTOR, 'ids-select[formcontrolname="estado_civil"]')
    combobox_estado_civil = botao_estado_civil.find_element(By.CSS_SELECTOR, "div[role='combobox']")
    action.move_to_element(combobox_estado_civil).pause(random.uniform(0.02, 0.3)).click(combobox_estado_civil).perform()
    #human_sleep()

    # lista de estados civis
    list_estados_civis = driver.find_elements(By.XPATH, '//span[@class="ids-option__text"]')
    estado_civil_cliente_atual = str(df_atual['estado_civil']).strip().lower()
    print(f"Estado Civil do cliente atual: {estado_civil_cliente_atual}")
    for estado_civil in list_estados_civis:
        estado_civil_texto = estado_civil.text.strip().lower()
        #eliminar (a) do texto
        estado_civil_texto = estado_civil_texto.replace('(a)', '').strip()
        print(estado_civil_texto)
        
        if estado_civil_texto == estado_civil_cliente_atual:
            print(f"Estado Civil '{estado_civil_texto}' encontrado. Selecionando...")
            action.move_to_element(estado_civil).pause(random.uniform(0.02, 0.3)).click(estado_civil).perform()
            human_sleep()
            break

    ### 2.4 Residência no Exterior
    reside_exterior = str(df_atual['residencia_exterior']).strip().lower()

    if reside_exterior == 'sim':
        botao_reside_exterior_Sim = shadow_root1.find_element(By.CSS_SELECTOR, 'input[value="S"][formcontrolname="residencia_exterior"]')
        if not botao_reside_exterior_Sim.is_selected():
            action.move_to_element(botao_reside_exterior_Sim).pause(random.uniform(0.02, 0.3)).click(botao_reside_exterior_Sim).perform()
            human_sleep()
            print("Residência no exterior marcada como 'Sim'.")
        else:
            print("Residência no exterior já está marcada como 'Sim'.")

    elif reside_exterior == 'nao' or reside_exterior == 'não':
        botao_reside_exterior_Nao = shadow_root1.find_element(By.CSS_SELECTOR, 'input[value="N"][formcontrolname="residencia_exterior"]')
        if not botao_reside_exterior_Nao.is_selected():
            action.move_to_element(botao_reside_exterior_Nao).pause(random.uniform(0.02, 0.3)).click(botao_reside_exterior_Nao).perform()
            human_sleep()
            print("Residência no exterior marcada como 'Não'.")
        else:
            print("Residência no exterior já está marcada como 'Não'.")
    else:
        print("Valor inválido para residência no exterior. Use 'Sim' ou 'Não'.")

    print("Residência no exterior selecionada com sucesso.")
    ### 2.5 Pessoa Politicamente Exposta
    pessoa_politicamente_exposta = str(df_atual['PEP']).strip().lower()

    if pessoa_politicamente_exposta == 'sim':
        botao_pessoa_politicamente_exposta_Sim = shadow_root1.find_element(By.CSS_SELECTOR, 'input[value="S"][formcontrolname="indicador_politicamente_exposto"]')
        if not botao_pessoa_politicamente_exposta_Sim.is_selected():
            action.move_to_element(botao_pessoa_politicamente_exposta_Sim).pause(random.uniform(0.02, 0.3)).click(botao_pessoa_politicamente_exposta_Sim).perform()
            human_sleep()
            print("Pessoa Politicamente Exposta marcada como 'Sim'.")
        else:
            print("Pessoa Politicamente Exposta já está marcada como 'Sim'.")

    elif pessoa_politicamente_exposta == 'nao' or pessoa_politicamente_exposta == 'não':
        botao_pessoa_politicamente_exposta_Nao = shadow_root1.find_element(By.CSS_SELECTOR, 'input[value="N"][formcontrolname="indicador_politicamente_exposto"]')
        if not botao_pessoa_politicamente_exposta_Nao.is_selected():
            action.move_to_element(botao_pessoa_politicamente_exposta_Nao).pause(random.uniform(0.02, 0.3)).click(botao_pessoa_politicamente_exposta_Nao).perform()
            human_sleep()
            print("Pessoa Politicamente Exposta marcada como 'Não'.")
        else:
            print("Pessoa Politicamente Exposta já está marcada como 'Não'.")
    else:
        print("Valor inválido para Pessoa Politicamente Exposta. Use 'Sim' ou 'Não'.")

    ### 2.6 Tipo de documento (rg / cnh / rne)
    tipo_documento = str(df_atual['tipo_documento']).strip().lower()

    botao_tipo_documento = shadow_root1.find_element(By.CSS_SELECTOR, 'ids-select[formcontrolname="tipo_documento"]')
    combobox_tipo_documento = botao_tipo_documento.find_element(By.CSS_SELECTOR, "div[role='combobox']")
    action.move_to_element(combobox_tipo_documento).pause(random.uniform(0.02, 0.3)).click(combobox_tipo_documento).perform()


    # Seleciona a opção de tipo de documento:
    if tipo_documento == 'rg':
        #seleciona rg
        opcao_rg = driver.find_element(By.XPATH, '//span[@class="ids-option__text" and text()="RG"]')
        action.move_to_element(opcao_rg).pause(random.uniform(0.02, 0.3)).click(opcao_rg).perform()
        human_sleep()
    elif tipo_documento == 'cnh':
        #seleciona cnh
        opcao_cnh = driver.find_element(By.XPATH, '//span[@class="ids-option__text" and text()="CNH"]')
        action.move_to_element(opcao_cnh).pause(random.uniform(0.02, 0.3)).click(opcao_cnh).perform()
        human_sleep()
    elif tipo_documento == 'rne':
        #seleciona rne
        opcao_rne = driver.find_element(By.XPATH, '//span[@class="ids-option__text" and text()="RNE"]')
        action.move_to_element(opcao_rne).pause(random.uniform(0.02, 0.3)).click(opcao_rne).perform()
        human_sleep()
    else:
        print("Tipo de documento não reconhecido no dataframe.")

    print(f"Tipo de documento selecionado com sucesso: {tipo_documento}")

    #2.7 Número do documento
    num_doc_atual = str(df_atual['numero_documento']).strip()
    print(f"Número do documento do cliente atual: {num_doc_atual}")
    input_numero_documento = shadow_root1.find_element(By.CSS_SELECTOR, 'input[formcontrolname="numero_documento"]')
    #input_numero_documento.clear()
    action.move_to_element(input_numero_documento).pause(random.uniform(0.02, 0.3)).click(input_numero_documento).perform()
    human_type(input_numero_documento, num_doc_atual)
    #human_sleep()


    #2.8 Órgão emissor
    orgao_emissor_atual = str(df_atual['orgao_expedidor']).strip()
    print(f"Órgão emissor do cliente atual: {orgao_emissor_atual}")
    input_orgao_emissor = shadow_root1.find_element(By.CSS_SELECTOR, 'input[formcontrolname="orgaoExpedidor"]')
    #input_orgao_emissor.clear()
    action.move_to_element(input_orgao_emissor).pause(random.uniform(0.02, 0.3)).click(input_orgao_emissor).perform()
    human_type(input_orgao_emissor, orgao_emissor_atual)



    #2.9 UF emissor
    uf_emissor_atual = str(df_atual['uf_expedidor']).strip().upper()
    print(f"UF emissor do cliente atual: {uf_emissor_atual}")
    botao_uf_emissor = shadow_root1.find_element(By.CSS_SELECTOR, 'input[formcontrolname="UFexpedidor"]')
    action.move_to_element(botao_uf_emissor).pause(random.uniform(0.02, 0.3)).click(botao_uf_emissor).perform()
    human_sleep()
    action.move_to_element(botao_uf_emissor).pause(random.uniform(0.02, 0.3)).send_keys(uf_emissor_atual).perform()
    human_sleep()


    # 2.10 Data de expedicao 
    # Obs.: Padrao dd/mm/aaaa
    data_expedicao_atual = str(df_atual['data_expedicao']).strip()
    print(f"Data de expedição do cliente atual: {data_expedicao_atual}")
    input_data_expedicao = shadow_root1.find_element(By.CSS_SELECTOR, 'input[formcontrolname="data_emissao_documento"]')
    #input_data_expedicao.clear()
    action.move_to_element(input_data_expedicao).pause(random.uniform(0.02, 0.3)).click(input_data_expedicao).perform()
    human_type(input_data_expedicao, data_expedicao_atual)
    human_sleep()

    #2.11 CEP
    cep_atual = str(df_atual['CEP']).strip()
    print(f"CEP do cliente atual: {cep_atual}")
    input_cep = shadow_root1.find_element(By.CSS_SELECTOR, 'input[formcontrolname="cep"]')
    #input_cep.clear()
    action.move_to_element(input_cep).pause(random.uniform(0.02, 0.3)).click(input_cep).perform()
    human_type(input_cep, cep_atual)
    human_sleep()

    #2.12 Número da residência
    numero_residencia_atual = str(df_atual['numero']).strip()
    print(f"Número da residência do cliente atual: {numero_residencia_atual}")
    input_numero_residencia = shadow_root1.find_element(By.CSS_SELECTOR, 'input[formcontrolname="numero"]')
    #input_numero_residencia.clear()
    action.move_to_element(input_numero_residencia).pause(random.uniform(0.02, 0.3)).click(input_numero_residencia).perform()
    human_type(input_numero_residencia, numero_residencia_atual)
    human_sleep()

    #2.13 Complemento
    complemento_atual = str(df_atual['complemento']).strip()
    #complemento_atual = ''
    #se complemento for NaN, vazio ou None, não preencher
    if complemento_atual.lower() in ['nan', 'none', '']:
        complemento_atual = ''
        print("Complemento está vazio. Pulando preenchimento.")
    else:
        print(f"Complemento do cliente atual: {complemento_atual}")
        input_complemento = shadow_root1.find_element(By.CSS_SELECTOR, 'input[formcontrolname="complemento"]')
        #input_complemento.clear()
        action.move_to_element(input_complemento).pause(random.uniform(0.02, 0.3)).click(input_complemento).perform()
        human_type(input_complemento, complemento_atual)
        human_sleep()


    #2.14 celular
    celular_atual = str(df_atual['celular']).strip()
    print(f"Celular do cliente atual: {celular_atual}")
    input_celular = shadow_root1.find_element(By.CSS_SELECTOR, 'input[formcontrolname="celular"]')
    #input_celular.clear()
    action.move_to_element(input_celular).pause(random.uniform(0.02, 0.3)).click(input_celular).perform()
    human_type(input_celular, celular_atual)
    #human_sleep()

    #2.15 email
    email_atual = str(df_atual['email']).strip()
    print(f"Email do cliente atual: {email_atual}")
    input_email = shadow_root1.find_element(By.CSS_SELECTOR, 'input[formcontrolname="email"]')
    #input_email.clear()
    action.move_to_element(input_email).pause(random.uniform(0.02, 0.3)).click(input_email).perform()
    human_type(input_email, email_atual)
    human_sleep()

    #2.16 Profissão
    profissao_atual = str(df_atual['profissao_cliente']).strip()
    print(f"Profissão do cliente atual: {profissao_atual}")
    input_profissao = shadow_root1.find_element(By.CSS_SELECTOR, 'input[formcontrolname="profissao"]')
    #input_profissao.clear()
    action.move_to_element(input_profissao).pause(random.uniform(0.02, 0.3)).click(input_profissao).perform()
    human_type(input_profissao, profissao_atual)
    #select primeiro item da lista de sugestão
    human_sleep()
    sugestoes_profissao = driver.find_elements(By.XPATH, '//ids-option[@class="ids-option ng-star-inserted"]')
    if sugestoes_profissao:
        action.move_to_element(sugestoes_profissao[0]).pause(random.uniform(0.02, 0.3)).click(sugestoes_profissao[0]).perform()
        human_sleep()

    print("Profissão preenchida com sucesso.")
    #2.17 Renda mensal
    renda_mensal_atual = str(df_atual['renda_mensal']).strip()
    renda_mensal_atual = renda_mensal_atual+'.00'  # Adiciona .00 ao final para formatar como valor monetário
    print(f"Renda mensal do cliente atual: {renda_mensal_atual}")
    input_renda_mensal = shadow_root1.find_element(By.CSS_SELECTOR, 'input[formcontrolname="valor_renda"]')
    #input_renda_mensal.clear()
    action.move_to_element(input_renda_mensal).pause(random.uniform(0.02, 0.3)).click(input_renda_mensal).perform()
    human_type(input_renda_mensal, renda_mensal_atual)
    human_sleep()

    #2.18 Patrimônio
    patrimonio_atual = str(df_atual['patrimonio']).strip()
    patrimonio_atual = patrimonio_atual+'.00'  # Adiciona .00 ao final para formatar como valor monetário
    print(f"Patrimônio do cliente atual: {patrimonio_atual}")
    input_patrimonio = shadow_root1.find_element(By.CSS_SELECTOR, 'input[formcontrolname="valor_patrimonio_total"]')
    #input_patrimonio.clear()
    action.move_to_element(input_patrimonio).pause(random.uniform(0.02, 0.3)).click(input_patrimonio).perform()
    human_type(input_patrimonio, patrimonio_atual)
    #human_sleep()

    print("Dados pessoais preenchidos com sucesso.")

    ### 2.19 Continuar para a próxima etapa ### (botão)
    botao_continuar = shadow_root1.find_elements(By.CSS_SELECTOR, 'button[idsmainbutton]')
    action.move_to_element(botao_continuar[0]).pause(random.uniform(0.02, 0.3)).click(botao_continuar[0]).perform()
    human_sleep()
    print("Clicado em 'Continuar', aguardando próxima tela...")


    #========================================================================================================
    # Pagamento de boleto - Etapa 3
    #========================================================================================================

    # select root do shadow DOM da próxima etapa

    time.sleep(4.5)

    # >>> Acessa o Shadow host
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "mf-iparceiros-contratacao")))
    host1 = driver.find_element(By.CSS_SELECTOR, "mf-iparceiros-contratacao")
    shadow_root2 = driver.execute_script("return arguments[0].shadowRoot", host1)

    #Forma de pagamento:

    #Boleto (Padrão)
    #WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[value="BB"][formcontrolname="forma_pagamento"]'))) # Melhoruar espera - Wait + root shadow
    time.sleep(0.5)
    botao_boleto = shadow_root2.find_element(By.CSS_SELECTOR, 'input[value="BB"][formcontrolname="forma_pagamento"]')
    if not botao_boleto.is_selected():
        action.move_to_element(botao_boleto).pause(random.uniform(0.2, 0.7)).click(botao_boleto).perform()
        human_sleep(0.1, 0.3)
        print("Forma de pagamento marcada como 'Boleto'.")

    #preencher dados da Conta Bancária para débito automático conforme o banco selecionado no dataframe
    banco_atual = str(df_atual['banco']).strip().lower()
    agencia_atual = str(df_atual['agencia']).strip().lower()
    conta_corrente_atual = str(df_atual['conta_corrente']).strip().lower()
    digito_conta_atual = str(df_atual['digito_conta']).strip().lower()
    print(f"Banco do cliente atual: {banco_atual}")
    print(f"Agência do cliente atual: {agencia_atual}")
    print(f"Conta corrente do cliente atual: {conta_corrente_atual}")
    print(f"Dígito da conta do cliente atual: {digito_conta_atual}")

    #Banco:
    botao_banco = shadow_root2.find_element(By.CSS_SELECTOR, 'input[formcontrolname="nome_banco_encerramento"]')
    action.move_to_element(botao_banco).pause(random.uniform(0.1, 0.5)).click(botao_banco).perform()
    human_sleep()
    action.move_to_element(botao_banco).pause(random.uniform(0.1, 0.5)).send_keys(banco_atual).perform()
    human_sleep()

    sugestoes_banco = driver.find_elements(By.XPATH, '//ids-option[@class="ids-option ng-star-inserted"]')
    if sugestoes_banco:
        action.move_to_element(sugestoes_banco[0]).pause(random.uniform(0.1, 0.5)).click(sugestoes_banco[0]).perform()
        human_sleep()



    #agencia
    input_agencia = shadow_root2.find_element(By.CSS_SELECTOR, 'input[formcontrolname="agencia_encerramento"]')
    #input_agencia.clear()
    action.move_to_element(input_agencia).pause(random.uniform(0.1, 0.5)).click(input_agencia).perform()
    human_type(input_agencia, agencia_atual)
    human_sleep()

    #conta corrente
    input_conta_corrente = shadow_root2.find_element(By.CSS_SELECTOR, 'input[formcontrolname="conta_encerramento"]')
    #input_conta_corrente.clear()
    action.move_to_element(input_conta_corrente).pause(random.uniform(0.1, 0.5)).click(input_conta_corrente).perform()
    human_type(input_conta_corrente, conta_corrente_atual)
    #human_sleep()



    #digito
    input_digito_conta = shadow_root2.find_element(By.CSS_SELECTOR, 'input[formcontrolname="digito_encerramento"]')
    #input_digito_conta.clear()
    action.move_to_element(input_digito_conta).pause(random.uniform(0.1, 0.5)).click(input_digito_conta).perform()
    human_type(input_digito_conta, digito_conta_atual)
    #human_sleep()

    # marcar a opção "Li e aceito os termos de uso e a política de privacidade"
    # 1 - Ciência prazo
    checkbox_cienciaprazo = shadow_root2.find_element(By.CSS_SELECTOR, 'input[formcontrolname="cienciaGarantiaPrazo"]')
    if not checkbox_cienciaprazo.is_selected():
        action.move_to_element(checkbox_cienciaprazo).pause(random.uniform(0.01, 0.05)).click(checkbox_cienciaprazo).perform()
        human_sleep()
        print("Ciência do prazo aceita.")


    # 2 regras de cancelamento
    checkbox_regras_cancelamento = shadow_root2.find_element(By.CSS_SELECTOR, 'input[formcontrolname="cienciaRegrasCancelamento"]')
    if not checkbox_regras_cancelamento.is_selected():
        action.move_to_element(checkbox_regras_cancelamento).pause(random.uniform(0.01, 0.05)).click(checkbox_regras_cancelamento).perform()
        human_sleep()
        print("Regras de cancelamento aceitas.")

    # 3 Informei o cliente:
    checkbox_informei_cliente = shadow_root2.find_element(By.CSS_SELECTOR, 'input[formcontrolname="cienciaRegrasCRP"]')
    if not checkbox_informei_cliente.is_selected():
        action.move_to_element(checkbox_informei_cliente).pause(random.uniform(0.01, 0.5)).click(checkbox_informei_cliente).perform()
        human_sleep()
        print("Informei o cliente aceito.")

    #==========================================
    ####         Botão Contratar           ####
    #==========================================
    #selecionar botao pelo texto "contratar"
    time.sleep(3) # webdriverwait não funciona aqui -
    contratar = shadow_root2.find_element(By.CLASS_NAME, 'btn-contratar')
    botao_contratar = contratar.find_elements(By.TAG_NAME, 'button')
    # action.move_to_element(botao_contratar[0]).pause(random.uniform(0.2, 0.7)).click(botao_contratar[0]).perform()
    # human_sleep(0.8, 1.4)
    #mover o mouse para o botão e esperar 2 segundos
    action.move_to_element(botao_contratar[0]).pause(2).perform()

    #fim - Contratar


    

#Iniciar o driver
iniciar_driver()


load_df_clientes()


inserir_dados_cliente()

buscar_consorcio_cliente()


#EXEMPLO
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#                                                               #   #   #   #   #       ---      Configuração da Interface Gráfica      ---    #   #   #   #   #  
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# ================= Cores personalizadas =================
cor_fundo_janela   = "#0A1D03"  # fundo da janela principal
cor_fundo_frame    = "#73BA5B"  # fundo dos blocos (LabelFrame)
cor_botao_fundo    = "#0A4928"  # fundo dos botões
cor_botao_texto    = "white"    # texto dos botões
cor_texto_label    = "black"    # texto fixo dentro dos frames
cor_fundo_entry    = "white"    # fundo dos campos de entrada


# ================= Janela principal =================
root = tk.Tk()
root.title("Automação Comunicação Legitimidade")  # Título da janela
root.configure(bg=cor_fundo_janela)

# Adicione esta linha! Defina uma largura e altura mínimas.
root.minsize(width=430, height=480) 

# Configure a coluna principal para se expandir com a janela
root.columnconfigure(0, weight=1)


# ================= Título da Aplicação =================
# 1. Crie o Label do título, apontando para a janela 'root'
titulo_label = tk.Label(
    root, 
    text=" Comunicação de Legitimidade",  # Texto do título
    font=("Helvetica", 12, "bold"),  # Define a Fonte: ("Nome", Tamanho, "Estilo")
    bg=cor_fundo_janela,             # Cor de fundo igual à da janela
    fg=cor_botao_texto               # Cor do texto (usei a mesma dos botões)
)

# 2. Posicione o título no topo com um espaçamento
titulo_label.pack(pady=(10, 10)) # Adiciona 10 pixels de espaço em cima e 20 embaixo





def carregar_ultimo_email():
    if os.path.exists("credenciais.json"):
        try:
            with open("credenciais.json", "r") as f:
                dados = json.load(f)
                email_entry.insert(0, dados.get("email", ""))
        except Exception as e:
            print(f"Erro ao carregar e-mail salvo: {e}")
            #messagebox.showerror("Erro ao Carregar E-mail", f"Não foi possível carregar o e-mail salvo: {e}")

def salvar_credenciais():
    global EMAIL_GLOBAL, SENHA_GLOBAL
    EMAIL_GLOBAL = email_entry.get()
    SENHA_GLOBAL = senha_entry.get()
    with open("credenciais.json", "w") as f:
        json.dump({"email": EMAIL_GLOBAL}, f)
    #messagebox.showinfo("Credenciais Salvas", "E-mail e senha armazenados com sucesso.")
    print("E-mail e senha armazenados com sucesso.")






# =================== Frame: Credenciais & Login ===================
cred_frame = tk.LabelFrame(root, text="Logar Admin e Salesforce", bg=cor_fundo_frame, fg=cor_texto_label, padx=10, pady=4)
cred_frame.pack(padx=15, pady=5, fill="x")

# Configurar colunas dentro do frame de credenciais
cred_frame.columnconfigure(1, weight=1) # Faz a coluna do Entry expandir

# Linha 0: Botão
btn_salvar_credenciais = tk.Button(cred_frame, text="Salvar Credenciais & Logar", command=salvar_credenciais_logar,
                                  bg=cor_botao_fundo, fg=cor_botao_texto)

# columnspan=2 faz o botão ocupar as duas colunas
btn_salvar_credenciais.grid(row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=(7, 4))


# Linha 1: E-mail
tk.Label(cred_frame, text="E-mail:", bg=cor_fundo_frame, fg=cor_texto_label).grid(row=1, column=0, sticky="w", padx=5, pady=2)
email_entry = tk.Entry(cred_frame, bg=cor_fundo_entry)
email_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=2) # sticky="ew" faz o campo expandir horizontalmente

# Linha 2: Senha
tk.Label(cred_frame, text="Senha:", bg=cor_fundo_frame, fg=cor_texto_label).grid(row=2, column=0, sticky="w", padx=5, pady=2)
senha_entry = tk.Entry(cred_frame, show="*", bg=cor_fundo_entry)
senha_entry.grid(row=2, column=1, sticky="ew", padx=5, pady=2)





carregar_ultimo_email()



#   =================== Buscar caso na Fila ===================
fila_frame = tk.LabelFrame(root, text="Buscar Caso para Analisar na Fila - Admin", bg=cor_fundo_frame, fg=cor_texto_label, padx=10, pady=2)
fila_frame.pack(padx=15, pady=2, fill="x")

# Configurar colunas dentro do frame de credenciais
fila_frame.columnconfigure(1, weight=1) # Faz a coluna do Entry expandir

botao_buscar = tk.Button(fila_frame, text="Buscar Caso", command=buscar_proxima_conta_fila,
          bg=cor_botao_fundo, fg=cor_botao_texto)
botao_buscar.grid(row=0, column=0, columnspan=2, pady=3, padx=5, sticky="ew")

fila_status_label = tk.Label(fila_frame, text="...", font=("TkDefaultFont", 9, "italic"),
                            bg=cor_fundo_frame, fg=cor_texto_label)
fila_status_label.grid(row=1, column=0, columnspan=2, pady=2)

# Id Conta
tk.Label(fila_frame, text="Id Conta:", bg=cor_fundo_frame, fg=cor_texto_label).grid(row=2, column=0, sticky="w", padx=5, pady=2)
id_conta_entry = tk.Entry(fila_frame, bg=cor_fundo_entry)
id_conta_entry.grid(row=2, column=1, sticky="ew", padx=20, pady=2)










# =================== Frame: Feedback ===================
feedback_frame = tk.LabelFrame(root, text="Feedback da Análise", bg=cor_fundo_frame, fg=cor_texto_label, padx=10, pady=5)
feedback_frame.pack(padx=15, pady=5, fill="x")

# Configura a coluna do Combobox para expandir
feedback_frame.columnconfigure(1, weight=1)

# Opções para o campo de lista
opcoes_feedback = ['Suspeito de fraude', 'Sim, Fraude']

# Rótulo (Label) para o campo
tk.Label(feedback_frame, text="Feedback:", bg=cor_fundo_frame, fg=cor_texto_label).grid(row=0, column=0, sticky="w", padx=5, pady=5)

# Criação do Combobox (o campo de lista)
feedback_combobox = ttk.Combobox(feedback_frame, values=opcoes_feedback, state="readonly")
feedback_combobox.grid(row=0, column=1, sticky="ew", padx=5, pady=5)

# Define um valor padrão (o primeiro da lista) para ser exibido
feedback_combobox.current(0)





# =================== Frame: Baixar PDF e Enviar Comunicações no SF ===================
zerada_frame = tk.LabelFrame(root, text="Baixar PDF e Enviar Comunicações no SF", bg=cor_fundo_frame, fg=cor_texto_label, padx=10, pady=2)
zerada_frame.pack(padx=15, pady=5, fill="x")

tk.Button(zerada_frame, text="Processar Replica Parcelado", command=confirmar_e_realizar_comunicacoes,
          bg=cor_botao_fundo, fg=cor_botao_texto).pack(pady=5, padx=5, fill="x")
zerada_status_label = tk.Label(zerada_frame, text="...", font=("TkDefaultFont", 9, "italic"),
                               bg=cor_fundo_frame, fg=cor_texto_label)
zerada_status_label.pack(pady=5)






# =================== Inicia GUI =================== #

root.mainloop()



# Fecha o navegador quando a aplicação é encerrada
if driver:
    driver.quit()





#...
#===================================================================================================================================================#
### ### ###                          """"                               Fim do Código                          """                        ### ### ###
#===================================================================================================================================================#
