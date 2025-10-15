
# Módulos da Biblioteca Padrão do Python
import os                   # Interage com o sistema operacional (manipulação de arquivos, variáveis de ambiente).
import sys                  # Fornece acesso a variáveis e funções específicas do interpretador (ex.: argumentos de linha de comando).
import time                 # Funções relacionadas ao tempo (pausas, medição de tempo de execução).
import csv                  # Implementa classes para ler e escrever dados tabulares no formato CSV.
import json                 # Codifica e decodifica dados no formato JSON.
import re                   # Suporte a Expressões Regulares (Regex) para busca e manipulação de strings.
import logging              # Configuração de logs para depuração, rastreamento e monitoramento.
from pathlib import Path    # Abstração orientada a objetos para o sistema de arquivos (manipulação de caminhos/diretórios).
from decimal import Decimal, ROUND_DOWN # Oferece aritmética de ponto flutuante com precisão configurável.
from datetime import datetime, timedelta, date # Classes para manipulação de datas, horas e intervalos de tempo.
from dataclasses import dataclass, fields # Ferramentas para criar classes simples que armazenam dados de forma concisa.
import math                 # Funções matemáticas (cálculos trigonométricos, logarítmicos, etc.).
import random               # Geração de números pseudo-aleatórios.
import traceback            # Ferramentas para lidar e exibir informações de rastreamento de erro (stack traces).
import threading            # Threading para execução de tarefas em background (usado pela GUI)

# ---

# Bibliotecas de Terceiros (Necessitam de instalação via pip, ex.: pip install pandas selenium numpy)
import pandas as pd         # Ferramenta essencial para análise e manipulação de dados estruturados (tabelas/DataFrames).
import numpy as np          # Suporte para arrays e matrizes de alta performance, fundamental para computação numérica.
import tkinter as tk        # Biblioteca padrão para criação de Interfaces Gráficas (GUIs).
from tkinter import ttk, scrolledtext, messagebox # Componentes avançados (widgets), área de texto com scroll e caixas de diálogo do Tkinter.
import undetected_chromedriver as uc # type: ignore # Um driver para Selenium que tenta evitar ser detectado por sites.


# ---

# Módulos e Componentes Específicos do Selenium WebDriver
from selenium import webdriver      # O módulo principal para controlar navegadores via Selenium.
from selenium.webdriver.chrome.options import Options # Permite configurar opções específicas do navegador Chrome.
from selenium.webdriver.common.by import By             # Define métodos para localizar elementos HTML (ex.: ID, XPATH, CSS_SELECTOR).
from selenium.webdriver.common.action_chains import ActionChains # Realiza uma cadeia de ações avançadas do usuário (mouse, teclado).
from selenium.webdriver.support.ui import WebDriverWait # Fornece um mecanismo para esperar por elementos em uma página antes de interagir.
from selenium.webdriver.support import expected_conditions as EC # Define condições pré-definidas para a espera (ex.: elemento visível, clicável).
from selenium.common.exceptions import TimeoutException # Exceção disparada quando uma espera (WebDriverWait) atinge o tempo limite.
from selenium.common.exceptions import NoSuchElementException # Exceção disparada quando um elemento não é encontrado na página.
from selenium.common.exceptions import InvalidSessionIdException, WebDriverException # Exceções gerais relacionadas a problemas de comunicação ou estado da sessão do driver.
from selenium.webdriver.common.keys import Keys
# ---
import atexit
import unicodedata

# # Configuração de Opções (Mantida como estava no original para contexto, mas movida para o final)
# options = Options()
# options.add_argument('--disable-backgrounding-occluded-windows')  # Impede que abas em 2º plano sejam pausadas
# options.add_argument('--no-sandbox')
# options.add_experimental_option("detach", True)  # Evita que a aba feche com o script

#=======================================================================================================================



# CONFIGURAÇÃO - ajuste conforme seu ambiente
#PROFILE_DIR = r"C:/selenium/chrome-profile"   # seu user-data-dir
PROFILE_DIR = str((Path.home() / ".selenium" / "chrome-profile").resolve())

USER_AGENT = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/141.0.7390.55 Safari/537.36")
COOKIES_FILE = Path("cookies_saved.json")
#START_URL = "https://canal360i.cloud.itau.com.br/login/iparceiros"   # verifique se abrirá direto no login tem problema...

#COOKIES_FILE = APP_DATA / "cookies_saved.json"    #Ainda nao testei

APP_DATA = Path.home() / ".selecao_cotas"
APP_DATA.mkdir(parents=True, exist_ok=True)
# Controle de reutilização de sessão do Chrome (defina True para tentar reutilizar o perfil)
# base_dir é usado como argumento --user-data-dir (pode ser o mesmo que PROFILE_DIR ou outro caminho)

# Variáveis globais
driver = None
wait = None
action = None
df_clientes = None
df_atual = None
cliente_atual = None
cpf_atual = None
grupo_encontrado = False
root = None
driver_started = False
driver_thread = None
_driver_keepalive_evt = threading.Event()



#=======================================================================================================================
# Funções auxiliares
#=======================================================================================================================

def driver_ativo(drv):
    try:
        _ = drv.current_url  # dispara se o driver já morreu
        return True
    except:
        return False
    
# utilitário: delays "humanos"
def human_sleep(a=0.05, b=0.5):
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

# utilitário: converte pd.Series em dict com tipos JSON-safe
def json_safe_dict_from_series(s: pd.Series) -> dict:
    def norm(v):
        if v is None or (hasattr(pd, "isna") and pd.isna(v)):
            return None
        if isinstance(v, (np.integer,)):
            return int(v)
        if isinstance(v, (np.floating,)):
            f = float(v)
            if math.isnan(f) or math.isinf(f):
                return None
            return f
        if isinstance(v, (pd.Timestamp, datetime, date)):
            return v.isoformat()
        if isinstance(v, bool):
            return bool(v)
        return str(v)
    return {k: norm(v) for k, v in s.to_dict().items()}

# utilitário: verifica se o driver está ativo (sessão válida)
def ensure_driver_alive(driver):
    try:
        # Qualquer comando simples serve como "ping"
        driver.execute_script("return 1")
        return True
    except (InvalidSessionIdException, WebDriverException):
        return False


def _shutdown():
    global driver
    try:
        if driver is not None:
            driver.quit()
    except: pass

if not getattr(sys, "frozen", False):  # só em modo dev
    atexit.register(_shutdown)

def on_close():
    try:
        _driver_keepalive_evt.set()   # libera o thread
    except: pass
    try:
        if driver: driver.quit()
    except: pass
    root.destroy()


def is_nan(x):
    try:
        return x is None or (isinstance(x, float) and math.isnan(x)) or (isinstance(x, np.floating) and np.isnan(x))
    except Exception:
        return False

def to_str_safe(x):
    """Converte qualquer coisa para string segura, sem quebrar em NaN/float64/etc."""
    if is_nan(x):
        return ""
    if isinstance(x, (np.integer, )):
        return str(int(x))
    if isinstance(x, (int, )):
        return str(x)
    if isinstance(x, (np.floating, float)):
        # cuidado: pode ser 21500.0 etc. Se quiser inteiro sem .0, formate aqui:
        if float(x).is_integer():
            return str(int(x))
        return str(x)
    if x is None:
        return ""
    return str(x).strip()

def only_digits(x):
    return re.sub(r'\D+', '', to_str_safe(x))

def normalize_cep(x):
    d = only_digits(x)
    if len(d) == 7:
        # caso clássico: perdeu o zero à esquerda → prefixa 0
        d = '0' + d
    if len(d) < 8:
        d = d.zfill(8)     # preserva zeros à esquerda
    elif len(d) > 8:
        d = d[-8:]         # pega os 8 finais se veio lixo junto
    return f"{d[:5]}-{d[5:]}" if len(d) == 8 else to_str_safe(x)

_DATE_PATS = ['%d/%m/%Y','%d-%m-%Y','%Y-%m-%d','%Y/%m/%d','%d%m%Y','%d.%m.%Y','%d %m %Y','%d/%m/%y']


def normalize_date_br(x):
    
    s = to_str_safe(x)
    if not s: return ""
    s2 = re.sub(r'[.\s-]', '/', s)  # uniformiza separador como '/'
    for pat in _DATE_PATS:
        p = pat.replace('-', '/').replace('.', '/').replace(' ', '/')
        try:
            dt = datetime.strptime(s2, p)
            return dt.strftime('%d/%m/%Y')
        except ValueError:
            continue
    return s  # se não reconhecer, devolve como veio

def normalize_inscricao_municipal(x):
    d = only_digits(x)
    if not d:
        return ""
    if len(d) < 11:
        d = d.zfill(11)     # completa com zeros à esquerda
    elif len(d) > 11:
        d = d[-11:]         # mantém os 11 finais se vier maior
    return d



def _js(driver, script, *args):
    return driver.execute_script(script, *args)

def ensure_uf_dropdown_closed_and_selected(driver, timeout=6):
    """
    Garante que UFexpedidor foi selecionado por clique e o dropdown está fechado.
    Retorna True se conseguiu fechar/selecionar; False caso contrário.
    """
    # 1) host do componente
    host = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "mf-iparceiros-cadastrocliente"))
    )

    # 2) localizar container (ids-select/ids-combobox/ids-input) da UF
    container = _js(driver, """
      const host = arguments[0];
      const root = host && host.shadowRoot;
      if(!root) return null;
      const el = root.querySelector('ids-select[formcontrolname="UFexpedidor"]')
             || root.querySelector('ids-combobox[formcontrolname="UFexpedidor"]')
             || (root.querySelector('input[formcontrolname="UFexpedidor"]')?.closest('ids-input, ids-fieldset, ids-combobox, ids-select')||null);
      return el || null;
    """, host)
    if not container:
        return False

    def open_dropdown():
        _js(driver, """
          const cont = arguments[0];
          const r = cont.shadowRoot || cont;
          const t = r.querySelector('[role="combobox"]')
                  || r.querySelector('button[aria-haspopup="listbox"]')
                  || r.querySelector('.ids-trigger')
                  || r.querySelector('.ids-select__trigger')
                  || cont;
          try { t.click(); } catch(e){}
          return true;
        """, container)

    def click_first_option():
        # tenta no shadowRoot e no documento (algumas libs rendem a lista fora)
        clicked = _js(driver, """
          const cont = arguments[0];
          const sr = cont.shadowRoot || cont;

          function visible(els){ return els.filter(e => e && e.offsetParent !== null); }

          let txts = visible(Array.from(sr.querySelectorAll('span.ids-option__text')));
          if(!txts.length) txts = visible(Array.from(document.querySelectorAll('span.ids-option__text')));

          let el = txts[0];
          if(!el){
            let opts = visible(Array.from(sr.querySelectorAll('ids-option[role="option"], .ids-option, [role="option"]')));
            if(!opts.length) opts = visible(Array.from(document.querySelectorAll('ids-option[role="option"], .ids-option, [role="option"]')));
            el = opts[0] || null;
          }
          if(!el) return false;

          const clickable = el.closest?.('ids-option, [role="option"], .ids-option') || el;
          try { clickable.click(); } catch(e){ return false; }
          return true;
        """, container)
        return bool(clicked)

    def finalize_close():
        _js(driver, """
          const cont = arguments[0];
          try{ cont.dispatchEvent(new Event('input',{bubbles:true})); }catch(e){}
          try{ cont.dispatchEvent(new Event('change',{bubbles:true})); }catch(e){}
          try{ cont.dispatchEvent(new CustomEvent('ids-change',{bubbles:true})); }catch(e){}

          const r = cont.shadowRoot || cont;
          const trg = r.querySelector('[role="combobox"]')
                   || r.querySelector('button[aria-haspopup="listbox"]')
                   || r.querySelector('.ids-trigger')
                   || r.querySelector('.ids-select__trigger')
                   || cont;

          try{
            const target = (trg.shadowRoot && trg.shadowRoot.activeElement) || trg;
            if(target && target.dispatchEvent){
              target.dispatchEvent(new KeyboardEvent('keydown',{key:'Escape',bubbles:true}));
              target.dispatchEvent(new KeyboardEvent('keyup',{key:'Escape',bubbles:true}));
            }
          }catch(e){}

          try{ (r.activeElement||document.activeElement)?.blur?.(); }catch(e){}
          try{ document.body && document.body.click(); }catch(e){}
        """, container)

    def is_closed():
        return _js(driver, """
          const cont = arguments[0];
          const r = cont.shadowRoot || cont;
          const combo = r.querySelector('[role="combobox"]') || r;
          const expanded = combo && combo.getAttribute('aria-expanded') === 'true';
          // tenta detectar listbox visível
          const visible = (root) => Array.from(root.querySelectorAll('[role="listbox"], .ids-listbox, ids-listbox'))
                                         .some(e => e && e.offsetParent !== null);
          const anyOpen = visible(document) || visible(r);
          return (!expanded && !anyOpen);
        """, container) is True

    # === sequência do guard ===
    open_dropdown()
    # pequena tolerância de render
    time.sleep(0.15)
    clicked = click_first_option()
    finalize_close()

    # Wait até fechar mesmo
    try:
        WebDriverWait(driver, timeout).until(lambda d: is_closed())
        return clicked or True
    except TimeoutException:
        # último recurso: ESC global + clique fora
        try:
            driver.switch_to.active_element.send_keys(Keys.ESCAPE)
        except Exception:
            pass
        _js(driver, "try{ document.body && document.body.click(); }catch(e){};")
        # tenta mais uma espera curta
        try:
            WebDriverWait(driver, 2).until(lambda d: is_closed())
            return clicked or True
        except TimeoutException:
            return False


def _executable_dir() -> Path:
    """
    Retorna a pasta do executável em todos os cenários:
    - Dev: pasta do arquivo .py
    - PyInstaller onefile/onedir: pasta do executável
    - .app no macOS: .../MeuApp.app/Contents/MacOS
    """
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent

def find_data_file(name: str) -> Path:
    """
    Procura o arquivo em lugares razoáveis:
    - CWD (onde o usuário executou)
    - pasta do executável
    - raiz do .app (subindo 3 níveis a partir de Contents/MacOS)
    - pasta do .py (modo dev)
    Levanta FileNotFoundError se não achar.
    """
    exe_dir = _executable_dir()
    candidates = [
        Path.cwd(),
        exe_dir,
        exe_dir.parent.parent.parent if ".app" in str(exe_dir) else None,  # raiz do .app
        Path(__file__).resolve().parent
    ]
    for base in filter(None, candidates):
        p = (base / name)
        if p.exists():
            return p
    raise FileNotFoundError(f"Arquivo '{name}' não encontrado. Procurei em:\n" +
                            "\n".join(str((c or Path()).resolve() / name) for c in candidates if c))

# - Load DataFrame de clientes
def load_df_clientes():
    global df_atual, df_clientes
    arq = find_data_file('base_clientes.xlsx')  # <— em vez de Path('base_clientes.xlsx')
    try:
        import openpyxl  # garante engine disponível no runtime
    except ImportError as e:
        raise RuntimeError(
            "Dependência ausente: 'openpyxl'. Instale e reempacote o app."
        ) from e

    df_clientes = pd.read_excel(arq, dtype={'cpf_cnpj': str}, engine='openpyxl')
    if 'status' not in df_clientes.columns:
        raise RuntimeError("A planilha não possui a coluna obrigatória 'status'.")

    df_atual = df_clientes[df_clientes['status'] == "Pendente"].copy()
    return df_clientes





# Atualizar status do cliente no xlsx
def atualizar_status_cliente(cpf_cliente, novo_status, caminho='base_clientes.xlsx'):
    arq = Path(caminho)
    if not arq.exists():
        print(f"❌ Erro: Arquivo '{arq.name}' não encontrado.")
        return
    try:
        # tenta ler Excel
        df = pd.read_excel(arq, dtype={'cpf_cnpj': str}, engine='openpyxl')
        is_excel = True
    except Exception as e:
        print("⚠️ openpyxl indisponível. Usando CSV como fallback.", e)
        # fallback: sincronize com CSV irmão
        csv_path = arq.with_suffix('.csv')
        df = pd.read_csv(csv_path, sep=';', dtype={'cpf_cnpj': str})
        is_excel = False

    alvo = re.sub(r'\D', '', str(cpf_cliente))
    col = df['cpf_cnpj'].astype(str).str.replace(r'\D', '', regex=True)

    if alvo in col.values:
        df.loc[col == alvo, 'status'] = novo_status
        try:
            if is_excel:
                df.to_excel(arq, index=False, engine='openpyxl')
            else:
                df.to_csv(arq.with_suffix('.csv'), sep=';', index=False)
        except Exception as e:
            print("❌ Falha ao salvar. Tentando CSV:", e)
            df.to_csv(arq.with_suffix('.csv'), sep=';', index=False)
        print(f"✅ Status de {cpf_cliente} atualizado para '{novo_status}'.")
    else:
        print(f"⚠️ CPF/CNPJ {cpf_cliente} não encontrado. Nada feito.")






#=======================================================================================================================
#                  FUNÇÃO 0 - Iniciar Driver com as configurações iniciais
#=======================================================================================================================
# Função INICIAL - iniciar driver com perfil, user-agent e stealth

# #Google Chrome Version 141 (testar sem versão 141)
def iniciar_driver():
    global driver, wait, action, driver_started

    if driver_started and driver is not None and driver_ativo(driver):
        print("Info", "O navegador já está iniciado.", parent=root)
        return
    options = uc.ChromeOptions()

    # perfil existente e UA (mantive seu uso)
    options.add_argument(f"--user-data-dir={PROFILE_DIR}")
    options.add_argument(f"--user-agent={USER_AGENT}")

    # flags úteis
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--lang=pt-BR")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-infobars")
    # options.add_argument("--start-maximized")  # opcional

    # caminho explícito do binário do Chrome no macOS
    options.binary_location = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"

    print("Iniciando o Chrome com undetected-chromedriver (forçando version_main=141)...")

    try:
        # força compatibilidade com sua versão do Chrome e usa subprocess no mac
        driver = uc.Chrome(
            options=options,
            version_main=141,    # <--- força para Chrome 141
            use_subprocess=True  # importante no macOS / Alterar para FALSE para TESTAR
        )
        driver_started = True

        wait = WebDriverWait(driver, 15)
        action = ActionChains(driver)
        print("Driver iniciado.")

        # ajuste de tamanho (mais "real")
        try:
            driver.set_window_size(1199, 889)
        except Exception:
            pass
        human_sleep(0.1, 1.1)

        # INJETAR stealth JS (mantive seu script)
        driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            "source": """
                Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
                window.chrome = window.chrome || { runtime: {} };
                Object.defineProperty(navigator, 'plugins', { get: () => [1, 2, 3, 4, 5] });
                Object.defineProperty(navigator, 'languages', { get: () => ['pt-BR','pt','en-US','en'] });
            """
        })

        print("Stealth JS injetado.")

        # override UA via CDP
        try:
            driver.execute_cdp_cmd("Network.setUserAgentOverride", {"userAgent": USER_AGENT})
        except Exception:
            pass

        # teste rápido
        driver.get("https://www.google.com")
        print("Título inicial:", driver.title)

        save_cookies(driver, COOKIES_FILE)
        human_sleep(1.0, 2.0)
        print("Fluxo finalizado sem exceções aparentes")
        #messagebox.showinfo("Info", "Driver iniciado com sucesso!")

        return driver

    except Exception as e:
        print("Erro ao criar/iniciar driver:")
        traceback.print_exc()
        # dica para debug: se quiser, expose logs:
        try:
            # tenta fechar com segurança
            driver.quit()
        except Exception:
            pass
        raise  # relança para o fluxo chamar saber que houve falha


#=======================================================================================================================
#                  FUNÇÃO EXTRA - Coletar e salvar todos os grupos disponíveis em CSV
#=======================================================================================================================
def _strip_accents(s: str) -> str:
    return ''.join(c for c in unicodedata.normalize('NFKD', s) if not unicodedata.combining(c))



#=======================================================================================================================
#                  FUNÇÃO 1 - Iniciar inserindo dados do cliente (CPF, data nascimento, tipo do produto)
#=======================================================================================================================
#Inicio - Inserir CPF / data nascimento / tipo do produto e deixar para usuario inserir o reCaptcha

def inserir_dados_cliente_js():
    global driver, df_atual, cliente_atual, cpf_atual

    if not ensure_driver_alive(driver):
        raise RuntimeError("WebDriver inválido. Recrie o driver.")
    load_df_clientes()
    if df_atual is None or len(df_atual) == 0:
        raise RuntimeError("df_atual vazio.")

    cliente_atual = df_atual.iloc[0]
    cpf_atual_str    = str(cliente_atual['cpf_cnpj']).strip()
    cpf_atual = re.sub(r'[.\-\/]', '', cpf_atual_str)  # só números
    tipo_cliente = str(cliente_atual['tipo_cliente']).strip().lower()   # "cpf" | "cnpj"
    data_nas     = str(cliente_atual['data_nasc_fundacao']).strip()
    tp_produto   = str(cliente_atual['tp_produto']).strip().lower()

    # ---------- ETAPA 1: clicar CPF/CNPJ (robusto no input) ----------
    want = "cpf" if "cpf" in tipo_cliente else "cnpj"
    radio_val = "F" if want == "cpf" else "J"      # geralmente F = Física(CPF), J = Jurídica(CNPJ)

    clicked = driver.execute_script(r"""
      const radioVal = arguments[0];

      // procura em document e no shadowRoot (se existir)
      const host = document.querySelector('mf-iparceiros-cadastrocliente');
      const roots = [document, host && host.shadowRoot].filter(Boolean);

      function tryClickIn(root){
        // seletores mais prováveis para o rádio
        const sel = [
          `input[name="tipoPessoa"][value="${radioVal}"]`,
          `input.ids-radio-button__input[name="tipoPessoa"][value="${radioVal}"]`,
          // fallback por id (caso único) – não confie, mas tenta:
          `#tipoPessoa[value="${radioVal}"]`
        ];
        let el = null;
        for (const s of sel){
          el = root.querySelector(s);
          if (el) break;
        }
        if (!el) return false;

        // garante visibilidade e clica via JS (não intercepta)
        try { el.scrollIntoView({block:'center', inline:'center'}); } catch(e){}
        el.click();
        // dispara eventos esperados pelo form
        el.dispatchEvent(new Event('input',  {bubbles:true}));
        el.dispatchEvent(new Event('change', {bubbles:true}));
        return true;
      }

      for (const r of roots){
        if (tryClickIn(r)) return true;
      }

      // último fallback: tenta label pelo texto (document + shadow)
      function clickLabelByText(root, txt){
        const labs = Array.from(root.querySelectorAll('label'));
        const lbl = labs.find(l => (l.textContent||'').toLowerCase().includes(txt));
        if (lbl){
          try { lbl.scrollIntoView({block:'center'}); } catch(e){}
          lbl.click();
          return true;
        }
        return false;
      }
      const key = (radioVal === 'F') ? 'cpf' : 'cnpj';
      if (host && host.shadowRoot && clickLabelByText(host.shadowRoot, key)) return true;
      if (clickLabelByText(document, key)) return true;

      return false;
    """, radio_val)

    if not clicked:
        raise RuntimeError("Não foi possível clicar no botão CPF/CNPJ.")

    # respiro para os campos aparecerem
    time.sleep(0.2)

    # poll curto até os inputs existirem (no shadow ou no document)
    deadline = time.time() + 5.0
    while time.time() < deadline:
        ok = driver.execute_script("""
          const host = document.querySelector('mf-iparceiros-cadastrocliente');
          const root = (host && host.shadowRoot) ? host.shadowRoot : document;
          const cpf  = root.querySelector('input[formcontrolname="cpfCnpj"]');
          const dt   = root.querySelector('input[formcontrolname="dtaNascimentoFundacao"]');
          const prod = root.querySelector('#codigoProduto, ids-select#codigoProduto, ids-select[formcontrolname="codigoProduto"]');
          return !!(cpf && dt && prod);
        """)
        if ok: break
        time.sleep(0.05)


    # ---------- ETAPA 2: preencher CPF, Data e Tipo de Produto ----------
    js_fill = r"""
    return(function fill(data){
      function getRoot(){
        const host = document.querySelector('mf-iparceiros-cadastrocliente');
        return (host && host.shadowRoot) ? host.shadowRoot : document;
      }
      const root = getRoot();

      function nativeSetValue(el, value){
        const proto =
          el instanceof HTMLInputElement ? HTMLInputElement.prototype :
          el instanceof HTMLTextAreaElement ? HTMLTextAreaElement.prototype :
          el.__proto__;
        const desc = proto && Object.getOwnPropertyDescriptor(proto, "value");
        if (desc && desc.set) desc.set.call(el, value); else el.value = value;
        el.dispatchEvent(new Event("input", {bubbles:true}));
        el.dispatchEvent(new Event("change", {bubbles:true}));
      }
      function setInput(sel, val, opts={}){
        const el = root.querySelector(sel) || document.querySelector(sel);
        if(!el) return {ok:false, sel};
        nativeSetValue(el, (val ?? "").toString());
        if (opts.blur) el.dispatchEvent(new Event("blur", {bubbles:true}));
        return {ok:true};
      }
      function setIdsSelectByText(selectSel, wanted){
        const normalize = s => (s||"").toString().trim().toLowerCase()
          .normalize('NFD').replace(/[\u0300-\u036f]/g,'');
        const want = normalize(wanted);
        const el = root.querySelector(selectSel) || document.querySelector(selectSel);
        if(!el) return {ok:false, reason:"select-not-found", selectSel};

        // tentativa programática
        try{
          if('value' in el) el.value = want;
          if('selectedValue' in el) el.selectedValue = want;
          el.dispatchEvent(new Event("input",{bubbles:true}));
          el.dispatchEvent(new Event("change",{bubbles:true}));
        }catch(_){}

        function openDd(){
          const r = el.shadowRoot || el;
          const cands=['div[role="combobox"]','button[aria-haspopup="listbox"]','.ids-trigger','.ids-select__trigger','ids-trigger','#codigoProduto'];
          for(const sel of cands){
            const t = r.querySelector(sel) || el.querySelector(sel);
            if(t){ t.click(); return t; }
          }
          el.click(); return el;
        }
        function options(){
          const spansDoc = Array.from(document.querySelectorAll('span.ids-option__text, ids-option[title]'));
          const spansLocal = Array.from((root||document).querySelectorAll('span.ids-option__text, ids-option[title]'));
          return spansLocal.concat(spansDoc);
        }
        function tryClick(){
          const all = options();
          // por texto do span
          let tgt = all.find(n => n.matches && n.matches('span.ids-option__text') && normalize(n.textContent)===want);
          // por atributo title do ids-option
          if(!tgt){
            const opts = all.filter(n => n.matches && n.matches('ids-option[title]'));
            const o = opts.find(n => normalize(n.getAttribute('title'))===want);
            if (o) tgt = o;
          }
          if (tgt){
            const clickEl = tgt.closest && tgt.closest('ids-option') ? tgt.closest('ids-option') : tgt;
            clickEl.click();
            el.dispatchEvent(new Event("input",{bubbles:true}));
            el.dispatchEvent(new Event("change",{bubbles:true}));
            try{ el.dispatchEvent(new CustomEvent("ids-change",{bubbles:true})); }catch(_){}
            try{ document.body && document.body.click(); }catch(_){}
            return true;
          }
          return false;
        }

        const trg = openDd();
        if (!tryClick()){
          void el.offsetHeight;
          if (!tryClick()){
            document.dispatchEvent(new KeyboardEvent('keydown',{key:'ArrowDown',bubbles:true}));
            document.dispatchEvent(new KeyboardEvent('keyup',{bubbles:true, key:'ArrowDown'}));
            document.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true}));
            document.dispatchEvent(new KeyboardEvent('keyup',{key:'Enter',bubbles:true}));
          }
        }
        return {ok:true};
      }

      if (data.cpf)      setInput('input[formcontrolname="cpfCnpj"]', data.cpf, {blur:true});
      if (data.data_nas) setInput('input[formcontrolname="dtaNascimentoFundacao"]', data.data_nas, {blur:true});

      if (data.tp_produto){
        const map = {
          "imoveis": "imóveis",
          "veiculos leves": "veículos leves",
          "motocicletas": "motocicletas",
          "veiculos pesados": "veículos pesados"
        };
        const want = map[data.tp_produto] || data.tp_produto;
        setIdsSelectByText('#codigoProduto, ids-select#codigoProduto, ids-select[formcontrolname="codigoProduto"]', want);
      }
      return {ok:true, step:"filled"};
    })(arguments[0]);
    """

    payload2 = {
        "cpf": cpf_atual,
        "data_nas": data_nas,
        "tp_produto": tp_produto
    }
    res = driver.execute_script(js_fill, payload2)
    print("Resultado inserir_dados_cliente_fast:", res)
    return res


def guardar_grupos_disponiveis():
    global driver, action, APP_DATA
    print("Iniciando a coleta dos grupos disponíveis...")

    # coleta
    grupos_coletados = []

    # Tipo de consórcio selecionado
    tipo_consorcio = driver.find_element(
        By.XPATH, '//div[@role="combobox" and @id="codigoProduto"]'
    ).get_attribute('innerText').strip().lower()

    # normaliza para nome de arquivo
    tipo_consorcio = _strip_accents(tipo_consorcio)                  # remove acentos
    tipo_consorcio = re.sub(r'[^\w\s-]', '', tipo_consorcio)         # remove tudo que não for letra/número/_
    tipo_consorcio = tipo_consorcio.strip().replace(' ', '_')         # troca espaços por _

    # Expandir para 50 linhas
    botao_linhas = driver.find_element(By.XPATH, '//div[@role="combobox" and @id="pageSizeId"]')
    action.move_to_element(botao_linhas).pause(random.uniform(0.1, 0.5)).click(botao_linhas).perform()
    human_sleep(0.5, 1.5)
    opcao_50 = driver.find_element(By.XPATH, '//span[@class="ids-option__text" and text()="50"]')
    action.move_to_element(opcao_50).pause(random.uniform(0.1, 0.5)).click(opcao_50).perform()
    human_sleep(1, 3)
    print("Tabela atualizada para mostrar 50 linhas.")

    while True:
        tabela = driver.find_element(By.XPATH, '//*[@aria-describedby="tabelaGrupos"]')
        linhas_tabela = tabela.find_elements(By.XPATH, './/tbody/tr')

        for linha in linhas_tabela:
            colunas = linha.find_elements(By.TAG_NAME, 'td')
            botao_grupo = colunas[0].find_element(By.TAG_NAME, 'button')
            numero_grupo = (botao_grupo.text or '').strip()
            if numero_grupo:
                grupos_coletados.append(str(numero_grupo))

        # próxima página?
        try:
            botao_proxima_pagina = driver.find_element(By.XPATH, '//button[@id="nextPageId"]')
            if botao_proxima_pagina.get_attribute("disabled"):
                print("Fim das páginas. Nenhuma próxima disponível.")
                break
            action.move_to_element(botao_proxima_pagina).pause(random.uniform(0.2, 0.6)).click(botao_proxima_pagina).perform()
            human_sleep(1.5, 3)
            print("Indo para a próxima página...")
        except Exception as e:
            print("Não foi possível ir para a próxima página:", e)
            break

    # remove duplicados desta sessão
    grupos_coletados = list(dict.fromkeys(grupos_coletados))  # preserva ordem e unicidade
    print(f"Coleta concluída. Total únicos nesta sessão: {len(grupos_coletados)}")
    print(grupos_coletados)

    # ===== salvamento sem duplicar =====
    APP_DATA.mkdir(parents=True, exist_ok=True)
    nome_arquivo = APP_DATA / f'grupos_{tipo_consorcio}.csv'

    existentes_set = set()
    if nome_arquivo.exists():
        try:
            df_existentes = pd.read_csv(nome_arquivo, sep=';', dtype={'grupo': str})
            # normaliza (tira espaços e NaN)
            df_existentes['grupo'] = df_existentes['grupo'].astype(str).str.strip()
            existentes_set = set(df_existentes['grupo'].dropna().tolist())
        except Exception as e:
            print(f"Aviso: não foi possível ler o CSV existente ({e}). Vou recriar do zero.")

    # filtra apenas novos
    novos = [g for g in grupos_coletados if g not in existentes_set]

    if nome_arquivo.exists() and len(novos) > 0:
        # append sem header
        pd.DataFrame(novos, columns=['grupo']).to_csv(nome_arquivo, sep=';', index=False, mode='a', header=False)
        print(f"Acrescentados {len(novos)} novos grupos (sem duplicar).")
    elif not nome_arquivo.exists():
        # cria arquivo com header
        pd.DataFrame(grupos_coletados, columns=['grupo']).to_csv(nome_arquivo, sep=';', index=False)
        print(f"Arquivo criado com {len(grupos_coletados)} grupos.")
    else:
        print("Nenhum grupo novo para acrescentar (todos já estavam salvos).")

    print(f"Grupos salvos em: {nome_arquivo}")





#=======================================================================================================================
#                  FUNÇÃO 2 - PRINCIPAL - Buscar consórcio do cliente e selecionar a melhor opção
#=======================================================================================================================


def buscar_consorcio_cliente():
    global grupo_encontrado, cpf_atual, cliente_atual
    global driver
    
    #list_grupos = ['050127', '50130', '020257', '020269', '020267']  # Exemplo de lista de grupos para ignorar <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    #carregar lista de grupos do arquivo CSV
    tipo_consorcio = str(cliente_atual['tp_produto']).strip().lower()
    nome_arquivo = f'grupos_{tipo_consorcio.replace(" ", "_")}.csv'
    df_grupos_ignorar = pd.read_csv(nome_arquivo, sep=';')
    list_grupos = df_grupos_ignorar['grupo'].tolist()
    print(f"Lista de grupos a ignorar carregada do arquivo {nome_arquivo}: {list_grupos}")
    print("Iniciando a busca pelo melhor consórcio para o cliente...")

    
    # Variável para controlar se um grupo foi encontrado
    grupo_encontrado = False

    #validar CPF/CNPJ preenchido com a variável global cpf_atual
    input_cpf = driver.find_element(By.CSS_SELECTOR, 'input[formcontrolname="cpfCnpj"]')
    valor_cpf_preenchido = input_cpf.get_attribute('value').strip()
    valor_cpf_preenchido = re.sub(r'[.\-/]', '', valor_cpf_preenchido)

    if valor_cpf_preenchido != cpf_atual:
        print(f"⚠️ CPF na tela ({valor_cpf_preenchido}) diferente do esperado ({cpf_atual})")
        raise RuntimeError("CPF na tela diferente do esperado. Verifique o preenchimento.")
    else:
        print(f"✅ CPF validado: {cpf_atual}")


    
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
        
        
        action.move_to_element(botao_grupo).pause(random.uniform(0.2, 0.7)).click(botao_grupo).perform()
        human_sleep(4, 5)


        #### > Clicar em exibir Créditos Disponíveis
        WebDriverWait(driver, 12).until(EC.presence_of_element_located((By.XPATH, '//span[contains(text(), " exibir créditos disponíveis ")]')))
        botao_exibir_creditos = driver.find_element(By.XPATH, '//span[contains(text(), " exibir créditos disponíveis ")]')
        action.move_to_element(botao_exibir_creditos).pause(random.uniform(0.2, 0.7)).click(botao_exibir_creditos).perform()
        human_sleep(2, 2.5)

        
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
                    human_sleep(2, 2.5)
                    
                    # Retirar Seguro (*Talvez mover para função separada ou trocar por um Loop de tentativas)

                    #Retira somente de CPF
                    #verificar se o tipo de documento do cliente é CPF antes de iniciar o loop
                    if str(cliente_atual.get('tipo_cliente') or '').strip().upper() == 'CPF':
                        #RETIRAR SEGURO - LOOP DE VERIFICAÇÃO
                        try:
                            botao_seguro = wait.until(EC.element_to_be_clickable((By.XPATH, '//input[@formcontrolname="checkSeguro"]')))
                            estado_atual = botao_seguro.get_attribute('aria-pressed')
                            print(f"🔍 Estado atual do seguro antes do loop de verificação: {estado_atual}")
                                
                            # loop para garantir que botao foi clicado e estado alterado
                            tentativas = 0
                            max_tentativas = 15

                            for tentativa in range(max_tentativas):
                                estado_atual = botao_seguro.get_attribute('aria-pressed')
                                print(f"🔍 Verificação {tentativa + 1}/{max_tentativas}: Estado atual do seguro: {estado_atual}")

                                if estado_atual == 'false':
                                    print("✅ O seguro está DESATIVADO. Prosseguindo com o processamento dos clientes...")
                                    break
                                else:
                                    print("⚠️ O seguro ainda está ATIVADO. Tentando desativar novamente...")
                                    # clicar no elemento atual
                                    action.move_to_element(botao_seguro).pause(random.uniform(0.01, 0.08)).click(botao_seguro).perform()
                                    human_sleep(0.5, 1.0) # Dá um tempo para a página processar o clique

                                    # tentar re-obter o elemento (em caso de re-render)
                                    try:
                                        botao_seguro = driver.find_element(By.XPATH, '//input[@formcontrolname="checkSeguro"]')
                                    except Exception:
                                        try:
                                            botao_seguro = WebDriverWait(driver, 5).until(
                                                EC.presence_of_element_located((By.XPATH, '//input[@formcontrolname="checkSeguro"]'))
                                            )
                                        except Exception:
                                            print("❌ Não foi possível localizar o botão de seguro após o clique.")
                                            break

                                    # última tentativa: reportar falha se ainda estiver ativo
                                    if tentativa == max_tentativas - 1:
                                        estado_atual = botao_seguro.get_attribute('aria-pressed')
                                        if estado_atual != 'false':
                                            print("❌ Falha: O estado do seguro ainda não é 'false' após várias tentativas. Necessário revisão manual.")
                        except TimeoutException:
                            print("❌ Erro: Tempo esgotado. O botão de seguro não foi encontrado ou não se tornou clicável em 10 segundos.")
                        except Exception as e:
                            print(f"❌ Ocorreu um erro inesperado ao interagir com o botão de seguro: {e}")
                    else:
                        human_sleep(1, 1.5)
    

                    ### >>> Clicar em CONTRATAR COTA
                    WebDriverWait(driver, 10).until(
                                                    EC.presence_of_element_located((By.XPATH, '//span[contains(text(), " contratar cota ")]'))
                                                  )
                    botao_contratar_cota = driver.find_element(By.XPATH, '//span[contains(text(), " contratar cota ")]')
                    action.move_to_element(botao_contratar_cota).pause(random.uniform(0.3, 0.7)).click(botao_contratar_cota).perform()
                    human_sleep(1.5, 2)
                    print("Clicado em CONTRATAR COTA, aguardando próxima tela...")


                    #===================================================
                    #===================================================
                    #===================================================
                    # Chamar a função para preencher os dados pessoais
                    #... continuar código Preencher os dados do cliente na próxima tela
                    
                    if str(cliente_atual.get('tipo_cliente') or '').strip().upper() == 'CPF':
                      sucesso_preenchimento = preencher_dados_pessoais()

                    else:
    
                      sucesso_preenchimento = preencher_dados_PJ()

                    #Checar se o preenchimento foi bem sucedido e atualizar o status no CSV
                    if sucesso_preenchimento:
                        print("✅ Dados pessoais preenchidos com sucesso.")
                        atualizar_status_cliente(cpf_atual, "Finalizado")
                    else:
                        
                        print("❌ Falha ao preencher os dados pessoais.")
                        atualizar_status_cliente(cpf_atual, "Erro")



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
#                                  ##### FUNÇÃO 3- Preencher dados pessoais  #####
#=======================================================================================================================

#3.1 - CPF
def preencher_cliente_s_profi_js(cliente_series):
    global driver

    if not ensure_driver_alive(driver):
        raise RuntimeError("WebDriver inválido. Recrie o driver.")

    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "mf-iparceiros-cadastrocliente"))
    )

    js = r"""
    return(function fill(data){
      function q(s, r=document){ return r.querySelector(s); }
      const host = q("mf-iparceiros-cadastrocliente");
      if(!host || !host.shadowRoot) return {ok:false, step:"shadowRoot"};
      const root = host.shadowRoot;

      function nativeSetValue(el, value){
        if(!el) return false;
        const proto =
          el instanceof HTMLInputElement ? HTMLInputElement.prototype :
          el instanceof HTMLTextAreaElement ? HTMLTextAreaElement.prototype :
          el.__proto__;
        const desc = proto && Object.getOwnPropertyDescriptor(proto, "value");
        if (desc && desc.set) desc.set.call(el, value); else el.value = value;
        el.dispatchEvent(new Event("input", {bubbles:true}));
        el.dispatchEvent(new Event("change", {bubbles:true}));
        return true;
      }
      function setInput(sel, val, opts={}){
        const el = q(sel, root);
        if(!el) return {ok:false, sel};
        nativeSetValue(el, (val ?? "").toString());
        if(opts.blur) el.dispatchEvent(new Event("blur", {bubbles:true}));
        return {ok:true, el};
      }

      // ---- helpers para combos ----
      function openTriggerIn(el){
        const r = el.shadowRoot || el;
        const cands=['[role="combobox"]','button[aria-haspopup="listbox"]','.ids-trigger','.ids-select__trigger','ids-trigger'];
        for(const sel of cands){
          const t = r.querySelector(sel) || el.querySelector(sel);
          if(t){ try{t.click();}catch(e){} return t; }
        }
        try{ el.click(); }catch(e){}
        return el;
      }
      function finalizeClose(el, trigger){
        try{ el.dispatchEvent(new Event("input",{bubbles:true})); }catch(e){}
        try{ el.dispatchEvent(new Event("change",{bubbles:true})); }catch(e){}
        try{ el.dispatchEvent(new CustomEvent("ids-change",{bubbles:true})); }catch(e){}
        try{
          const trg = trigger || el;
          const target = (trg.shadowRoot && trg.shadowRoot.activeElement) || trg;
          if (target && target.dispatchEvent){
            target.dispatchEvent(new KeyboardEvent('keydown',{key:'Escape',bubbles:true}));
            target.dispatchEvent(new KeyboardEvent('keyup',{key:'Escape',bubbles:true}));
          }
        }catch(e){}
        try{ (el.shadowRoot||el).activeElement && (el.shadowRoot||el).activeElement.blur(); }catch(e){}
        try{ document.body && document.body.click(); }catch(e){}
      }

      // Clicar SEMPRE na 1ª opção visível do combobox por formcontrolname
      function selectFirstOptionByFcAlways(fc){
        // tenta localizar container do campo
        const container =
          q(`ids-select[formcontrolname="${fc}"]`, root) ||
          q(`ids-combobox[formcontrolname="${fc}"]`, root) ||
          (q(`input[formcontrolname="${fc}"]`, root)?.closest('ids-input, ids-fieldset, ids-combobox, ids-select')||null);

        if(!container) return {ok:false, reason:"container-not-found", fc};

        const trigger = openTriggerIn(container);

        // coleta opções locais e globais (lista pode estar fora do shadowRoot)
        function getVisible(elts){
          return elts.filter(e => e && e.offsetParent !== null);
        }
        function allOptions(){
          const sr = container.shadowRoot || container;
          const localTxt = Array.from(sr.querySelectorAll('span.ids-option__text'));
          const localOpt = Array.from(sr.querySelectorAll('ids-option[role="option"], .ids-option, [role="option"]'));
          const docTxt   = Array.from(document.querySelectorAll('span.ids-option__text'));
          const docOpt   = Array.from(document.querySelectorAll('ids-option[role="option"], .ids-option, [role="option"]'));
          // prioriza opções com texto (muitas libs encapsulam o click no span)
          return getVisible(localTxt).concat(getVisible(docTxt), getVisible(localOpt), getVisible(docOpt));
        }

        let opts = allOptions();
        if(!opts.length){
          // força reflow e tenta novamente (render preguiçosa)
          void container.offsetHeight;
          opts = allOptions();
        }

        const first = opts[0];
        if(first){
          // se veio um span de texto, sobe para o option clicável
          const clickable = first.closest?.('ids-option, [role="option"], .ids-option') || first;
          try{ clickable.click(); }catch(e){ return {ok:false, reason:"click-failed", fc}; }
          finalizeClose(container, trigger);
          return {ok:true, clicked:true, fc};
        }

        // fallback: fecha mesmo sem opção (para não ficar aberto)
        finalizeClose(container, trigger);
        return {ok:false, reason:"no-visible-options", fc};
      }

      // ===== setIdsSelect (inalterado) =====
      function setIdsSelect(fc, txt){
        const el = q(`ids-select[formcontrolname="${fc}"]`, root);
        if(!el) return {ok:false, comp:fc, reason:"ids-select-not-found"};

        const normalize = s => (s||"").toString().toLowerCase().trim()
          .replace(/\([^)]*\)/g, '')
          .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
          .replace(/\s+/g, ' ');

        const want = normalize(txt);
        if(!want) return {ok:true};

        try{
          if('value' in el) el.value = want;
          if('selectedValue' in el) el.selectedValue = want;
          el.dispatchEvent(new Event("input",{bubbles:true}));
          el.dispatchEvent(new Event("change",{bubbles:true}));
          try{ el.dispatchEvent(new CustomEvent("ids-change",{bubbles:true, detail:{value:want}})); }catch(_){}
        }catch(_){}

        const trigger = openTriggerIn(el);

        function getOptions(){
          const spansDoc = Array.from(document.querySelectorAll('span.ids-option__text'));
          const spansLocal = Array.from((root||document).querySelectorAll('span.ids-option__text'));
          return spansLocal.concat(spansDoc);
        }
        const normalizeTxt = s => (s||"").toString().toLowerCase().trim()
          .replace(/\([^)]*\)/g,'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,' ');
        function clickByText(){
          const all = getOptions();
          let tgtSpan = all.find(s => normalizeTxt(s.textContent) === want) ||
                        all.find(s => normalizeTxt(s.textContent).includes(want));
          if (tgtSpan){
            const opt = tgtSpan.closest('ids-option') || tgtSpan;
            opt.click();
            finalizeClose(el, trigger);
            return true;
          }
          return false;
        }

        if (!clickByText()){
          void el.offsetHeight;
          if (!clickByText()){
            try{
              const r = el.shadowRoot || el;
              (r.querySelector('[role="combobox"]') || r.querySelector('button[aria-haspopup="listbox"]') || el).focus();
            }catch(_){}
            let tries = 10, picked=false;
            while(tries-- > 0 && !picked){
              document.dispatchEvent(new KeyboardEvent('keydown',{key:'ArrowDown',bubbles:true}));
              document.dispatchEvent(new KeyboardEvent('keyup',{key:'ArrowDown',bubbles:true}));
              if (clickByText()){ picked=true; break; }
            }
            if(!picked){
              document.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true}));
              document.dispatchEvent(new KeyboardEvent('keyup',{key:'Enter',bubbles:true}));
              finalizeClose(el, trigger);
            }
          }
        }
        return {ok:true, clicked:true};
      }
      // ===== fim setIdsSelect =====

      const lower=s=>(s??"").toString().trim().toLowerCase();
      const upper=s=>(s??"").toString().trim().toUpperCase();

      if (data.genero) setIdsSelect("sexo", data.genero);
      if (data.nacionalidade){
        const nac = (data.nacionalidade||"").toString().replace(/\([^)]*\)/g,'').trim();
        setIdsSelect("nacionalidade", nac);
      }
      if (data.estado_civil){
        const ec=(data.estado_civil||"").toString().replace("(a)","").trim();
        setIdsSelect("estado_civil", ec);
      }
      if (data.residencia_exterior){
        const vn=["sim","s","yes","y"].includes(lower(data.residencia_exterior))?"S":"N";
        const el=q(`input[formcontrolname="residencia_exterior"][value="${vn}"]`, root);
        if(el && !el.checked){ el.click(); el.dispatchEvent(new Event("change",{bubbles:true}));}
      }
      if (data.PEP){
        const vn=["sim","s","yes","y"].includes(lower(data.PEP))?"S":"N";
        const el=q(`input[formcontrolname="indicador_politicamente_exposto"][value="${vn}"]`, root);
        if(el && !el.checked){ el.click(); el.dispatchEvent(new Event("change",{bubbles:true}));}
      }
      if (data.tipo_documento) setIdsSelect("tipo_documento", data.tipo_documento);
      if (data.numero_documento) setInput('input[formcontrolname="numero_documento"]', data.numero_documento);
      if (data.orgao_expedidor) setInput('input[formcontrolname="orgaoExpedidor"]', data.orgao_expedidor);

      // >>> AQUI: UFexpedidor por clique na 1ª opção (sempre verdadeira)
      if (data.uf_expedidor){
        // tenta selecionar por dropdown (1ª opção visível)
        const rUF = selectFirstOptionByFcAlways("UFexpedidor");
        if(!rUF.ok){
          // fallback: se for apenas input simples, ainda preenche e aplica blur
          setInput('input[formcontrolname="UFexpedidor"]', upper(data.uf_expedidor), {blur:true});
        }
      }

      if (data.data_expedicao) setInput('input[formcontrolname="data_emissao_documento"]', data.data_expedicao);
      if (data.CEP) setInput('input[formcontrolname="cep"]', data.CEP, {blur:true});
      if (data.numero) setInput('input[formcontrolname="numero"]', data.numero);
      if (data.complemento && !["","nan","none"].includes(lower(data.complemento))) setInput('input[formcontrolname="complemento"]', data.complemento);
      if (data.celular) setInput('input[formcontrolname="celular"]', data.celular);
      if (data.email) setInput('input[formcontrolname="email"]', data.email);
      if (data.renda_mensal){
        const v = data.renda_mensal.toString().includes('.')?data.renda_mensal:(data.renda_mensal+'.00');
        setInput('input[formcontrolname="valor_renda"]', v);
      }
      if (data.patrimonio){
        const v = data.patrimonio.toString().includes('.')?data.patrimonio:(data.patrimonio+'.00');
        setInput('input[formcontrolname="valor_patrimonio_total"]', v);
      }

      // segurança extra: fecha qualquer listbox que tenha sobrado aberto
      try{
        document.dispatchEvent(new KeyboardEvent('keydown',{key:'Escape',bubbles:true}));
        document.dispatchEvent(new KeyboardEvent('keyup',{key:'Escape',bubbles:true}));
        document.body && document.body.click();
      }catch(e){}

      return {ok:true, step:"sem_profissao"};
    })(arguments[0]);
    """

    safe = json_safe_dict_from_series(cliente_series)
    payload = {k: safe.get(k) for k in [
        "genero","nacionalidade","estado_civil","residencia_exterior","PEP",
        "tipo_documento","numero_documento","orgao_expedidor","uf_expedidor",
        "data_expedicao","CEP","numero","complemento","celular","email",
        "renda_mensal","patrimonio"
    ]}
    res = driver.execute_script(js, payload)

    #garantir que o UF esta selecionado (fechado)
    ensure_uf_dropdown_closed_and_selected(driver, timeout=6)
    return res



#3.2  CPF
def preencher_profissao_js(profissao_texto: str, timeout=10):
    """
    Preenche o campo 'profissão' dentro do shadow DOM de mf-iparceiros-cadastrocliente.
    - Aguarda o campo existir (com polling)
    - Tenta abrir seção/accordion se necessário
    - Aceita variações de seletor (ids-combobox, ids-input etc.)
    - Digita parcial + escolhe 1ª opção visível
    """
    global driver
    if not ensure_driver_alive(driver):
        raise RuntimeError("WebDriver inválido. Recrie o driver.")

    host = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "mf-iparceiros-cadastrocliente"))
    )

    # helper JS: tentar achar o input da profissão por diferentes seletores/estruturas
    def try_get_prof_input():
        return driver.execute_script("""
          const host = arguments[0];
          const root = host && host.shadowRoot;
          if (!root) return null;

          // 1) tenta input direto
          let el = root.querySelector('input[formcontrolname="profissao"]');
          if (el) return el;

          // 2) pode estar dentro de ids-combobox/fieldset
          el = root.querySelector('ids-combobox[formcontrolname="profissao"] input, ids-input[formcontrolname="profissao"] input');
          if (el) return el;

          // 3) fallback por label: procura label com texto contendo 'profiss' e acha input irmão
          const labs = Array.from(root.querySelectorAll('label, .ids-label, ids-label'));
          const lab = labs.find(l => (l.textContent||'').toLowerCase().includes('profiss'));
          if (lab) {
            // tenta achar um input próximo
            let c = lab.closest('ids-fieldset, ids-input, ids-combobox, .field, .form-group, div') || root;
            let cand = c.querySelector('input');
            if (cand) return cand;
          }

          return null;
        """, host)

    # Primeiro tenta achar; se não achar, tenta abrir seções/accordions e esperar
    inp = try_get_prof_input()
    if not inp:
        # Tenta abrir seção de "Dados Profissionais" (ou similar) dentro do shadow root
        driver.execute_script("""
          const host = arguments[0];
          const root = host && host.shadowRoot;
          if (!root) return;

          // tenta clicar em acordions que escondem profissão
          const texts = ['profiss', 'dados profissionais', 'ocupação', 'atividade'];
          const btns  = Array.from(root.querySelectorAll('button, [role="button"], ids-accordion-panel, .accordion, ids-accordion'));
          for (const b of btns) {
            const t = (b.textContent || '').toLowerCase();
            if (texts.some(x => t.includes(x))) {
              try { b.click(); } catch(e) {}
            }
          }
        """, host)

        # aguarda até timeout com polling curto
        deadline = time.time() + timeout
        while time.time() < deadline and not inp:
            time.sleep(0.2)
            inp = try_get_prof_input()

    if not inp:
        # último recurso: force scroll dentro do shadow para “despertar” lazy rendering
        driver.execute_script("""
          const host = arguments[0];
          const root = host && host.shadowRoot;
          if (!root) return;
          try {
            const last = root.querySelector('form, .container, .content, ids-panel, div:last-child') || root;
            last.scrollIntoView({block: 'end'});
          } catch(e){}
        """, host)
        time.sleep(0.3)
        inp = try_get_prof_input()

    if not inp:
        raise RuntimeError("Input de profissão não encontrado no shadowRoot.")

    full = (profissao_texto or "").strip()
    if len(full) == 0:
        return {"ok": False, "reason": "profissao_vazia"}

    partial_base = full[:-2] if len(full) >= 2 else ""
    penultimo = full[-2] if len(full) >= 2 else (full[-1] if len(full) == 1 else "")

    # Seta valor parcial via setter nativo e dispara eventos
    driver.execute_script("""
      const inp = arguments[0], val = arguments[1];
      function nativeSetValue(el, value){
        const proto = el instanceof HTMLInputElement ? HTMLInputElement.prototype
                   : el instanceof HTMLTextAreaElement ? HTMLTextAreaElement.prototype
                   : el.__proto__;
        const desc = proto && Object.getOwnPropertyDescriptor(proto, "value");
        if (desc && desc.set) desc.set.call(el, value); else el.value = value;
        el.dispatchEvent(new Event("input",  {bubbles:true}));
        el.dispatchEvent(new Event("change", {bubbles:true}));
      }
      try { inp.scrollIntoView({block:'center'}); } catch(e){}
      inp.focus();
      nativeSetValue(inp, "");
      nativeSetValue(inp, val);
    """, inp, partial_base)
    
    time.sleep(0.2)

    if penultimo:
        inp.send_keys(penultimo)

    # abre dropdown se houver
    try:
        driver.execute_script("""
          const host = arguments[0];
          const root = host.shadowRoot;
          const inp  = arguments[1];
          const c = inp.closest("ids-input, ids-fieldset, ids-combobox") || inp;
          const r = c.shadowRoot || c;
          const t = r && (r.querySelector('div[role="combobox"]')
                       || r.querySelector('button[aria-haspopup="listbox"]')
                       || r.querySelector('.ids-trigger')
                       || r.querySelector('.ids-select__trigger'));
          if (t) t.click();
        """, host, inp)
    except Exception:
        pass

    # clica a primeira opção visível
    def click_first_option():
        try:
            opts_txt = WebDriverWait(driver, 5).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "span.ids-option__text"))
            )
            for el in opts_txt:
                if el.is_displayed():
                    el.click()
                    return True
        except Exception:
            pass
        try:
            opts = driver.find_elements(By.CSS_SELECTOR, "ids-option[role='option'], .ids-option")
            for el in opts:
                if el.is_displayed():
                    el.click()
                    return True
        except Exception:
            pass
        return False

    if not click_first_option():
        try:
            inp.send_keys(Keys.ARROW_DOWN)
            inp.send_keys(Keys.ENTER)
        except Exception:
            return {"ok": False, "reason": "no-options-visible"}

    return {"ok": True, "typed": partial_base + penultimo}




#3.3 - CPF e CNPJ
def preencher_pagamento_boleto_js(cliente_series, timeout=8):
    global driver
    
    if not ensure_driver_alive(driver):
        raise RuntimeError("WebDriver inválido. Recrie o driver.")
    
    time.sleep(0.2)

    banco_atual          = str(cliente_series.get('banco') or '').strip()
    agencia_atual        = str(cliente_series.get('agencia') or '').strip()
    conta_corrente_atual = str(cliente_series.get('conta_corrente') or '').strip()
    digito_conta_atual   = str(cliente_series.get('digito_conta') or '').strip()

    WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "mf-iparceiros-contratacao"))
    )

    js = r"""
      return (function run(data){
        try{
          // === guards ===
          const host = document.querySelector('mf-iparceiros-contratacao');
          if(!host) return { ok:false, step:'no-host' };
          const root = host.shadowRoot;
          if(!root) return { ok:false, step:'no-shadow' };

          // === selecionar BOLETO (tenta BB e BO) ===
          function clickRadio(val){
            const inp = root.querySelector(`input[formcontrolname="forma_pagamento"][value="${val}"]`);
            if(!inp) return false;
            try { inp.scrollIntoView({block:'center'}); } catch(e){}
            if(!inp.checked){
              try{
                inp.click();
                inp.dispatchEvent(new Event('input',{bubbles:true}));
                inp.dispatchEvent(new Event('change',{bubbles:true}));
              }catch(e){}
            }
            return !!inp.checked;
          }
          let boleto = clickRadio('BB') || clickRadio('BO');
          if(!boleto) return { ok:false, step:'boleto-not-found' };

          // pequena espera para UI reagir
          // (sincrono: só da uma chance a máscara/DOM)
          // o Python ainda pode aguardar fora com time.sleep(0.2) se quiser
          // mas aqui seguimos direto

          // === checar campos dependentes ===
          function fields(){
            const banco  = root.querySelector('input[formcontrolname="nome_banco_encerramento"], input[formcontrolname="nome_banco"]');
            const ag     = root.querySelector('input[formcontrolname="agencia_encerramento"], input[formcontrolname="agencia"]');
            const conta  = root.querySelector('input[formcontrolname="conta_encerramento"], input[formcontrolname="conta"]');
            const digito = root.querySelector('input[formcontrolname="digito_encerramento"], input[formcontrolname="digito"]');
            return { banco, ag, conta, digito };
          }
          let { banco, ag, conta, digito } = fields();
          if(!(banco && ag && conta && digito)){
            return { ok:false, step:'fields-missing' };
          }

          // === helpers ===
          function nativeSetValue(el, value){
            if(!el) return;
            const proto =
              el instanceof HTMLInputElement ? HTMLInputElement.prototype :
              el instanceof HTMLTextAreaElement ? HTMLTextAreaElement.prototype :
              el.__proto__;
            const desc = proto && Object.getOwnPropertyDescriptor(proto, "value");
            if (desc && desc.set) desc.set.call(el, value); else el.value = value;
            el.dispatchEvent(new Event("input", {bubbles:true}));
            el.dispatchEvent(new Event("change", {bubbles:true}));
          }

          // Banco + tentativa de sugerir 1ª opção
          (function fillBanco(){
            if(!banco) return;
            try { banco.scrollIntoView({block:'center'}); } catch(e){}
            nativeSetValue(banco, "");
            nativeSetValue(banco, (data.banco||""));
            try{
              const c = banco.closest("ids-input, ids-fieldset, ids-combobox") || banco;
              const r = c && c.shadowRoot ? c.shadowRoot : c;
              const t = r && (r.querySelector('div[role="combobox"]')
                          || r.querySelector('button[aria-haspopup="listbox"]')
                          || r.querySelector('.ids-trigger')
                          || r.querySelector('.ids-select__trigger'));
              if (t) t.click();
            }catch(e){}
            setTimeout(()=>{
              try{
                const doc = Array.from(document.querySelectorAll('ids-option, .ids-option, span.ids-option__text'));
                const loc = Array.from((root||document).querySelectorAll('ids-option, .ids-option, span.ids-option__text'));
                const all = doc.concat(loc).filter(o => o && (o.offsetParent !== null));
                if (all.length){
                  const first = all[0].closest && all[0].closest('ids-option') ? all[0].closest('ids-option') : all[0];
                  first && first.click();
                  banco.dispatchEvent(new Event("input",{bubbles:true}));
                  banco.dispatchEvent(new Event("change",{bubbles:true}));
                }
              }catch(e){}
            }, 30);
          })();

          // Agência, Conta, Dígito
          nativeSetValue(ag,     (data.agencia||""));
          nativeSetValue(conta,  (data.conta||""));
          nativeSetValue(digito, (data.digito||""));

          // Checkboxes de ciência
          function check(sel){
            const el = root.querySelector(sel);
            if (!el) return;
            if (!el.checked){
              try{
                el.click();
                el.dispatchEvent(new Event('input',{bubbles:true}));
                el.dispatchEvent(new Event('change',{bubbles:true}));
              }catch(e){}
            }
          }
          check('input[formcontrolname="cienciaGarantiaPrazo"]');
          check('input[formcontrolname="cienciaRegrasCancelamento"]');
          check('input[formcontrolname="cienciaRegrasCRP"]');

          return { ok:true, step:'boleto-preenchido' };
        }catch(e){
          return { ok:false, step:'exception', message: String(e && e.message || e) };
        }
      })(arguments[0]);
    """

    try:
        res = driver.execute_script(js, {
            "banco": banco_atual,
            "agencia": agencia_atual,
            "conta": conta_corrente_atual,
            "digito": digito_conta_atual
        })
    except Exception as e:
        print("JavascriptException em preencher_pagamento_boleto_js:", e)
        return {"ok": False, "step": "js-exception", "message": str(e)}

    print("Resultado preencher_pagamento_boleto_fast:", repr(res))
    return res




### 3.1 Preencher dados para CNPJ > JS
def preencher_cliente_CNPJ_js(cliente_series):
    global driver
    if not ensure_driver_alive(driver):
        raise RuntimeError("WebDriver inválido. Recrie o driver.")
    time.sleep(1.5)

    WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "mf-iparceiros-cadastrocliente"))
    )

        # >>> Acessa o Shadow host
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "mf-iparceiros-cadastrocliente")))
    host1 = driver.find_element(By.CSS_SELECTOR, "mf-iparceiros-cadastrocliente")
    shadow_root1 = driver.execute_script("return arguments[0].shadowRoot", host1)
    print("Shadow DOM acessado com sucesso.")

    WebDriverWait(driver, 10).until(
      lambda d: len(shadow_root1.find_elements(By.CSS_SELECTOR, 'input[formcontrolname="CNAE"][id="atividade"]')) > 0
    )
    time.sleep(0.1)
    # ---- payload seguro + normalizações (INCLUI data_nasc_fundacao) ----
    payload = {
        "atividade_principal": to_str_safe(cliente_series.get("atividade_principal")),

        "inscricao_municipal": normalize_inscricao_municipal(cliente_series.get("inscricao_municipal")),
        "inscricao_estadual":  to_str_safe(cliente_series.get("inscricao_estadual")),

        "residencia_exterior": to_str_safe(cliente_series.get("residencia_exterior")).lower(),

        "renda_mensal":        to_str_safe(cliente_series.get("renda_mensal")),
        "patrimonio":          to_str_safe(cliente_series.get("patrimonio")),

        "PEP":                 to_str_safe(cliente_series.get("PEP")).lower(),

        # CEP EMPRESA normalizado
        "cep_empresa":         normalize_cep(cliente_series.get("CEP")),
        "numero_empresa":      to_str_safe(cliente_series.get("numero")),
        "complemento_empresa": to_str_safe(cliente_series.get("complemento")),

        "email":               to_str_safe(cliente_series.get("email")),
        "celular":             to_str_safe(cliente_series.get("celular")),
        "telefone":            to_str_safe(cliente_series.get("telefone")),

        # >>>: data de fundação em formato BR
        "data_nasc_fundacao":  normalize_date_br(cliente_series.get("data_nasc_fundacao")),

        # SÓCIO — com normalizações
        "cpf_socio":           only_digits(cliente_series.get("cpf_socio")),
        "data_nasc_socio":     normalize_date_br(cliente_series.get("data_nasc_socio")),
        "nome_socio":          to_str_safe(cliente_series.get("nome_socio")),
        "perc_participacao":   only_digits(cliente_series.get("perc_participacao")),
        "data_entrada_socio":  normalize_date_br(cliente_series.get("data_entrada_socio")),

        # RG do sócio: somente números
        "rg_socio":            only_digits(cliente_series.get("rg_socio")),
        "orgao_emissor_socio": to_str_safe(cliente_series.get("orgao_emissor_socio")),
        "Socio_e_PEP":         to_str_safe(cliente_series.get("Socio_e_PEP")).lower(),

        "uf_expe_socio":       to_str_safe(
                                  cliente_series.get("uf_expe_socio")
                                  or cliente_series.get("uf_expe_socio".replace("expe","expedidor"))
                              ).upper(),
        "data_expe_socio":     normalize_date_br(
                                  cliente_series.get("data_expe_socio")
                                  or cliente_series.get("data_esp_socio")   # fallback p/ seu nome antigo
                              ),
        # CEP SÓCIO normalizado
        "CEP_Socio":           normalize_cep(cliente_series.get("CEP_Socio")),
        "num_residencia_socio":to_str_safe(cliente_series.get("num_residencia_socio")),
        "complemento_socio":   to_str_safe(cliente_series.get("complemento_socio")),
        "celular_socio":       to_str_safe(cliente_series.get("celular_socio")),
    }

    js = r"""
        return (function fill(data){
          function q(s, r=document){ return r.querySelector(s); }

          // --- Guards iniciais (host/root) ---
          const host = q("mf-iparceiros-cadastrocliente");
          if (!host) return { ok:false, step:"no-host" };
          const root = host.shadowRoot;
          if (!root) return { ok:false, step:"no-shadow" };

          // ------- helpers base -------
          function nativeSetValue(el, value){
            if(!el) return;
            const proto =
              el instanceof HTMLInputElement ? HTMLInputElement.prototype :
              el instanceof HTMLTextAreaElement ? HTMLTextAreaElement.prototype :
              el.__proto__;
            const desc = proto && Object.getOwnPropertyDescriptor(proto, "value");
            if (desc && desc.set) desc.set.call(el, value); else el.value = value;
            el.dispatchEvent(new Event("input", {bubbles:true}));
            el.dispatchEvent(new Event("change", {bubbles:true}));
          }
          function setInput(sel, val, opts={}){
            const el = root.querySelector(sel);
            if(!el) return {ok:false, sel};
            try { el.scrollIntoView({block:'center'}); } catch(e){}
            nativeSetValue(el, (val ?? "").toString());
            if (opts.blur) el.dispatchEvent(new Event("blur", {bubbles:true}));
            return {ok:true};
          }
          function setInputNth(sel, idx, val, opts={}){
            const els = root.querySelectorAll(sel);
            const el = (els && els.length > idx) ? els[idx] : null;
            if(!el) return {ok:false, sel, idx};
            try { el.scrollIntoView({block:'center'}); } catch(e){}
            nativeSetValue(el, (val ?? "").toString());
            if (opts.blur) el.dispatchEvent(new Event("blur", {bubbles:true}));
            return {ok:true};
          }

          // ------- datas (dd/mm/aaaa) -------
          function typeMaskedDigits(el, digits){
            if (!el) return;
            try{
              el.focus();
              const proto = el instanceof HTMLInputElement ? HTMLInputElement.prototype : el.__proto__;
              const desc  = proto && Object.getOwnPropertyDescriptor(proto, "value");
              if (desc && desc.set) desc.set.call(el, ""); else el.value = "";
              el.dispatchEvent(new Event("input", {bubbles:true}));
              let acc = "";
              for (const ch of (digits || "")){
                if (!/[0-9]/.test(ch)) continue;
                acc += ch;
                if (desc && desc.set) desc.set.call(el, acc); else el.value = acc;
                el.dispatchEvent(new Event("input", {bubbles:true}));
              }
              el.dispatchEvent(new Event("change", {bubbles:true}));
            }catch(e){}
          }
          function brDateToDigits(brDate){
            if (!brDate) return "";
            const d = String(brDate).replace(/\D+/g,"");
            if (d.length === 8){
              const yyyy = d.slice(0,4), mm = d.slice(4,6), dd = d.slice(6,8);
              const looksISO = (+yyyy > 1900 && +mm >= 1 && +mm <= 12 && +dd >= 1 && +dd <= 31);
              return looksISO ? (dd+mm+yyyy) : d; // yyyy-mm-dd -> ddmmyyyy
            }
            return d;
          }
          function setDateMasked(sel, brDate){
            const el = root.querySelector(sel);
            if (!el) return {ok:false, sel};
            typeMaskedDigits(el, brDateToDigits(brDate));
            return {ok:true};
          }
          function setInputLast(sel, val, opts={}){
            const els = root.querySelectorAll(sel);
            const el = (els && els.length) ? els[els.length - 1] : null;
            if(!el) return {ok:false, sel, idx:"last"};
            try { el.scrollIntoView({block:'center'}); } catch(e){}
            nativeSetValue(el, (val ?? "").toString());
            if (opts.blur) el.dispatchEvent(new Event("blur", {bubbles:true}));
            return {ok:true};
          }
          function setDateMaskedLast(sel, brDate){
            const els = root.querySelectorAll(sel);
            const el = (els && els.length) ? els[els.length - 1] : null;
            if(!el) return {ok:false, sel, idx:"last"};
            typeMaskedDigits(el, brDateToDigits(brDate));
            return {ok:true};
          }
          function setDateMaskedNth(sel, idx, brDate){
            const els = root.querySelectorAll(sel);
            const el = (els && els.length > idx) ? els[idx] : null;
            if (!el) return {ok:false, sel, idx};
            typeMaskedDigits(el, brDateToDigits(brDate));
            return {ok:true};
          }
          function clickRadioNth(formSel, valueVN, idx){
            const radios = root.querySelectorAll(`input[formcontrolname="${formSel}"][value="${valueVN}"]`);
            const el = (radios && radios.length > idx) ? radios[idx] : null;
            if(!el) return {ok:false, formSel, valueVN, idx};
            try { el.scrollIntoView({block:'center'}); } catch(e){}
            if(!el.checked){
              try{
                el.click();
                el.dispatchEvent(new Event('input',{bubbles:true}));
                el.dispatchEvent(new Event('change',{bubbles:true}));
              }catch(e){}
            }
            return {ok:true};
          }
          function clickRadio(formSel, valueVN){
            const el = root.querySelector(`input[formcontrolname="${formSel}"][value="${valueVN}"]`);
            if(!el) return {ok:false, formSel, valueVN};
            if(!el.checked){
              try { el.scrollIntoView({block:'center'}); } catch(e){}
              try{
                el.click();
                el.dispatchEvent(new Event("input",{bubbles:true}));
                el.dispatchEvent(new Event("change",{bubbles:true}));
              }catch(e){}
            }
            return {ok:true};
          }

          // CNAE
          function fillCNAE(val){
            const inp = root.querySelector('input[formcontrolname="CNAE"]#atividade');
            if(!inp) return {ok:false, field:"CNAE"};
            try { inp.scrollIntoView({block:'center'}); } catch(e){}
            nativeSetValue(inp, "");
            nativeSetValue(inp, (val||"").toString());
            try{
              const host = inp.closest("ids-input, ids-fieldset, ids-combobox") || inp;
              const r = host && host.shadowRoot ? host.shadowRoot : host;
              const t = r && (r.querySelector('div[role="combobox"]')
                          || r.querySelector('button[aria-haspopup="listbox"]')
                          || r.querySelector('.ids-trigger')
                          || r.querySelector('.ids-select__trigger'));
              if (t) t.click();
            }catch(e){}
            function options(){
              const doc = Array.from(document.querySelectorAll('ids-option, .ids-option, span.ids-option__text'));
              const loc = Array.from((root||document).querySelectorAll('ids-option, .ids-option, span.ids-option__text'));
              return loc.concat(doc);
            }
            setTimeout(()=>{
              try{
                const all = options().filter(o => o && (o.offsetParent !== null));
                if (all.length){
                  const first = all[0].closest && all[0].closest('ids-option') ? all[0].closest('ids-option') : all[0];
                  first && first.click();
                  inp.dispatchEvent(new Event("input",{bubbles:true}));
                  inp.dispatchEvent(new Event("change",{bubbles:true}));
                }
              }catch(e){}
            }, 30);
            return {ok:true};
          }

          // ------- Empresa -------
          try{
            if (data.atividade_principal) fillCNAE(data.atividade_principal);

            if (data.inscricao_municipal) setInput('input[formcontrolname="inscricao_municipal"]', data.inscricao_municipal);
            if (data.inscricao_estadual)  setInput('input[formcontrolname="inscricao_estadual"]',  data.inscricao_estadual);

            if (data.residencia_exterior){
              const vn = ["sim","s","yes","y"].includes((data.residencia_exterior||"").toLowerCase()) ? "S" : "N";
              clickRadio("residencia_exterior", vn);
            }

            if (data.renda_mensal){
              const v = data.renda_mensal.toString().includes('.') ? data.renda_mensal : (data.renda_mensal + '.00');
              setInput('input[formcontrolname="valor_faturamento_medio"]', v);
            }
            if (data.patrimonio){
              const v = data.patrimonio.toString().includes('.') ? data.patrimonio : (data.patrimonio + '.00');
              setInput('input[formcontrolname="valor_capital"]', v);
            }
            if (data.PEP){
              const vn = ["sim","s","yes","y"].includes((data.PEP||"").toLowerCase()) ? "S" : "N";
              clickRadioNth("indicador_politicamente_exposto", vn, 0);
            }

            if (data.data_nasc_fundacao){
              setDateMasked('input[formcontrolname="data_nascimento_fundacao"]', data.data_nasc_fundacao);
            }

            if (data.cep_empresa){
              setInput('input[formcontrolname="cep"]', data.cep_empresa, {blur:true});
            }
            if (data.numero_empresa) setInput('input[formcontrolname="numero"]', data.numero_empresa);
            if (data.complemento_empresa){
              const comp = (data.complemento_empresa||"").toLowerCase();
              if (comp !== "" && comp !== "nan" && comp !== "none"){
                setInput('input[formcontrolname="complemento"]', data.complemento_empresa);
              }
            }
            if (data.email)    setInput('input[formcontrolname="email"]',   data.email);
            if (data.celular)  setInput('input[formcontrolname="celular"]', data.celular);
            if (data.telefone) setInput('input[formcontrolname="telefone"]',data.telefone);

            // ------- Sócio -------
            if (data.cpf_socio) setInputNth('input[formcontrolname="cpf_cnpj"]', 0, data.cpf_socio);

            if (data.data_nasc_socio){
              setDateMaskedNth('input[formcontrolname="data_nascimento_fundacao"]', 0, data.data_nasc_socio);
            }
            if (data.data_entrada_socio){
              setDateMaskedNth('input[formcontrolname="dt_entrada"]', 0, data.data_entrada_socio);
            }

            if (data.nome_socio)        setInputNth('input[formcontrolname="nome_completo_razao_social"]', 0, data.nome_socio);
            if (data.perc_participacao) setInputNth('input[formcontrolname="pe_participacao"]', 0, data.perc_participacao);

            if (data.rg_socio)            setInput('input[formcontrolname="numero_documento"]', data.rg_socio);
            if (data.orgao_emissor_socio) setInput('input[formcontrolname="orgaoExpedidor"]', data.orgao_emissor_socio);

            if (data.uf_expe_socio){
              setInputLast('input[formcontrolname="UFexpedidor"]', data.uf_expe_socio);
            }

            if (data.data_expe_socio){
              setDateMaskedLast('input[formcontrolname="data_emissao_documento"]', data.data_expe_socio);
            }

            if (data.Socio_e_PEP){
              const vnS = ["sim","s","yes","y"].includes((data.Socio_e_PEP||"").toLowerCase()) ? "S" : "N";
              clickRadioNth("indicador_politicamente_exposto", vnS, 1);
            }

            if (data.CEP_Socio){
              setInputNth('input[formcontrolname="cep"]', 1, data.CEP_Socio, {blur:true});
            }
            if (data.num_residencia_socio){
              setInputNth('input[formcontrolname="numero"]', 1, data.num_residencia_socio);
            }
            if (data.complemento_socio){
              const cs = (data.complemento_socio||"").toLowerCase();
              if (cs !== "" && cs !== "nan" && cs !== "none"){
                setInputNth('input[formcontrolname="complemento"]', 1, data.complemento_socio);
              }
            }
            if (data.celular_socio){
              setInputNth('input[formcontrolname="celular"]', 1, data.celular_socio);
            }
          }catch(e){
            return { ok:false, step:"exception", message: String(e && e.message || e) };
          }

          return {ok:true, step:"cnpj-fast"};
        })(arguments[0]);
        """


    res = driver.execute_script(js, payload)
    print("Resultado preencher_cliente_CNPJ_js:", res)
    return res





###
#=======================================================================================================================
# Finalizar preenchimento dos dados do cliente e finalizar a proposta

# 3 (Principal CNPJ)
def preencher_dados_PJ ():
    global driver, action, df_atual, grupo_encontrado, wait, cpf_atual, cliente_atual
    
    wait = WebDriverWait(driver, 10)
    wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "mf-iparceiros-cadastrocliente")))
    #verificar se o shadow host está presente
    # >>> Acessa o Shadow host
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "mf-iparceiros-cadastrocliente")))
    host1 = driver.find_element(By.CSS_SELECTOR, "mf-iparceiros-cadastrocliente")
    shadow_root1 = driver.execute_script("return arguments[0].shadowRoot", host1)
    print("Shadow DOM acessado com sucesso.")




    #========================================================================================================

    # Clicar em Continuar para ir para preenchimento de dados bancarios e gerar boleto - Etapa 2
    time.sleep(0.1)
    res = None
    try:
        res = preencher_cliente_CNPJ_js(cliente_atual)
        print("Resultado do preenchimento dados do CNPJ:", repr(res))
    except Exception as e:
        print("JS crashed:", e)
        return False

    if not (isinstance(res, dict) and res.get("ok") is True):
        print(f"Falha ao preencher dados do CNPJ. Necessário revisão manual. Detalhes: {repr(res)}")
        return False


    
    time.sleep(0.1)

    #========================================================================================================



    ### CONTINUAR para a próxima etapa ### (botão)
    try:
        
        btn_continuar = shadow_root1.find_element(By.CSS_SELECTOR, ".cadastro .btn-enviar button.ids-main-button")
        action.move_to_element(btn_continuar).pause(random.uniform(0.02, 0.2)).click(btn_continuar).perform()
        human_sleep()
        print("Clicado em 'Continuar', aguardando próxima tela...")
    except Exception as e:
        
        print("⚠️ Necessário revisão manual para este cliente.")
          # Sai da função sem continuar para a próxima etapa

    time.sleep(1)

    #Verificar se o botão CONTINUAR ainda está presente (o que indica que não avançou) e tentar clicar novamente
    try:
        botao_continuar_check = shadow_root1.find_element(By.CSS_SELECTOR, ".cadastro .btn-enviar button.ids-main-button")
        if botao_continuar_check and botao_continuar_check.is_displayed():
            print("O botão 'Continuar' ainda está presente. Tentando clicar novamente...")
            time.sleep(1)
            action.move_to_element(btn_continuar).pause(random.uniform(0.02, 0.2)).click(btn_continuar).perform()
            human_sleep()
            print("Clicado em 'Continuar' novamente, aguardando próxima tela...")
    except Exception as e:
        
        pass  # Se der erro, ignora e continua
    
    time.sleep(1)



    
    #                                   Pagamento de boleto - Etapa 3
    #========================================================================================================


    print("Preenchendo dados de pagamento via boleto...")

    res3 = preencher_pagamento_boleto_js(cliente_atual)
    print("Resultado do preenchimento do pagamento via boleto:", repr(res3))
    if not (isinstance(res3, dict) and res3.get("ok") is True):
        print("❌ Falha ao preencher os dados de pagamento via boleto. Detalhes:", repr(res3))
        return False

    #--------------------------------------------------------------------------------------------------------



    time.sleep(2)

    # >>> Acessa o Shadow host
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "mf-iparceiros-contratacao")))
    host1 = driver.find_element(By.CSS_SELECTOR, "mf-iparceiros-contratacao")
    shadow_root2 = driver.execute_script("return arguments[0].shadowRoot", host1)

    print("Shadow DOM da etapa de contratação acessado com sucesso.")
        
    #==========================================
    ####         Botão Contratar           ####
    #==========================================
    try:
        #selecionar botao pelo texto "contratar"
        time.sleep(1) # webdriverwait não funciona aqui -
        contratar = shadow_root2.find_element(By.CLASS_NAME, 'btn-contratar')
        botao_contratar = contratar.find_elements(By.TAG_NAME, 'button')
        # action.move_to_element(botao_contratar[0]).pause(random.uniform(0.2, 0.7)).click(botao_contratar[0]).perform()
        # human_sleep(0.8, 1.4)
        #mover o mouse para o botão e esperar 2 segundos
        action.move_to_element(botao_contratar[0]).pause(2).perform()
        
        print("Finalizando....")

        return True # Retorna True indicando que tudo foi preenchido com sucesso e pode prosseguir para o próximo cliente
    
    except Exception as e:
        print(f"❌ Erro ao tentar localizar ou clicar no botão 'Contratar': {e}")
        print("⚠️ Necessário revisão manual para este cliente.")
        return False    #Retorna False indicando que houve um problema e não deve prosseguir para o próximo cliente 
    
    
# 3 (Principal CPF)
def preencher_dados_pessoais():
    global driver, action, df_atual, grupo_encontrado, wait, cpf_atual, cliente_atual

    print("Iniciando o preenchimento dos dados pessoais do cliente...")

    #Esperar todos elementos carregarem
    time.sleep(1.5)
    wait = WebDriverWait(driver, 10)
    wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "mf-iparceiros-cadastrocliente")))
    #verificar se o shadow host está presente
    # >>> Acessa o Shadow host
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "mf-iparceiros-cadastrocliente")))
    host1 = driver.find_element(By.CSS_SELECTOR, "mf-iparceiros-cadastrocliente")
    shadow_root1 = driver.execute_script("return arguments[0].shadowRoot", host1)
    print("Shadow DOM acessado com sucesso.")


    #                                          PREENCHER DADOS PESSOAIS 
    #--------------------------------------------------------------------------------------------------------
    
    #preencher dados sem profissão:

    res = preencher_cliente_s_profi_js(cliente_atual)
    print("Resultado do preenchimento sem profissão:", res)
    if not isinstance(res, dict) or not res.get("ok"):
        print("Falha ao preencher dados pessoais (sem profissão). Necessário revisão manual.")
        return False

    time.sleep(0.2)
    #descer a tela ate o final
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(0.1)

    #preencher profissão:
    res2 = preencher_profissao_js(str(cliente_atual.get('profissao_cliente') or '').strip())
    print("Resultado do preenchimento da profissão:", res2)
    if not isinstance(res2, dict) or not res2.get("ok"):
        print("Falha ao preencher a profissão. Necessário revisão manual.")
        return False

    human_sleep(1.5, 2)
    #--------------------------------------------------------------------------------------------------------






    ### 2.19 CONTINUAR para a próxima etapa ### (botão)
    try:
        botao_continuar = shadow_root1.find_elements(By.CSS_SELECTOR, 'button[idsmainbutton]')
        action.move_to_element(botao_continuar[0]).pause(random.uniform(0.02, 0.2)).click(botao_continuar[0]).perform()
        human_sleep()
        print("Clicado em 'Continuar', aguardando próxima tela...")
    except Exception as e:
        
        print("⚠️ Necessário revisão manual para este cliente.")
        return  # Sai da função sem continuar para a próxima etapa
    
    time.sleep(1)

    #Verificar se o botão CONTINUAR ainda está presente (o que indica que não avançou) e tentar clicar novamente
    try:
        botao_continuar_check = shadow_root1.find_elements(By.CSS_SELECTOR, 'button[idsmainbutton]')
        if botao_continuar_check and botao_continuar_check[0].is_displayed():
            print("O botão 'Continuar' ainda está presente. Tentando clicar novamente...")
            time.sleep(1)
            action.move_to_element(botao_continuar_check[0]).pause(random.uniform(0.02, 0.2)).click(botao_continuar_check[0]).perform()
            human_sleep()
            print("Clicado em 'Continuar' novamente, aguardando próxima tela...")
    except Exception as e:
        
        pass  # Se der erro, ignora e continua

    #========================================================================================================

    # select root do shadow DOM da próxima etapa

    time.sleep(2)

    # >>> Acessa o Shadow host
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "mf-iparceiros-contratacao")))
    host1 = driver.find_element(By.CSS_SELECTOR, "mf-iparceiros-contratacao")
    shadow_root2 = driver.execute_script("return arguments[0].shadowRoot", host1)
    print("Shadow DOM da etapa de contratação acessado com sucesso.")

    time.sleep(0.2)

    #                                   Pagamento de boleto - Etapa 3
    #========================================================================================================


    print("Preenchendo dados de pagamento via boleto...")

    res3 = preencher_pagamento_boleto_js(cliente_atual)
    print("Resultado do preenchimento do pagamento via boleto:", repr(res3))
    if not (isinstance(res3, dict) and res3.get("ok") is True):
        print("❌ Falha ao preencher os dados de pagamento via boleto. Detalhes:", repr(res3))
        return False

    #--------------------------------------------------------------------------------------------------------


    time.sleep(0.2)


    #==========================================
    ####         Botão Contratar           ####
    #==========================================
    try:
        #selecionar botao pelo texto "contratar"
        time.sleep(1) # webdriverwait não funciona aqui -
        contratar = shadow_root2.find_element(By.CLASS_NAME, 'btn-contratar')
        botao_contratar = contratar.find_elements(By.TAG_NAME, 'button')
        # action.move_to_element(botao_contratar[0]).pause(random.uniform(0.2, 0.7)).click(botao_contratar[0]).perform()
        # human_sleep(0.8, 1.4)
        #mover o mouse para o botão e esperar 2 segundos
        action.move_to_element(botao_contratar[0]).pause(2).perform()
        
        print("Finalizando....")

        return True # Retorna True indicando que tudo foi preenchido com sucesso e pode prosseguir para o próximo cliente
    
    except Exception as e:
        print(f"❌ Erro ao tentar localizar ou clicar no botão 'Contratar': {e}")
        print("⚠️ Necessário revisão manual para este cliente.")
        return False    #Retorna False indicando que houve um problema e não deve prosseguir para o próximo cliente 





# FIM




def _start_driver_bg():
    global driver_thread, _driver_keepalive_evt

    try:
        # desabilita botão UI
        root.after(0, lambda: btn_iniciar_driver.configure(state="disabled"))

        iniciar_driver()  # cria driver com use_subprocess=True (mantém stealth)

        # driver iniciado: avisa o usuário
        root.after(0, lambda: messagebox.showinfo("OK", "Driver iniciado!"))

        # aguarda até que o app peça para encerrar (isso mantém o thread vivo)
        _driver_keepalive_evt.wait()   # <-- bloqueia aqui até on_close setar o evento

    except Exception as e:
        root.after(0, lambda: btn_iniciar_driver.configure(state="normal"))
        root.after(0, lambda: messagebox.showerror("Erro", str(e)))


def start_driver_thread():
    global driver_thread, _driver_keepalive_evt
    _driver_keepalive_evt.clear()
    # driver_thread = threading.Thread(target=_start_driver_bg, daemon=True)
    # driver_thread.start()
    threading.Thread(target=_start_driver_bg, daemon=True).start()





    
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#                                                               #   #   #   #   #       ---      Configuração da Interface Gráfica      ---    #   #   #   #   #  
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------




# ==================== Cores personalizadas ====================
cor_fundo_janela   = "#040312"
cor_fundo_frame    = "#0C1B39"
cor_botao_fundo    = "#4691BF"
cor_botao_hover    = "#5988D9"
cor_botao_texto    = "white"
cor_texto_label    = "white"
cor_fundo_entry    = "#12121B"
cor_texto_entry    = "#E0E0E0"


# app.py (logo no começo do main())
import socket
def _ensure_single_instance(port=49666):
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.bind(("127.0.0.1", port))
    except OSError:
        # já tem uma instância
        return None
    s.listen(1)
    return s  # manter a referência viva



def main():
    global btn_iniciar_driver, root
    
    guard = _ensure_single_instance()
    if guard is None:
        return  # já existe outra janela aberta
    
    # ... resto do seu Tk ...
    if tk._default_root is not None:
        return
    
    # ==================== Janela principal ====================
    root = tk.Tk()
    root.title("Automação Seleção de Consórcio")
    root.configure(bg=cor_fundo_janela)
    root.minsize(width=430, height=480)
    root.columnconfigure(0, weight=1)
    
    root.protocol("WM_DELETE_WINDOW", on_close)

    # ==================== Tema ttk ====================
    style = ttk.Style()
    style.theme_use("clam")  # Força o tema compatível com macOS

    style.configure(
        "Custom.TButton",
        background=cor_botao_fundo,
        foreground=cor_botao_texto,
        font=("Helvetica", 12, "bold"),
        borderwidth=0,
        focusthickness=3,
        focuscolor=cor_fundo_frame,
        padding=6,
        )
    style.map(
        "Custom.TButton",
        background=[("active", cor_botao_hover)],
        foreground=[("active", cor_botao_texto)]
        )

    # ==================== Título ====================
    titulo_label = tk.Label(
        root, 
        text="Simular e Contratar Consórcio",
        font=("Helvetica", 14, "bold"),
        bg=cor_fundo_janela,
        fg=cor_botao_texto
        )
    titulo_label.pack(pady=(10, 10))




    # =================== Frame: Iniciar Driver ===================
    cred_frame = tk.LabelFrame(root, text="Abrir Chrome", bg=cor_fundo_frame, fg=cor_texto_label, padx=10, pady=4)
    cred_frame.pack(padx=15, pady=5, fill="x")
    cred_frame.columnconfigure(1, weight=1)

    btn_iniciar_driver = ttk.Button(cred_frame, text="Abrir Navegador", style="Custom.TButton")
    btn_iniciar_driver.configure(command=start_driver_thread)
    btn_iniciar_driver.grid(row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=(7, 4))





    # =================== Frame: Buscar e guardar Grupos (Antigos) ===================
    cred_frame = tk.LabelFrame(root, text="Guardar os códigos de Grupos para desconsiderar", bg=cor_fundo_frame, fg=cor_texto_label, padx=10, pady=4)
    cred_frame.pack(padx=15, pady=5, fill="x")
    cred_frame.columnconfigure(1, weight=1)

    btn_salvar_grupos = ttk.Button(
        cred_frame, text="Salvar os códigos de Grupos", style="Custom.TButton",
        command=guardar_grupos_disponiveis
                        )
    btn_salvar_grupos.grid(row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=(7, 4))








    # =================== Frame: Inserir Cliente ===================
    fila_frame = tk.LabelFrame(root, text="Inserir CPF/Data Nasc e Tipo Cons. do Cliente", bg=cor_fundo_frame, fg=cor_texto_label, padx=10, pady=2)
    fila_frame.pack(padx=15, pady=2, fill="x")
    fila_frame.columnconfigure(1, weight=1)

    botao_buscar = ttk.Button(
    fila_frame, text="Inserir dados Cliente", style="Custom.TButton",
    command=inserir_dados_cliente_js 
                )
    botao_buscar.grid(row=0, column=0, columnspan=2, pady=3, padx=5, sticky="ew")

    fila_status_label = tk.Label(fila_frame, text="...", font=("TkDefaultFont", 9, "italic"),
                            bg=cor_fundo_frame, fg=cor_texto_label)
    fila_status_label.grid(row=1, column=0, columnspan=2, pady=2)





    # =================== Frame: Função Principal ===================
    busar_Consorcio_frame = tk.LabelFrame(root, text="Principal: Busca Brugo > Credito e Preenche dados cliente", bg=cor_fundo_frame, fg=cor_texto_label, padx=10, pady=2)
    busar_Consorcio_frame.pack(padx=15, pady=5, fill="x")

    btn_buscar_grupo = ttk.Button(
    busar_Consorcio_frame, text="Buscar Consórcio", style="Custom.TButton",
    command=buscar_consorcio_cliente 
                )
    btn_buscar_grupo.pack(pady=5, padx=5, fill="x")

    bsc_consc_status_label = tk.Label(busar_Consorcio_frame, text="...", font=("TkDefaultFont", 9, "italic"),
                            bg=cor_fundo_frame, fg=cor_texto_label)
    bsc_consc_status_label.pack(pady=5)




    # =================== Inicia GUI ===================
    root.mainloop()

    # =================== Encerramento ===================
    # try:
    #     if driver:
    #         driver.quit()
    # except NameError:
    #     pass

if __name__ == "__main__":
    import multiprocessing as mp
    mp.freeze_support()   # OK manter mesmo no macOS
    main()
    
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
