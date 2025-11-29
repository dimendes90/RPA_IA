# Módulos da Biblioteca Padrão do Python
import os                   # Interage com o sistema operacional (manipulação de arquivos, variáveis de ambiente).
import sys                  # Fornece acesso a variáveis e funções específicas do interpretador (ex.: argumentos de linha de comando).

import multiprocessing

try:
    multiprocessing.set_start_method('spawn', force=True)
    print("Modo 'spawn' do multiprocessing ativado.")
except (RuntimeError, ValueError) as e:

    print(f"Aviso ao configurar 'spawn' (ignorar se for processo-filho): {e}")



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
from tkinter import scrolledtext, messagebox # Componentes avançados (widgets), área de texto com scroll e caixas de diálogo do Tkinter.
import undetected_chromedriver as uc # type: ignore # Um driver para Selenium que tenta evitar ser detectado por sites.
import tkinter.ttk as ttk

# ---

# Módulos e Componentes Específicos do Selenium WebDriver
from selenium import webdriver      # O módulo principal para controlar navegadores via Selenium.
from selenium.webdriver.chrome.options import Options # Permite configurar opções específicas do navegador Chrome.
from selenium.webdriver.common.by import By             # Define métodos para localizar elementos HTML (ex.: ID, XPATH, CSS_SELECTOR).
from selenium.webdriver.common.action_chains import ActionChains # Realiza uma cadeia de ações avançadas do usuário (mouse, teclado).
from selenium.webdriver.support.ui import WebDriverWait # Fornece um mecanismo para esperar por elementos em uma página antes de interagir.
from selenium.webdriver.support.ui import WebDriverWait as W
from selenium.webdriver.support import expected_conditions as EC # Define condições pré-definidas para a espera (ex.: elemento visível, clicável).
from selenium.common.exceptions import TimeoutException # Exceção disparada quando uma espera (WebDriverWait) atinge o tempo limite.
from selenium.common.exceptions import NoSuchElementException, NoSuchWindowException, ElementClickInterceptedException, StaleElementReferenceException # Exceção disparada quando um elemento não é encontrado na página.
from selenium.common.exceptions import InvalidSessionIdException, WebDriverException # Exceções gerais relacionadas a problemas de comunicação ou estado da sessão do driver.
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.chrome.service import Service as ChromeService

import atexit
import unicodedata
from difflib import get_close_matches
from typing import Iterable, Union
import openpyxl
import subprocess
import platform, shutil
from urllib.parse import urlparse


# # Configuração de Opções (Mantida como estava no original para contexto, mas movida para o final)
# options = Options()
# options.add_argument('--disable-backgrounding-occluded-windows')  # Impede que abas em 2º plano sejam pausadas
# options.add_argument('--no-sandbox')
# options.add_experimental_option("detach", True)  # Evita que a aba feche com o script

#=======================================================================================================================

arch = platform.machine().lower()  # 'arm64' ou 'x86_64'
os.environ["UC_USER_DATA_DIR"] = os.path.expanduser(f"~/Library/Application Support/undetected_chromedriver_{arch}")
os.environ["UC_DATA_DIR"] = os.environ["UC_USER_DATA_DIR"]

# base_uc = Path(os.path.expanduser("~/Library/Application Support"))
# # Alguns UC usam nome com sufixo, outros sem
# candidate1 = base_uc / f"undetected_chromedriver_{arch}"
# candidate2 = base_uc / "undetected_chromedriver"

# if candidate1.exists():
#     uc_cache = candidate1
# else:
#     uc_cache = candidate2 



# CONFIGURAÇÃO - ajuste conforme seu ambiente
#PROFILE_DIR = r"C:/selenium/chrome-profile"   # seu user-data-dir

# PROFILE_DIR = str((Path.home() / ".selenium" / "chrome-profile").resolve())
# USER_AGENT = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
#             "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/142 Safari/537.36")


PROFILE_DIR = os.path.expanduser("~/Library/Application Support/SelecionaCotas_Profile")

# FORÇAR tudo para o mesmo diretório
os.environ["UC_USER_DATA_DIR"] = PROFILE_DIR
os.environ["UC_DATA_DIR"] = PROFILE_DIR


#USER_AGENT = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
#             "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/142 Safari/537.36")


#COOKIES_FILE = Path("cookies_saved.json")


#START_URL = "https://canal360i.cloud.itau.com.br/login/iparceiros"   # verifique se abrirá direto no login tem problema...

# APP_DATA = Path.home() / ".selecao_cotas"
# APP_DATA.mkdir(parents=True, exist_ok=True)

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
#root = None
driver_started = False
driver_thread = None
_driver_keepalive_evt = threading.Event()





def ui_alert(title, msg, kind="info"):
    def _show():
        if 'root' not in globals() or not root or not root.winfo_exists():
            print(f"[ALERTA] {title}: {msg}")
            return
        try:
            root.lift(); root.focus_force()
        except: pass
        top = tk.Toplevel(root); top.withdraw(); top.transient(root)
        try: top.attributes("-topmost", True)
        except: pass
        top.update_idletasks(); top.deiconify(); top.lift(); top.focus_force()
        try:
            if kind == "warning":   messagebox.showwarning(title, msg, parent=top)
            elif kind == "error":   messagebox.showerror(title, msg, parent=top)
            else:                   messagebox.showinfo(title, msg, parent=top)
        finally:
            try: top.destroy()
            except: pass
    try:
        root.after(0, _show)
    except Exception:
        print(f"[ALERTA] {title}: {msg}")





def on_quit():
    # Pare threads/driver aqui
    try:
        if driver: driver.quit()
    except: pass
    try: root.quit()
    except: pass
    try: root.destroy()
    except: pass






# ====== MODAL (para erros/confirm) ======
def show_modal(title, message, kind="info"):
    """Janela modal simples; bloqueia até clicar OK."""
    def _open():
        win = tk.Toplevel(root)
        win.title(title)
        win.transient(root)
        win.grab_set()
        win.configure(bg="#0C1B39")
        win.resizable(False, False)

        # estética básica
        pad = {"padx": 16, "pady": 10}
        lbl = tk.Label(win, text=message, fg="white", bg="#0C1B39",
                       font=("Helvetica", 12))
        lbl.pack(**pad)

        btn = ttk.Button(win, text="OK", command=win.destroy)
        btn.pack(pady=(0, 14))
        win.update_idletasks()

        # centraliza no root
        rw, rh = root.winfo_width(), root.winfo_height()
        rx, ry = root.winfo_rootx(), root.winfo_rooty()
        ww, wh = win.winfo_width(), win.winfo_height()
        win.geometry(f"{ww}x{wh}+{rx + (rw-ww)//2}+{ry + (rh-wh)//2}")

        win.lift()
        win.focus_force()
        root.wait_window(win)
    root.after(0, _open)

#=======================================================================================================================
# Funções auxiliares
#=======================================================================================================================

def driver_ativo(drv):
    try:
        _ = drv.current_url  # dispara se o driver já morreu
        return True
    except:
        return False


def _get_app_path() -> Path:
    """
    Retorna o diretório base do executável/script, lidando com 
    ambientes interativos (como Jupyter) e PyInstaller.
    """
    
    # 1. Caso PyInstaller (Executável Empacotado)
    if getattr(sys, 'frozen', False):
        # Em Mac, tenta encontrar a pasta que contém o .app
        if sys.platform == 'darwin':
            current_path = Path(sys.executable)
            for parent in current_path.parents:
                if parent.suffix == '.app':
                    # Retorna o diretório ao lado do .app
                    return parent.parent 
        
        # Para outros sistemas ou fallback, usa o diretório do executável
        return Path(sys.executable).parent

    # 2. Caso Script Normal ou Ambiente Interativo
    try:
        # Tenta usar __file__ (funciona se for um script normal)
        # O 'globals()' aqui é apenas para verificar a variável em certos contextos, 
        # mas a simples tentativa de acesso é o suficiente.
        return Path(__file__).resolve().parent 
    except NameError:
        # Ocorre em ambientes interativos (Jupyter, iPython, etc.)
        # Retorna o Diretório de Trabalho Atual como fallback
        return Path.cwd() 


def _find_user_file(name: str) -> Path:
    """
    Procura o arquivo do usuário SOMENTE na mesma pasta do executável/script.
    Aceita .xlsx e .csv.
    Se 'name' vier sem extensão, tenta ambos.
    """
    app_dir = _get_app_path()  # assumindo que você já tem essa função
    base = Path(name)

    # lista de candidatos a testar
    candidates = []

    if base.suffix:  
        # o usuário já passou uma extensão
        if base.suffix.lower() not in ('.xlsx', '.csv'):
            raise FileNotFoundError(
                f"Extensão '{base.suffix}' não suportada. Use .xlsx ou .csv."
            )
        candidates.append(app_dir / base.name)
    else:
        # sem extensão: tenta ambos
        candidates.append(app_dir / (base.name + '.xlsx'))
        candidates.append(app_dir / (base.name + '.csv'))

    # debug opcional
    print("Procurando em:", app_dir)
    for c in candidates:
        print("Testando:", c)

    # procura o primeiro que existir e seja suportado
    for c in candidates:
        if c.exists():
            # segurança extra: só aceita .xlsx ou .csv mesmo
            if c.suffix.lower() in ('.xlsx', '.csv'):
                return c
            else:
                # isso aqui teoricamente não dispara, mas deixei por robustez
                raise FileNotFoundError(
                    f"Arquivo '{c.name}' encontrado, mas extensão não suportada ({c.suffix})."
                )

    # se nenhum encontrado:
    raise FileNotFoundError(
        "Arquivo não encontrado. O arquivo deve estar na mesma pasta do executável:\n"
        f"- {str(app_dir)}\n"
        "Nomes tentados:\n  " + "\n  ".join(str(c) for c in candidates)
    )


# Base para arquivos do projeto / recursos (drivers, etc.)
APP_BASE = _get_app_path()

# Base para dados gravados pelo app (logs etc.)
if getattr(sys, "frozen", False):
    # .app rodando no cliente → usa pasta no HOME (evita permissão em /Applications)
    APP_DATA = Path.home() / "SelecionaCotasLogs"
else:
    # rodando .py → usa a pasta do projeto
    APP_DATA = APP_BASE

APP_DATA.mkdir(parents=True, exist_ok=True)

LOG_FILE = APP_DATA / "driver_log.txt"

def log(msg: str):
    try:
        with LOG_FILE.open("a", encoding="utf-8") as f:
            f.write(msg + "\n")
    except Exception:
        pass




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
    try:
        cookies = driver.get_cookies()
    except Exception as e:
        print("Falha ao obter cookies do driver:", e)
        return False

    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cookies, f, indent=2, ensure_ascii=False)
    print(f"Cookies salvos em {path}")
    return True


# carregar cookies de arquivo. O driver precisa estar em uma página do mesmo domínio
def load_cookies(driver, path: Path, domain_url: str = "https://accounts.google.com"):
    """
    - driver: webdriver já criado
    - path: Path do arquivo JSON de cookies
    - domain_url: URL que será aberta para permitir adicionar cookies (domínio base)
    """
    if not path.exists():
        print("Arquivo de cookies não existe:", path)
        return False

    # navega para a origem do cookie (necessário para add_cookie)
    try:
        driver.get(domain_url)
    except WebDriverException as e:
        print("Falha ao navegar para domínio base antes de carregar cookies:", e)
        return False

    with open(path, "r", encoding="utf-8") as f:
        cookies = json.load(f)

    added = 0
    for ck in cookies:
        # sanitize do cookie para evitar exceções
        safe = {}
        # campos permitidos pelo selenium: name, value, path, domain, secure, httpOnly, expiry
        safe["name"] = ck.get("name")
        safe["value"] = ck.get("value")
        if ck.get("path"):
            safe["path"] = ck.get("path")
        # expiry: muitas libs gravam float; selenium espera int (ou não enviar)
        expiry = ck.get("expiry") or ck.get("expires")  # alguns formatos
        if expiry:
            try:
                safe["expiry"] = int(expiry)
            except Exception:
                # descarta expiry se estiver inválido
                pass
        if ck.get("secure") is not None:
            safe["secure"] = bool(ck.get("secure"))
        if ck.get("httpOnly") is not None:
            safe["httpOnly"] = bool(ck.get("httpOnly"))

        # NÃO envie sameSite (pode quebrar em algumas versões do selenium)
        # domain: se cookie é para .google.com e estamos em accounts.google.com, ok.
        # se o add_cookie falhar por domain, tentamos novamente sem o campo domain.
        if "domain" in ck and ck.get("domain"):
            safe["domain"] = ck.get("domain")

        try:
            driver.add_cookie(safe)
            added += 1
        except Exception as e:
            # tentativa de fallback: remover domain e expiry e tentar novamente
            fallback = {k: v for k, v in safe.items() if k not in ("domain", "expiry")}
            try:
                driver.add_cookie(fallback)
                added += 1
            except Exception as e2:
                # ignora cookie problemático, mas loga aviso
                print("Warning: cookie add failed:", safe.get("name"), "->", e2)

    # depois de adicionar cookies, recarrega para que tenham efeito
    try:
        driver.refresh()
    except Exception:
        pass

    print(f"Cookies carregados de {path} (adicionados: {added})")
    return added > 0



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
        _driver_keepalive_evt.set()
    except: pass
    try:
        if driver: driver.quit()
    except: pass
    try:
        root.quit()   # sai do loop (não destrói a janela ainda)
    except: pass
    # root.destroy() será chamado pelo on_quit() quando for finalizar de vez


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

def _normalize_tipo(s: str) -> str:
    s = (s or "").strip().lower()
    s = _strip_accents(s)                     # remove acentos
    s = re.sub(r'[^\w\s-]', '', s)            # remove pontuação estranha
    return s.replace(' ', '_')                # troca espaço por _


def _resolve_grupos_csv(tp_produto: str) -> Path | None:
    base = _normalize_tipo(tp_produto)
    fname = f"grupos_{base}.csv"
    candidatos = [
        APP_DATA / fname,                     # onde você salva na função guardar_grupos_disponiveis
        _executable_dir() / fname,            # ao lado do executável (PyInstaller)
        Path.cwd() / fname,                   # diretório atual
    ]
    # fallback legados (se já existirem com acento/sem normalizar)
    legacy = [
        APP_DATA / f"grupos_{tp_produto.replace(' ', '_')}.csv",
        _executable_dir() / f"grupos_{tp_produto.replace(' ', '_')}.csv",
        Path.cwd() / f"grupos_{tp_produto.replace(' ', '_')}.csv",
    ]
    for p in candidatos + legacy:
        if p.exists():
            return p
    return None


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
    Retorna a pasta base do executável/script de forma robusta:
    - App empacotado (PyInstaller): pasta do executável
    - Script normal: pasta do arquivo principal
    - Jupyter/REPL: diretório atual
    """
    # PyInstaller (onefile/onedir)
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent

    # Script normal (quando __main__.__file__ existe)
    main_file = getattr(sys.modules.get("__main__"), "__file__", None)
    if main_file:
        return Path(main_file).resolve().parent

    # Fallback 1: caminho passado na linha de comando
    if sys.argv and sys.argv[0]:
        try:
            return Path(sys.argv[0]).resolve().parent
        except Exception:
            pass

    # Fallback 2: ambiente interativo (Jupyter/REPL)
    return Path.cwd()




def find_data_file(name: str) -> Path:
    exe_dir = _executable_dir()
    # se estiver dentro de um .app, subir 3 níveis chega na raiz do app; e subir 4 pode chegar no dist/
    app_root = exe_dir.parent.parent.parent if ".app" in str(exe_dir) else None
    dist_root = app_root.parent if app_root else None

    candidates = [
        Path.cwd(),            # onde o usuário executou (terminal)
        exe_dir,               # ao lado do binário (Contents/MacOS)
        app_root,              # raiz do .app
        dist_root,             # pasta dist/ (um nível acima da raiz do .app)
        Path.home() / ".selecao_cotas",  # pasta de dados do usuário
    ]

    for base in filter(None, candidates):
        p = base / name
        if p.exists():
            return p

    tried = "\n".join(str((c or Path()).resolve() / name) for c in candidates if c)
    raise FileNotFoundError(
        f"Arquivo '{name}' não encontrado. Procurei em:\n{tried}\n"
        "Coloque-o ao lado do executável (.app/Contents/MacOS), na raiz do .app, "
        "na pasta dist/ ou em ~/.selecao_cotas."
    )



def load_df_clientes():
    global df_atual, df_clientes
    try:
        arq = _find_user_file('base_clientes.xlsx')
    except FileNotFoundError as e:
        ui_alert(
            "Arquivo não encontrado",
            "Não localizei o arquivo 'base_clientes.xlsx' na mesma pasta do executável.\n"
            "Coloque a planilha na pasta e tente novamente.",
            "error"
        )
        # Levanta de novo pra quem chamou saber que não tem como continuar
        raise RuntimeError("Planilha base_clientes.xlsx ausente.") from e

    # garante engine disponível
    try:
        import openpyxl
    except ImportError as e:
        ui_alert(
            "Dependência ausente",
            "A biblioteca 'openpyxl' não está disponível. Reinstale a aplicação.",
            "error"
        )
        raise RuntimeError("Dependência 'openpyxl' ausente.") from e

    # tenta ler
    df_clientes = pd.read_excel(arq, dtype={'cpf_cnpj': str}, engine='openpyxl')
    if 'cpf_cnpj' in df_clientes.columns:
        linhas_antes = len(df_clientes)
        df_clientes.drop_duplicates(subset=['cpf_cnpj'], keep='first', inplace=True)
        linhas_depois = len(df_clientes)
        if linhas_antes > linhas_depois:
            print(f"Alerta: {linhas_antes - linhas_depois} linhas duplicadas de CPF/CNPJ removidas.")

    # valida coluna obrigatória
    if 'status' not in df_clientes.columns:
        ui_alert("Erro", "Coluna obrigatória 'status' não encontrada na planilha!", "error")
        raise RuntimeError("A planilha não possui a coluna obrigatória 'status'.")

    # filtra pendentes
    df_atual = df_clientes[df_clientes['status'] == "Pendente"].copy()

    return df_clientes



def _atomic_replace(src_tmp: Path, dst: Path):
    # troca o arquivo de forma segura (evita corrupção se der crash no meio)
    src_tmp.replace(dst)


def atualizar_status_cliente(cpf_cliente, novo_status, caminho='base_clientes.xlsx'):
    try:
        # usa a MESMA busca da leitura (pasta do exe, CWD, raiz do .app, pasta do .py)
        arq = find_data_file(caminho)
    except FileNotFoundError as e:
        print(f"❌ Erro: {e}")
        return

    # tenta ler Excel; se faltar openpyxl, cai pro CSV irmão
    is_excel = True
    try:
        df = pd.read_excel(arq, dtype={'cpf_cnpj': str}, engine='openpyxl')
    except Exception as e:
        print("⚠️ Falha ao ler Excel; tentando CSV irmão:", e)
        is_excel = False
        arq_csv = arq.with_suffix('.csv')
        if not arq_csv.exists():
            print(f"❌ CSV irmão não encontrado: {arq_csv.name}")
            return
        df = pd.read_csv(arq_csv, sep=';', dtype={'cpf_cnpj': str})

    alvo = re.sub(r'\D', '', str(cpf_cliente))
    col = df['cpf_cnpj'].astype(str).str.replace(r'\D', '', regex=True)

    if alvo not in col.values:
        print(f"⚠️ CPF/CNPJ {cpf_cliente} não encontrado. Nada feito.")
        return

    df.loc[col == alvo, 'status'] = novo_status

    try:
        if is_excel:
            # salva de forma atômica na MESMA pasta do arquivo localizado
            tmp = arq.with_suffix('.tmp.xlsx')
            df.to_excel(tmp, index=False, engine='openpyxl')
            _atomic_replace(tmp, arq)
        else:
            tmp = arq.with_suffix('.tmp.csv')
            df.to_csv(tmp, sep=';', index=False)
            _atomic_replace(tmp, arq.with_suffix('.csv'))
        print(f"✅ Status de {cpf_cliente} atualizado para '{novo_status}' em {arq.name}.")
    except PermissionError as e:
        print(f"❌ Permissão negada ao salvar (arquivo pode estar aberto no Excel): {e}")
    except Exception as e:
        print(f"❌ Falha ao salvar: {e}")


def _ensure_active_window(driver, wait_secs=5):
    """Garante que há ao menos 1 janela e foca na última."""
    t0 = time.time()
    while time.time() - t0 < wait_secs:
        try:
            handles = driver.window_handles
            if handles:
                try:
                    driver.switch_to.window(handles[-1])
                    # tocar no título força o alvo estar OK
                    _ = driver.title
                    return True
                except NoSuchWindowException:
                    # tenta o primeiro handle disponível
                    for h in handles:
                        try:
                            driver.switch_to.window(h)
                            _ = driver.title
                            return True
                        except NoSuchWindowException:
                            continue
        except WebDriverException:
            pass
        time.sleep(0.1)
    return False

def _safe_get(driver, url, retries=1):
    """Faz driver.get com recuperação de janela em caso de NoSuchWindow."""
    try:
        if not _ensure_active_window(driver, wait_secs=5):
            raise NoSuchWindowException("Nenhuma janela ativa para navegar.")
        driver.get(url)
        return True
    except NoSuchWindowException:
        if retries > 0:
            # tenta recuperar e repetir
            if _ensure_active_window(driver, wait_secs=3):
                return _safe_get(driver, url, retries=retries-1)
        raise




#=======================================================================================================================
#                  FUNÇÃO 0 - Iniciar Driver com as configurações iniciais
#=======================================================================================================================
# Função INICIAL - iniciar driver com perfil, user-agent e stealth


def get_chrome_major_macos(bin_path="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"):
    try:
        out = subprocess.check_output([bin_path, "--version"], text=True).strip()
        m = re.search(r"(\d+)\.", out)
        return int(m.group(1)) if m else None
    except Exception:
        return None

def harden_network(driver, block_fonts=False):
    """
    Bloqueia trackers/ads comuns e (opcional) webfonts.
    Use block_fonts=True só se a UI não depender de ícones via fonte.
    """
    try:
        driver.execute_cdp_cmd("Network.enable", {})
        urls = [
            "*googletagmanager.com/*", "*google-analytics.com/*",
            "*doubleclick.net/*", "*hotjar.com/*", "*facebook.net/*",
            "*segment.io/*", "*fullstory.com/*", "*intercomcdn.com/*",
        ]
        if block_fonts:
            urls += ["*.woff", "*.woff2", "*.ttf"]
        driver.execute_cdp_cmd("Network.setBlockedURLs", {"urls": urls})
    except Exception:
        # Se falhar (p.ex. versão Chrome/CDP), só segue sem bloquear
        pass


def tune_chrome_runtime(driver):
    """Desativa o tradutor, gestor de senha, etc., e esconde a barra de 'automation'."""
    try:
        # prefs de runtime (algumas coisas precisam ser via options, outras sobrevivem via CDP)
        driver.execute_cdp_cmd(
            "Browser.setDownloadBehavior",
            {"behavior": "allow", "downloadPath": "/tmp"}
        )
    except Exception:
        pass

def build_chrome_options() -> uc.ChromeOptions:
    options = uc.ChromeOptions()
    options.add_argument(f"--user-data-dir={PROFILE_DIR}")
    options.add_argument("--profile-directory=Default")
    options.add_argument("--lang=pt-BR")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-background-timer-throttling")
    options.add_argument("--disable-backgrounding-occluded-windows")
    options.add_argument("--disable-renderer-backgrounding")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-default-apps")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disk-cache-size=104857600")
    options.binary_location = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"

    prefs = {
        "translate_whitelists": {},
        "translate": {"enabled": False},
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False,
    }
    options.add_experimental_option("prefs", prefs)
    options.set_capability("pageLoadStrategy", "eager")

    return options

def _pos_configurar_driver(drv):
    """
    Configura wait, action, janela, garante que há janela ativa,
    carrega cookies e abre google.com.
    """
    global driver, wait, action, driver_started

    driver = drv
    driver_started = True
    wait = WebDriverWait(driver, 15)
    action = ActionChains(driver)

    try:
        driver.set_window_size(1199, 889)
    except Exception:
        pass

    if not _ensure_active_window(driver, wait_secs=5):
        raise RuntimeError("Chrome iniciou sem janelas ativas (UC).")

    # cookies + primeira navegação
    try:
        load_cookies(driver, Path("cookies_saved.json"),
                     domain_url="https://accounts.google.com")
    except Exception:
        # se falhar, não mata o driver
        print("Aviso: falha ao carregar cookies_saved.json", file=sys.stderr)

    _safe_get(driver, "https://www.google.com", retries=1)
    print("Título inicial:", driver.title)
    return driver


if getattr(sys, "frozen", False):
    RES_BASE = Path(sys._MEIPASS)
else:
    RES_BASE = APP_BASE

FALLBACK_ARM_DRIVER = RES_BASE / "drivers" / "chromedriver-arm64"


def _tentar_fallback_arm():
    if not FALLBACK_ARM_DRIVER.exists():
        raise RuntimeError(
            f"Fallback ARM driver não encontrado em: {FALLBACK_ARM_DRIVER}"
        )

    log(f"Usando fallback ARM em {FALLBACK_ARM_DRIVER}")
    print("Iniciando fallback com chromedriver ARM fixo:", FALLBACK_ARM_DRIVER)

    options = build_chrome_options()
    service = ChromeService(executable_path=str(FALLBACK_ARM_DRIVER))

    drv = webdriver.Chrome(service=service, options=options)
    return _pos_configurar_driver(drv)



#############
#################
######################  ____  INICIAR CHROME:

def iniciar_driver(max_retries_uc: int = 2):
    """
    Inicializa o Chrome com undetected_chromedriver de forma robusta:
    - tenta iniciar normalmente
    - se der 'Bad CPU type', apaga o driver/caches e tenta de novo
    - se continuar falhando, tenta fallback ARM fixo (se existir)
    """
    global driver, wait, action, driver_started

    
    log(f"tentando iniciar driver")

    arch = platform.machine().lower()  # esperado 'arm64'

    base_uc = Path(os.path.expanduser("~/Library/Application Support"))
    candidate1 = base_uc / f"undetected_chromedriver_{arch}"
    candidate2 = base_uc / "undetected_chromedriver"

    if candidate1.exists():
        uc_cache = candidate1
    else:
        uc_cache = candidate2
    
    def _tentar_uc(tag_tentativa: str):
        options = build_chrome_options()
        major = get_chrome_major_macos(options.binary_location)

        print(f"[{tag_tentativa}] Chrome major detectado:", major)
        print(f"[{tag_tentativa}] Iniciando UC com cache: {uc_cache} (arch={arch})")

        drv = uc.Chrome(
            options=options,
            version_main=major,
            use_subprocess=False,
            headless=False,
        )
        return _pos_configurar_driver(drv)

    tentativas = 0
    ultimo_erro = None

    # ---------- TENTATIVAS COM UC (com limpeza de Bad CPU) ----------
    while tentativas < max_retries_uc:
        tag = f"UC tentativa {tentativas + 1}"
        try:
            driver = _tentar_uc(tag)
            log(f"driver iniciado")
            return driver

        except Exception as exc:
            ultimo_erro = exc
            msg = str(exc)
            
            log(f"{tag} falhou: {msg}")
            log(traceback.format_exc())

            print(f"{tag} falhou:", msg)
            traceback.print_exc(file=sys.stdout)

            if "Bad CPU type" in msg:
                print("Detectado 'Bad CPU type' — limpando driver/caches incompatíveis.")

                # 1) tentar pegar caminho do executável quebrado
                m = re.search(r"executable:\s*'([^']+)'", msg)
                if m:
                    bad_exec = Path(m.group(1))
                    bad_dir = bad_exec.parent
                    print("Removendo diretório do driver incompatível:", bad_dir)
                    shutil.rmtree(bad_dir, ignore_errors=True)

                # 2) limpar também o cache por arquitetura
                if uc_cache.exists():
                    print("Removendo cache UC:", uc_cache)
                    shutil.rmtree(uc_cache, ignore_errors=True)

                tentativas += 1
                continue  # tenta de novo com NOVO options

            # se não é Bad CPU, repetir igual não adianta
            break

    # ---------- GUARD RAIL: FALLBACK ARM FIXO ----------
    print("UC falhou após tentativas configuradas. Tentando fallback ARM fixo...")
    try:
        driver = _tentar_fallback_arm()
        return driver
    except Exception as fallback_exc:
        log(f"{tag} falhou: {msg}")
        log(traceback.format_exc())

        print("Fallback ARM também falhou:", fallback_exc)
        traceback.print_exc(file=sys.stdout)

        # Monta mensagem final amigável
        msg_final = (
            "Falha ao iniciar o ChromeDriver após todas as tentativas.\n\n"
            "Possíveis causas:\n"
            "- Chrome instalado para arquitetura diferente (Intel x ARM);\n"
            "- Rosetta não instalada (para Chrome Intel);\n"
            "- Fallback ARM não empacotado corretamente.\n\n"
            "Soluções comuns para o usuário:\n"
            "1) Certificar que o Google Chrome instalado é a versão para Apple Silicon (ARM);\n"
            "2) Se usar Chrome Intel, instalar o Rosetta;\n"
            "3) Reabrir o app após essas mudanças."
        )
        raise RuntimeError(msg_final) from (fallback_exc or ultimo_erro)

#old 17/11/25
# def iniciar_driver():
#     global driver, wait, action, driver_started

#     arch = platform.machine().lower()
#     uc_cache = os.path.expanduser(
#         f"~/Library/Application Support/undetected_chromedriver_{arch}"
#     )

#     options = build_chrome_options()
#     major = get_chrome_major_macos(options.binary_location)
#     print("Chrome major detectado:", major)

#     try:
#         print(f"Iniciando UC com cache: {uc_cache} (arch={arch})")
#         driver = uc.Chrome(
#             options=options,
#             version_main=major,
#             use_subprocess=False,
#             headless=False,
#         )

#         driver_started = True
#         wait = WebDriverWait(driver, 15)
#         action = ActionChains(driver)

#         try:
#             driver.set_window_size(1199, 889)
#         except Exception:
#             pass

#         if not _ensure_active_window(driver, wait_secs=5):
#             raise RuntimeError("Chrome iniciou sem janelas ativas (UC).")

#         load_cookies(driver, Path("cookies_saved.json"),
#                      domain_url="https://accounts.google.com")

#         _safe_get(driver, "https://www.google.com", retries=1)
#         print("Título inicial:", driver.title)
#         return driver

#     except Exception as exc:
#         print("Erro ao criar/iniciar driver:")
#         traceback.print_exc(file=sys.stdout)

#         if "Bad CPU type" in str(exc):
#             try:
#                 print("Limpando cache de UC desta arquitetura e tentando novamente:", uc_cache)
#                 shutil.rmtree(uc_cache, ignore_errors=True)
#             except Exception:
#                 pass

#             # 👉 CRIA **NOVO** ChromeOptions PARA A SEGUNDA TENTATIVA
#             options2 = build_chrome_options()
#             driver = uc.Chrome(
#                 options=options2,
#                 version_main=major,
#                 use_subprocess=False,
#                 headless=False,
#             )
#             driver_started = True
#             wait = WebDriverWait(driver, 15)
#             action = ActionChains(driver)
#             driver.get("https://www.google.com")
#             print("Título inicial:", driver.title)
#             return driver

#         # se não era "Bad CPU type", propaga o erro original
#         raise


#old 14/nov/25
# def iniciar_driver():
#     global driver, wait, action, driver_started

#     arch = platform.machine().lower()
#     uc_cache = os.path.expanduser(f"~/Library/Application Support/undetected_chromedriver_{arch}")
#     # os.environ["UC_USER_DATA_DIR"] = uc_cache
#     # os.environ["UC_DATA_DIR"] = uc_cache

#     options = uc.ChromeOptions()
#     options.add_argument(f"--user-data-dir={PROFILE_DIR}")
#     #options.add_argument(f"--user-agent={USER_AGENT}")
#     options.add_argument("--profile-directory=Default") #NEW
#     options.add_argument("--lang=pt-BR")
#     options.add_argument("--no-sandbox")
#     options.add_argument("--disable-dev-shm-usage")
#     options.add_argument("--disable-infobars")
#     options.add_argument("--disable-background-timer-throttling")
#     options.add_argument("--disable-backgrounding-occluded-windows")
#     options.add_argument("--disable-renderer-backgrounding")
#     options.add_argument("--disable-notifications")
#     options.add_argument("--disable-default-apps")
#     options.add_argument("--start-maximized")
#     options.add_argument("--disable-blink-features=AutomationControlled")
#     # ⚠️ mantenha cache pequeno/saudável (NÃO zere):
#     options.add_argument("--disk-cache-size=104857600")
#     options.binary_location = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"

#     # ⚠️ REMOVA estas duas linhas problemáticas:
#     # options.add_experimental_option("excludeSwitches", ["enable-automation"])
#     # options.add_experimental_option("useAutomationExtension", False)

#     # Prefs úteis (ok manter)
#     prefs = {
#         "translate_whitelists": {},
#         "translate": {"enabled": False},
#         "credentials_enable_service": False,
#         "profile.password_manager_enabled": False,
#     }
#     options.add_experimental_option("prefs", prefs)

#     # Use a forma moderna de setar pageLoadStrategy (sem desired_capabilities):
#     options.set_capability("pageLoadStrategy", "eager")

#     major = get_chrome_major_macos(options.binary_location)
#     print("Chrome major detectado:", major)

#     try:
#         print(f"Iniciando UC com cache: {uc_cache} (arch={arch})")

#         driver = uc.Chrome(
#             options=options,
#             version_main=major,
#             use_subprocess=False,   
#             headless=False
#         )

#         driver_started = True
#         wait = WebDriverWait(driver, 15)
#         action = ActionChains(driver)

#         try:
#             driver.set_window_size(1199, 889)
#         except Exception:
#             pass

#         # (Opcional) seus boosts por CDP:
#         # harden_network(driver, block_fonts=False)
#         # tune_chrome_runtime(driver)

#         if not _ensure_active_window(driver, wait_secs=5):
#             raise RuntimeError("Chrome iniciou sem janelas ativas (UC).")

#         load_cookies(driver, Path("cookies_saved.json"), domain_url="https://accounts.google.com")

#         _safe_get(driver, "https://www.google.com", retries=1)
#         print("Título inicial:", driver.title)
#         return driver

#     except Exception as exc:
#         print("Erro ao criar/iniciar driver:")
#         traceback.print_exc(file=sys.stdout)

#         if "Bad CPU type" in f"{exc}":
#             try:
#                 print("Limpando cache de UC desta arquitetura e tentando novamente:", uc_cache)
#                 shutil.rmtree(uc_cache, ignore_errors=True)
#             except Exception:
#                 pass
#             driver = uc.Chrome(
#                 options=options,
#                 version_main=major,
#                 use_subprocess=False,
#                 headless=False
#             )
#             driver_started = True
#             wait = WebDriverWait(driver, 15)
#             action = ActionChains(driver)
#             driver.get("https://www.google.com")
#             print("Título inicial:", driver.title)
#             return driver

#         raise





#############
#################
###################### ===== Helpers rápidos =====

#aux
def _normalize_text(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = unicodedata.normalize("NFD", s)
    return "".join(ch for ch in s if not unicodedata.combining(ch))

def _to_list(x: Union[str, Iterable[str], None]) -> list[str]:
    if x is None:
        return []
    if isinstance(x, str):
        return [x]
    return list(x)

def ensure_tab_with_title(
    driver,
    desired_titles: Union[str, Iterable[str]],
    timeout: float = 6,
    try_url_contains: Union[str, Iterable[str], None] = None
) -> bool:
    """
    Garante que o driver esteja numa aba cujo título contenha ALGUM dos 'desired_titles'.
    - desired_titles: string ou lista de strings (ex.: ["simulação e contratação", "plataforma 360i"])
    - try_url_contains: string ou lista de strings para fallback por URL
    Retorna True se encontrar e fizer o switch; False caso contrário.
    """
    if driver is None:
        raise RuntimeError("WebDriver é None.")

    needles = [_normalize_text(t) for t in _to_list(desired_titles) if t]
    if not needles:
        raise ValueError("Parâmetro 'desired_titles' não pode ser vazio.")

    # 1) garantir que existam janelas
    try:
        handles = driver.window_handles
    except WebDriverException:
        raise RuntimeError("WebDriver inválido / processo do navegador encerrou.")

    if not handles:
        raise RuntimeError("Nenhuma aba/janela encontrada no navegador.")

    def title_matches(s: str) -> bool:
        nt = _normalize_text(s)
        return any(n in nt for n in needles)

    # 2) fast check na aba atual
    try:
        if title_matches(driver.title):
            return True
    except Exception:
        pass

    # 3) varrer todas as abas e checar título
    for h in handles:
        try:
            driver.switch_to.window(h)
            end = time.time() + 0.11
            while time.time() < end:
                t = driver.title
                if t is not None:
                    break
                time.sleep(0.03)
            if title_matches(driver.title):
                return True
        except NoSuchWindowException:
            continue

    # 4) fallback por URL (SPA etc.)
    url_needles = [u.lower().strip() for u in _to_list(try_url_contains) if u]
    if url_needles:
        for h in driver.window_handles:
            try:
                driver.switch_to.window(h)
                u = (driver.current_url or "").lower()
                if any(k in u for k in url_needles):
                    # dá uma chance do title atualizar
                    try:
                        WebDriverWait(driver, timeout/2).until(
                            lambda d: title_matches(d.title) or True
                        )
                    except Exception:
                        pass
                    return True
            except NoSuchWindowException:
                continue

    # 5) última tentativa: esperar um pouco pelo título em cada aba
    for h in driver.window_handles:
        try:
            driver.switch_to.window(h)
            try:
                WebDriverWait(driver, timeout).until(lambda d: title_matches(d.title))
                return True
            except TimeoutException:
                continue
        except NoSuchWindowException:
            continue

    return False



def js_click(driver, element):
    driver.execute_script("arguments[0].click();", element)

def js_query_all_text(driver, css):
    return driver.execute_script(
        "return Array.from(document.querySelectorAll(arguments[0])).map(e => (e.textContent||'').trim());",
        css
    )


def js_query_value(driver, css):
    return driver.execute_script(
        "const e=document.querySelector(arguments[0]); return e? (e.value||'') : null;", css
    )


def esperar_sumir_loading(driver, timeout=12):
    try:
        W(driver, timeout).until(EC.invisibility_of_element_located(
            (By.CSS_SELECTOR, ".loading,.spinner,.is-loading")))
    except Exception:
        pass  # se não existir spinner, beleza


def get_valor_maximo_float(driver):
    # Lê por JS para reduzir roundtrips
    txt = driver.execute_script("""
        const h5 = document.querySelector('.valores-min-max h5');
        return h5 ? (h5.textContent||'').trim() : '';
    """)
    m = re.search(r'R\$[\s]*([\d\.,]+)', txt or '')
    if not m:
        return None
    return float(m.group(1).replace('.', '').replace(',', '.'))




def _js_is_in_viewport(driver, el):
    return driver.execute_script("""
        const el = arguments[0];
        if (!el) return false;
        const r = el.getBoundingClientRect();
        const vw = (window.innerWidth || document.documentElement.clientWidth);
        const vh = (window.innerHeight || document.documentElement.clientHeight);
        // consideramos visível se centro está no viewport
        const cx = r.left + r.width/2;
        const cy = r.top  + r.height/2;
        return cx >= 0 && cy >= 0 && cx <= vw && cy <= vh;
    """, el) or False

def _js_scroll_to_center(driver, el):
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center', inline:'center', behavior:'instant'});", el)
        # se houver header fixo, compensa 80px para cima
        driver.execute_script("window.scrollBy(0, -80);")
    except Exception:
        pass

def _wait_overlay_option_50(driver, timeout=5):
    """
    Espera a opção '50' no overlay (portais). Retorna o elemento <span> ou None.
    """
    end = time.time() + timeout
    while time.time() < end:
        try:
            # Muitas libs geram opções fora do combobox, direto no body:
            opts = driver.find_elements(By.XPATH, '//span[contains(@class,"ids-option__text") and normalize-space()="50"]')
            if opts:
                return opts[0]
        except StaleElementReferenceException:
            pass
        time.sleep(0.05)
    return None

def set_page_size_50(driver, max_attempts=3, after_click_wait_spinner=True):
    """
    Torna o seletor de page size visível, abre o dropdown e escolhe '50'.
    Tenta algumas vezes, rolando e garantindo overlay carregado.
    """
    # 1) garantir presença do combobox
    combo = W(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'div[role="combobox"]#pageSizeId'))
    )

    for attempt in range(1, max_attempts+1):
        try:
            # 2) garantir visibilidade no viewport
            if not _js_is_in_viewport(driver, combo):
                _js_scroll_to_center(driver, combo)
                # pequena folga para layout estabilizar
                time.sleep(0.05)

            # 3) abrir o dropdown (JS click não depende de estar perfeitamente “clickable”)
            driver.execute_script("arguments[0].click();", combo)

            # 4) esperar a opção 50 aparecer no overlay
            opt_50 = _wait_overlay_option_50(driver, timeout=3)
            if not opt_50:
                # algumas UIs precisam de um 2º clique para abrir
                driver.execute_script("arguments[0].click();", combo)
                opt_50 = _wait_overlay_option_50(driver, timeout=3)

            if not opt_50:
                # ainda não abriu? rola de novo e tenta próxima volta
                _js_scroll_to_center(driver, combo)
                continue

            # 5) garantir que a própria opção está visível e clicar
            if not _js_is_in_viewport(driver, opt_50):
                _js_scroll_to_center(driver, opt_50)
            driver.execute_script("arguments[0].click();", opt_50)

            # 6) aguardar fechamento do overlay e (opcional) spinner sumir
            # overlay costuma sumir: verifique se a opção desapareceu
            try:
                W(driver, 2).until(EC.staleness_of(opt_50))
            except TimeoutException:
                pass

            if after_click_wait_spinner:
                try:
                    W(driver, 8).until(EC.invisibility_of_element_located(
                        (By.CSS_SELECTOR, ".loading,.spinner,.is-loading")
                    ))
                except TimeoutException:
                    pass

            print("Tabela atualizada para mostrar 50 linhas.")
            return True

        except (TimeoutException, StaleElementReferenceException):
            # tenta de novo
            continue

    # fallback final: tente usar teclado (algumas libs navegam com setas/enter)
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", combo)
        combo.click()
        time.sleep(0.1)
        # manda HOME e depois algumas setas para garantir chegar na última opção
        combo.send_keys(Keys.HOME)
        for _ in range(10):
            combo.send_keys(Keys.ARROW_DOWN)
        combo.send_keys(Keys.END)   # em muitas UIs, END vai para maior (ex.: 50/100)
        combo.send_keys(Keys.ENTER)
        print("Tabela atualizada (fallback teclado).")
        return True
    except Exception:
        pass

    print("⚠️ Não consegui definir 50 linhas (fora do viewport/overlay teimoso). Role a tela levemente e tente novamente.")
    return False



def ler_grupos_visiveis(driver):
    # Retorna lista de strings (números visualizados)
    return driver.execute_script("""
        const btns = document.querySelectorAll('[aria-describedby="tabelaGrupos"] tbody tr td:first-child button');
        return Array.from(btns).map(b => (b.textContent||'').trim());
    """)



def clicar_grupo_por_numero(driver, numero_str):
    # Clica por JS no botão do grupo == numero_str
    return driver.execute_script("""
        const alvo = arguments[0];
        const btns = document.querySelectorAll('[aria-describedby="tabelaGrupos"] tbody tr td:first-child button');
        for (const b of btns) {
            if (((b.textContent||'').trim()) === alvo) { b.click(); return true; }
        }
        return false;
    """, str(numero_str))


def abrir_exibir_creditos(driver, timeout=15):

    span = W(driver, 15).until(EC.presence_of_element_located((By.XPATH, XPATH_EXIBIR_CREDITOS)))
    btn  = span.find_element(By.XPATH, "./ancestor::button[1]")
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
    driver.execute_script("window.scrollBy(0, -64);")

    try:
        btn.click()
    except Exception:
        driver.execute_script("arguments[0].click();", btn)

    # # aguarda carregar e ficar clicável
    # btn = W(driver, timeout).until(
    #     EC.element_to_be_clickable((By.XPATH, XPATH_EXIBIR_CREDITOS))
    # )
    # _scroll_center(driver, btn)
    # try:
    #     _real_click(driver, btn)
    # except Exception:
    #     driver.execute_script("arguments[0].click();", btn)

    # aguarda a tabela de créditos aparecer (ou spinner sumir)
    try:
        W(driver, 12).until(EC.presence_of_element_located(
            (By.XPATH, "//p[normalize-space()='créditos disponíveis']/following-sibling::div/table")
        ))
    except TimeoutException:
        # fallback: espere spinner sumir e tente localizar de novo rapidamente
        try:
            W(driver, 6).until(EC.invisibility_of_element_located(
                (By.CSS_SELECTOR, ".loading,.spinner,.is-loading")
            ))
        except Exception:
            pass
        W(driver, 4).until(EC.presence_of_element_located(
            (By.XPATH, "//p[normalize-space()='créditos disponíveis']/following-sibling::div/table")
        ))



def ler_tabela_creditos(driver):
    """
    Lê a tabela 'créditos disponíveis' em UM roundtrip via JS.
    Retorna lista de dicts: [{'codigo':'','nome':'','credito':float,'parcela':float}, ...]
    """
    rows = driver.execute_script("""
        const table = document.evaluate("//p[normalize-space()='créditos disponíveis']/following-sibling::div/table",
            document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
        if (!table) return [];
        return Array.from(table.querySelectorAll('tbody tr')).map(tr => {
            const tds = tr.querySelectorAll('td');
            const codigo = (tds[0]?.textContent||'').trim();
            const nome   = (tds[1]?.textContent||'').trim();
            const creditoTxt = (tds[3]?.textContent||'').trim();
            const parcelaTxt = (tds[4]?.textContent||'').trim();
            function parseBR(v){return parseFloat((v||'').replace(/\./g,'').replace(',','.'))||0;}
            return { codigo, nome, credito: parseBR(creditoTxt), parcela: parseBR(parcelaTxt) };
        });
    """)
    return rows or []

def clicar_codigo_credito(driver, codigo_bem):
    # Clica no <u> da coluna do código correspondente
    ok = driver.execute_script("""
        const alvo = arguments[0];
        const table = document.evaluate("//p[normalize-space()='créditos disponíveis']/following-sibling::div/table",
            document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
        if (!table) return false;
        const trs = table.querySelectorAll('tbody tr');
        for (const tr of trs) {
            const td0 = tr.querySelector('td');
            if (!td0) continue;
            const code = (td0.textContent||'').trim();
            if (code === alvo) {
                const u = td0.querySelector('u') || td0;
                u.click();
                return true;
            }
        }
        return false;
    """, str(codigo_bem))
    return bool(ok)

def clicar_contratar_cota(driver):
    el = W(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//span[contains(normalize-space(), "contratar cota")]'))
    )
    js_click(driver, el)

def voltar_para_grupos(driver):
    el = W(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//p[contains(normalize-space(), "voltar para grupos")]'))
    )
    js_click(driver, el)
    esperar_sumir_loading(driver, timeout=8)


def proxima_pagina(driver):
    # Tenta clicar no "próxima página" se existir e estiver habilitado
    return driver.execute_script("""
        const btn = document.querySelector('button#nextPageId');
        if (!btn) return "NAO_EXISTE";
        if (btn.disabled) return "DESABILITADO";
        btn.click(); return "OK";
    """)

def ir_para_pagina_1_e_filtrar(driver):
    # Campo "ir para a página"
    ok = driver.execute_script("""
        const input = document.querySelector('input[aria-label="ir para a página"]');
        if (!input) return false;
        input.focus();
        input.value = "1";
        input.dispatchEvent(new Event('input', {bubbles:true}));
        input.dispatchEvent(new KeyboardEvent('keydown', {key:'Enter', bubbles:true}));
        return true;
    """)
    esperar_sumir_loading(driver, timeout=10)
    # Botão Filtrar
    btn_ativo = driver.execute_script("""
        const b = document.getElementById('btnFiltrar');
        if (!b) return "SEM_BOTAO";
        if (b.disabled) return "DESABILITADO";
        b.click(); return "OK";
    """)
    return ok, btn_ativo


def toggle_seguro_off_if_cpf(driver, is_cpf: bool, max_tentativas=8):
    if not is_cpf:
        return
    try:
        btn = W(driver, 6).until(
            EC.presence_of_element_located((By.XPATH, '//input[@formcontrolname="checkSeguro"]'))
        )
    except Exception:
        return
    for _ in range(max_tentativas):
        estado = btn.get_attribute('aria-pressed')
        if estado == 'false':
            return
        js_click(driver, btn)
        time.sleep(0.25)  # pequena folga pra UI reagir
        try:
            btn = driver.find_element(By.XPATH, '//input[@formcontrolname="checkSeguro"]')
        except Exception:
            try:
                btn = W(driver, 3).until(
                    EC.presence_of_element_located((By.XPATH, '//input[@formcontrolname="checkSeguro"]'))
                )
            except Exception:
                break



def ler_grupos_visiveis(driver):
    js = """
    return Array.from(
      document.querySelectorAll('[aria-describedby="tabelaGrupos"] tbody tr td:first-child button')
    ).map(b => (b.textContent||'').trim());
    """
    return driver.execute_script(js)

def clicar_grupo_por_texto(driver, numero):
    js = """
    const btns = document.querySelectorAll('[aria-describedby="tabelaGrupos"] tbody tr td:first-child button');
    for (const b of btns) {
        if ((b.textContent||'').trim() === arguments[0]) { b.click(); return true; }
    }
    return false;
    """
    return driver.execute_script(js, str(numero))


def _strip_accents(s: str) -> str:
    return ''.join(c for c in unicodedata.normalize('NFKD', s) if not unicodedata.combining(c))


ALLOWED_TP = [
    "imoveis",
    "veiculos leves",
    "motocicletas",
    "veiculos pesados",
]

# sinônimos/erros comuns -> alvo
TP_SYNONYMS = {
    # imóveis
    "imovel": "imoveis",
    "imóveis": "imoveis",
    "imóvel": "imoveis",
    "imo": "imoveis",
    "imoveis residenciais": "imoveis",
    "imoveis comerciais": "imoveis",

    # veículos leves
    "carro": "veiculos leves",
    "carros": "veiculos leves",
    "automovel": "veiculos leves",
    "automóvel": "veiculos leves",
    "veiculo": "veiculos leves",
    "veículos": "veiculos leves",
    "veiculo leve": "veiculos leves",
    "veículo leve": "veiculos leves",
    "veiculos leve": "veiculos leves",
    "veículos leves": "veiculos leves",
    "leves": "veiculos leves",
    "car": "veiculos leves",

    # motocicletas
    "moto": "motocicletas",
    "motocicleta": "motocicletas",
    "motocycle": "motocicletas",
    "motoca": "motocicletas",

    # veículos pesados
    "caminhao": "veiculos pesados",
    "caminhão": "veiculos pesados",
    "caminhoes": "veiculos pesados",
    "caminhões": "veiculos pesados",
    "truck": "veiculos pesados",
    "veiculo pesado": "veiculos pesados",
    "veículo pesado": "veiculos pesados",
    "veiculos pesado": "veiculos pesados",
    "pesados": "veiculos pesados",
    "veiculos pesados": "veiculos pesados",
}

def _norm(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = " ".join(s.split())  # colapsa espaços
    return s

def resolve_tp_produto(raw_value, cutoff=0.74, fallback=None, alert_on_fail=True):
    """
    Tenta resolver o tipo de produto para um dos ALLOWED_TP.
    - Usa normalização, sinônimos e fuzzy matching.
    - Se não encontrar:
        - dispara ui_alert (se alert_on_fail=True)
        - retorna (None) ou (fallback) se informado.
    Retorna: string normalizada ou None/fallback.
    """
    val = _norm(raw_value)

    # 1) já é válido?
    if val in ALLOWED_TP:
        return val

    # 2) sinônimos diretos
    if val in TP_SYNONYMS:
        return TP_SYNONYMS[val]

    # 3) fuzzy contra allowed + chaves de sinônimos
    universe = list(set(ALLOWED_TP + list(TP_SYNONYMS.keys())))
    match = get_close_matches(val, universe, n=1, cutoff=cutoff)
    if match:
        m = match[0]
        # se casou sinônimo, converte ao alvo
        if m in TP_SYNONYMS:
            return TP_SYNONYMS[m]
        return m  # já é um allowed

    # 4) falha total
    if alert_on_fail:
        ui_alert("Erro", "Tipo de produto preenchido errado", "info")
    return fallback  # pode ser None (caller decide o que fazer)


def bring_chrome_to_front():

    """Garante que o Google Chrome esteja ativo e na frente no macOS."""
    if platform.system() != "Darwin":
        return  # só aplica no macOS

    try:
        script = '''
        tell application "System Events"
            set frontmost of process "Google Chrome" to true
        end tell
        '''
        subprocess.run(["osascript", "-e", script], check=True)
        print("🪟 Chrome trazido para frente.")
    except Exception as e:
        print("⚠️ Falha ao trazer Chrome para frente:", e)


def _ci(text: str) -> str:
    # lower sem acentos (para XPaths case/diacritic-insensitive)
    import unicodedata
    t = unicodedata.normalize("NFD", text)
    t = "".join(ch for ch in t if not unicodedata.combining(ch))
    return t.lower()

XPATH_EXIBIR_CREDITOS = (
    "//div[contains(@class,'cdk-overlay-pane')]"
    "//button[contains(@class,'ids-main-button')]"
    "/span[normalize-space()='exibir créditos disponíveis']"
)


def _scroll_center(driver, el):
    driver.execute_script("arguments[0].scrollIntoView({block:'center', inline:'center'});", el)
    driver.execute_script("window.scrollBy(0, -64);")  # compensa header fixo

def _real_click(driver, el):
    # click real com MouseEvent, evita alguns bloqueios de framework
    driver.execute_script("""
        const el = arguments[0];
        const r = el.getBoundingClientRect();
        const x = r.left + r.width/2, y = r.top + r.height/2;
        const opts = {bubbles:true, cancelable:true, clientX:x, clientY:y, view:window};
        el.dispatchEvent(new MouseEvent('pointerdown', opts));
        el.dispatchEvent(new MouseEvent('mousedown', opts));
        el.dispatchEvent(new MouseEvent('mouseup', opts));
        el.dispatchEvent(new MouseEvent('click', opts));
    """, el)

def _table_grupos_present(driver):
    els = driver.find_elements(By.CSS_SELECTOR, '[aria-describedby="tabelaGrupos"] tbody tr')
    return len(els) > 0

def wait_exibir_creditos_or_state(driver, timeout=12):
    """
    Espera a 'tela' do grupo abrir.
    Retorna True se o botão aparecer ou se detectarmos estado alternativo confiável.
    """
    end = time.time() + timeout
    last_exc = None
    while time.time() < end:
        try:
            # 1) botão visível/clicável?
            btn = driver.find_element(By.XPATH, XPATH_EXIBIR_CREDITOS)
            if btn.is_displayed():
                return True
        except Exception as e:
            last_exc = e

        # 2) às vezes a tabela de grupos fica oculta/é substituída
        #    (se sumiu, já é um bom sinal de navegação)
        if not _table_grupos_present(driver):
            # dá mais 0.5s para a UI montar; se o botão aparecer nesse intervalo, ok.
            try:
                W(driver, 0.5).until(EC.presence_of_element_located((By.XPATH, XPATH_EXIBIR_CREDITOS)))
                return True
            except Exception:
                # pode haver outra tela intermediária; seguimos tentando até o timeout
                pass

        time.sleep(0.15)
    # estourou
    return False

def _get_table_container(driver):
    # container rolável da tabela (ajuste se o seu for outro)
    el = driver.execute_script("""
        const t = document.querySelector('[aria-describedby="tabelaGrupos"]');
        if (!t) return null;
        // sobe até achar o rolável
        let n = t;
        while (n && n !== document.body) {
            const s = getComputedStyle(n);
            if (/(auto|scroll)/.test(s.overflowY) && n.scrollHeight > n.clientHeight) return n;
            n = n.parentElement;
        }
        return document.scrollingElement || document.documentElement;
    """)
    return el

def _scroll_container_to_btn(driver, numero_grupo):
    # acha o botão do grupo comparando por int (ignora zeros à esquerda)
    return driver.execute_script("""
        const alvoInt = parseInt(String(arguments[0]).trim(), 10);
        const table = document.querySelector('[aria-describedby="tabelaGrupos"]');
        if (!table) return null;

        function findBtn() {
          const btns = table.querySelectorAll('tbody tr td:first-child button');
          for (const b of btns) {
            const t = (b.textContent||'').trim();
            const val = parseInt(t, 10);
            if (!Number.isNaN(val) && val === alvoInt) return b;
          }
          return null;
        }

        let btn = findBtn();
        if (btn) { btn.scrollIntoView({block:'center', inline:'center'}); return btn; }

        // lista pode ser virtualizada → role o container rolável
        function getContainer(el){
          let n = el;
          while (n && n !== document.body) {
            const s = getComputedStyle(n);
            if (/(auto|scroll)/.test(s.overflowY) && n.scrollHeight > n.clientHeight) return n;
            n = n.parentElement;
          }
          return document.scrollingElement || document.documentElement;
        }
        const cont = getContainer(table);
        const step = Math.max(200, Math.floor(cont.clientHeight*0.8));
        for (let i=0;i<20;i++){
          cont.scrollTop = Math.min(cont.scrollTop + step, cont.scrollHeight);
          btn = findBtn();
          if (btn){
            btn.scrollIntoView({block:'center', inline:'center'});
            return btn;
          }
        }
        return null;
    """, str(numero_grupo))

def _normalizar_int(texto: str):
    """Tenta extrair um inteiro de um texto qualquer."""
    if texto is None:
        return None
    texto = texto.strip()
    try:
        return int(texto)
    except ValueError:
        numeros = re.sub(r'\D', '', texto)
        return int(numeros) if numeros else None




def confirmar_grupo_na_janela(driver, alvo: int, tentar_reclicar=True, timeout=8) -> bool:
    """
    Etapa 1 de confirmação:
    - Lê o campo codGrupo na janela de detalhes
    - Se for diferente do grupo alvo, fecha e tenta reabrir UMA vez (se tentar_reclicar=True)
    - Retorna True se conseguir garantir que o grupo é o correto, senão False
    """
    tentativas = 0

    while True:
        tentativas += 1
        try:
            elem = WebDriverWait(driver, timeout).until(
                EC.visibility_of_element_located((By.ID, "codGrupo"))
            )
            id_janela_grupo = elem.text
        except TimeoutException:
            print("⚠️ Não consegui localizar o campo 'codGrupo' na janela de detalhes.")
            return False

        grupo_lido = _normalizar_int(id_janela_grupo)
        print(f"[CONF 1] codGrupo na janela: {id_janela_grupo!r} (normalizado: {grupo_lido})")

        if grupo_lido == alvo:
            print(f"✅ Grupo confirmado na janela de detalhes: {grupo_lido}")
            return True

        # Se chegou aqui, grupo está diferente
        print(f"⚠️ Grupo da janela ({grupo_lido}) diferente do alvo ({alvo}).")

        # Tenta fechar a janela
        try:
            botao_fechar = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "#ids-modal-0-close"))
            )
            botao_fechar.click()
            # Se tiver um spinner/aplicação usa, mantemos o padrão:
            try:
                esperar_sumir_loading(driver, timeout=8)
            except Exception:
                pass
        except TimeoutException:
            print("⚠️ Não consegui clicar no botão de fechar da janela de grupo.")
            return False

        # Se já tentamos reabrir uma vez ou não queremos reclicar, encerra
        if not (tentar_reclicar and tentativas == 1):
            print("⚠️ Não será feita nova tentativa de reabrir o mesmo grupo.")
            return False

        # Tentar clicar novamente no mesmo grupo
        print("🔁 Tentando reabrir o mesmo grupo para confirmar...")
        if not clicar_grupo_e_esperar_tela(driver, alvo, timeout=12, retry=0):
            print("❌ Falha ao reabrir o grupo após inconsistência.")
            return False
        # volta para o while e lê de novo o codGrupo


def confirmar_grupo_apos_creditos(driver, alvo: int, timeout=8) -> bool:
    """
    Etapa 2 de confirmação:
    - Depois de clicar em 'Exibir créditos' e a tela aparecer,
      confere o número do grupo no botão ao lado do texto 'grupo'
    - Retorna True se o grupo confere, False se for diferente ou não localizar.
    """
    try:
        botao_elemento = WebDriverWait(driver, timeout).until(
            EC.visibility_of_element_located((
                By.XPATH,
                "//p[normalize-space(translate(text(),'GRUPO','grupo'))='grupo']/following-sibling::button"
            ))
        )
        numero_texto = botao_elemento.find_element(By.TAG_NAME, 'u').text
    except TimeoutException:
        print("⚠️ Não achei o botão com o número do grupo na tela de créditos.")
        return False
    except Exception as e:
        print(f"⚠️ Erro ao localizar número do grupo na tela de créditos: {e}")
        return False

    grupo_lido = _normalizar_int(numero_texto)
    print(f"[CONF 2] grupo na tela de créditos: {numero_texto!r} (normalizado: {grupo_lido})")

    if grupo_lido == alvo:
        print(f"✅ Grupo confirmado na tela de créditos: {grupo_lido}")
        return True

    print(f"⚠️ Grupo da tela de créditos ({grupo_lido}) diferente do alvo ({alvo}).")
    return False


def _safe_real_click(driver, el):
    # tenta click “real”, cai para .click() e por fim JS .click()
    try:
        driver.execute_script("""
            const el = arguments[0];
            const r = el.getBoundingClientRect();
            const x = r.left + r.width/2, y = r.top + r.height/2;
            const opts = {bubbles:true, cancelable:true, clientX:x, clientY:y, view:window};
            el.dispatchEvent(new MouseEvent('pointerdown', opts));
            el.dispatchEvent(new MouseEvent('mousedown', opts));
            el.dispatchEvent(new MouseEvent('mouseup', opts));
            el.dispatchEvent(new MouseEvent('click', opts));
        """, el)
        return True
    except Exception:
        try:
            el.click()
            return True
        except Exception:
            try:
                driver.execute_script("arguments[0].click();", el)
                return True
            except Exception:
                return False


def clicar_grupo_e_esperar_tela(driver, numero_grupo: int, timeout=12, retry=2):
    for attempt in range(retry+1):
        btn = _scroll_container_to_btn(driver, numero_grupo)
        if not btn:
            # tente rolar a janela um pouco e repetir
            driver.execute_script("window.scrollBy(0, 200);")
            time.sleep(0.2)
            continue

        # clique robusto
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
        driver.execute_script("window.scrollBy(0, -64);")
        try:
            btn.click()
        except Exception:
            driver.execute_script("arguments[0].click();", btn)

        if wait_exibir_creditos_or_state(driver, timeout=timeout):
            return True

        # pequena folga e re-tentativa
        time.sleep(0.3)

    return False

def abrir_creditos_ate_popular(driver, max_open_retries=3, wait_rows_timeout=10):
    for _ in range(max_open_retries):
        # achar o span dentro do overlay e clicar no button ancestral
        try:
            span = W(driver, 15).until(EC.presence_of_element_located((By.XPATH, XPATH_EXIBIR_CREDITOS)))
        except TimeoutException:
            return False
        btn = span.find_element(By.XPATH, "./ancestor::button[1]")
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
        driver.execute_script("window.scrollBy(0, -64);")
        try:
            btn.click()
        except Exception:
            driver.execute_script("arguments[0].click();", btn)

        # aguarda tabela aparecer
        try:
            W(driver, 12).until(EC.presence_of_element_located(
                (By.XPATH, "//p[normalize-space()='créditos disponíveis']/following-sibling::div/table")
            ))
        except TimeoutException:
            pass

        # espera linhas > 0
        end = time.time() + wait_rows_timeout
        while time.time() < end:
            rows = driver.execute_script("""
                const t = document.evaluate("//p[normalize-space()='créditos disponíveis']/following-sibling::div/table",
                    document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                return t ? t.querySelectorAll('tbody tr').length : -1;
            """)
            if rows and rows > 0:
                return True
            time.sleep(0.25)

        # re-tenta reabrir
        time.sleep(0.3)

    return False




#=======================================================================================================================
#                  FUNÇÃO 1 - Iniciar inserindo dados do cliente (CPF, data nascimento, tipo do produto)
#=======================================================================================================================
#Inicio - Inserir CPF / data nascimento / tipo do produto e deixar para usuario inserir o reCaptcha


def inserir_dados_cliente_js():
    global driver, df_atual, cliente_atual, cpf_atual, list_grupos

    if not ensure_driver_alive(driver):
      raise RuntimeError("WebDriver inválido. Recrie o driver.")
    
    # tentar garantir que estamos na aba "simulação e contratação"
    ok = ensure_tab_with_title(
    driver,
    desired_titles=[
            "simulação e contratação",
            "contratação e simulação",
            "plataforma 360i",
        ],
        timeout=6,
        try_url_contains=["simulacao", "contratacao", "360i"]
        )

    if not ok:
      # pode lançar erro, ou mostrar alerta, ou abrir nova aba manualmente
      # eu sugiro avisar o usuário e interromper o fluxo
      import tkinter as tk
      from tkinter import messagebox
      try:
          root = tk.Tk(); root.withdraw()
          messagebox.showwarning(
              "Aba não encontrada",
              "Não localizei uma aba com o título 'simulação e contratação'.\n"
              "Verifique se a aba está aberta no Chrome e tente novamente."
          )
          root.destroy()
      except Exception:
          # fallback para console
          print("Aba 'simulação e contratação' não encontrada.")
      raise RuntimeError("Aba alvo não encontrada.")


    # carregar clientes (com tratamento de erro interno/alerta)
    try:
        load_df_clientes()
    except RuntimeError as e:
        # load_df_clientes já chamou ui_alert apropriado (arquivo ausente, etc.)
        print(f"[ERRO] load_df_clientes falhou: {e}")
        return  # aborta cedo sem tentar preencher nada

    if df_atual is None or len(df_atual) == 0:
        ui_alert("Arquivo vazio", "O arquivo esta vazio ou não foi localizado na pasta", "info")
        raise RuntimeError("df_atual vazio.")

    cliente_atual = df_atual.iloc[0]
    cpf_atual_str    = str(cliente_atual['cpf_cnpj']).strip()
    cpf_atual = re.sub(r'[.\-\/]', '', cpf_atual_str)  # só números
    tipo_cliente = str(cliente_atual['tipo_cliente']).strip().lower()   # "cpf" | "cnpj"
    data_nas     = str(cliente_atual['data_nasc_fundacao']).strip()
    tp_produto_cliente   = str(cliente_atual['tp_produto']).strip().lower()

    raw_tp = tp_produto_cliente
    tp_produto = resolve_tp_produto(raw_tp, cutoff=0.74, fallback=None, alert_on_fail=True)

    if not tp_produto:
        print(f"[WARN] tp_produto inválido: {raw_tp!r}. Prosseguindo sem preencher produto.")
        raise
    else:
        print("tp_produto resolvido:", tp_produto)


    # ---------- ETAPA 1: clicar CPF/CNPJ ----------
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
    time.sleep(0.18)

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
    
    tp = str(cliente_atual['tp_produto'] or '').strip().lower()
    
    raw_tp = tp
    tp_produto = resolve_tp_produto(raw_tp, cutoff=0.74, fallback=None, alert_on_fail=True)

    name_arquivo_grupo = (f'grupos_{_normalize_tipo(tp_produto)}.csv')
    try:
        csv_path = _find_user_file(f'grupos_{_normalize_tipo(tp_produto)}.csv')
        print(f"Usando lista de grupos: {csv_path}")
        df_grupos_ignorar = pd.read_csv(csv_path, sep=';', dtype={'grupo': str})
        #list_grupos = df_grupos_ignorar['grupo'].astype(str).str.strip().tolist()
        list_grupos = [int(x) for x in df_grupos_ignorar['grupo'].astype(str).str.strip().tolist()]
    except FileNotFoundError:
        ui_alert("Não Encontrado", f"Não localizei o arquivo {name_arquivo_grupo} na pasta", "info")
        raise False
    
    return res







#=======================================================================================================================
#                  FUNÇÃO 2 - PRINCIPAL - Buscar consórcio do cliente e selecionar a melhor opção
#=======================================================================================================================

###
def clicar_reCaptcha():
    #Tentar recaptcha simples:
    # Acessar iframe do recaptcha e clicar no botao
    iframe_Rv2 = driver.find_element(By.CSS_SELECTOR, "iframe")
    driver.switch_to.frame(iframe_Rv2)
    driver.find_element(By.ID, "recaptcha-anchor").click()
    driver.switch_to.default_content()
    
    esperar_sumir_loading(driver, timeout=10)
    
    time.sleep(1.5)

    # Botão Filtrar
    btn_ativo = driver.execute_script("""
        const b = document.getElementById('btnFiltrar');
        if (!b) return "SEM_BOTAO";
        if (b.disabled) return "DESABILITADO";
        b.click(); return "OK";
    """)
    return btn_ativo


def buscar_consorcio_cliente():
    global grupo_encontrado, cpf_atual, cliente_atual, driver

    print("Iniciando a busca pelo melhor consórcio para o cliente...")
    
    list_grupos_set = set(int(x) for x in list_grupos)

    bring_chrome_to_front()

    
    # 1) Valor máximo (rápido por JS)
    valor_maximo_float = get_valor_maximo_float(driver)
    if valor_maximo_float is None:
        print("⚠️ Não foi possível extrair o valor máximo.")
    else:
        print(f"Valor máximo extraído: R$ {valor_maximo_float:.2f}")

    # 2) Garantir 50 linhas (cliques via JS + espera spinner)
    ok = set_page_size_50(driver)
    if not ok:
        # se quiser, faça uma rolagem global e tente outra vez
        driver.execute_script("window.scrollBy(0, window.innerHeight/2);")
        set_page_size_50(driver)
    print("Tabela atualizada para mostrar 50 linhas.")

    grupo_encontrado = False
    grupo_localizado = True

    tentativas = 0
    while grupo_localizado:

        if tentativas > 50:
            ui_alert("Limite", "Limite de controle", "info")
            break

        tentativas+1

        # 3) Garantir tabela presente e estável
        W(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '[aria-describedby="tabelaGrupos"] tbody')))
        esperar_sumir_loading(driver, timeout=8)

        # 4) Varredura de grupos em lote (JS) e filtro em Python
        numeros = ler_grupos_visiveis(driver)
        print(f"Grupos visíveis nesta página: {numeros[:10]}{' ...' if len(numeros)>10 else ''}")
        # Converter p/ int e procurar o primeiro que esteja no set
        alvo = None
        for n in numeros:
            try:
                ni = int(n)
            except:
                continue
            if ni in list_grupos_set:   # <-- use set para O(1)
                alvo = ni
                break

        if alvo is None:
            # nenhuma correspondência nesta página: tentar paginação
            print("Nenhum grupo-alvo na página atual.")
            status = proxima_pagina(driver)
            if status == "OK":
                esperar_sumir_loading(driver, timeout=8)
                print("Indo para a próxima página...")
                continue


            elif status == "DESABILITADO":
                print("Próxima página desabilitada — voltando p/ página 1 e (se possível) refiltrando.")
                ok, filtrar = ir_para_pagina_1_e_filtrar(driver)
                if filtrar == "OK":
                    esperar_sumir_loading(driver, timeout=10)
                    continue
                else:
                    btn_ativo = clicar_reCaptcha()
                    if btn_ativo == "OK":
                        continue
                    else:
                        ui_alert("reCAPTCHA", "Resolva o reCAPTCHA e clique em Buscar Consórcio!", "info")
                        break
            
            
            else:  # NAO_EXISTE
                print("Sem paginação — voltando p/ página 1 e (se possível) refiltrando.")
                ok, filtrar = ir_para_pagina_1_e_filtrar(driver)
                if filtrar == "OK":
                    esperar_sumir_loading(driver, timeout=10)
                    continue
                else:
                    ui_alert("reCAPTCHA", "Resolva o reCAPTCHA e clique em Buscar Consórcio!", "info")
                    break

    
        # 5) Clicar no grupo alvo e esperar a tela abrir
        if not clicar_grupo_e_esperar_tela(driver, alvo, timeout=12, retry=2):
            print(f"Falha ao abrir detalhes do grupo {alvo}; tentando próxima página.")
            status = proxima_pagina(driver)
            if status == "OK":
                esperar_sumir_loading(driver, timeout=8)
                continue
            elif status == "DESABILITADO":
                print("Próxima página desabilitada — voltando p/ página 1 e refiltrando.")
                ok, filtrar = ir_para_pagina_1_e_filtrar(driver)
                if filtrar == "OK":
                    esperar_sumir_loading(driver, timeout=10)
                    continue
                else:
                    ui_alert("reCAPTCHA", "Resolva o reCAPTCHA e clique em Buscar Consórcio!", "info")
                    break
            else:
                # NAO_EXISTE
                ok, filtrar = ir_para_pagina_1_e_filtrar(driver)
                if filtrar == "OK":
                    esperar_sumir_loading(driver, timeout=10)
                    continue
                else:
                    ui_alert("reCAPTCHA", "Resolva o reCAPTCHA e clique em Buscar Consórcio!", "info")
                    break

        # Etapa 1 de confirmação: garantir que a janela que abriu é o mesmo grupo
        if not confirmar_grupo_na_janela(driver, alvo, tentar_reclicar=True):
            print(f"❌ Inconsistência de grupo na janela de detalhes para {alvo}. Ignorando este grupo.")
            # Remove esse grupo da lista de candidatos para evitar loop infinito
            list_grupos_set.discard(alvo)
            voltar_para_grupos(driver)
            # segue no while, procurando outro grupo
            continue


        # 6) Abrir “exibir créditos disponíveis” e garantir linhas
        if not abrir_creditos_ate_popular(driver, max_open_retries=3, wait_rows_timeout=10):
            print("⚠️ Não consegui popular a tabela de créditos. Voltando para grupos…")
            voltar_para_grupos(driver)
            continue

        # Etapa 2 de confirmação: depois de exibir créditos, conferir o grupo na tela
        if not confirmar_grupo_apos_creditos(driver, alvo, timeout=8):
            print(f"❌ Inconsistência de grupo na tela de créditos para {alvo}. Voltando para grupos e ignorando este grupo.")
            list_grupos_set.discard(alvo)
            voltar_para_grupos(driver)
            # segue no while; vai procurar outro grupo/página
            continue




        # 7) Ler créditos (lote, via JS) e escolher melhor dentro do limite
        creditos = ler_tabela_creditos(driver)
        print(f"Total de linhas de crédito: {len(creditos)}")

        melhor = None
        for c in creditos:
            # c = {'codigo','nome','credito','parcela'}
            if valor_maximo_float is not None and c['parcela'] > valor_maximo_float:
                continue
            if melhor is None or c['credito'] > melhor['credito']:
                melhor = c

        if not melhor:
            print(f"❌ Nenhum crédito <= R$ {valor_maximo_float} no grupo {alvo}. Voltando para grupos...")
            voltar_para_grupos(driver)
            # segue no while (tentar próximo grupo/página)
            continue

        print("✅ Melhor opção:", melhor)

        # 8) Clicar no código do bem selecionado
        if not clicar_codigo_credito(driver, melhor['codigo']):
            print("Falha ao clicar no código do bem; voltando para grupos.")
            voltar_para_grupos(driver)
            continue

        # 9) Desativar seguro (se CPF)
        is_cpf = str(cliente_atual.get('tipo_cliente') or '').strip().upper() == 'CPF'
        toggle_seguro_off_if_cpf(driver, is_cpf=is_cpf)

        # 10) Contratar cota
        clicar_contratar_cota(driver)
        esperar_sumir_loading(driver, timeout=10)
        print("Clicado em CONTRATAR COTA, aguardando próxima tela...")

        time.sleep(0.1)
        # 11) Preenchimento específico
        if is_cpf:
            time.sleep(0.05)
            sucesso_preenchimento = preencher_dados_pessoais()
        else:
            sucesso_preenchimento = preencher_dados_PJ()

        if sucesso_preenchimento:
            print("✅ Dados pessoais preenchidos com sucesso.")
            atualizar_status_cliente(cpf_atual, "Finalizado")
            ui_alert("Finalizado", "Infos Preenchidas - Clique em CONTRATAR!", "info")
            grupo_encontrado = True
            grupo_localizado = False
        else:
            print("❌ Falha ao preencher os dados pessoais.")
            atualizar_status_cliente(cpf_atual, "Erro")
            ui_alert("erro", "Ops, algo deu errado; verifique e preencha manualmente ou reinicie!", "info")
            # Se quiser tentar outro grupo, volte para grupos aqui:
            # voltar_para_grupos(driver)
            grupo_localizado = False  # ou True se quiser continuar a buscar

    # fim while


###

#=======================================================================================================================
#                                  ##### FUNÇÃO 3- Preencher dados pessoais  #####
#=======================================================================================================================

#Aux..
def _get_shadow_root1(driver, timeout=10):
    host = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "mf-iparceiros-cadastrocliente"))
    )
    return driver.execute_script("return arguments[0].shadowRoot;", host)

def _wait_overlays_gone(driver, timeout=8):
    # Ajuste os seletores conforme seu app
    overlay_sel = ".cdk-overlay-backdrop, .loading, ids-spinner, [aria-busy='true']"
    end = time.time() + timeout
    while time.time() < end:
        try:
            overs = driver.find_elements(By.CSS_SELECTOR, overlay_sel)
            if not any(o.is_displayed() for o in overs):
                return True
        except Exception:
            return True
        time.sleep(0.15)
    return True  # não bloqueia; apenas tenta

def _btn_enabled(btn):
    try:
        dis = btn.get_attribute("disabled")
        aria = btn.get_attribute("aria-disabled")
        return (dis in (None, "", "false")) and (aria in (None, "", "false"))
    except Exception:
        return True

def _try_click_strategies(driver, btn):
    try:
        ActionChains(driver).move_to_element(btn).pause(random.uniform(0.05, 0.2)).click(btn).perform()
        return True
    except ElementClickInterceptedException:
        try:
            driver.execute_script("arguments[0].click();", btn)
            # depois do clique, garantir combos fechados
            ensure_key_dropdowns_closed(driver)

            return True
        except Exception:
            pass
    except Exception:
        pass

    # último recurso: disparo de evento e tecla Enter
    try:
        driver.execute_script("""
          const el = arguments[0];
          try { el.scrollIntoView({block:'center'}); } catch(e){}
          try { el.focus(); } catch(e){}
          const evt = new MouseEvent('click', {bubbles:true, cancelable:true, composed:true});
          el.dispatchEvent(evt);
        """, btn)
        return True
    except Exception:
        pass

    try:
        btn.send_keys(Keys.ENTER)
        return True
    except Exception:
        return False



def _js_neutral_click(driver):
    """
    Dispara um clique sintético em uma área neutra (backdrop/container/body),
    com a sequência pointerdown/mousedown/pointerup/mouseup/click.
    """
    js = r"""
    (function(){
      function fireClick(el){
        if(!el) return false;
        const r = el.getBoundingClientRect();
        const x = Math.floor(r.left + (r.width  || 100)/2);
        const y = Math.floor(r.top  + (r.height || 100)/2);
        const opts = {bubbles:true, composed:true, cancelable:true, clientX:x, clientY:y, view:window};
        for (const t of ['pointerdown','mousedown','pointerup','mouseup','click']) {
          try { el.dispatchEvent(new MouseEvent(t, opts)); } catch(e){}
        }
        return true;
      }

      // candidatos do overlay/container e áreas neutras
      const candidates = [
        document.querySelector('.cdk-overlay-backdrop'),
        document.querySelector('.cdk-overlay-container'),
        document.querySelector('.container.ng-star-inserted'),
        document.querySelector('main'),
        document.body,
        document.documentElement
      ].filter(Boolean);

      // tenta clicar no primeiro que conseguir
      for (const el of candidates){
        if (fireClick(el)) return true;
      }
      return false;
    })();
    """
    try:
        return bool(driver.execute_script(js))
    except Exception:
        return False

def _actions_neutral_click(driver):
    """
    Clique físico via Actions no <body> com leve offset.
    Útil quando JS é bloqueado por alguma política.
    """
    try:
        body = driver.find_element(By.TAG_NAME, "body")
        ActionChains(driver).move_to_element_with_offset(body, 5, 5).click().perform()
        return True
    except Exception:
        return False

def _ensure_combo_closed_by_fc(driver, fc):
    """
    Fecha o dropdown de um formcontrolname específico (nacionalidade/UFexpedidor),
    incluindo Escape, blur e clique neutro.
    """
    js = r"""
    return (function(fc){
      function q(s, r=document){ return r.querySelector(s); }
      const host = q("mf-iparceiros-cadastrocliente");
      if(!host || !host.shadowRoot) return {ok:false, reason:"no-host"};
      const root = host.shadowRoot;

      const container =
        q(`ids-select[formcontrolname="${fc}"]`, root) ||
        q(`ids-combobox[formcontrolname="${fc}"]`, root) ||
        (q(`input[formcontrolname="${fc}"]`, root)?.closest('ids-input, ids-fieldset, ids-combobox, ids-select')||null);
      if(!container) return {ok:false, reason:"no-container"};

      const r = container.shadowRoot || container;
      const trigger = r.querySelector('[role="combobox"], [aria-haspopup="listbox"], .ids-trigger, button') || container;

      function isOpen(){
        // pode estar local ou no overlay global
        const localOpen = r.querySelector('[role="listbox"], .ids-listbox, ids-listbox');
        const overlayOpen = Array.from(document.querySelectorAll('[role="listbox"], .ids-listbox, ids-listbox'))
                                 .find(el => el.offsetParent !== null);
        return !!(localOpen || overlayOpen);
      }

      // força fechar
      try{
        trigger.dispatchEvent(new KeyboardEvent('keydown',{key:'Escape',bubbles:true}));
        trigger.dispatchEvent(new KeyboardEvent('keyup',{key:'Escape',bubbles:true}));
      }catch(_){}

      try{ (r.activeElement||document.activeElement)?.blur?.(); }catch(_){}
      try{ document.body && document.body.click(); }catch(_){}

      return {ok:true, closed: !isOpen()};
    })(arguments[0]);
    """
    try:
        return driver.execute_script(js, fc)
    except Exception:
        return {"ok": False, "error": "script-failed"}

def close_dropdowns_aggressively(driver):
    """
    Tática completa:
    - ESC global
    - blur ativo
    - fechar nacionalidade e UFexpedidor
    - clique neutro (JS) e, se precisar, clique via Actions
    """
    # ESC global
    try:
        body = driver.find_element(By.TAG_NAME, "body")
        body.send_keys(Keys.ESCAPE)
    except Exception:
        pass

    # blur ativo (no shadow do host)
    try:
        host = driver.find_element(By.CSS_SELECTOR, "mf-iparceiros-cadastrocliente")
        driver.execute_script("""
          const host = arguments[0];
          try {
            const r = host.shadowRoot || host;
            (r.activeElement || document.activeElement)?.blur?.();
          } catch(e){}
        """, host)
    except Exception:
        pass

    # fecha combos críticos
    for fc in ("nacionalidade", "UFexpedidor"):
        try:
            _ensure_combo_closed_by_fc(driver, fc)
        except Exception:
            pass

    # clique neutro via JS; se falhar, tente Actions
    ok = _js_neutral_click(driver)
    if not ok:
        _actions_neutral_click(driver)




def ensure_key_dropdowns_closed(driver):
  # Fecha especificamente os dois campos críticos
  for fc in ("nacionalidade", "UFexpedidor"):
      try:
          _ensure_combo_closed_by_fc(driver, fc)
      except Exception:
          pass


def clicar_continuar_robusto(driver, max_tries=12, wait_next=3):
    """
    Clica no botão 'Continuar' (button[idsmainbutton]) com retentativas e verifica avanço.
    Retorna True se avançou, False caso contrário.
    """
    url0 = driver.current_url
    next_markers = [
        "[data-etapa='proxima']",
        ".etapa-seguinte",
        "mf-iparceiros-documentos",      # exemplo provável de próxima etapa
        "[data-step='documentos']",       # ajuste conforme seu DOM
    ]

    for i in range(1, max_tries+1):
        try:
            # garantir que nenhum combo chave ficou aberto (nacionalidade/UFexpedidor)
            ensure_key_dropdowns_closed(driver)

            shadow_root1 = _get_shadow_root1(driver)
            botoes = shadow_root1.find_elements(By.CSS_SELECTOR, "button[idsmainbutton]")
            if not botoes:
                # seletores alternativos
                botoes = shadow_root1.find_elements(By.CSS_SELECTOR, 'button[type="submit"], ids-button[variant="primary"] button, button.primary')
            if not botoes:
                print("Botão 'Continuar' não encontrado.")
                return False

            btn = botoes[0]

            # esperar overlays saírem
            _wait_overlays_gone(driver, timeout=6)

            # garantir visível/habilitado
            driver.execute_script("try{arguments[0].scrollIntoView({block:'center'});}catch(e){}", btn)
            if not btn.is_displayed() or not _btn_enabled(btn):
                time.sleep(0.05)

            # tentativas de clique
            ok_click = _try_click_strategies(driver, btn)
            # FECHAR DROPDOWNS E OVERLAYS ANTES DE TUDO
            close_dropdowns_aggressively(driver)

            shadow_root1 = _get_shadow_root1(driver)
            botoes = shadow_root1.find_elements(By.CSS_SELECTOR, "button[idsmainbutton]") or ...
            ...
            ok_click = _try_click_strategies(driver, btn)

            # GARANTE FECHAMENTO APÓS O CLIQUE (caso a lista reabra por efeito colateral)
            close_dropdowns_aggressively(driver)

            if not ok_click:
                # refoca e tenta de novo
                try:
                    driver.execute_script("try{arguments[0].focus();}catch(e){}", btn)
                except Exception:
                    pass
                ok_click = _try_click_strategies(driver, btn)

            # mesmo após o clique, espere sinais de avanço
            advanced = False
            t0 = time.time()
            while time.time() - t0 < wait_next and not advanced:
                time.sleep(0.05)

                # 1) sumiu o botão (ou ficou invisível)
                try:
                    sr_now = _get_shadow_root1(driver, timeout=3)
                    again = sr_now.find_elements(By.CSS_SELECTOR, "button[idsmainbutton]")
                    if not again or (again and not again[0].is_displayed()):
                        advanced = True
                        break
                except Exception:
                    # se deu erro ao reler, pode ser troca de view → considerar avançado
                    advanced = True
                    break

                # 2) URL mudou
                if driver.current_url != url0:
                    advanced = True
                    break

                # 3) apareceu marcador da próxima etapa
                try:
                    for sel in next_markers:
                        if driver.find_elements(By.CSS_SELECTOR, sel):
                            advanced = True
                            break
                    if advanced:
                        break
                except Exception:
                    pass

            if advanced:
                print(f"Avançou após tentativa {i}.")
                return True

            # backoff e re-tentativa
            time.sleep(0.1 * i)

        except StaleElementReferenceException:
            time.sleep(0.25)
        except Exception:
            time.sleep(0.25)

    print("Não foi possível avançar após múltiplas tentativas.")
    return False


def verificar_algum_campo_texto_preenchido(driver, tentativas=3, intervalo=0.5):
    """
    Verifica se ALGUM campo de texto (input) chave foi preenchido.
    Tenta algumas vezes com intervalo.
    """
    seletores = [
        'input[formcontrolname="numero_documento"]',
        'input[formcontrolname="cep"]',
        'input[formcontrolname="celular"]',
        'input[formcontrolname="email"]',
    ]

    js_ler_valor = """
    const sel = arguments[0];
    const host = document.querySelector("mf-iparceiros-cadastrocliente");
    if (!host || !host.shadowRoot) return null;
    const root = host.shadowRoot;
    const el = root.querySelector(sel);
    if (!el) return null;
    return el.value || el.textContent || "";
    """

    for tentativa in range(1, tentativas + 1):
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "mf-iparceiros-cadastrocliente"))
            )
        except Exception as e:
            print(f"[CHECK CAMPOS] Shadow host não encontrado na tentativa {tentativa}: {e}")
            return False

        algum_preenchido = False

        for sel in seletores:
            try:
                valor = driver.execute_script(js_ler_valor, sel)
            except Exception as e:
                print(f"[CHECK CAMPOS] Erro ao ler seletor {sel} na tentativa {tentativa}: {e}")
                continue

            print(f"[CHECK CAMPOS] Tentativa {tentativa} seletor {sel}: valor={repr(valor)}")

            if valor and str(valor).strip():
                algum_preenchido = True
                break

        if algum_preenchido:
            print("✅ Confirmação extra: pelo menos um campo de texto foi preenchido com sucesso.")
            return True

        if tentativa < tentativas:
            print(f"⚠️ Nenhum campo de texto parece preenchido. Aguardando {intervalo}s e tentando novamente...")
            time.sleep(intervalo)
        else:
            print("❌ Após múltiplas tentativas, nenhum campo de texto parece preenchido.")
            return False

    return False

def conferir_preenchimento_com_possivel_retry(driver, cliente_atual, max_reexec=2):
    """
    1) Checa se algum campo de texto foi preenchido.
    2) Se não estiver, tenta reaplicar preencher_cliente_s_profi_js até max_reexec vezes.
    3) Depois de cada reaplicação, checa de novo.
    Retorna True se em algum momento achar campo preenchido, senão False.
    """

    # 1ª checagem normal
    if verificar_algum_campo_texto_preenchido(driver, tentativas=2, intervalo=0.5):
        return True

    # Se chegou aqui, não achou nada preenchido. Vamos tentar reaplicar o preenchimento.
    for i in range(max_reexec):
        print(f"⚠️ Campos ainda vazios; tentando reaplicar preencher_cliente_s_profi_js (tentativa extra {i+1}/{max_reexec})...")

        res_retry = preencher_cliente_s_profi_js(cliente_atual)
        print("Resultado do retry do preenchimento sem profissão:", res_retry)

        if not isinstance(res_retry, dict) or not res_retry.get("ok"):
            print("❌ Retry do preenchimento retornou erro (ok=False).")
            # aqui você pode decidir já abortar ou mesmo assim checar a tela;
            # vou manter checagem para o caso do JS ter escrito algo mesmo retornando estranho

        time.sleep(0.3)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(0.1)

        if verificar_algum_campo_texto_preenchido(driver, tentativas=2, intervalo=0.5):
            print("✅ Após retry, campos de texto foram preenchidos com sucesso.")
            return True

    print("❌ Mesmo após reaplicar o preenchimento, nenhum campo de texto parece preenchido.")
    return False



#========== Preencher: =============#
#===================================#

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

      // segurança extra: fechar nacionalidade e UF e dar um clique neutro
        (function(){
        function q(s, r=document){ return r.querySelector(s); }
        const host = q("mf-iparceiros-cadastrocliente");
        const root = host && host.shadowRoot;

        function forceCloseByFc(fc){
            if(!root) return;
            const container =
            q(`ids-select[formcontrolname="${fc}"]`, root) ||
            q(`ids-combobox[formcontrolname="${fc}"]`, root) ||
            (q(`input[formcontrolname="${fc}"]`, root)?.closest('ids-input, ids-fieldset, ids-combobox, ids-select')||null);
            if(!container) return;
            const r = container.shadowRoot || container;
            const trigger = r.querySelector('[role="combobox"], [aria-haspopup="listbox"], .ids-trigger, button') || container;

            try{
            trigger.dispatchEvent(new KeyboardEvent('keydown',{key:'Escape',bubbles:true}));
            trigger.dispatchEvent(new KeyboardEvent('keyup',{key:'Escape',bubbles:true}));
            }catch(_){}
            try{ (r.activeElement||document.activeElement)?.blur?.(); }catch(_){}
        }

        forceCloseByFc("nacionalidade");
        forceCloseByFc("UFexpedidor");

        // clique neutro no container/backdrop/body
        (function(){
            function fireClick(el){
            if(!el) return false;
            const r = el.getBoundingClientRect();
            const x = Math.floor(r.left + (r.width  || 100)/2);
            const y = Math.floor(r.top  + (r.height || 100)/2);
            const opts = {bubbles:true, composed:true, cancelable:true, clientX:x, clientY:y, view:window};
            for (const t of ['pointerdown','mousedown','pointerup','mouseup','click']) {
                try { el.dispatchEvent(new MouseEvent(t, opts)); } catch(e){}
            }
            return true;
            }
            const candidates = [
            document.querySelector('.cdk-overlay-backdrop'),
            document.querySelector('.cdk-overlay-container'),
            document.querySelector('.container.ng-star-inserted'),
            document.querySelector('main'),
            document.body,
            document.documentElement
            ].filter(Boolean);
            for (const el of candidates){ if (fireClick(el)) break; }
        })();
        })();


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

def preencher_profissao_js(timeout=10, max_attempts=3, filtro_texto="out", alvo_opcao="outros"):
    """
    Preenche a PROFISSÃO selecionando 'outros' SEM digitar o texto completo.
    Fluxo:
      - encontra input de profissão no shadow DOM: input[formcontrolname="profissao"]
      - garante foco real (JS + click) e visibilidade
      - digita 'out' lentamente (humano)
      - ArrowDown para abrir a lista
      - seleciona 'outros' no listbox correto (por aria-controls do input)
      - fallback: primeira opção, ou Enter
    Retorna: { ok: bool, value?: str, tries: int, reason?: str }
    """

    global driver

    # ---------------- helpers ----------------
    def _normalize(s):
        s = (s or "").strip().lower()
        s = unicodedata.normalize("NFD", s)
        return "".join(ch for ch in s if not unicodedata.combining(ch))

    def _type_like_human(el, text, dmin=0.05, dmax=0.11):
        # digitação “humana”
        for ch in text:
            el.send_keys(ch)
            time.sleep(random.uniform(dmin, dmax))

    def _value_of(el):
        try:
            return (driver.execute_script("return arguments[0].value || '';", el) or "").strip()
        except Exception:
            return ""

    def _get_prof_input(host_el):
        # busca input de profissão dentro do shadow
        return driver.execute_script("""
          const host = arguments[0];
          const root = host && host.shadowRoot;
          if (!root) return null;

          // prioriza exatamente formcontrolname="profissao"
          let el = root.querySelector('input[formcontrolname="profissao"]');
          if (el) return el;

          el = root.querySelector('ids-input[formcontrolname="profissao"] input');
          if (el) return el;

          el = root.querySelector('ids-combobox[formcontrolname="profissao"] input');
          if (el) return el;

          return null;
        """, host_el)

    def _ensure_focus(el):
        # traz para a tela, foca e clica; confirma foco
        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        except Exception:
            pass
        try:
            driver.execute_script("arguments[0].focus();", el)
        except Exception:
            pass
        try:
            el.click()
        except Exception:
            # último recurso: click via JS
            try:
                driver.execute_script("arguments[0].click();", el)
            except Exception:
                pass

        # confere foco (dentro do shadow o activeElement às vezes é o host; então testamos seleção de texto)
        try:
            driver.execute_script("""
              const el = arguments[0];
              try { el.setSelectionRange(0, 0); } catch(e) {}
            """, el)
        except Exception:
            pass

    def _listbox_id_for(el):
        return el.get_attribute("aria-controls")  # ex: ids-listbox-0

    def _options_for_input(el):
        lbid = _listbox_id_for(el)
        if not lbid:
            return []
        # procura apenas dentro do listbox desse input
        opts = driver.find_elements(
            By.CSS_SELECTOR,
            f'#{lbid} ids-option[role="option"], '
            f'#{lbid} .ids-option, '
            f'#{lbid} [role="option"]'
        )
        return [o for o in opts if o.is_displayed()]

    def _select_option_in_listbox(el, alvo_txt, timeout_s=5):
        alvo_norm = _normalize(alvo_txt)
        end = time.time() + timeout_s
        while time.time() < end:
            vis = _options_for_input(el)
            if vis:
                # tenta match exato por texto (normalizado)
                for o in vis:
                    if _normalize(o.text or "") == alvo_norm:
                        try:
                            o.click()
                        except Exception:
                            driver.execute_script("arguments[0].click();", o)
                        return True
                # fallback: primeira opção
                try:
                    vis[0].click()
                except Exception:
                    driver.execute_script("arguments[0].click();", vis[0])
                return True
            time.sleep(0.1)
        return False

    # ---------------- início ----------------
    # 1) host / shadow root
    host = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "mf-iparceiros-cadastrocliente"))
    )

    # 2) input de profissão
    inp = _get_prof_input(host)
    if not inp:
        return {"ok": False, "reason": "input_profissao_nao_encontrado", "tries": 0}

    # 3) tentativas completas
    for attempt in range(1, max_attempts + 1):
        try:
            # foca e garante visibilidade
            _ensure_focus(inp)
            time.sleep(0.08)

            # o campo começa limpo (como você disse); digita "out"
            _type_like_human(inp, filtro_texto)
            time.sleep(0.12)

            # abre a lista com seta para baixo
            try:
                inp.send_keys(Keys.ARROW_DOWN)
            except Exception:
                pass

            # espera o listbox desse input aparecer
            t0 = time.time()
            while time.time() - t0 < 1.8:
                if _options_for_input(inp):
                    break
                time.sleep(0.08)

            # tenta selecionar "outros" dentro do listbox CORRETO
            if _select_option_in_listbox(inp, alvo_opcao, timeout_s=4):
                time.sleep(0.18)
                val = _value_of(inp)
                if val:
                    return {"ok": True, "value": val, "tries": attempt}

            # fallback: Enter (seleciona primeira opção desse listbox que estiver aberta)
            try:
                inp.send_keys(Keys.ENTER)
                time.sleep(0.2)
            except Exception:
                pass

            val = _value_of(inp)
            if val:
                return {"ok": True, "value": val, "tries": attempt}

            # último recurso: abrir e Enter novamente
            try:
                inp.send_keys(Keys.ARROW_DOWN)
                inp.send_keys(Keys.ENTER)
                time.sleep(0.2)
            except Exception:
                pass

            val = _value_of(inp)
            if val:
                return {"ok": True, "value": val, "tries": attempt}

            # se falhou, re-obter referências (evitar stale)
            try:
                host = driver.find_element(By.CSS_SELECTOR, "mf-iparceiros-cadastrocliente")
            except Exception:
                host = WebDriverWait(driver, timeout).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "mf-iparceiros-cadastrocliente"))
                )
            inp = _get_prof_input(host)

        except StaleElementReferenceException:
            time.sleep(0.2)
            try:
                host = driver.find_element(By.CSS_SELECTOR, "mf-iparceiros-cadastrocliente")
            except Exception:
                host = WebDriverWait(driver, timeout).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "mf-iparceiros-cadastrocliente"))
                )
            inp = _get_prof_input(host)

        except Exception:
            time.sleep(0.25)

    return {"ok": False, "reason": "falha_na_selecao", "tries": max_attempts}



#3.3 - CPF e CNPJ

def preencher_pagamento_boleto_js ():
    global driver
    
    # 1. Pré-verificações (mantidas por segurança)
    if not ensure_driver_alive(driver):
        raise RuntimeError("WebDriver inválido. Recrie o driver.")
    
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "mf-iparceiros-contratacao"))
    )
    time.sleep(2)
    
    js = r"""
      return (function run(){
        try{
          // === Configurações e Guards ===
          const host = document.querySelector('mf-iparceiros-contratacao');
          if(!host) return { ok:false, step:'no-host' };
          const root = host.shadowRoot;
          if(!root) return { ok:false, step:'no-shadow' };

          // === Função de Clique em Radio Button (CORRIGIDA) ===
          function clickRadio(selector, expectedValue){
            
            // --- BLOCO RESTAURADO: Localiza o INPUT ---
            let inp = null;
            if (expectedValue) {
              // Localiza por formcontrolname="forma_pagamento" e valor (usado para Boleto)
              inp = root.querySelector(`input[formcontrolname="forma_pagamento"][value="${expectedValue}"]`);
            } else if (selector) {
              // Localiza por seletor (usado para No Resgate)
              inp = root.querySelector(selector);
            }
            // ------------------------------------------

            if(!inp) return { checked: false, value: expectedValue || selector, step: 'not-found' };
            
            // Se já estiver checado, retorna sucesso imediatamente
            if (inp.checked) return { checked: true, value: expectedValue || selector, step: 'already-checked' };

            // --- BLOCO DE CLIQUE MELHORADO: Tenta o Label, senão o Input ---
            const labelId = inp.id;
            let targetElement = inp; // O default é o input (fallback)

            if (labelId) {
                // Procura o label que está FOR esse input no Shadow Root
                const associatedLabel = root.querySelector(`label[for="${labelId}"]`);
                // Se encontrou o label, clica no label (mais stealth/robusto)
                if (associatedLabel) {
                    targetElement = associatedLabel;
                }
            }
            // -----------------------------------------------------------------

            // Tenta clicar no elemento alvo (pode ser o input ou o label)
            try { targetElement.scrollIntoView({block:'center'}); } catch(e){}
            try{
                targetElement.click(); // CLICA no elemento mais externo/visível
                // Dispara eventos no INPUT de rádio (o elemento 'inp'), independentemente de onde clicamos
                inp.dispatchEvent(new Event('input',{bubbles:true}));
                inp.dispatchEvent(new Event('change',{bubbles:true}));
            }catch(e){
                return { checked: false, value: expectedValue || selector, step: 'click-failed' };
            }
            
            // Retorna o status final de checked
            return { checked: !!inp.checked, value: expectedValue || selector, step: 'clicked' };
          }
          // === FIM da Função de Clique em Radio Button ===


          // 1. Selecionar BOLETO (tenta BB e depois BO)
          let resBoleto = clickRadio(null, 'BB');
          if (!resBoleto.checked) {
            resBoleto = clickRadio(null, 'BO');
          }

          // 2. Selecionar 'no resgate'
          
          let resResgate;
          // Localiza o H5 que contém o texto (mais robusto que querySelector aninhado)
          const h5Resgate = Array.from(root.querySelectorAll('h5')).find(el => el.textContent.trim() === 'no resgate');

          if (h5Resgate) {
              // 1. Encontra o elemento <label> pai do <h5>
              const label = h5Resgate.closest('label');
              if (label && label.htmlFor) {
                  // 2. Localiza o INPUT usando o atributo 'for' do label
                  const resgateInputSelector = `input#${label.htmlFor}`;
                  resResgate = clickRadio(resgateInputSelector, null);
              }
          }
          // Se falhou ou não encontrou pelo texto, cai no fallback (menos específico)
          if (!resResgate) {
              resResgate = clickRadio('input[formcontrolname="dados_encerramento"]', null); 
          }
          

          // === Seleção dos Checkboxes
          function check(formControlName){
            const el = root.querySelector(`input[formcontrolname="${formControlName}"]`);
            if (!el) return { checked: false, name: formControlName, step: 'not-found' };

            // Se já estiver checado, ignora
            if (el.checked) return { checked: true, name: formControlName, step: 'already-checked' };

            try{
                // Tenta encontrar o label (melhor para cliques)
                const label = root.querySelector(`label[for="${el.id}"]`);
                const target = label || el;

                // Clica no elemento (label ou input)
                target.click(); 

                // Dispara eventos no INPUT para garantir que o framework reaja
                el.dispatchEvent(new Event('input',{bubbles:true}));
                el.dispatchEvent(new Event('change',{bubbles:true}));
                
                return { checked: !!el.checked, name: formControlName, step: 'clicked' };
            }catch(e){
                return { checked: false, name: formControlName, step: 'click-failed', message: String(e) };
            }
          }

          // Execução sequencial
          let resGarantiaPrazo      = check('cienciaGarantiaPrazo');
          let resRegrasCancelamento = check('cienciaRegrasCancelamento');
          let resRegrasCRP          = check('cienciaRegrasCRP');

          // Verificação de sucesso para os checkboxes
          let checkboxesOk = resGarantiaPrazo.checked && resRegrasCancelamento.checked && resRegrasCRP.checked;

         



          return { 
              ok: resBoleto.checked && (resResgate && resResgate.checked), // Certifica que resResgate não é undefined
              status: { 
                  boleto: resBoleto, 
                  resgate: resResgate 
              }
          };
          
        }catch(e){
          // Retorna a exceção se ocorrer algum erro grave no script
          return { ok:false, step:'exception', message: String(e && e.message || e) };
        }
      })()
    """

    attempt = 0
    max_attempts=12
    # O timeout do loop será baseado no número de tentativas, não no time total
    while attempt < max_attempts:
        attempt += 1
        print(f"Tentativa {attempt}/{max_attempts} de selecionar 'Boleto' e 'No Resgate' via JS...")
        
        try:
            # Executa o JavaScript
            res = driver.execute_script(js)
            
            # 2. Verifica o resultado da execução do JS
            if res.get('ok') is True:
                print("✅ Ambos 'Boleto' e 'No Resgate' selecionados com sucesso na tentativa:", attempt)
                return {"ok": True, "step": "success", "attempts": attempt}
            
            # Se não deu OK, imprime o status detalhado (para debug)
            status_boleto = res.get('status', {}).get('boleto', {})
            status_resgate = res.get('status', {}).get('resgate', {})
            
            print(f"   Status Boleto: {'CHECKED' if status_boleto.get('checked') else 'FAILED'} (Step: {status_boleto.get('step')})")
            print(f"   Status Resgate: {'CHECKED' if status_resgate.get('checked') else 'FAILED'} (Step: {status_resgate.get('step')})")
            
            # 3. Pausa mais humana/aleatória para nova tentativa
            # Se falhou, esperamos um pouco mais antes de tentar de novo, simulando um humano
            time.sleep(random.uniform(1.0, 1.1))
            
        except Exception as e:
            print(f"⚠️ JavascriptException na tentativa {attempt}:", e)
            time.sleep(random.uniform(1, 1.5)) # Pausa maior em caso de erro grave

    # Se o loop terminar sem sucesso
    print(f"❌ Falha ao selecionar 'Boleto' e 'No Resgate' após {max_attempts} tentativas.")
    return {"ok": False, "step": "max-attempts-reached", "attempts": max_attempts, "final_result": res}




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






#=======================================================================================================================
### PRINCIPAL - Preencher 
#=======================================================================================================================


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


    
    time.sleep(0.2)

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
            action.move_to_element(btn_continuar).pause(random.uniform(0.05, 0.2)).click(btn_continuar).perform()
            human_sleep()
            print("Clicado em 'Continuar' novamente, aguardando próxima tela...")
    except Exception as e:
        
        pass  # Se der erro, ignora e continua
    
    time.sleep(1.1)



    
    #                                   Pagamento de boleto - Etapa 3
    #========================================================================================================


    print("Preenchendo dados de pagamento via boleto...")
    time.sleep(1.11)
    res3 = preencher_pagamento_boleto_js()
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
    time.sleep(1.29)
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
        ui_alert("erro", "Falha ao preencher dados pessoais Necessário revisão manual.", "info")
        return False

    time.sleep(0.09)
    #descer a tela ate o final
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(0.1)

    #Verificacao:  Confirmação extra com possível retry do preenchimento
    if not conferir_preenchimento_com_possivel_retry(driver, cliente_atual, max_reexec=1):
        print("❌ Verificação extra falhou mesmo após retry.")
        ui_alert(
            "erro",
            "Falha ao confirmar o preenchimento automático dos dados pessoais. "
            "Verifique o formulário e preencha manualmente, por favor.",
            "info"
        )
        return False


    #preencher profissão:
    res2 = preencher_profissao_js()
    print("Resultado do preenchimento da profissão:", res2)
    if not isinstance(res2, dict) or not res2.get("ok"):
        print("Falha ao preencher a profissão. Necessário revisão manual.")
        ui_alert("erro", "Falha ao preencher dados pessoais (Profissao) Necessário revisão manual.", "info")
        return False

    time.sleep(0.7)
    #--------------------------------------------------------------------------------------------------------







    ### 2.19 CONTINUAR para a próxima etapa ### (botão)

    ok_next = clicar_continuar_robusto(driver, max_tries=3, wait_next=8)
    if not ok_next:
        print("Necessário revisão manual para este cliente (não avançou).")
        ui_alert("erro", "Necessário revisão manual para este cliente (não avançou).", "info")
        return False
    
    # try:
    #     botao_continuar = shadow_root1.find_elements(By.CSS_SELECTOR, 'button[idsmainbutton]')
    #     action.move_to_element(botao_continuar[0]).pause(random.uniform(0.02, 0.2)).click(botao_continuar[0]).perform()
    #     human_sleep()
    #     print("Clicado em 'Continuar', aguardando próxima tela...")
    # except Exception as e:
        
    #     print("⚠️ Necessário revisão manual para este cliente.")
    #     return  # Sai da função sem continuar para a próxima etapa
    
    # time.sleep(1)

    # #Verificar se o botão CONTINUAR ainda está presente (o que indica que não avançou) e tentar clicar novamente
    # try:
    #     botao_continuar_check = shadow_root1.find_elements(By.CSS_SELECTOR, 'button[idsmainbutton]')
    #     if botao_continuar_check and botao_continuar_check[0].is_displayed():
    #         print("O botão 'Continuar' ainda está presente. Tentando clicar novamente...")
    #         time.sleep(1)
    #         action.move_to_element(botao_continuar_check[0]).pause(random.uniform(0.02, 0.2)).click(botao_continuar_check[0]).perform()
    #         human_sleep()
    #         print("Clicado em 'Continuar' novamente, aguardando próxima tela...")
    # except Exception as e:
        
    #     pass  # Se der erro, ignora e continua

    #========================================================================================================

    # select root do shadow DOM da próxima etapa

    time.sleep(1.9)

    # >>> Acessa o Shadow host
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "mf-iparceiros-contratacao")))
    host1 = driver.find_element(By.CSS_SELECTOR, "mf-iparceiros-contratacao")
    shadow_root2 = driver.execute_script("return arguments[0].shadowRoot", host1)
    print("Shadow DOM da etapa de contratação acessado com sucesso.")

    time.sleep(1)

    #                                   Pagamento de boleto - Etapa 3
    #========================================================================================================


    print("Preenchendo dados de pagamento via boleto...")

    res3 = preencher_pagamento_boleto_js()
    print("Resultado do preenchimento do pagamento via boleto:", repr(res3))
    if not (isinstance(res3, dict) and res3.get("ok") is True):
        print("❌ Falha ao preencher os dados de pagamento via boleto. Detalhes:", repr(res3))
        ui_alert("Erro", "Algo aconteceu de errado ao preencher o boleto, finalize manualmente ou reinicie o processo", "Info")
        raise False

    #--------------------------------------------------------------------------------------------------------




    #==========================================
    ####         Botão Contratar           ####
    #==========================================
    try:
        #selecionar botao pelo texto "contratar"
        time.sleep(0.05) # webdriverwait não funciona aqui -
        contratar = shadow_root2.find_element(By.CLASS_NAME, 'btn-contratar')
        botao_contratar = contratar.find_elements(By.TAG_NAME, 'button')
        action.move_to_element(botao_contratar[0]).pause(random.uniform(0.1, 0.2)).click(botao_contratar[0]).perform()
        # human_sleep(0.8, 1.4)
        #mover o mouse para o botão e esperar 2 segundos
        #action.move_to_element(botao_contratar[0]).pause(2).perform()
        
        print("Finalizando....")

        return True # Retorna True indicando que tudo foi preenchido com sucesso e pode prosseguir para o próximo cliente
    
    except Exception as e:
        ui_alert("Erro ao tentar localizar ou clicar no botão 'Contratar' - revise manualmente")
        return False    #Retorna False indicando que houve um problema e não deve prosseguir para o próximo cliente 





# FIM...






def _start_driver_bg():
    global driver_thread, _driver_keepalive_evt

    try:
        root.after(0, lambda: btn_iniciar_driver.configure(state="disabled"))
        iniciar_driver()  # cria o driver
        #root.after(0, lambda: messagebox.showinfo("OK", "Driver iniciado!"))
        ui_alert("OK", "Driver iniciado!", "info")
        _driver_keepalive_evt.wait()
    except Exception as e:
        msg = f"{e.__class__.__name__}: {e}"
        root.after(0, lambda: btn_iniciar_driver.configure(state="normal"))
        #root.after(0, lambda m=msg: ui_alert("Erro", m, "error"))
        root.after(0, lambda m=msg: show_modal("Erro", m, "error"))




def start_driver_thread():
    btn_iniciar_driver.configure(state="disabled", text="Iniciando...")
    t = threading.Thread(target=_thread_driver, daemon=True)
    t.start()


    
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#                                                               #   #   #   #   #       ---      Configuração da Interface Gráfica      ---    #   #   #   #   #  
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------




# ==================== Cores personalizadas ====================

cor_fundo_janela   = "#040312"
cor_fundo_frame    = "#0C1B39"
cor_botao_fundo    = "#3E8BBB"
cor_botao_hover    = "#4E7ED2"
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

def _thread_driver():
    global driver_started
    try:
        iniciar_driver()
        ui_alert("OK", "Driver iniciado!", "info")
        root.after(0, lambda: btn_inserir.configure(state="normal"))
        root.after(0, lambda: btn_buscar.configure(state="normal"))
        root.after(0, lambda: btn_iniciar_driver.configure(text="Driver iniciado"))
    except Exception as e:
        ui_alert("Erro", f"Falha ao iniciar driver:\n{e}", "error")
        root.after(0, lambda: btn_iniciar_driver.configure(
            state="normal", text="Iniciar Navegador"
        ))

def main():
    global root

    # 1) trava de instância única — ANTES de criar janela/threads
    guard = _ensure_single_instance()
    if guard is None:
        sys.exit(0)

    # 2) cria a janela
    root = tk.Tk()
    root.title("Automação Seleção de Consórcio")
    root.configure(bg=cor_fundo_janela)
    root.minsize(width=360, height=280)
    root.protocol("WM_DELETE_WINDOW", on_close)

    # 3) monta toda a UI aqui
    build_ui(root)


    # 4 loop de mensagens
    root.mainloop()

def build_ui(root):
    global btn_iniciar_driver, btn_inserir, btn_buscar

    # --- estilos ---
    style = ttk.Style()
    style.theme_use("clam")
    style.configure("Custom.TButton",
                    background=cor_botao_fundo, foreground=cor_botao_texto,
                    font=("Helvetica", 12, "bold"), borderwidth=0, padding=6)
    style.map("Custom.TButton",
              background=[("active", cor_botao_hover)],
              foreground=[("active", cor_botao_texto)])

    # --- título ---
    tk.Label(root, text="Simular e Contratar Consórcio",
             font=("Helvetica", 15, "bold"),
             bg=cor_fundo_janela, fg=cor_botao_texto).pack(pady=(30, 20))

    # === ÚNICO CONJUNTO DE BOTÕES (na ordem desejada) ===
    btn_iniciar_driver = ttk.Button(
        root, text="0 - Iniciar Navegador", style="Custom.TButton",
        command=start_driver_thread
    )
    btn_iniciar_driver.pack(padx=15, pady=8, fill="x")

    btn_inserir = ttk.Button(
        root, text="1 - Inserir dados Cliente", style="Custom.TButton",
        command=inserir_dados_cliente_js, state="disabled"
    )
    btn_inserir.pack(padx=15, pady=8, fill="x")

    btn_buscar = ttk.Button(
        root, text="2 - Buscar Consórcio", style="Custom.TButton",
        command=buscar_consorcio_cliente, state="disabled"
    )
    btn_buscar.pack(padx=15, pady=8, fill="x")

    # # === Seções apenas informativas (sem botões) ===
    # sec1 = tk.LabelFrame(root, text="Inserir CPF/CNPJ; Dt Nasc e Tp Prod. Cliente",
    #                      bg=cor_fundo_frame, fg=cor_texto_label, padx=10, pady=6)
    # sec1.pack(padx=15, pady=6, fill="x")
    # tk.Label(sec1, text="...", font=("TkDefaultFont", 9, "italic"),
    #          bg=cor_fundo_frame, fg=cor_texto_label).pack(anchor="w")

    # sec2 = tk.LabelFrame(root, text="Principal: Busca Grupo > Crédito > Preenche dados",
    #                      bg=cor_fundo_frame, fg=cor_texto_label, padx=10, pady=6)
    # sec2.pack(padx=15, pady=6, fill="x")
    # tk.Label(sec2, text="...", font=("TkDefaultFont", 9, "italic"),
    #          bg=cor_fundo_frame, fg=cor_texto_label).pack(anchor="w")

    # traz à frente
    root.after(0, root.lift)
    root.after(0, lambda: root.attributes('-topmost', True))
    root.after(100, lambda: root.attributes('-topmost', False))




if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()
    
    
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
