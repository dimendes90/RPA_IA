#app.py 

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
from selenium.webdriver.support import expected_conditions as EC # Define condições pré-definidas para a espera (ex.: elemento visível, clicável).
from selenium.common.exceptions import TimeoutException # Exceção disparada quando uma espera (WebDriverWait) atinge o tempo limite.
from selenium.common.exceptions import NoSuchElementException # Exceção disparada quando um elemento não é encontrado na página.
from selenium.common.exceptions import InvalidSessionIdException, WebDriverException # Exceções gerais relacionadas a problemas de comunicação ou estado da sessão do driver.
from selenium.webdriver.common.keys import Keys

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
    """
    app_dir = _get_app_path()
    candidate = app_dir / name
    
    if candidate.exists():
        # Verificação de segurança adicional para o .xlsx
        if candidate.suffix == '.xlsx':
            return candidate
        else:
            raise FileNotFoundError(f"Arquivo '{name}' encontrado, mas não é um .xlsx válido.")
    
    raise FileNotFoundError(
        f"Arquivo '{name}' não encontrado. O arquivo deve estar na mesma pasta do executável:\n"
        f"- {str(app_dir)}"
    )






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

# def _executable_dir() -> Path: -> OLD
#     if getattr(sys, 'frozen', False):
#         return Path(sys.executable).resolve().parent  # .../SelecionaCotas.app/Contents/MacOS
#     return Path(__file__).resolve().parent

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


# - Load DataFrame de clientes
def load_df_clientes():
    global df_atual, df_clientes
    arq = _find_user_file('base_clientes.xlsx')  # <— em vez de Path('base_clientes.xlsx')
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




#=======================================================================================================================
#                  FUNÇÃO 0 - Iniciar Driver com as configurações iniciais
#=======================================================================================================================
# Função INICIAL - iniciar driver com perfil, user-agent e stealth

# #Google Chrome Version 141 (testar sem versão 141)
def iniciar_driver():
  global driver, wait, action, driver_started

  if driver_started and driver is not None and driver_ativo(driver):
      print("Info: O navegador já está iniciado.")
      return

  options = uc.ChromeOptions()

  # MANTER o mesmo PROFILE_DIR (use sua constante PROFILE_DIR)
  # isso preserva cookies, extensões e comportamento "real" entre execuções
  options.add_argument(f"--user-data-dir={PROFILE_DIR}")
  options.add_argument(f"--user-agent={USER_AGENT}")

  # Stealth / invisibilidade (mantidos do original)
  options.add_argument("--disable-blink-features=AutomationControlled")
  options.add_argument("--lang=pt-BR")
  options.add_argument("--no-sandbox")
  options.add_argument("--disable-dev-shm-usage")
  options.add_argument("--disable-infobars")

  # Limites de processos e ajustes para reduzir consumo de RAM
  options.add_argument("--renderer-process-limit=3")
  options.add_argument("--process-per-site")
  options.add_argument("--disable-site-isolation-trials")
  options.add_argument("--disk-cache-size=1")  # reduz cache em disco

  # Menos "ruído" em background (pouco impacto em stealth)
  options.add_argument("--disable-background-networking")
  options.add_argument("--disable-extensions")
  options.add_argument("--disable-sync")

  ### > verificar se NAO atraplaha a invisibilidade
  options.add_argument("--disable-cache")
  options.add_argument("--disk-cache-size=0")
  options.add_argument("--media-cache-size=0")
  options.add_argument("--disable-application-cache")

  # Caminho explícito do Chrome (macOS)
  options.binary_location = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"

  print("Iniciando o Chrome com undetected-chromedriver (version_main=141)...")

  try:
      # use_subprocess=False é preferível no macOS + onefile
      driver = uc.Chrome(
          options=options,
          version_main=141,       # força a major correta do Chrome (sua versão: 141.x)
          use_subprocess=False,
          headless=False
      )
      driver_started = True
      wait = WebDriverWait(driver, 15)
      action = ActionChains(driver)

      # janela realista
      try:
          driver.set_window_size(1199, 889)
      except Exception:
          pass

      # pequeno delay humano
      human_sleep(0.1, 1.1)

      # stealth JS injetado (mantém invisibilidade)
      driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
          "source": """
              try {
                Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
              } catch(e) {}
              window.chrome = window.chrome || { runtime: {} };
              try {
                Object.defineProperty(navigator, 'plugins', { get: () => [1,2,3,4,5] });
              } catch(e) {}
              try {
                Object.defineProperty(navigator, 'languages', { get: () => ['pt-BR','pt','en-US','en'] });
              } catch(e) {}
          """
      })
      # override UA via CDP (faça em try/except)
      try:
          driver.execute_cdp_cmd("Network.setUserAgentOverride", {"userAgent": USER_AGENT})
      except Exception:
          pass

      # Limpar cache/cookies iniciais para não inflar memória
      try:
          driver.execute_cdp_cmd("Network.enable", {})
          driver.execute_cdp_cmd("Network.clearBrowserCache", {})
          driver.execute_cdp_cmd("Network.clearBrowserCookies", {})
      except Exception:
          pass

      # Teste rápido de navegação
      driver.get("https://www.google.com")
      print("Título inicial:", driver.title)

      # salvar cookies (se quiser)
      try:
          save_cookies(driver, COOKIES_FILE)
      except Exception:
          pass

      print("Fluxo finalizado sem exceções aparentes")
      #status_var.set("Driver iniciado.")
      #ui_alert("OK", "Driver iniciado!", "info")
      #toast("Driver iniciado.", "success")
      
      return driver

  except Exception as exc:
      # Log completo para debugging; NÃO force driver.quit() pois driver pode não existir
      import traceback, sys
      print("Erro ao criar/iniciar driver:")
      traceback.print_exc(file=sys.stdout)
      # re-raise para o thread chamar a UI com erro (se desejar)
      raise
  






#=======================================================================================================================
#                  FUNÇÃO EXTRA - Coletar e salvar todos os grupos disponíveis em CSV
#=======================================================================================================================


# def guardar_grupos_disponiveis():
#     global driver, action, APP_DATA
#     print("Iniciando a coleta dos grupos disponíveis...")

#     # coleta
#     grupos_coletados = []

#     # Tipo de consórcio selecionado
#     tipo_consorcio = driver.find_element(
#         By.XPATH, '//div[@role="combobox" and @id="codigoProduto"]'
#     ).get_attribute('innerText').strip().lower()

#     # normaliza para nome de arquivo
#     tipo_consorcio = _strip_accents(tipo_consorcio)                  # remove acentos
#     tipo_consorcio = re.sub(r'[^\w\s-]', '', tipo_consorcio)         # remove tudo que não for letra/número/_
#     tipo_consorcio = tipo_consorcio.strip().replace(' ', '_')         # troca espaços por _

#     # Expandir para 50 linhas
#     botao_linhas = driver.find_element(By.XPATH, '//div[@role="combobox" and @id="pageSizeId"]')
#     action.move_to_element(botao_linhas).pause(random.uniform(0.1, 0.5)).click(botao_linhas).perform()
#     human_sleep(0.5, 1.5)
#     opcao_50 = driver.find_element(By.XPATH, '//span[@class="ids-option__text" and text()="50"]')
#     action.move_to_element(opcao_50).pause(random.uniform(0.1, 0.5)).click(opcao_50).perform()
#     human_sleep(1, 3)
#     print("Tabela atualizada para mostrar 50 linhas.")

#     while True:
#         tabela = driver.find_element(By.XPATH, '//*[@aria-describedby="tabelaGrupos"]')
#         linhas_tabela = tabela.find_elements(By.XPATH, './/tbody/tr')

#         for linha in linhas_tabela:
#             colunas = linha.find_elements(By.TAG_NAME, 'td')
#             botao_grupo = colunas[0].find_element(By.TAG_NAME, 'button')
#             numero_grupo = (botao_grupo.text or '').strip()
#             if numero_grupo:
#                 grupos_coletados.append(str(numero_grupo))

#         # próxima página?
#         try:
#             botao_proxima_pagina = driver.find_element(By.XPATH, '//button[@id="nextPageId"]')
#             if botao_proxima_pagina.get_attribute("disabled"):
#                 print("Fim das páginas. Nenhuma próxima disponível.")
#                 break
#             action.move_to_element(botao_proxima_pagina).pause(random.uniform(0.2, 0.6)).click(botao_proxima_pagina).perform()
#             human_sleep(1.5, 3)
#             print("Indo para a próxima página...")
#         except Exception as e:
#             print("Não foi possível ir para a próxima página:", e)
#             break

#     # remove duplicados desta sessão
#     grupos_coletados = list(dict.fromkeys(grupos_coletados))  # preserva ordem e unicidade
#     print(f"Coleta concluída. Total únicos nesta sessão: {len(grupos_coletados)}")
#     print(grupos_coletados)

#     # ===== salvamento sem duplicar =====
#     APP_DATA.mkdir(parents=True, exist_ok=True)
#     nome_arquivo = APP_DATA / f'grupos_{tipo_consorcio}.csv'

#     existentes_set = set()
#     if nome_arquivo.exists():
#         try:
#             df_existentes = pd.read_csv(nome_arquivo, sep=';', dtype={'grupo': str})
#             # normaliza (tira espaços e NaN)
#             df_existentes['grupo'] = df_existentes['grupo'].astype(str).str.strip()
#             existentes_set = set(df_existentes['grupo'].dropna().tolist())
#         except Exception as e:
#             print(f"Aviso: não foi possível ler o CSV existente ({e}). Vou recriar do zero.")

#     # filtra apenas novos
#     novos = [g for g in grupos_coletados if g not in existentes_set]

#     if nome_arquivo.exists() and len(novos) > 0:
#         # append sem header
#         pd.DataFrame(novos, columns=['grupo']).to_csv(nome_arquivo, sep=';', index=False, mode='a', header=False)
#         print(f"Acrescentados {len(novos)} novos grupos (sem duplicar).")
#     elif not nome_arquivo.exists():
#         # cria arquivo com header
#         pd.DataFrame(grupos_coletados, columns=['grupo']).to_csv(nome_arquivo, sep=';', index=False)
#         print(f"Arquivo criado com {len(grupos_coletados)} grupos.")
#     else:
#         print("Nenhum grupo novo para acrescentar (todos já estavam salvos).")

#     print(f"Grupos salvos em: {nome_arquivo}")


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





#=======================================================================================================================
#                  FUNÇÃO 2 - PRINCIPAL - Buscar consórcio do cliente e selecionar a melhor opção
#=======================================================================================================================


def buscar_consorcio_cliente():
    global grupo_encontrado, cpf_atual, cliente_atual
    global driver
    
    # === carregar lista de grupos a ignorar ===
    tp = str(cliente_atual['tp_produto'] or '').strip().lower()
    try:
        csv_path = _find_user_file(f'grupos_{_normalize_tipo(tp)}.csv')
        print(f"Usando lista de grupos: {csv_path}")
        df_grupos_ignorar = pd.read_csv(csv_path, sep=';', dtype={'grupo': str})
        list_grupos = df_grupos_ignorar['grupo'].astype(str).str.strip().tolist()
    except FileNotFoundError:
        print(f"⚠️ Não encontrei grupos para '{tp}'. Vou seguir sem ignorar nenhum grupo.")
        list_grupos = []

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

    # Expandir para 50 linhas
    botao_linhas = driver.find_element(By.XPATH, '//div[@role="combobox" and @id="pageSizeId"]')
    action.move_to_element(botao_linhas).pause(random.uniform(0.1, 0.2)).click(botao_linhas).perform()
    human_sleep(0.1, 0.2)
    opcao_50 = driver.find_element(By.XPATH, '//span[@class="ids-option__text" and text()="50"]')
    action.move_to_element(opcao_50).pause(random.uniform(0.1, 0.2)).click(opcao_50).perform()
    human_sleep(0.1, 0.2)
    print("Tabela atualizada para mostrar 50 linhas.")


    grupo_localizado = True
    ### - LOOP Principal - Roda ate ACHAR o grupo (muda para false <-***
    while grupo_localizado:
      print("inciando processo....")
      ### Selecionar Tabela e interagir com dropdowns - Grupos ### ### ###
      
      WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH, '//*[@aria-describedby="tabelaGrupos"]')))
      tabela = driver.find_element(By.XPATH, '//*[@aria-describedby="tabelaGrupos"]') # localizar a tabela
      linhas_tabela = tabela.find_elements(By.XPATH, './/tbody/tr') # localizar todas as linhas da tabela, exceto o cabeçalho
      human_sleep(0.2, 0.4)
      print(len(linhas_tabela))
      #===========================================
      ##########  1 - Busca de GRUPO #############
      #===========================================

      # Percorrer as linhas da tabela
      for linha in linhas_tabela:
          time.sleep(0.04)
          try:
            colunas = linha.find_elements(By.TAG_NAME, 'td')
          except Exception as e:
              time.sleep(2)
              colunas = linha.find_elements(By.TAG_NAME, 'td')
          
          botao_grupo = colunas[0].find_element(By.TAG_NAME, 'button')
          numero_grupo = botao_grupo.text.strip()
          print(f"Número do grupo: {numero_grupo}")
          
          #Verificar se o grupo esta na Lista, se não tiver pula pata o proximo
          if numero_grupo not in list_grupos:
              print(f"Grupo {numero_grupo} NÃO está na lista de grupos. Pulando...")
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
                          # - LOOP DE VERIFICAÇÃO
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

                      #Alterando booleano para assegurar que o loop principal nao vai rodar
                      grupo_localizado = False

                      #===================================================
                      #===================================================
                      #===================================================
                      # Chamar a função para preencher os dados pessoais
                      
                      if str(cliente_atual.get('tipo_cliente') or '').strip().upper() == 'CPF':
                        sucesso_preenchimento = preencher_dados_pessoais()

                      else:
      
                        sucesso_preenchimento = preencher_dados_PJ()

                      #Checar se o preenchimento foi bem sucedido e atualizar o status no CSV
                      if sucesso_preenchimento:
                          print("✅ Dados pessoais preenchidos com sucesso.")
                          atualizar_status_cliente(cpf_atual, "Finalizado")
                          #status_var.set("Finalizado.")
                          ui_alert("Finalizado", "Infos Preenchidas - Clique em CONTRATAR!", "info")


                          #toast("Infos preenchidas — clique em CONTRATAR!", "success", timeout=6000)

                      else:
                          
                          print("❌ Falha ao preencher os dados pessoais.")
                          atualizar_status_cliente(cpf_atual, "Erro")
                          #status_var.set("Erro")
                          ui_alert("erro", "Ops, algo deu errado, verifique e preencha manualmente ou reinicie o processo!", "info")

                          #toast("Algo deu errado, termine manualmente ou reiniciei o processo", "Ops", timeout=6000)


                      #===================================================
                      #===================================================
                      #===================================================

                      break
                  
              # Fim do loop de busca por grupos
          else:
              print(f"❌ Nenhuma linha de crédito foi encontrada com parcela menor ou igual a R$ {valor_maximo_float}.")
              #garante que a variavel para manter o loop principal e True
              grupo_localizado = True

          # Se grupo Não encontrado, clicar em voltar e tentar o próximo grupo
          if not grupo_encontrado:
              print(f"⚠️ Nenhuma opção válida encontrada no grupo {numero_grupo}. Voltando para a lista de grupos...")
              botao_voltar = driver.find_element(By.XPATH, '//p[contains(text(), " voltar para grupos")]')
              action.move_to_element(botao_voltar).pause(random.uniform(0.2, 0.7)).click(botao_voltar).perform()
              human_sleep(0.8, 1.4)

              #garante que a variavel para manter o loop principal e True
              grupo_localizado = True
              print("Recarregando a lista de grupos após voltar...")
              break 

          else:
              print("✅ Grupo e crédito selecionados com sucesso. Saindo do loop de grupos.")
              #Alterando booleano para assegurar que o loop principal nao vai rodar
              grupo_localizado = False
              break  # Sai do loop de grupos se um grupo válido foi encontrado      

        
      # fim loop de grupos

      ### >>> Pular Pagina >>>

      print("Nao encontrado na pagina indo para a proxima")

      # próxima página? - double check se existe ou se desativado
      try:
          botao_proxima_pagina = driver.find_element(By.XPATH, '//button[@id="nextPageId"]')
          botao_prox_pagina_existe = True
      except Exception as e:
          print("Não foi possível ir para a próxima página:")
          human_sleep(1.5, 2.5)
          botao_prox_pagina_existe = False



      if botao_prox_pagina_existe:
        
          if botao_proxima_pagina.get_attribute("disabled"):
              print("Botao desativaado - Voltando para Pg1")
              #Voltar para a Pagina (ir para a pagina)
              ir_para_pg = driver.find_element(By.CSS_SELECTOR, 'input[aria-label="ir para a página"]')
              action.move_to_element(ir_para_pg)\
                          .pause(random.uniform(0.2, 0.3))\
                          .click(ir_para_pg)\
                          .pause(random.uniform(0.1, 0.3))\
                          .send_keys(Keys.CONTROL + 'a')\
                          .send_keys(Keys.DELETE)\
                          .send_keys('1')\
                          .perform()
              action.move_to_element(ir_para_pg).send_keys(Keys.ENTER).perform()
              #Verificar se botao Filtrar esta ativo:
              # Filtrar 

              human_sleep(1, 1.5)
              print("verificando botão Filtrar - inicar tudo novamente")
              btn_filtrar = driver.find_element(By.ID, 'btnFiltrar')

              if btn_filtrar.is_enabled():
                  print("Botão 'Filtrar' está ATIVO. Clicando e voltando para o loop")
                  action.move_to_element(btn_filtrar).pause(random.uniform(0.1, 0.2)).click(btn_filtrar).perform()
                  human_sleep(5, 6)
                  
              else:
                  print("Botão 'Filtrar' está INATIVO (desabilitado).")
                  #status_var.set("reCAPTCHA")
                  ui_alert("reCAPTCHA", "Resolva o reCAPTCHA e clique em Buscar Consorcio!", "info")

                  #toast("reCAPTCHA", "Resolva o reCAPATCHA e clique em Buscar Consorcio Novamente", "reCAPTCHA", timeout=6000)
                  break 

          else:
              #ATIVO SEGUINDO para PROXIMA PAGINA
              action.move_to_element(botao_proxima_pagina).pause(random.uniform(0.1, 0.2)).click(botao_proxima_pagina).perform()
              human_sleep(1, 1.5)
              print("Indo para a próxima página...")


      else:
          print("Botao nao existe - Voltando pagina 1 ")
          #Voltar para a Pagina (ir para a pagina)
          ir_para_pg = driver.find_element(By.CSS_SELECTOR, 'input[aria-label="ir para a página"]')
          action.move_to_element(ir_para_pg)\
                      .pause(random.uniform(0.2, 0.3))\
                      .click(ir_para_pg)\
                      .pause(random.uniform(0.1, 0.3))\
                      .send_keys(Keys.CONTROL + 'a')\
                      .send_keys(Keys.DELETE)\
                      .send_keys('1')\
                      .perform()
          action.move_to_element(ir_para_pg).send_keys(Keys.ENTER).perform()
          
          #Verificar se botao Filtrar esta ativo:
          # Filtrar 

          human_sleep(1, 1.5)
          print("verificando botão Filtrar - inicar tudo novamente")
          btn_filtrar = driver.find_element(By.ID, 'btnFiltrar')

          if btn_filtrar.is_enabled():
              print("Botão 'Filtrar' está ATIVO. Clicando e voltando para o loop")
              action.move_to_element(btn_filtrar).pause(random.uniform(0.2, 0.5)).click(btn_filtrar).perform()
              human_sleep(4.5, 5.5)
              
          else:
              print("Botão 'Filtrar' está INATIVO (desabilitado).")
              #status_var.set("Erro")
              #toast("reCAPTCHA", "Resolva o reCAPATCHA e clique em Buscar Consorcio Novamente", "reCAPTCHA", timeout=6000)
              ui_alert("reCAPTCHA", "Resolva o reCAPTCHA e clique em Buscar Consorcio!", "info")
              break         





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

def preencher_pagamento_boleto_js (timeout=30, max_attempts=3):
    global driver
    
    # 1. Pré-verificações (mantidas por segurança)
    if not ensure_driver_alive(driver):
        raise RuntimeError("WebDriver inválido. Recrie o driver.")
    
    WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "mf-iparceiros-contratacao"))
    )
    time.sleep(0.2)
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
            time.sleep(random.uniform(1.0, 2.5))
            
        except Exception as e:
            print(f"⚠️ JavascriptException na tentativa {attempt}:", e)
            time.sleep(random.uniform(1.5, 3.0)) # Pausa maior em caso de erro grave

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
    time.sleep(0.1)
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
    root.minsize(width=400, height=300)
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
