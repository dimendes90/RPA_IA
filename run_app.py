# run_app.py / launch 

import os
import ssl
import certifi
import multiprocessing as mp
import sys # Importe sys para checar a plataforma

# Configurações de certificado (OK)
os.environ["SSL_CERT_FILE"] = certifi.where()
ssl._create_default_https_context = lambda *a, **k: ssl.create_default_context(cafile=certifi.where())

# Suprimir logs de multiprocessing (OK)
try:
    import multiprocessing.util, logging
    multiprocessing.util.log_to_stderr().setLevel(logging.ERROR)
except Exception:
    pass


if __name__ == "__main__":
    # 1. Trava de segurança do PyInstaller 
    #    Deve ser a primeira coisa no _main_.
    mp.freeze_support()
    
    # 2. Forçar 'spawn' (essencial no macOS para evitar "fork bomb")
    #    Isso DEVE vir depois do freeze_support() e ANTES
    #    de importar qualquer coisa que use multiprocessing (como o app.py)
    if sys.platform == 'darwin': # 'darwin' é o nome do kernel do macOS
        try:
            mp.set_start_method('spawn', force=True)
            print("Modo 'spawn' do multiprocessing ativado.")
        except (RuntimeError, ValueError) as e:
            # Isso pode acontecer se o script for executado de novo
            # (o que não deve, mas é bom previnir)
            print(f"Aviso ao configurar 'spawn' (ignorar se for processo-filho): {e}")
    
    # 3.   IMPORTE e execute o aplicativo principal.
    try:
        from app import main # <-- Selecao_cotas.py
        main()
    except Exception as e:
        # Se a importação ou execução falhar, mostre um erro
        print(f"Erro ao carregar ou executar o app: {e}")
        #  Mostrar erro em uma janela se o app falhar ao carregar
        if 'tk' not in sys.modules:
            try:
                import tkinter as tk
                from tkinter import messagebox
                root = tk.Tk()
                root.withdraw() # Não mostrar a janela principal do tk
                messagebox.showerror("Erro Crítico", f"Falha ao iniciar o aplicativo:\n{e}")
                root.destroy()
            except ImportError:
                pass # Se o tkinter falhar, o print no console é o fallback
        sys.exit(1) # Sair com código de erro
