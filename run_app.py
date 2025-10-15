#importa o arquivo principal
import multiprocessing as mp

# Fix global de certificados para urllib/ssl
import os, ssl, certifi
os.environ["SSL_CERT_FILE"] = certifi.where()
ssl._create_default_https_context = lambda *a, **k: ssl.create_default_context(cafile=certifi.where())

from Selecao_cotas import main

if __name__ == "__main__":
    mp.freeze_support()
    try:
        mp.set_start_method("spawn")
    except RuntimeError:
        pass
    main()
