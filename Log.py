# --- Configuração do Logger ---
log_filename = os.path.join('Log_Replica.log')

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - [%(funcName)s:%(lineno)d] - %(message)s',
    encoding='utf-8', 
    handlers=[
        logging.FileHandler(log_filename, mode='w'), 
        
        logging.StreamHandler()
    ]
)

logging.info("Aplicação iniciada. Logger configurado.")
