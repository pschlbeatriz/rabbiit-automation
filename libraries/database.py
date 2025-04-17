from pymongo.mongo_client import MongoClient
from pymongo.server_api import ServerApi
from datetime import datetime, timedelta
from variables.config import MONGO_URI, env
from libraries.PrimeLogger import logger
from collections import defaultdict

class ColaboradoresRabbiit:
    def __init__(self):
        logger.log_info("Iniciando Class ColaboradoresPonto")
        uri = MONGO_URI
        self.client = MongoClient(uri, server_api=ServerApi('1'))
        self.db = self.client['bot_ponto']
        self.confronto_rabbiit = self.db['confronto_rabbiit'] if env == 'prd' else self.db['confronto_rabbiit_dev']
    
    def insere_registros(self, dados):
        logger.log_info("Inserindo registros no banco de dados")
        for record in dados:
            if not record.get('data_exec', ''):
                record['data_exec'] = datetime.now().strftime("%Y-%m-%d")
        self.confronto_rabbiit.insert_many(dados)
        logger.log_info("Registros inseridos no banco de dados")
