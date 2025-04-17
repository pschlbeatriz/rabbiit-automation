import os
from datetime import datetime, timedelta
from dotenv import load_dotenv
from pathlib import Path
import socket
import shutil

env = 'dev'
load_dotenv()

################# RPA ###########################
HOSTNAME = socket.gethostname() if env == 'prd' else 'Local'
ROBOT_NAME = 'bot-confronto-rabbiit'

################# Default ###########################
ROOT_RPA = Path(os.path.dirname(os.path.abspath(__file__))).parent

############ Timeout ###############################
DEFAULT_SELENIUM_TIMEOUT = '40 seconds'
DEFAULT_DOWNLOAD_TIMEOUT = '60 seconds'
DEFAULT_KEYWORD_TIMEOUT = '60 seconds'
SMALL_TIMEOUT = 5

############  Diretórios Locais   ###################
if not os.path.exists(os.path.join(ROOT_RPA, 'PrintScreen')):
    os.makedirs(os.path.join(ROOT_RPA, 'PrintScreen'))
SCREENSHOT_DIRECTORY = os.path.join(ROOT_RPA, 'PrintScreen')

if not os.path.exists(os.path.join(ROOT_RPA, "Download")):
    os.makedirs(os.path.join(ROOT_RPA, "Download"))
else:
    shutil.rmtree(os.path.join(ROOT_RPA, "Download"))
    os.makedirs(os.path.join(ROOT_RPA, "Download"))
DOWNLOAD_DIRECTORY = os.path.join(ROOT_RPA, "Download")

if not os.path.exists(os.path.join(ROOT_RPA, "outputs")):
    os.makedirs(os.path.join(ROOT_RPA, "outputs"))
OUTPUT_DIRECTORY = os.path.join(ROOT_RPA, "outputs")

if not os.path.exists(os.path.join(ROOT_RPA, "resultados")):
    os.makedirs(os.path.join(ROOT_RPA, "resultados"))
RESULTADOS_DIRECTORY = os.path.join(ROOT_RPA, "resultados")

############ MONGO ##################
MONGO_URI = os.getenv('MONGO_URI')

########### RABBIIT ##############
EMAIL_RABBIIT = os.getenv('EMAIL_RABBIIT')
PASSWORD_RABBIIT = os.getenv('PASSWORD_RABBIIT')
ACCOUNT_ID_RABBIIT = os.getenv('ACCOUNT_ID_RABBIIT')
ARQ_COLAB = os.path.join("resources", "Colab_Gestores.xlsx")

######### MS GRAPH #################
TENANT_ID = os.getenv('TENANT_ID')
CLIENT_ID = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
SENDER_EMAIL = os.getenv('SENDER_EMAIL')
SENDER_PWD = os.getenv('SENDER_PWD')
EMAIL_DIRETOR = "john.doe@example.com.br"
EMAIL_DIRETOR2 = "jane.doe@example.com.br"
EMAIL_TESTE = "tommy.tester@example.com.br"
EMAILS_METRICAS = "sam.sample@example.com.br,mary.jane@example.com.br"
EMAILS_SUS = "tommy.tester@example.com.br"

#### Retentativas ####
RETENTATIVAS = 3
