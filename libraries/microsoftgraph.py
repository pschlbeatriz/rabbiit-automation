import requests
from msal import ConfidentialClientApplication
from variables.config import TENANT_ID, CLIENT_ID, CLIENT_SECRET, SENDER_EMAIL, SENDER_PWD, EMAIL_DIRETOR, EMAIL_TESTE, EMAILS_METRICAS, EMAILS_SUS, env
from datetime import datetime, timedelta
from libraries.PrimeLogger import logger
import traceback
import os
import base64
from libraries.rabbiit import ApiRabbiit
from collections import defaultdict

class MSGraphBase:
    def __init__(self):
        self.tenant_id = TENANT_ID
        self.client_id = CLIENT_ID
        self.client_secret = CLIENT_SECRET
        self.graph_base_url = 'https://graph.microsoft.com/v1.0'
        self.scope = ['https://graph.microsoft.com/.default']
        self.authority = f"https://login.microsoftonline.com/{self.tenant_id}"
        self.app = ConfidentialClientApplication(
            self.client_id,
            client_credential=self.client_secret,
            authority=self.authority
        )
        self.subject = "Atenção: Divergência nas marcações – Rabbiit x PontoMais"
        self.subject_metricas = "Relatório de Resultados - Bot Confronto PontoMais x Rabbiit"
        self.sender = SENDER_EMAIL
        self.sender_pwd = SENDER_PWD
        self._limpa_cache_autenticacao()
        self.token = self._get_token_user_pass()
        self.emails_metricas = EMAILS_METRICAS
        self.emails_sus = EMAILS_SUS
        self.my_id = self._get_user_id(self.sender)
        self.date_body = ""
    
    def _limpa_cache_autenticacao(self):
        logger.log_info("Limpando cache de autenticacao")
        accounts = self.app.get_accounts()
        for account in accounts:
            self.app.remove_account(account)
    
    def _get_token(self):
        logger.log_info("Capturando token")
        result = self.app.acquire_token_for_client(scopes=self.scope)
        if 'access_token' in result:
            return result['access_token']
        else:
            raise Exception(f"Erro ao obter token de acesso: {result.get('error_description', 'Erro desconhecido')}")
        
    def _get_token_user_pass(self):
        logger.log_info("Capturando token")
        result = self.app.acquire_token_by_username_password(
            username=self.sender,
            password=self.sender_pwd,
            scopes=self.scope
        )
        if 'access_token' in result:
            return result['access_token']
        else:
            raise Exception(f"Erro ao obter token de acesso: {result.get('error_description', 'Erro desconhecido')}")
        
    def _get_user_id(self, user_email=None):
        logger.log_info("Capturando ID do usuario")
        if user_email:
            response = requests.get(
                f"{self.graph_base_url}/users/{user_email}",
                headers={"Authorization": f"Bearer {self.token}"}
            )
            user_data = response.json()
            if "error" in user_data:
                raise Exception(f"Erro ao obter informações do usuário: {user_data['error']['message']}")
            return user_data['id']
        else:
            response = requests.get(
                f"{self.graph_base_url}/me",
                headers={"Authorization": f"Bearer {self.token}"}
            )
            user_data = response.json()
            if "error" in user_data:
                raise Exception(f"Erro ao obter informações do usuário: {user_data['error']['message']}")
            return user_data['id']
    
    def _monta_paragraphic(self, data, hora_pontomais, hora_rabbiit, nome):
        divergencias_por_nome = defaultdict(list)

        for nome_colab, date, hora_pm, hora_rb in zip(nome[0], data, hora_pontomais, hora_rabbiit):
            divergencias_por_nome[nome_colab].append((date, hora_pm, hora_rb))

        paragrafo_divergencia = []
        for nome_colab, divergencias in divergencias_por_nome.items():
            paragrafo_divergencia.append(f"<br><p><b>{nome_colab}</b>:</p>")
            datas, horas_pm, horas_rb = divergencias[0]
            for date, hora_pm, hora_rb in zip(datas, horas_pm, horas_rb):
                hr_pm = datetime.strptime(hora_pm, "%H:%M:%S")
                hr_rb = datetime.strptime(hora_rb, "%H:%M:%S")
                diferenca = max(hr_pm, hr_rb) - min(hr_pm, hr_rb)
                paragrafo_divergencia.append(
                    f"<ul><li>{datetime.strptime(date, '%Y-%m-%d').strftime('%d/%m/%Y')}, Marcação PontoMais: {hora_pm} | Marcação Rabbiit: {hora_rb} <b>[Diferença: {diferenca}]</b></li></ul>"
                )
        return paragrafo_divergencia
    
    def _monta_message(self, data, hora_pontomais, hora_rabbiit, nome, lider=False, nome_lider=None, coordenador=False, nome_coordenador=None):
        logger.log_info("Montando mensagem de alerta")

        if lider or coordenador: 
            alertas = list(zip(data, hora_pontomais, hora_rabbiit))

            if coordenador:
                paragrafo_divergencia = self._monta_paragraphic(data, hora_pontomais, hora_rabbiit, nome)
                return f'''
                    <p style="color: cian;"><b>Atenção: Divergência nas marcações – Rabbiit x PontoMais</b></p>
                    <br>
                    <p>Olá, {nome_coordenador}!</p>
                    <p>Identificamos divergências nas marcações de ponto dos seguintes colaboradores do seu time, conforme registrado nos sistemas PontoMais e Rabbiit:</p>
                    <br>
                    {''.join(paragrafo_divergencia)}
                    <br>
                    <p><b>Ação necessária:</b></p>
                    <p>Solicitamos que oriente os colaboradores para que realizem corretamente as marcações de ponto em ambos os sistemas. Caso necessário, você pode ajustar as marcações diretamente no Rabbiit.</p>
                    <br>
                    <p><b>Prazo:</b></p>
                    <p>Os ajustes devem ser realizados até às <b>10h da manhã do 2º dia útil do mês.</b></p>
                    <p>Contamos com sua colaboração para manter as informações alinhadas e evitar impactos nos registros de jornada e fluxo de faturamento.</p>
                    <br>
                    <br>
                    <p><i>Esta é uma mensagem automática, favor não responder.</i></p>
                '''
            elif lider:
                paragrafo_divergencia = paragrafo_divergencia = self._monta_paragraphic(data, hora_pontomais, hora_rabbiit, nome)
                return f'''
                    <p style="color: cian;"><b>Atenção: Divergência nas marcações – Rabbiit x PontoMais</b></p>
                    <br>
                    <p>Olá, {nome_lider}!</p>
                    <p>Identificamos divergências nas marcações de ponto dos seguintes colaboradores do seu time, conforme registrado nos sistemas PontoMais e Rabbiit:</p>
                    <br>
                    <hr>
                    {''.join(paragrafo_divergencia)}  
                    <br>         
                    <hr>
                    <br>
                    <p><b>Ação necessária:</b></p>
                    <p>Solicitamos que oriente os colaboradores para que realizem corretamente as marcações de ponto em ambos os sistemas. Caso necessário, você pode ajustar as marcações diretamente no Rabbiit.</p>
                    <br>
                    <p><b>Prazo:</b></p>
                    <p>Os ajustes devem ser realizados até às <b>10h da manhã do 2º dia útil do mês.</b></p>
                    <p>Contamos com sua colaboração para manter as informações alinhadas e evitar impactos nos registros de jornada e fluxo de faturamento.</p>
                    <br>
                    <br>
                    <p><i>Esta é uma mensagem automática, favor não responder.</i></p>
                '''
        else:    
            alertas = list(zip(data, hora_pontomais, hora_rabbiit))
            primeiros_dias_uteis = ApiRabbiit()._get_util_days() 
            msg_limite = '''<br>
            <p><b>Importante: Hoje é o último dia para realizar os ajustes das marcações divergentes!</b></p>'''

            #  &nbsp; por conta que o teams não reconhece o CSS necessário para tamanho de margem
            paragrafo_divergencia = [f'<ul><li><b>{data}, Marcação PontoMais: {hora_pontomais} | Marcação Rabbiit: {hora_rabbiit}'\
                                    f'{";" if i < len(alertas)-1 else "."}</b></li></ul>' for i, (data, hora_pontomais, hora_rabbiit) in enumerate(alertas)]  # realiza o looping na lista de alertas, no último insere um ponto final.
            return f'''
                <p><b>Atenção: Divergência nas marcações de ponto</b></p>
                <br>
                Olá, {nome.split(' ')[0].capitalize()}!</b>
                <br>
                <p>Identificamos divergências entre as marcações registradas no PontoMais e no Rabbiit nas seguintes datas:</p>        
                <br>
                {''.join(paragrafo_divergencia)}
                <br>
                <p>Para garantir a conformidade das informações, solicitamos que revise e corrija as marcações no sistema o quanto antes.</p>
                {msg_limite if datetime.now().strftime("%d") == primeiros_dias_uteis[0] else ""}
                <p>Caso tenha qualquer dúvida, perceba alguma inconsistência ou precise de ajuda, não hesite em procurar seu líder ou coordenador para esclarecer.</p>
                <br>
                <br>
                <p><i>Esta é uma mensagem automática, favor não responder.</i></p>
            '''
        

class MSGraphTeams(MSGraphBase):
    def __init__(self):
        super().__init__()
    
    def _monta_chat_body(self, user_id):
        logger.log_info("Montando Chat Body Teams")
        chat_body = {
            "chatType": "oneOnOne",
            "members": [
                {
                    "@odata.type": "#microsoft.graph.aadUserConversationMember",
                    "roles": ["owner"],
                    "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{self.my_id}')"
                },
                {
                    "@odata.type": "#microsoft.graph.aadUserConversationMember",
                    "roles": ["owner"],
                    "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{user_id}')"
                }
            ]
        }
        return chat_body
    
    def _send_message_to_teams(self, 
                    data, 
                    hora_pontomais, 
                    hora_rabbiit, 
                    nome, 
                    user_email,
                    lider=False, 
                    nome_lider=None, 
                    coordenador=False, 
                    nome_coordenador=None):
        """
        Envia mensagem no teams
        Args:
            data (str): Data do alerta
            hora_pontomais (str): Hora do ponto mais
            hora_rabbiit (str): Hora do rabbiit
            nome (str): Nome do colaborador
            user_email (str): E-mail do colaborador
            lider (bool): Se é lider ou não
            nome_lider (str): Nome do lider
            coordenador (bool): Se é coordenador ou não
            nome_coordenador (str): Nome do coordenador
        """
        user_email = user_email if env == "prd" else EMAIL_TESTE
        user_id = self._get_user_id(user_email)
        chat_body = self._monta_chat_body(user_id)
        chat_response = requests.post(
            f"{self.graph_base_url}/chats",
            headers={"Authorization": f"Bearer {self.token}", "Content-Type": "application/json"},
            json=chat_body
        )
        logger.log_info(f"Chat criado. Response: {chat_response.status_code}")
        if chat_response.status_code != 201:
            raise Exception(f"Erro ao criar chat: {chat_response.json()}")

        chat_id = chat_response.json()['id']

        # Criar a mensagem no chat
        message = self._monta_message(data, hora_pontomais, hora_rabbiit, nome, lider, nome_lider, coordenador, nome_coordenador)
        message_body = {
                "body": {
                    "contentType": "html",
                    "content": message
                }
            }
        
        post_message_url = f"{self.graph_base_url}/chats/{chat_id}/messages"
        message_response = requests.post(
            post_message_url,
            headers={"Authorization": f"Bearer {self.token}", "Content-Type": "application/json"},
            json=message_body
        )

        logger.log_info(f"Message response: {message_response.status_code}, {message_response.json()}")

        if message_response.status_code == 201:
            logger.log_info(f"Mensagem via teams enviada para {user_email}")
        else:
            raise Exception(f"Erro ao enviar mensagem para {user_email}: {message_response.json()}")

class MSGraphEmail(MSGraphBase):
    def __init__(self):
        super().__init__()

    @staticmethod
    def _prepara_anexos(file_paths):
        anexos = []
        for file in file_paths:
            if not file:
                continue
            with open(file, "rb") as f:
                file_data = f.read()
            file_name = os.path.basename(file)
            attachment = {
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": file_name,
                "contentType": "application/octet-stream",
                "contentBytes": base64.b64encode(file_data).decode("utf-8")
            }
            anexos.append(attachment)
        
        return anexos
        
    def _monta_body_email_anexos(self, email, anexos, msg):
        message = {
            "message": {
                    "subject": self.subject_metricas,
                    "body": {
                        "contentType": "HTML",
                        "content": msg
                    },
                    "toRecipients": [
                        {
                            "emailAddress": {
                                "address": email.strip()
                            }
                        }
                    ],
                    "attachments": anexos
                }
            }
        return message
    
    def _monta_body_email(self, email, msg, lider=False):
        logger.log_info("Montando body email")
        message = {
        "message": {
                "subject": self.subject_lider if lider else self.subject,
                "body": {
                    "contentType": "HTML",
                    "content": msg
                },
                "toRecipients": [
                    {
                        "emailAddress": {
                            "address": email.strip()
                        }
                    }
                ],
                "importance": "high" if not lider else "normal"
            }
        }
        return message

    # Função para enviar e-mail
    def _send_email(self, 
                    data, 
                    hora_pontomais,
                    hora_rabbiit, 
                    nome, 
                    user_email,
                    lider=False, 
                    nome_lider=None, 
                    coordenador=False, 
                    nome_coordenador=None):
        """
        Envia o email
        Args:
            data (str): Data do alerta,
            hora_pontomais (str): Hora do ponto mais,
            hora_rabbiit (str): Hora do rabbiit,
            nome (str): Nome do colaborador,
            user_email (str): E-mail do colaborador,
            lider (bool): Se é lider ou não,
            nome_lider (str): Nome do lider,
            coordenador (bool): Se é coordenador ou não,
            nome_coordenador (str): Nome do coordenador
        """
        user_email = user_email if env == "prd" else EMAIL_TESTE
        msg_email = self._monta_message(data, hora_pontomais, hora_rabbiit, nome, lider, nome_lider, coordenador, nome_coordenador)
        message = self._monta_body_email(user_email, msg_email)
        response = requests.post(
            f"{self.graph_base_url}/users/{self.my_id}/sendMail",
            headers={"Authorization": f"Bearer {self.token}", "Content-Type": "application/json"},
            json=message
        )
        if response.status_code == 202:
            logger.log_info(f"E-mail enviado com sucesso para {user_email}!")
        else:
            raise Exception(f"Erro ao enviar e-mail para {user_email}: {response.json()}")
        
    def _send_email_anexo(self, user_email, body_email):
        user_email = user_email if env == "prd" else EMAIL_TESTE
        response = requests.post(
            f"{self.graph_base_url}/users/{self.my_id}/sendMail",
            headers={"Authorization": f"Bearer {self.token}", "Content-Type": "application/json"},
            json=body_email
        )
        if response.status_code == 202:
            logger.log_info(f"E-mail enviado com sucesso para {user_email}!")
        else:
            raise Exception(f"Erro ao enviar e-mail para {user_email}: {response.json()}")
    
    @staticmethod
    def _monta_msg_anexo(metricas):
            return f"""
        <html>
        <body style="background-color: #ffffff; font-family: Arial, sans-serif; color: #333;">
            <div style="width: 90%; max-width: 600px; margin: auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px;">
                <h2 style="color: #1e90ff; text-align: center;">Relatório de Execuções Bot Confronto PontoMais x Rabbiit</h2>
                
                <p style="font-size: 14px; color: #333333; margin: 0; text-align: center; padding: 0;">
                    <i>E-mail de relatório automático. Favor não responder.</i>
                </p>

                <p style="color: #333; font-size: 14px;">
                    Prezados,
                </p>

                <p style="color: #333; font-size: 14px;">
                    O processo de verificação de divergências PontoMais x Rabbiit por colaborador foi finalizado. <br>
                    Colaboradores alertados por divergências:
                </p>
                <table style="width: 100%; border-collapse: collapse; margin-top: 10px;">
                    {len(metricas)}
                </table>
                <p style="color: #333; font-size: 14px; margin-top: 20px;">
                    Em anexo, a planilha de resultados contendo os descritivos de cada colaborador alertado e a planilha de exceções.
                </p>

                <p style="color: #333; font-size: 14px;">
                    Em caso de dúvidas ou para mais informações, por favor, entre em contato com nossa equipe de sustentação.
                </p>

                <p style="color: #333; font-size: 14px;">
                    <br><br>
                </p>
            </div>
        </body>
        </html>
        """
    
    def _monta_body_empty(self, email, msg):
        message = {
            "message": {
                    "subject": self.subject_empty,
                    "body": {
                        "contentType": "HTML",
                        "content": msg
                    },
                    "toRecipients": [
                        {
                            "emailAddress": {
                                "address": email.strip()
                            }
                        }
                    ]
                }
            }
        return message
        
    def _send_email_empty(self, user_email, body_email):
        user_email = user_email if env == "prd" else EMAIL_TESTE
        response = requests.post(
            f"{self.graph_base_url}/users/{self.my_id}/sendMail",
            headers={"Authorization": f"Bearer {self.token}", "Content-Type": "application/json"},
            json=body_email
        )
        if response.status_code == 202:
            logger.log_info(f"E-mail enviado com sucesso para {user_email}!")
        else:
            raise Exception(f"Erro ao enviar e-mail para {user_email}: {response.json()}")
        
    def _send_email_error(self, user_email, msg):
        message = {
            "message": {
                    "subject": "Erro Fatal Bot Ponto",
                    "body": {
                        "contentType": "Text",
                        "content": msg
                    },
                    "toRecipients": [
                        {
                            "emailAddress": {
                                "address": user_email.strip()
                            }
                        }
                    ]
                }
            }
        response = requests.post(
            f"{self.graph_base_url}/users/{self.my_id}/sendMail",
            headers={"Authorization": f"Bearer {self.token}", "Content-Type": "application/json"},
            json=message
        )
        if response.status_code == 202:
            logger.log_info(f"E-mail enviado com sucesso para {user_email}!")
        else:
            raise Exception(f"Erro ao enviar e-mail para {user_email}: {response.json()}")
            
class MSGraph(MSGraphEmail, MSGraphTeams):
    def __init__(self):
        logger.log_info("Iniciando Class MSGraph")
        super().__init__()

    def envia_email_erro(self, msg):
        emails = self.emails_sus.split(',')
        for email in emails:
            self._send_email_error(email, msg)
    
    def envia_metricas(self, registros=[], *anexos):
        if not registros:
            logger.log_info("Nao foram encontrados registros para essa execucao")
            return
        logger.log_info("Enviando emails com as metricas")
        emails = self.emails_metricas.split(',')
        for email in emails:
            lista_anexos = self._prepara_anexos(anexos)
            msg_email = self._monta_msg_anexo(registros)
            body_email = self._monta_body_email_anexos(email, lista_anexos, msg_email)
            self._send_email_anexo(email, body_email)

    def envia_alertas(self, registro: dict | list, lider: bool=False, coordenador: bool=False):
        """
        Envio dos alertas via email e teams
        
        Args:
            registro (dict | list): dict com o registro do colaborador ou lista caso seja envio para lider
            lider (bool): True para envio ao lider, default=False
            coordenador (bool): True para envio ao coordenador, default=False
        """ 
        logger.log_info("Enviando alertas")
        try:
            
            data = [[dado['date'] for dado in reg['dados']] for reg in registro] if lider else [reg['date'] for reg in registro['dados']]
            nome = [[reg['nome'] for reg in registro]] if lider else registro['nome']
            hora_pontomais = [[dado['time_clock'] for dado in reg['dados']] for reg in registro] if lider else [reg['time_clock'] for reg in registro['dados']]
            hora_rabbiit = [[dado['time_total'] for dado in reg['dados']] for reg in registro] if lider else [reg['time_total'] for reg in registro['dados']]
            email_lider = registro[0]['email_gestor'] if lider else None
            nome_lider = email_lider.split('.')[0].capitalize() if lider else None
            email_coordenador = registro[0]['email_coordenador'] if coordenador else None
            nome_coordenador = email_coordenador.split('.')[0].capitalize() if coordenador else None
            
            if env == "prd":
                ## SUSTENTACAO ###
                self._send_email(
                    data=data,
                    hora_pontomais=hora_pontomais,
                    hora_rabbiit=hora_rabbiit,
                    nome=nome,
                    user_email=self.emails_sus,
                    lider=lider,
                    nome_lider=nome_lider
                )
                self._send_message_to_teams(
                    data=data,
                    hora_pontomais=hora_pontomais,
                    hora_rabbiit=hora_rabbiit,
                    nome=nome,
                    user_email=self.emails_sus,
                    lider=lider,
                    nome_lider=nome_lider
                )
                ## SUSTENTACAO ###


            if lider:
                #### LIDER ####
                self._send_email(
                    data=data,
                    hora_pontomais=hora_pontomais,
                    hora_rabbiit=hora_rabbiit,
                    nome=nome,
                    user_email=email_lider,
                    lider=True,
                    nome_lider=nome_lider
                )
                self._send_message_to_teams(
                    data=data,
                    hora_pontomais=hora_pontomais,
                    hora_rabbiit=hora_rabbiit,
                    nome=nome,
                    user_email=email_lider,
                    lider=True,
                    nome_lider=nome_lider
                )

            elif coordenador:
                #### COORDENADOR ####
                self._send_email(
                    data=data,
                    hora_pontomais=hora_pontomais,
                    hora_rabbiit=hora_rabbiit,
                    nome=nome,
                    user_email=email_coordenador,
                    coordenador=coordenador,
                    nome_coordenador=nome_coordenador
                )
                self._send_message_to_teams(
                    data=data,
                    hora_pontomais=hora_pontomais,
                    hora_rabbiit=hora_rabbiit,
                    nome=nome,
                    user_email=email_coordenador,
                    coordenador=coordenador,
                    nome_coordenador=nome_coordenador
                )

            else:
                #### COLABORADOR ####
                self._send_email(
                    data=data,
                    hora_pontomais=hora_pontomais,
                    hora_rabbiit=hora_rabbiit,
                    nome=nome,
                    user_email=registro['dados'][0]['user_email']
                )
                self._send_message_to_teams(
                    data=data,
                    hora_pontomais=hora_pontomais,
                    hora_rabbiit=hora_rabbiit,
                    nome=nome,
                    user_email=registro['dados'][0]['user_email']
                )

            
        except Exception as error:
            logger.log_info(f"Erro ao enviar alerta do colaborador registro: {registro}")
            logger.log_info(f'Detalhes do erro: {error}')
            logger.log_info(f'Traceback: {traceback.format_exc()}')
            raise Exception(f"Erro ao enviar alerta do colaborador: {registro}")
