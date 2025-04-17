import requests
import json
import datetime
from workalendar.america import Brazil

import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from variables.config import EMAIL_RABBIIT, PASSWORD_RABBIIT, ACCOUNT_ID_RABBIIT
from dateutil.relativedelta import relativedelta

class ApiRabbiit:
    def __init__ (self):
        self.email = EMAIL_RABBIIT
        self.password = PASSWORD_RABBIIT
        self.account_id = ACCOUNT_ID_RABBIIT
        self.api_url = "https://app.rabbiit.com/api/v1"

    
    def _authenticate(self):
        request=requests.post(f"{self.api_url}/auth",headers = {
        "Content-Type": "application/json"
        },
        params={
        "email": self.email,
        "password": self.password,
        "account_id": self.account_id
        })

        request_text=json.loads(request.text)
        api_token=request_text["token"]

        return api_token
    

    @staticmethod
    def _get_util_days():
        dias_uteis = []
        ano = datetime.datetime.now().year
        mes = datetime.datetime.now().month 
        for dia in range(1, 7):
            try:
                data = datetime.date(ano, mes, dia)
                if Brazil().is_working_day(data):
                    dias_uteis.append(data.strftime("%d"))
            except ValueError:
                break
        return dias_uteis
    

    def _get_time(self, date_start, date_end):
        time=requests.get(f"{self.api_url}/reports/daily?date_execution_end={date_end}&date_execution_start={date_start}&only_with_divergence=true",headers={
        "Content-Type": "application/json",
        "x-api-token": self._authenticate()
        })

        return json.loads(time.text)
    
    
    def get_inconsistences_points(self):
        dia_atual = datetime.datetime.now().strftime("%d")
        primeiros_dias_uteis = self._get_util_days()
        if dia_atual == primeiros_dias_uteis[0] or dia_atual == primeiros_dias_uteis[1]:
            month_year = (datetime.datetime.now() - relativedelta(months=1)).strftime("%Y-%m")
        else:
            month_year = datetime.datetime.now().strftime("%Y-%m")
        date_atual = datetime.datetime.now().strftime("%Y-%m-%d")
        start_date = month_year + "-01"

        inconsistences_points = self._get_time(start_date, date_atual) 
        inconsistences_points = inconsistences_points["data"]

        for point in inconsistences_points:
            time_rabbiit = point["time_total"]
            time_pontomais = point["time_clock"]
            horas_rabbiit, minutos_rabbiit, segundos_rabbiit = map(int, time_rabbiit.split(":"))
            horas_pontomais, minutos_pontomais, segundos_pontomais = map(int, time_pontomais.split(":"))
            
            if horas_rabbiit == horas_pontomais and abs(minutos_rabbiit - minutos_pontomais) == 0:
                inconsistences_points.remove(point)
            
        from collections import defaultdict
        grouped_points = defaultdict(list)
        for point in inconsistences_points:
            nome = point["user_name"]
            user_email = point["user_email"]
            grouped_points[nome].append(point)
        
        inconsistences_points = [{'nome': nome, 'dados': itens} for nome, itens in grouped_points.items()]   
        return inconsistences_points  
        


# teste = ApiRabbiit("api.rabbiit@primecontrol.com.br", "Rbt157@@", "562d3917307e").get_inconsistences_points()
