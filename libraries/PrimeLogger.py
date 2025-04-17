import json
import logging
from datetime import datetime, timezone
from variables import config
import os

LEVEL_MAP = {
    "Information": logging.INFO,
    "Warning": logging.WARNING,
    "Error": logging.ERROR
}

SourceContext = 'bot-confroto-rabbiit'

class JSONFormatter(logging.Formatter):
    def format(self, record):
        log_record = {
            "Timestamp": f"{datetime.now(timezone.utc).isoformat(timespec='microseconds')}",
            "Level": record.levelname,
            "MessageTemplate": record.getMessage(),
            "Properties": {
                "SourceContext": SourceContext,
                "ApplicationName": record.service_name
            }
        }
        return json.dumps(log_record)

class PrimeLogger:
    def __init__(self, service_name):
        self.service_name = service_name
        self.logger = logging.getLogger(service_name)
        handler = logging.StreamHandler()
        handler.setFormatter(JSONFormatter())
        self.logger.addHandler(handler)
        self.logger.setLevel(logging.INFO)

    def log_info(self, texto, level="Information", registro=None):
        """Responsável por logar no STDOUT e enviar para o OpenSearch.

        Args:
            texto (str): Texto do Log
            level (str, optional): Level com os valores "Information", "Error" e "Warning".
            registro (_dict_, optional): Dicionário com registros do processo. Defaults None.
        """
        try:
            log_level = LEVEL_MAP.get(level, logging.INFO)
            log_record = self._create_log_record(texto, level, registro)
            self.logger.log(log_level, log_record["Message"], extra={"service_name": self.service_name})
            self._log_to_console(log_record, registro)

        except Exception as e:
            self.log_exception(e)

    def log_exception(self, e):
        error_message = str(e)
        log_record = self._create_log_record(error_message, level="Error")
        self.logger.error(log_record["Message"], extra={"service_name": self.service_name})
        self._log_to_console(log_record)

    def _create_log_record(self, texto, level, registro=None):
        log_message = f"{texto}"
        log_level = LEVEL_MAP.get(level, logging.INFO)

        # log_record = {
        #     "Timestamp": f"{datetime.now(timezone.utc).isoformat(timespec='microseconds')}",
        #     "Level": logging.getLevelName(log_level),
        #     "MessageTemplate": log_message,
        #     "Properties": {
        #         "SourceContext":SourceContext,
        #         "ApplicationName": self.service_name
        #     }
        # }
        log_record = {"Timestamp": f"{datetime.now(timezone.utc).isoformat(timespec='microseconds')}",
                      "Message": log_message}
        self._save_log_to_json(log_record)
        return log_record

    def _log_to_console(self, log_record, registro=None):
        # if registro is not None:
        #     print(f"Level: {log_record['Level']} - Message: {log_record['MessageTemplate']} - Registro: {registro}")
        # else:
        #     print(f"Level: {log_record['Level']} - Message: {log_record['MessageTemplate']}")

        if registro is not None:
            print(f"Level: {log_record['Timestamp']} - Message: {log_record['Message']} - Registro: {registro}")
        else:
            print(f"Level: {log_record['Timestamp']} - Message: {log_record['Message']}")

    def _save_log_to_json(self, log_record):
        name_log = f'{config.ROBOT_NAME}_{datetime.now().strftime("%Y-%m-%d")}.json'
        filepath = os.path.join(config.OUTPUT_DIRECTORY, name_log)
        if os.path.exists(filepath):
            with open(filepath, 'a') as f:
                f.write(f'{json.dumps(log_record)}\n')
        else:
            with open(filepath, 'w') as f:
                json.dump(log_record, f)
                f.write('\n')

logger = PrimeLogger(
    service_name=config.ROBOT_NAME
)
