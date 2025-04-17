import pandas as pd
from datetime import datetime
from variables.config import RESULTADOS_DIRECTORY
import os
from libraries.PrimeLogger import logger
from variables.config import ARQ_COLAB

def escreve_planilha(registros):
    logger.log_info("Escrevendo planilha de resultados")
    if not registros:
        logger.log_info("Sem registros para escrever na planilha de resultados")
        return None
    sucessos = [d for d in registros if "descricao_erro" not in d]
    if not sucessos:
        logger.log_info("Nao foram encontrados casos de sucesso para essa execucao")
        return None         
    lista_formatada = []
    for registro in sucessos:
        dados = registro.get("dados", [])
        if not dados:
            continue

        for item in dados:

            hora_pm_str = dados[0].get("time_clock", "00:00:00")
            hora_rb_str = dados[0].get("time_total", "00:00:00")
            data_divergencia_str = item.get("date", registro.get("data_exec"))

            try:
                hora_pm = datetime.strptime(hora_pm_str, "%H:%M:%S")
                hora_rb = datetime.strptime(hora_rb_str, "%H:%M:%S")
                diferenca = abs(hora_pm - hora_rb)

            except Exception as e:
                logger.log_error(f"Erro ao calcular diferença de horas: {e}")
                diferenca = "Erro ao calcular"

            lista_formatada.append({
                "Data do Alerta": datetime.strptime(registro["data_exec"], "%Y-%m-%d").strftime("%d/%m/%Y"),
                "Data da Divergência": datetime.strptime(data_divergencia_str, "%Y-%m-%d").strftime("%d/%m/%Y"),
                "Nome do Colaborador": registro.get("nome", "N/A"),
                "Líder do Colaborador": registro.get("gestor", "N/A"),
                "Coordenador do Colaborador": registro.get("coordenador", "N/A"),
                "Marcação PontoMais": hora_pm_str,
                "Marcação Rabbiit": hora_rb_str,
                "Diferença": str(diferenca)
            })
    df = pd.DataFrame(lista_formatada)
    hoje = datetime.now().strftime("%d-%m-%Y")
    file_name = f"Relatorio Confronto Rabbiit {hoje}.xlsx"
    path = os.path.join(RESULTADOS_DIRECTORY, file_name)
    df.to_excel(path, index=False, engine="openpyxl")
    logger.log_info("Planilha de resultados criada com sucesso")
    return path

def escreve_planilha_excecoes(registros):
    logger.log_info("Escrevendo planilha de excecoes")
    if not registros:
        logger.log_info("Sem registros para escrever na planilha de exceções")
        return None
    erros = [d for d in registros if "descricao_erro" in d]
    if not erros:
        logger.log_info("Nao foram encontradas exececoes para esta execucao")
        return None
    lista_formatada = []
    for registro in erros:
        lista_formatada.append({
            "Data da execução": datetime.now().strftime("%d/%m/%Y"),
            "Nome do Colaborador": registro.get("nome", "N/A"),
            "E-mail Colaborador": registro.get("email", "N/A"),
            "Motivo da não execução": registro.get("descricao_erro", "Erro não informado")
        })

    df = pd.DataFrame(lista_formatada)
    hoje = datetime.now().strftime("%d-%m-%Y")
    file_name = f"Planilha de Exceções Confronto Rabbiit {hoje}.xlsx"
    path = os.path.join(RESULTADOS_DIRECTORY, file_name)
    df.to_excel(path, index=False, engine="openpyxl")
    logger.log_info("Planilha de execeções criada com sucesso")
    return path

def carrega_planilha_colab():
    try:
        logger.log_info("Carregando planilha colaboradores x gestores")
        # Carregando a planilha
        planilha = pd.read_excel(ARQ_COLAB, engine='openpyxl')
        
        # Cria o dicionário e define email como índice
        dict_infos = planilha.set_index("Email").to_dict(orient="index")
        return dict_infos
    except Exception as e:
        raise Exception(f"Falha ao carregar planilha {e}")
    
def verifica_gestores(alerta, gestores):
    try:
        alerta['gestor'] = gestores[alerta['dados'][0]['user_email']]['Gestor Direto']
        alerta['email_gestor'] = gestores[alerta['dados'][0]['user_email']]['Email Gestor Direto']
        alerta['coordenador'] = gestores[alerta['dados'][0]['user_email']]['Coordenador']
        alerta['email_coordenador'] = gestores[alerta['dados'][0]['user_email']]['Email Coordenador']
        alerta['diretor_bu'] = gestores[alerta['dados'][0]['user_email']]['Diretor']
        alerta['email_diretor_bu'] = gestores[alerta['dados'][0]['user_email']]['email Diretor']
        return alerta
    except:
        raise Exception("Colaborador não encontrado na planilha de Gestores")
