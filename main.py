from libraries.database import ColaboradoresRabbiit
from libraries.rabbiit import ApiRabbiit
from libraries.microsoftgraph import MSGraph
from libraries.PrimeLogger import logger
from libraries.excel import escreve_planilha, escreve_planilha_excecoes, carrega_planilha_colab, verifica_gestores
import traceback
import asyncio
from datetime import datetime

def main():
    db = ColaboradoresRabbiit()
    pontos_inconsistentes = ApiRabbiit().get_inconsistences_points()
    registros_colabs = carrega_planilha_colab()


    mg = MSGraph()
    lista_exceptions = []
    envio_lider = []
    for idx, registro in enumerate(pontos_inconsistentes):
        try:
            registro = verifica_gestores(registro, registros_colabs)
            mg.envia_alertas(registro)
            registro['data_exec'] = datetime.now().strftime("%Y-%m-%d")
            if not 'descricao_erro' in registro:
                envio_lider.append(registro)
        except Exception as error:
            logger.log_info(f"Erro ao processar registro: {registro}")
            logger.log_info(f"Detalhes do erro: {error}")
            logger.log_info(f"Traceback: {traceback.format_exc()}")
            registro['descricao_erro'] = str(error)
            lista_exceptions.append(registro)

    email_gestores = list({reg['email_gestor'] for reg in envio_lider})
    for email in email_gestores:
        try:
            alerta_lider = [registro for registro in envio_lider \
                                    if registro['email_gestor'] == email]
            mg.envia_alertas(alerta_lider, lider=True) if len(alerta_lider) > 0 else None
        except Exception as error:
            logger.log_info(f"Erro ao enviar alerta para o líder.")
            logger.log_info(f"Detalhes do erro: {error}")
            logger.log_info(f"Traceback: {traceback.format_exc()}")
    
    dia_atual = datetime.now().strftime("%d")
    primeiros_dias_uteis = ApiRabbiit()._get_util_days()
    if dia_atual == primeiros_dias_uteis[0] or dia_atual == primeiros_dias_uteis[1]:
        email_coordenadores = list({reg['email_coordenador'] for reg in envio_lider})
        for email in email_coordenadores:
            try:
                alerta_coordenador = [registro for registro in envio_lider \
                                        if registro['email_coordenador'] == email]
                mg.envia_alertas(alerta_coordenador, coordenador=True) if len(alerta_coordenador) > 0 else None
            except Exception as error:
                logger.log_info(f"Erro ao enviar alerta para o coordenador.")
                logger.log_info(f"Detalhes do erro: {error}")
                logger.log_info(f"Traceback: {traceback.format_exc()}")

    try:
        logger.log_info(f"Inserindo registros no banco de dados")
        db.insere_registros(pontos_inconsistentes)
        logger.log_info(f"Registros inseridos no banco de dados")
    except Exception as error:
        logger.log_info(f"Erro ao inserir registros no banco de dados")
        logger.log_info(f"Detalhes do erro: {error}")
        logger.log_info(f"Traceback: {traceback.format_exc()}")

    path_resultados = None
    path_excecoes = None
    try:
        path_resultados = escreve_planilha(pontos_inconsistentes)
    except Exception as error:
        logger.log_info(f"Erro ao escrever planilha de resultados")
        logger.log_info(f"Detalhes do erro: {error}")
        logger.log_info(f"Traceback: {traceback.format_exc()}")

    try:
        path_excecoes = escreve_planilha_excecoes(pontos_inconsistentes)
    except Exception as error:
        logger.log_info(f"Erro ao escrever planilha de exceções")
        logger.log_info(f"Detalhes do erro: {error}")
        logger.log_info(f"Traceback: {traceback.format_exc()}")

    if path_resultados or path_excecoes:
        try:
            mg.envia_metricas(
                pontos_inconsistentes,
                path_resultados,
                path_excecoes
            )
        except Exception as error:
            logger.log_info(f"Erro ao escrever planilha de resultados")
            logger.log_info(f"Detalhes do erro: {error}")
            logger.log_info(f"Traceback: {traceback.format_exc()}")
    else:
        logger.log_info(f"Nao foram encontrados registros para essa execucao")

if __name__ == "__main__":
    main()
