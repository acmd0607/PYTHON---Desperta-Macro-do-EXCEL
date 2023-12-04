import win32com.client as win32
import sys
import os
from datetime import datetime, timedelta
from time import sleep

def trata_erro(texto, arquivo_de_log):
    escreve_no_log(texto, arquivo_de_log)
    sys.exit(texto)

def escreve_no_log(texto, arquivo_de_log):
    arq_log = open(arquivo_de_log, "a", encoding='utf-8-sig')
    arq_log.write(texto)
    arq_log.close()
    return    

def processa_tempo(hora, minuto, dias_da_semana_int, excel_filein_path, macro_name, arquivo_de_log):   
    
    key = True
    while key:
        agora = datetime.now()

        if esta_na_hora(hora, minuto, agora) and esta_no_dia_da_semana(dias_da_semana_int, agora):
            key = False
            executa_macro(excel_filein_path, macro_name, arquivo_de_log)
        else:
            sleep(30)
    return

def esta_na_hora(hora, minuto, data_atual):
    if data_atual.hour == hora and data_atual.minute == minuto:
        return True
    return False

def esta_no_dia_da_semana(dias_da_semana, data_atual):
    if data_atual.weekday() == dias_da_semana:
        return True
    return False

def executa_macro(excel_filein_path, macro_name, arquivo_de_log):

    # Criar uma instância do EXCEL.
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False

    flag = True

    # Abrindo o arquivo Excel
    workbook = excel.Workbooks.Open(excel_filein_path, ReadOnly=False)
    
    texto = '   O arquivo EXCEL '
    texto = texto + excel_filein_path
    texto = texto + ' foi aberto com sucesso.\n'
    escreve_no_log(texto, arquivo_de_log)

    try:
        # Executar a macro -----

        texto = '   A macro com o nome: '
        texto = texto + macro_name
        texto = texto + ' está iniciando.\n'
        escreve_no_log(texto, arquivo_de_log)

        excel.Application.Run(macro_name)
        
        while flag:
            sleep(180)                 # tempo em segundos para executar a macro. Coloquei 3 minutos
            excel.Application.Wait()
            if excel.Application.Ready:
                texto = '   A macro com o nome: '
                texto = texto + macro_name
                texto = texto + ' finalizou a execução.\n'
                escreve_no_log(texto, arquivo_de_log)
                flag = False
            else:
                texto = '   A macro com o nome: '
                texto = texto + macro_name
                texto = texto + ' ainda está em execução.\n'
                escreve_no_log(texto, arquivo_de_log)
                flag = True

    except Exception as e:
        texto = '\n\n     ****** ERRO - Ocorreu um erro ao executar a MACRO: ' + str(e) + '\n\n'
        trata_erro(texto, arquivo_de_log)

    # Salvar e fechar o arquivo "excel_filein_path"
    workbook.Save()
    workbook.Close()

    # Fechar o EXCEL
    excel.Application.Quit()
    return

if __name__ == '__main__':          #Main Program

    # Coleto as informações enviadas pela linha de comando
    n = len(sys.argv)
    if n == 4:
        excel_filein_path = sys.argv[1]
        macro_name = sys.argv[2]
        arquivo_de_log = sys.argv[3]
    else:
        arquivo_de_log = "teste.txt"     # Não é necessário esse comando, pode deletar essa linha
        texto = '\n\n     ****** ERRO - FALTOU INFORMAÇÂO NA LINHA DE COMANDO\n\n'
        trata_erro(texto, arquivo_de_log)
    if not os.path.isfile(excel_filein_path):
        texto = '\n\n     ****** ERRO - NÃO ENCONTREI O ARQUIVO EXCEL ONDE TEM A MACRO\n\n'
        trata_erro(texto, arquivo_de_log)

    texto1 = '\n      ---------------------------------------------------------------'
    texto2 = '\n            EXECUTA UMA TAREFA PROGRAMADA EM UMA MACRO NO EXCEL \n'
    texto = texto1 + texto2 + '      ---------------------------------------------------------------\n\n'
    print(texto)
    escreve_no_log(texto2, arquivo_de_log)

    numero_de_dias = int(input('  Quantos dias seguidos você gostaria de executar a tarefa: '))
    texto = '\n   Quantos dias seguidos você gostaria de executar a tarefa: ' + str(numero_de_dias) + '\n'
    escreve_no_log(texto, arquivo_de_log)
    hora_string = input("  Qual o horario para executar a tarefa? (hh:mm): ")
    texto = '   Qual o horario para executar a tarefa? (hh:mm): ' + hora_string + '\n'
    escreve_no_log(texto, arquivo_de_log)

    hora = int(hora_string.split(':')[0])
    minuto = int(hora_string.split(':')[1])
    hora_minuto_tarefa = hora*60 + minuto

    agora = datetime.now() # guarda o tempo inicial
    hora_minuto_agora = agora.hour*60 + agora.minute
    
    if hora_minuto_agora < hora_minuto_tarefa:

        hoje = agora.weekday()
        dias_da_semana_int = agora.weekday()
        data_tarefa = agora.date()

        for i in range(0, numero_de_dias, 1):

            data_formatada = data_tarefa.strftime('%d/%m/%y')

            texto = '   Vou executar a tarefa por ' + str(numero_de_dias) + ' dia(s). Dia ' + str(i+1) + ': ' + str(data_formatada) + ' às: '
            texto = texto + str(hora) + ' horas e ' + str(minuto) + ' minutos\n'
            escreve_no_log(texto, arquivo_de_log)
            
            processa_tempo(hora, minuto, dias_da_semana_int, excel_filein_path, macro_name, arquivo_de_log)

            agora = datetime.now()

            texto = '   Tarefa finalizada na data: ' + str(agora.day) + '/' + str(agora.month) + '/' + str(agora.year)
            texto = texto + ' às ' + str(agora.hour) + ' horas ' + str(agora.minute) + ' minutos ' + str(agora.second) + ' segundos\n'
            escreve_no_log(texto, arquivo_de_log)

            dias_da_semana_int = +1
            if dias_da_semana_int > 6:
                dias_da_semana_int = 0
            data_tarefa = data_tarefa + timedelta(1)

    else:
        texto = '\n\n      ****** ERRO - Hora informada para executar a tarefa de hoje é menor que a hora atual\n\n'
        trata_erro(texto, arquivo_de_log)
