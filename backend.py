import datetime

logging = open('log.txt', 'a')
message = " Programma wordt afgesloten."

def date_log():
    dt = str(datetime.datetime.today()) + ': '
    return dt

def loglist(name_check, log_list):
    logging.write(date_log() + f'Check {name_check}: ' + str(log_list) + '\n')

def log(message):
    logging.write(date_log() + message + '\n')