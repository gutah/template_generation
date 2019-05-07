import codecs
import csv

def log_config():
    import logging 
    logger = logging.getLogger('CCEP')
    logger.setLevel(logging.DEBUG)
    fh = logging.FileHandler('template_generation.log')
    fh.setLevel(logging.DEBUG)
    ch = logging.StreamHandler()
    ch.setLevel(logging.ERROR)
    # create formatter and add it to the handlers
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    ch.setFormatter(formatter)
    fh.setFormatter(formatter)
    # add the handlers to logger
    logger.addHandler(ch)
    logger.addHandler(fh)
    return logger
log = log_config()
log.info("Log Configured")


base_path = r'C:\projetos\generate_templates'
template_path_masculine = base_path + r'\main_template\Template_Masculino.html'
template_path_feminine = base_path + r'\main_template\\Template_Feminino.html'
csv_base = base_path + r'base.csv'
path = base_path + r'\templates_gerados'

log.info("Base Path: " + base_path)
log.info("Tempalte Path Masculine: " + template_path_masculine)
log.info("Tempalte Path Feminine: " + template_path_feminine)
log.info("CSV Path: " + csv_base)
log.info("Path: " + path)

# Read Mail Template
def read_template():
    file_masculino = codecs.open(template_path_masculine,'r')
    file_feminino = codecs.open(template_path_feminine,'r')
    global template_masculine = file_masculino.read()
    global template_feminine = file_feminino.read()
    file_masculino.close()
    file_feminino.close()

def generate_mail(text, recipient, path, auto=True):
    import win32com.client as win32   

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.Bcc = "gustavo.maciel@cognizant.com;pedro.marin@cognizant.com"
    mail.Subject = 'CCEP 2019 - Cognizant Customer Excellence Program'
    mail.HtmlBody = text
    
    if auto:
        #mail.Display(True)
        mail.SaveAs(Path=path)
        mail.close()
    else:
        print("Erro")


# Generate Mail Template
def template_generation(row):
    file_name = ""
    contact_name = row[0]
    gender = row[1]
    survey_link = row[2]
    limit_date = row[3]
    contact_type = row[4]
    customer = row[5]
    vertical = row[6]
    contact_mail = row[7]
    file_name = "\CCEP Survey - " + vertical +" - "+ customer + " - " + contact_type + " - " + contact_name + ".msg"
    f_path = path+file_name
        
    if gender == 'M':
        file = template_masculine
        file = file.replace('contact_name', contact_name)
        file = file.replace('limit_date', limit_date)
        file = file.replace('survey_link', survey_link)
        generate_mail(file,contact_mail,f_path)
    else:
        file = template_feminine
        file = file.replace('contact_name', contact_name)
        file = file.replace('limit_date', limit_date)
        file = file.replace('survey_link', survey_link)
        generate_mail(file,contact_mail, f_path)

read_template()

# Read CSV
with open(base_path) as csv_file:
    csv_reader = csv.reader(csv_file, delimiter=';')
    line_count = 0
    for row in csv_reader: 
        template_generation(row)
        line_count +=1
csv_file.close()        