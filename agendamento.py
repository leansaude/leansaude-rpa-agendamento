#############################################################
# Identifica visitas médicas pendentes de agendamento e 
# realiza os agendamentos no Amplimed.
#############################################################

##################################
# BIBLIOTECAS
##################################
from seleniumwire import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from anticaptchaofficial.recaptchav2proxyless import*
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import pandas as pd
import os.path
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from urllib.parse import urlencode
from random import randint, choice
import sys
import re
from dotenv import load_dotenv
load_dotenv()

##################################
# CONSTANTES E VARIÁVEIS GLOBAIS
##################################
ALWAYS_CONFIRM_BEFORE_PROCEED = os.getenv('ALWAYS_CONFIRM_BEFORE_PROCEED')
ALWAYS_MANUALLY_SOLVE_CAPTCHA = os.getenv('ALWAYS_MANUALLY_SOLVE_CAPTCHA')
ENVIRONMENT = os.getenv('ENVIRONMENT')
SPREADSHEET_MANAGEMENT = {}
SPREADSHEET_MANAGEMENT['staging'] = os.getenv('SPREADSHEET_MANAGEMENT_STAGING')
SPREADSHEET_MANAGEMENT['production'] = os.getenv('SPREADSHEET_MANAGEMENT_PRODUCTION')
SPREADSHEET_HOSPITALS = os.getenv('SPREADSHEET_HOSPITALS')
RANGE_PATIENTS = os.getenv('RANGE_PATIENTS')
RANGE_VISITS = os.getenv('RANGE_VISITS')
RANGE_HOSPITALS = os.getenv('RANGE_HOSPITALS')
RANGE_PROFESSIONALS_HOSPITALS = os.getenv('RANGE_PROFESSIONALS_HOSPITALS')
RANGE_PROFESSIONALS = os.getenv('RANGE_PROFESSIONALS')
AMPLIMED_LOGIN_URL = os.getenv('AMPLIMED_LOGIN_URL')
AMPLIMED_LOGIN_EMAIL = os.getenv('AMPLIMED_LOGIN_EMAIL')
AMPLIMED_LOGIN_PASSWORD = os.getenv('AMPLIMED_LOGIN_PASSWORD')
AMPLIMED_PROCEDIMENTO_VISITA_ID = os.getenv('AMPLIMED_PROCEDIMENTO_VISITA_ID')
AMPLIMED_CONVENIO_ID = os.getenv('AMPLIMED_CONVENIO_ID')
ANTICAPTCHA_KEY = os.getenv('ANTICAPTCHA_KEY')
ANTICAPTCHA_WEBSITE_KEY = os.getenv('ANTICAPTCHA_WEBSITE_KEY')
AMPLIMED_AUTHORIZATION_KEY = None # @todo: persistir em os.environ['AMPLIMED_AUTHORIZATION_KEY'] para reuso em execução futura, caso ainda válida
STAGING_DOCTOR_CPF = os.getenv('STAGING_DOCTOR_CPF')
STAGING_AMPLIMED_DOCTOR_ID = os.getenv('STAGING_AMPLIMED_DOCTOR_ID')
STAGING_AMPLIMED_HOSPITAL_ID = os.getenv('STAGING_AMPLIMED_HOSPITAL_ID')
STAGING_AMPLIMED_PATIENT_ID = os.getenv('STAGING_AMPLIMED_PATIENT_ID')
WAIT_TIME_SECONDS = int(os.getenv('WAIT_TIME_SECONDS'))
MIN_SCHEDULE_HOUR = int(os.getenv('MIN_SCHEDULE_HOUR'))
MAX_SCHEDULE_HOUR = int(os.getenv('MAX_SCHEDULE_HOUR'))
MAX_GOOGLE_API_TRIES = int(os.getenv('MAX_GOOGLE_API_TRIES'))
chromeBrowser = None
nextVisitRowIndex = None

##################################
# FUNÇÕES AUXILIARES
##################################

# Realiza o processamento da visita (obtenção de dados, agendamento, inserção de linha de visita)
def processVisit(patientAmplimedId, carteirinha, inHospitalStayCode, hospitalId, deadline, firstVisit, currentDoctorName=None):
    # obtém código Amplimed do hospital de agendamento
    hospitalAmplimedId = getHospitalAmplimedId(hospitalId)
    if (hospitalAmplimedId == '' or not hospitalAmplimedId):
        print('-- Interrompendo processamento da visita por não ter sido localizado o ID Amplimed do hospital: ' + hospitalId + " --")
        return False
    
    # obtém o médico com quem agendar
    doctorCpf = getDoctor(hospitalId, firstVisit, currentDoctorName)
    if (doctorCpf == '' or not doctorCpf):
        print('-- Interrompendo processamento da visita pela ausência de médico atuando no hospital: ' + hospitalId + " --")
        return False

    doctorAmplimedId = getDoctorAmplimedId(doctorCpf)
    if (doctorAmplimedId == '' or not doctorAmplimedId):
        print('-- Interrompendo processamento da visita pois não foi localizado o ID Amplimed do médico com o CPF: ' + doctorCpf + " --")
        return False

    doctorName = getDoctorName(doctorCpf)
    if (doctorName == '' or not doctorName):
        print('-- Interrompendo processamento da visita pois não foi localizado o nome do médico com o CPF: ' + doctorCpf + " --")
        return False

    # checa data-limite da visita
    if not checkDeadline(deadline):
        print('-- Interrompendo processamento da visita pois data-limite está inconsistente: ' + deadline + " --")
        return False        
    
    # realiza o agendamento
    scheduleVisit(patientAmplimedId, doctorAmplimedId, deadline, hospitalAmplimedId)
    
    # insere a nova linha na planilha Visitas
    addVisitRow(carteirinha, inHospitalStayCode, doctorName, deadline)

    return True

##################################
# Checa se a data limite é consistente
def checkDeadline(deadline):
    # formato DD/MM/YYYY?
    if not re.match(r"^\d{2}\/\d{2}\/\d{4}$", deadline):
        return False

    # intervalos numéricos consistentes?
    parts = deadline.split('/')
    day = int(parts[0])
    month = int(parts[1])
    year = int(parts[2])
    if day < 1 or day > 31 :
        return False

    if month < 1 or month > 12 :
        return False

    if year < 2022 or year > 2050 :
        return False

    return True

##################################
# Obtém o médico com quem agendar a próxima visita do paciente, retornando o CPF do médico
def getDoctor(hospitalId, firstVisit, currentDoctorName):
    if firstVisit :
        # atribui a um médico aleatório que atenda no hospital
        # @todo: avaliar na prática impacto deste algoritmo
        candidates = getDoctorsForHospital(hospitalId)

        if len(candidates) == 0 :
            print('-- Nenhum médico encontrado para o hospital: ' + hospitalId + ' --')
            return ''

        return choice(candidates)
    else :
        # atribui ao último médico que visitou o paciente
        dfDoctor = dfProfessionals.loc[dfProfessionals['Nome do profissional']==currentDoctorName]
        return dfDoctor.iloc[0]['CPF']

##################################
# Obtém um array de CPFs de médicos ativos que atendem no hospital informado
def getDoctorsForHospital(hospitalId):
    hospitalId = str(hospitalId).zfill(10)
    dfSelectedRows = dfProfessionalsHospitals.loc[(dfProfessionalsHospitals['Código interno operadora']==hospitalId) &
                                                  (dfProfessionalsHospitals['Status Profissional']=='Ativo') &
                                                  (dfProfessionalsHospitals['Status Hospital atendimento']=='Sim')]
    return dfSelectedRows['CPF'].values

##################################
# Agenda a visita no Amplimed
def scheduleVisit(patientAmplimedId, doctorAmplimedId, deadline, hospitalAmplimedId):
    if (ENVIRONMENT == 'staging'):
        patientAmplimedId = STAGING_AMPLIMED_PATIENT_ID
        doctorAmplimedId = STAGING_AMPLIMED_DOCTOR_ID
        hospitalAmplimedId = STAGING_AMPLIMED_HOSPITAL_ID

    # 1ª chamada à API: cadastrar agendamento
    url = 'https://app.amplimed.com.br/pag/AGEnda_new/acoes/CRUDagendamento.php'
    startTime = getStartTime()
    endTime = getEndTime(startTime)
    dateForAmplimed = translateDate(deadline)
    
    params = {}
    params['action'] = 'ADD'
    params['dados[bloqueado]'] = 'false'
    params['dados[paciente]'] = patientAmplimedId
    params['dados[motivo]'] = ''
    params['dados[local]'] = hospitalAmplimedId
    params['dados[profissional]'] = doctorAmplimedId
    params['dados[h_inicio]'] = startTime
    params['dados[h_fim]'] = endTime
    params['dados[data]'] = dateForAmplimed
    params['dados[procedimento]'] = AMPLIMED_PROCEDIMENTO_VISITA_ID
    params['dados[status]'] = 'Agendado'
    params['dados[convenio]'] = AMPLIMED_CONVENIO_ID
    params['dados[valor]'] = '0'
    params['dados[plano]'] = '0'
    params['dados[desconto]'] = '0'
    params['dados[tipo_desconto]'] = '0'
    params['dados[valor final]'] = '0'
    params['dados[id_soli]'] = ''
    params['dados[obs_a]'] = ''
    params['dados[obs_p]'] = '<br>'
    params['dados[utiliza_integracao]'] = 'false'

    print("\nDados do agendamento:")
    print('Data: ' + dateForAmplimed)
    print('Hora inicial: ' + startTime)
    print('Hora final: ' + endTime)

    print("\nPreparando chamada à API: " + url)
    print('-- params: ' + urlencode(params) + ' --')

    response = callAmplimedApi(url, 'POST', urlencode(params))

    #print("\n-- response: " + response + " --")
    
    #decodedResponse = json.loads(response)
    #idEvento = decodedResponse['eventos'][0]

    #print('-- decodedResponse: ' + decodedResponse + ' --')
    #print('-- idEvento: ' + idEvento + ' --')
    
    #decodedResponse #remover
    #idEvento #remover
    
    # 2ª chamada à API: 'vincula-ag-app.php' (finalidade?) - SUPRIMIDO POR ORA
    #url = 'https://app.amplimed.com.br/pag/agenda/acoes/vincula-ag-app.php'
    
    #params = {}
    #params['usuclin'] = '14197'
    #params['id_evento_amplimed[]'] = idEvento
    #params['celular'] = '11999999999'
    #params['codu'] = doctorAmplimedId
    
    #response = callAmplimedApi(url, 'POST', urlencode(params))
    
    #response #remover
    
    # 3ª chamada à API: log do agendamento - SUPRIMIDO POR ORA
    #url = 'https://app.amplimed.com.br/pag/LOGs/acoes/CRUDlogs.php'
    
    #log = {}
    #translatedLog = urlencode(json.dumps(log))
    
    #example:
    ##translatedLog = '["{\"codu\":\"857316\",\"usuclin\":\"14197\",\"idEvent\":738,\"logTime\":\"2022-08-26 20:32:26\",\"action\":\"CEVENTO\",\"content\":\"{\\\"title\\\":\\\"TESTE Frederico da Silva Melo | Visita hospitalar | Bradesco Sa\\\\u00fade | 11999999999 |  R$ 0,00 | BENEF. PORTUGUESA SANTO ANDR\\\\u00c9 (SANTO ANDR\\\\u00c9-SP) | Francisco Costa|##|0\\\",\\\"start\\\":\\\"2022-08-30 16:00:00\\\",\\\"end\\\":\\\"2022-08-30 16:30:00\\\",\\\"color\\\":\\\"#ab00fa\\\",\\\"paciente\\\":null,\\\"profissional\\\":\\\"Francisco Costa\\\",\\\"codproc\\\":\\\"5\\\",\\\"status\\\":\\\"Agendado\\\",\\\"horatu\\\":\\\"\\\",\\\"codconvenio\\\":\\\"6\\\",\\\"codplano\\\":\\\"0\\\",\\\"obsagen\\\":\\\"\\\",\\\"codp\\\":\\\"52\\\",\\\"codu\\\":\\\"857316\\\",\\\"medicoSolicitante\\\":null,\\\"codVinculo\\\":0,\\\"id_solicitacao\\\":\\\"\\\",\\\"deuDesconto\\\":\\\"N\\\\u00e3o\\\",\\\"desconto\\\":\\\"\\\",\\\"tipoDesconto\\\":\\\"\\\",\\\"valorfinal\\\":0}\"}"]'
    
    #params = {}
    #params['logs'] = translatedLog
    #params['collection'] = 'eventoslogs'
    #params['action'] = 'SAVE_LOGS'
    
    #response = callAmplimedApi(url, 'POST', urlencode(params))   
    
    # 4ª chamada à API: 'movimentacaoHorarios.php' (finalidade?) - SUPRIMIDO POR ORA
    #url = 'https://app.amplimed.com.br/pag/ampliconsulta/acoes/movimentacaoHorarios.php'
    
    #dados = """[{"start":\"""" + dateForAmplimed + " " + startTime + """":00","end":\"""" + #dateForAmplimed + " " + endTime + """":00","codu":\"""" + doctorAmplimedId  + """","flag":"true"}]"""
    
    #params = {}
    #params['dados'] = dados
    #params['type'] = 'tabeladisp'
    #params['codu'] = doctorAmplimedId
    #params['fusoHorario'] = '-10800'

    #print('Preparando chamada à API: ' + url)
    #print('-- params: ' + urlencode(params) + ' --')
    
    #response = callAmplimedApi(url, 'POST', urlencode(params))

    #print('-- response: ' + response + ' --')
    
    #response #remover   

##################################
# Obtém o horário de início da visita (HH:MM), entre MIN_SCHEDULE_HOUR e MAX_SCHEDULE_HOUR
def getStartTime():
    hour = str(randint(MIN_SCHEDULE_HOUR, MAX_SCHEDULE_HOUR)).zfill(2)
    
    minute = randint(0, 1)
    if (minute == 0) :
        minute = '00'
    else :
        minute = '30'
    
    return hour + ':' + minute

##################################
# Adiciona 30 minutos ao horário de início fornecido e retorna no formato HH:MM
def getEndTime(startTime):
    parts = startTime.split(':')
    hour = int(parts[0])
    minute = int(parts[1])
    
    if (minute == 0) :
        minute = 30
    else :
        hour = hour+1
        minute = 0
        
    hour = str(hour).zfill(2)
    minute = str(minute).zfill(2)
    
    return hour + ':' + minute

##################################
# Adapta uma data no formato DD/MM/YYYY para YYYY-MM-DD
def translateDate(originalDate):
    parts = originalDate.split('/')
    return parts[2] + '-' + parts[1] + '-' + parts[0]

##################################
# Obtém o ID do hospital no Amplimed a partir do código do referenciado
def getHospitalAmplimedId(hospitalId):
    # completa com zeros à esquerda até 10 posições, para compatibilidade com a planilha "Rede credenciada"
    hospitalId = str(hospitalId).zfill(10)
    
    dfHospitalResult = dfHospitals.loc[dfHospitals['cod_referenciado']==hospitalId]

    if (len(dfHospitalResult.index) == 0):
        print('Hospital não localizado na lista de hospitais com atuação. Pesquisou pelo código: ' + hospitalId)
        return ''

    return dfHospitalResult.iloc[0]['cod_amplimed']

##################################
# Obtém o ID do médico no Amplimd a partir do CPF do médico
def getDoctorAmplimedId(doctorCpf):
    dfProfessionalResult = dfProfessionals.loc[dfProfessionals['CPF']==doctorCpf]

    if (len(dfProfessionalResult.index) == 0):
        print('Médico não localizado com o CPF: ' + doctorCpf)
        return ''

    return dfProfessionalResult.iloc[0]['profissional_cod_amplimed']

##################################
# Obtém o nome completo do médico a partir do CPF do médico
def getDoctorName(doctorCpf):
    dfProfessionalResult = dfProfessionals.loc[dfProfessionals['CPF']==doctorCpf]

    if (len(dfProfessionalResult.index) == 0):
        print('Médico não localizado com o CPF: ' + doctorCpf)
        return ''

    return dfProfessionalResult.iloc[0]['Nome do profissional']

##################################
# Adiciona nova linha de visita ao Google Sheet
def addVisitRow(carteirinha, inHospitalStayCode, doctorName, deadline):
    global nextVisitRowIndex

    updateVisitCell('B', carteirinha)
    updateVisitCell('C', inHospitalStayCode)
    updateVisitCell('I', deadline)
    updateVisitCell('J', doctorName)
    updateVisitCell('K', 'Agendada')
    
    nextVisitRowIndex = nextVisitRowIndex + 1

##################################
# Atualiza 1 célula na planilha Visitas
def updateVisitCell(column, value) :
    cellRangeToUpdate = "Visitas!" + column + str(nextVisitRowIndex)
    translatedValue = [[value]]
    result = sheet.values().update(spreadsheetId=SPREADSHEET_MANAGEMENT[ENVIRONMENT],
                                   range=cellRangeToUpdate, valueInputOption="USER_ENTERED",
                                   body={"values":translatedValue}).execute()
    print('Atualizou célula Visitas!' + column + str(nextVisitRowIndex) + ' com valor: ' + value)

##################################
# Obtém a chave de autorização das APIs Amplimed e salva em AMPLIMED_AUTHORIZATION_KEY
def getAmplimedAuthorizationKey():
    global AMPLIMED_AUTHORIZATION_KEY
    global chromeBrowser

    openAmplimed()

    if AMPLIMED_AUTHORIZATION_KEY :
        return
    
    if not chromeBrowser :
        print('Erro: chromeBrowser não definido.')
        return

    # extrai o AMPLIMED_AUTHORIZATION_KEY
    for request in chromeBrowser.requests :
        if request.headers['authorization'] :
            AMPLIMED_AUTHORIZATION_KEY = request.headers['authorization']
            print('Obtido token para chamadas à API Amplimed')
            break;

    print('AMPLIMED_AUTHORIZATION_KEY: ' + str(AMPLIMED_AUTHORIZATION_KEY))
    #AMPLIMED_AUTHORIZATION_KEY #remover

##################################
# Abre Amplimed no Chrome e efetua login (se necessário)
def openAmplimed():
    global chromeBrowser

    # stop if Amplimed already open
    if chromeBrowser:
        return
    
    options = Options()
    #options.add_argument('--headless')
    options.add_argument('window-size=2000,1000')
    chromeService = Service(ChromeDriverManager().install())
    chromeBrowser = webdriver.Chrome(options=options,service=chromeService)
    chromeBrowser.get(AMPLIMED_LOGIN_URL)
    time.sleep(10)

    if AMPLIMED_AUTHORIZATION_KEY:
        print('AMPLIMED_AUTHORIZATION_KEY já definida. Apenas abriu Chrome e navegou ao site do Amplimed, mas não irá efetuar login.')
        return
    
    print("Iniciando login no Amplimed")
    loginEmail = chromeBrowser.find_element(By.XPATH, '//*[@id="loginform"]/div[1]/div/div/input')
    loginEmail.send_keys(AMPLIMED_LOGIN_EMAIL)
    
    loginPassword = chromeBrowser.find_element(By.XPATH, '//*[@id="loginform"]/div[2]/div/div/input')
    loginPassword.send_keys(AMPLIMED_LOGIN_PASSWORD)

    # só executa anti-captcha se assim configurado
    if ALWAYS_MANUALLY_SOLVE_CAPTCHA != 'SIM' :
        print("Iniciando destravamento do Captcha")
        solver = recaptchaV2Proxyless()
        solver.set_verbose(1)
        solver.set_key(ANTICAPTCHA_KEY)
        solver.set_website_url(AMPLIMED_LOGIN_URL)
        solver.set_website_key(ANTICAPTCHA_WEBSITE_KEY)
        response = solver.solve_and_return_solution()

        if response != 0:
            print(response)
            chromeBrowser.execute_script(f"document.getElementById('g-recaptcha-response-100000').innerHTML = '{response}'")
            chromeBrowser.find_element(By.XPATH, '//*[@id="loginform"]/div[3]/div/button').click()
        else:
            print(solver.err_string)

        time.sleep(10)
    else : # ALWAYS_MANUALLY_SOLVE_CAPTCHA == 'SIM'
        print("--> AGUARDANDO 30 SEGUNDOS PARA EFETUAR LOGIN MANUAL NO AMPLIMED... <--")
        time.sleep(30)

    # navega para uma página que requeira alguma requisição POST contendo
    # o authorization header
    print('Navegando para agenda Amplimed para obter authorization header')
    #chromeBrowser.get("https://app.amplimed.com.br/agenda")
    wait = WebDriverWait(chromeBrowser, timeout=30)
    wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="navigation"]/ul/li[2]/a'))).click()
    time.sleep(10)

##################################
# Realiza uma chamada a um endpoint da API Amplimed
def callAmplimedApi(url, method, params):
    global AMPLIMED_AUTHORIZATION_KEY
    global chromeBrowser

    getAmplimedAuthorizationKey()
    
    if not AMPLIMED_AUTHORIZATION_KEY:
        sys.exit('Erro: AMPLIMED_AUTHORIZATION_KEY não definido.')

    if not chromeBrowser:
        sys.exit('Erro: chromeBrowser não definido.')
    
    request = '''var xhr = new XMLHttpRequest();
    xhr.open("''' + method + '''", "''' + url + '''", false);
    xhr.setRequestHeader('Content-type', 'application/x-www-form-urlencoded');
    xhr.setRequestHeader('authorization', "''' + AMPLIMED_AUTHORIZATION_KEY + '''");
    xhr.send("''' + params + '''");
    return xhr.response;'''

    return chromeBrowser.execute_script(request)


##################################
# OBTENÇÃO DE DADOS DA PLANILHA DE GERENCIAMENTO
##################################

# Inicia Google Spreadsheet Service
googleSpreadsheetService = build('sheets', 'v4')
sheet = googleSpreadsheetService.spreadsheets()

# Obtém Pacientes
for x in range(MAX_GOOGLE_API_TRIES):
    try:
        print('Tentativa ' + str(x+1) + ': obtenção de pacientes pela API Google Sheet')
        resultPatients = sheet.values().get(spreadsheetId = SPREADSHEET_MANAGEMENT[ENVIRONMENT],
                                            range = RANGE_PATIENTS).execute()
        break
    except Exception as e:
        print(e)
        continue

valuesPatients = resultPatients.get('values', [])
dfPatients = pd.DataFrame(valuesPatients[1:], columns=valuesPatients[0])
#dfPatients #remover
print('Lidos ' + str(len(dfPatients.index)) + ' registros de pacientes.')

# Obtém Visitas
for x in range(MAX_GOOGLE_API_TRIES):
    try:
        print('Tentativa ' + str(x+1) + ': obtenção de visitas pela API Google Sheet')
        resultVisits = sheet.values().get(spreadsheetId = SPREADSHEET_MANAGEMENT[ENVIRONMENT],
                                          range = RANGE_VISITS).execute()
        break
    except Exception as e:
        print(e)
        continue

valuesVisits = resultVisits.get('values', [])
dfVisits = pd.DataFrame(valuesVisits[1:], columns=valuesVisits[0])
#dfVisits #remover
print('Lidos ' + str(len(dfVisits.index)) + ' registros de visitas.')

# Calcula nextVisitRowIndex
dfVisitsColB = dfVisits[['Carteirinha']]
dfVisitsColB = dfVisitsColB.dropna()
nextVisitRowIndex = len(dfVisitsColB.index) + 2
#nextVisitRowIndex #remover
print('Posição da próxima visita a ser inserida:  ' + str(nextVisitRowIndex))

# Obtém Hospitais com atuação
for x in range(MAX_GOOGLE_API_TRIES):
    try:
        print('Tentativa ' + str(x+1) + ': obtenção de hospitais pela API Google Sheet')
        resultHospitals = sheet.values().get(spreadsheetId = SPREADSHEET_HOSPITALS,
                                             range = RANGE_HOSPITALS).execute()
        break
    except Exception as e:
        print(e)
        continue

valuesHospitals = resultHospitals.get('values', [])
dfHospitals = pd.DataFrame(valuesHospitals[1:], columns=valuesHospitals[0])
dfHospitals = dfHospitals.loc[dfHospitals['hospital_com_atuação']=='Sim']
#dfHospitals #remover
print('Lidos ' + str(len(dfHospitals.index)) + ' registros de hospitais com atuação.')

# Obtém cruzamento Profissionais x Hospitais
for x in range(MAX_GOOGLE_API_TRIES):
    try:
        print('Tentativa ' + str(x+1) + ': obtenção de cruzamento Profissionais x Hospitais pela API Google Sheet')
        resultProfessionalsHospitals = sheet.values().get(spreadsheetId = SPREADSHEET_HOSPITALS,
                                                        range = RANGE_PROFESSIONALS_HOSPITALS).execute()
        break
    except Exception as e:
        print(e)
        continue

valuesProfessionalsHospitals = resultProfessionalsHospitals.get('values', [])
dfProfessionalsHospitals = pd.DataFrame(valuesProfessionalsHospitals[1:], columns=valuesProfessionalsHospitals[0])
#dfProfessionalsHospitals #remover
print('Lidos ' + str(len(dfProfessionalsHospitals.index)) + ' registros de correlação profissionais x hospitais.')

# Obtém Profissionais (médicos) ativos
for x in range(MAX_GOOGLE_API_TRIES):
    try:
        print('Tentativa ' + str(x+1) + ': obtenção de profissionais pela API Google Sheet')
        resultProfessionals = sheet.values().get(spreadsheetId = SPREADSHEET_MANAGEMENT[ENVIRONMENT],
                                                range = RANGE_PROFESSIONALS).execute()
        break
    except Exception as e:
        print(e)
        continue

valuesProfessionals = resultProfessionals.get('values', [])
dfProfessionals = pd.DataFrame(valuesProfessionals[1:], columns=valuesProfessionals[0])
dfProfessionals = dfProfessionals.loc[dfProfessionals['Status']=='Ativo']
#dfProfessionals #remover
print('Lidos ' + str(len(dfProfessionals.index)) + ' registros de profissionais (médicos) ativos.')


##################################
# AGENDAMENTOS DE PRIMEIRA VISITA
##################################

print("\nAGENDAMENTOS DE PRIMEIRA VISITA")

# seleciona pacientes com status "Novo" (=sem visita "Realizada"), não possuam nenhuma visita "Agendada"
# e estejam cadastrados no Amplimed
dfPatientsWithoutFirstVisit = dfPatients.loc[(dfPatients['Status']=='Novo') & 
                                             (dfPatients['possui_alguma_visita_agendada']=='0') &
                                             (dfPatients['Status de cadastro na Amplimed']=='Cadastrado')]
#dfPatientsWithoutFirstVisit #remover
print('Localizados ' + str(len(dfPatientsWithoutFirstVisit.index)) + ' pacientes com 1ª visita pendente.')

# para cada paciente
for i in dfPatientsWithoutFirstVisit.index:
    # obtém variáveis das planilhas
    deadline = dfPatientsWithoutFirstVisit.loc[i,'data_limite_primeira_visita']
    hospitalId = dfPatientsWithoutFirstVisit.loc[i,'Código interno operadora']
    inHospitalStayCode = dfPatientsWithoutFirstVisit.loc[i,'Senha']
    patientAmplimedId = dfPatientsWithoutFirstVisit.loc[i,'ID Amplimed']
    carteirinha = dfPatientsWithoutFirstVisit.loc[i,'Carteirinha']

    print("\n[" + str(i+2) + "] Dados do paciente cuja 1ª visita será inserida: Carteirinha: " + carteirinha + ", Senha de internação: " + inHospitalStayCode + ", Código do hospital: " + hospitalId + ", Data-limite da visita: " + deadline)
    
    # processa 1ª visita
    if not processVisit(patientAmplimedId, carteirinha, inHospitalStayCode, hospitalId, deadline, firstVisit=True) :
        userInput = input('Prosseguir? (s/n)')
        if userInput == 'n' :
            sys.exit()
        continue
    
    # pausa alguns segundos para mimetizar interação humana
    time.sleep(WAIT_TIME_SECONDS)

    if ALWAYS_CONFIRM_BEFORE_PROCEED == 'SIM':
        userInput = input('Prosseguir? (s/n)')
        if userInput == 'n' :
            sys.exit()

##################################
# AGENDAMENTOS DE VISITAS DE SEGUIMENTO
##################################

print("\nAGENDAMENTOS DE VISITAS DE SEGUIMENTO")

# seleciona visitas com a indicação de agendamento da próxima visita
dfVisitsAwaitingNextVisit = dfVisits.loc[dfVisits['Data da proxima visita']=='Agendar próxima visita']

print('Localizados ' + str(len(dfVisitsAwaitingNextVisit.index)) + ' pacientes com visita de seguimento pendente.')

# para cada visita selecionada
for i in dfVisitsAwaitingNextVisit.index:
    # obtém variáveis das planilhas
    hospitalId = dfVisitsAwaitingNextVisit.loc[i,'cod_hospital_operadora']
    inHospitalStayCode = dfVisitsAwaitingNextVisit.loc[i,'Senha']
    patientAmplimedId = dfVisitsAwaitingNextVisit.loc[i,'ID Amplimed']
    carteirinha = dfVisitsAwaitingNextVisit.loc[i,'Carteirinha']
    currentDoctorName = dfVisitsAwaitingNextVisit.loc[i,'Profissional']
    deadline = dfVisitsAwaitingNextVisit.loc[i,'Data sugerida']

    print("\n[" + str(i+2) + "] Dados do paciente cuja visita de seguimento será inserida: Carteirinha: " + carteirinha + ", Senha de internação: " + inHospitalStayCode + ", Código do hospital: " + hospitalId + ", Data-limite da visita: " + deadline + ", Nome do médico: " + currentDoctorName)
    
    # processa nova visita
    if not processVisit(patientAmplimedId, carteirinha, inHospitalStayCode, hospitalId, deadline, False, currentDoctorName) :
        userInput = input('Prosseguir? (s/n)')
        if userInput == 'n' :
            sys.exit()        
        continue
    
    # pausa alguns segundos para mimetizar interação humana
    time.sleep(WAIT_TIME_SECONDS)

    if ALWAYS_CONFIRM_BEFORE_PROCEED == 'SIM':
        userInput = input('Prosseguir? (s/n)')
        if userInput == 'n' :
            sys.exit()

print("\nEXECUÇÃO ENCERRADA.")