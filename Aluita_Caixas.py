import configparser
from datetime import datetime
import json
import sys
import os
import pandas as pd
from tabulate import tabulate
import builtins
from robotpy.Robot import Robot

config = configparser.ConfigParser()
log = """
    <style> 
    table, th, td { border: 1px solid black; border-collapse: collapse; }
    th, td { padding: 5px; }
    </style>
"""

'''
    # find on all cells first date of the sheet
'''

def print(message):
    global log
    log += str(message) + "<br>"
    builtins.print(message)
    
def stringIsInDateFormat(string, dateFormat):
    try:
        date = pd.to_datetime(string, format=dateFormat)
        return date is not None
    except:
        return None

def getDateFirstDateOfSheet(sheet):
    # For each row and column, verify if the cell is a date, if is a date, return the date
    for row in sheet._values:
        for column in row:
            # if column is not None and is not nan
            if pd.isnull(column) == False:
                try:
                    #If type of column is datetime or string
                    if type(column) is str or type(column) is datetime:
                        #If column is string, consider BR date format on convert to date dd/mm/yyyy or dd-mm-yyyy or dd.mm.yyyy
                        if isinstance(column, str):
                            if stringIsInDateFormat(column, '%d/%m/%Y'):
                                return pd.to_datetime(column, format='%d/%m/%Y').strftime('%Y-%m-%d')
                            elif stringIsInDateFormat(column, '%d-%m-%Y'):
                                return pd.to_datetime(column, format='%d-%m-%Y').strftime('%Y-%m-%d')
                            elif stringIsInDateFormat(column, '%d.%m.%Y'):
                                return pd.to_datetime(column, format='%d.%m.%Y').strftime('%Y-%m-%d')
                        # if column is a date
                        else:
                            #convert date to sql date
                            return pd.to_datetime(column).strftime('%Y-%m-%d')
                except:
                    pass

    return None

def textHasAllWords(text, words):
    for word in words:
        if word.lower() not in text.lower():
            return False
    return True

def getParamInSectionWithFilters(section, text):
    try:
        for param in config[section].keys():
            #Split param by '#', the first part is the words must have in the text, the second part is the words must not have in the text. The words is separated by space
            words = config[section][param].split('#')
            if len(words) == 2:
                if textHasAllWords(text, words[0].split(' ')) and not textHasAllWords(text, words[1].split(' ')):
                    return str(param)
            elif len(words) == 1:
                if textHasAllWords(text, words[0].split(' ')):
                    return str(param)
    except:
        pass
    return None


def Aluita_Caixas(robot):
    # read config file aluita_caixas.ini using utf-8 encoding
    config.read('aluita_caixas.ini', encoding='utf-8')

    #Set inicial return log 'Aluita Caixas month/year'
    print('Aluita Caixas ' + str(robot['month']) + '/' + str(robot['year']))

    # set on section 'config', month(fill to 2) and year on config file
    config['paths']['month'] = str(robot['month']).zfill(2)
    config['paths']['year'] = str(robot['year'])

    #Converte o mes e ano para int para poder comparar nas datas
    robot['month'] = int(robot['month'])
    robot['year'] = int(robot['year'])

    # cashiers folder in config paths.boletinscaixa
    mainPath = config['paths']['boletinscaixa']

    #verify if folder exists
    if os.path.exists(mainPath):

        # Get cashiers in cashier section
        cashiers = config['cashiers'].keys()

        cashiersWithDifferences = 0

        # For each cashier
        for cashier in cashiers:
            # Get cashier filter
            cashier_filter = config['cashiers'][cashier]
            # Put 'CAIXA' before filter
            cashier_filter = 'caixa ' + cashier_filter
            # create filters with cashier split by ' '
            cashier_filter = cashier_filter.split(' ')

            # Find on folder the first file with all strings in cashier_filter and has not config config.edited_name in name
            for file in os.listdir(mainPath):
                if file.lower().endswith('.xlsx') and all(word.lower() in file.lower() for word in cashier_filter) and config['config']['edited_name'].lower() not in file.lower():                    
                    try:
                        df = None
                        
                        #if cashier is 'porto alegre'
                        if cashier == 'porto alegre':
                            # Open file and read all sheets without use headers, from column A to column F
                            df = pd.read_excel(
                                mainPath + file, sheet_name=None, header=None,
                                usecols=[0, 1, 2, 3, 4, 5],
                            )
                        else:
                            # Open file and read all sheets without use headers, from column A to column F
                            df = pd.read_excel(
                                mainPath + file, sheet_name=None, header=None, names=['A', 'B', 'C'])

                        #Dados para o sistema contabil
                        compact = []

                        #Totals differences
                        totalsDiffs = []

                        # For each sheet
                        for sheet in df:

                            # If sheet is not empty
                            if df[sheet].empty == False:
                                # Get the first date of the sheet
                                date = getDateFirstDateOfSheet(df[sheet])
                                # If date is not None and is in the same month and year
                                if date is not None and pd.to_datetime(date).month == robot['month'] and pd.to_datetime(date).year == robot['year']:
                                    totals = {
                                        'receipts': 0,
                                        'payments': 0,
                                        'receiptsInserted': 0,
                                        'paymentsInserted': 0,
                                    }

                                    valueTitle = ""
                                    valueType = ""

                                    # For each row
                                    for row in df[sheet]._values:
                                        # Remove all nans
                                        row = [x for x in row if pd.isnull(x) == False]

                                        if len(row) > 0:
                                            firstColumn = str(row[0])

                                            paymentTitle = None
                                            receiptTitle = None
                                            totalTitle = None
                                            title = None

                                            # Se o caixa for de poa
                                            if cashier == 'porto alegre':
                                                poaType = getParamInSectionWithFilters(
                                                    'POA', firstColumn)
                                                if poaType == 'payments':
                                                    paymentTitle = "Pagamentos"
                                                elif poaType == 'receipts' or poaType == 'antecipado':
                                                    receiptTitle = "Receitas"
                                            else:
                                                # Se nao for caixa poa verifica se é um titulo
                                                paymentTitle = getParamInSectionWithFilters(
                                                    'payments', firstColumn)
                                                receiptTitle = getParamInSectionWithFilters(
                                                    'receipts', firstColumn)
                                                totalTitle = getParamInSectionWithFilters(
                                                    'totals', firstColumn)
                                                title = getParamInSectionWithFilters(
                                                    'titles', firstColumn)

                                            # If column is a payment
                                            if paymentTitle is not None:
                                                valueType = "payment"
                                                valueTitle = paymentTitle
                                            # If column is a receipt
                                            elif receiptTitle is not None:
                                                valueType = "receipt"
                                                valueTitle = receiptTitle
                                            # If column is a total
                                            elif totalTitle is not None:
                                                valueType = "total"
                                                valueTitle = totalTitle

                                                totals[valueTitle] += float(row[-1])
                                            # If column is a title
                                            elif title is not None:
                                                #Reset valueType and valueTitle
                                                valueType = ""
                                                valueTitle = ""
                                            # If column is not a title and valueType is not empty and valueTitle is not empty
                                            elif title is None and valueType != "" and valueTitle != "":
                                                # description = valueTitle + first column + second column if has 3 columns)
                                                description = " ".join(
                                                    ['Receita' if valueType == 'receipt' else 'Pagamento'
                                                        , valueTitle
                                                        , firstColumn
                                                        , (str(row[1]) if len(row) == 3 else '')
                                                    ]
                                                )
                                                # value = last column
                                                value = row[-1]
                                                #if value is int, convert to float
                                                if isinstance(value, int):
                                                    value = float(value)

                                                # If description is not empty and value is not empty and value is float and first column is not equal last column
                                                if description != "" and value != "" and type(value) is float and str(firstColumn) != str(value):
                                                    #If valueType is receipt or payment
                                                    if valueType == "receipt" or valueType == "payment":
                                                        cashier_account = config['cashiers_accounts'][cashier]
                                                        config_section = valueType + 's_accounts'
                                                        
                                                        totals[valueType + 'sInserted'] += value
                                                        
                                                        #adiciona para a lista que será impressa em csv
                                                        compact.append({
                                                            '#cod_empresa': config['config']['enterprise_code'],
                                                            'data': date,
                                                            'debito': cashier_account if not config.has_option(config_section, 'debit') else config[config_section]['debit'],
                                                            'credito': cashier_account if not config.has_option(config_section, 'credit') else config[config_section]['credit'],
                                                            'historico padrao': config[valueType + 's_accounts']['history'],
                                                            'complemento historico': description,
                                                            'valor': value,
                                                        })
                                    #round totals
                                    totals['receiptsInserted'] = round(totals['receiptsInserted'], 2)
                                    totals['paymentsInserted'] = round(totals['paymentsInserted'], 2)
                                    totals['receipts'] = round(totals['receipts'], 2)
                                    totals['payments'] = round(totals['payments'], 2)

                                    #Verifica se nos totais o total inserido é igual ao total e se o totals não é 0, se for diferente, adiciona nas diferenças
                                    if str(totals['receiptsInserted']) != str(totals['receipts']) and totals['receipts'] != 0:
                                        totalsDiffs.append({
                                            'Nome da Aba': sheet,
                                            'Dia Considerado': date,
                                            'Tipo': 'Receitas',
                                            'Total inserido': totals['receiptsInserted'],
                                            'Total': totals['receipts'],
                                            'Diferença': totals['receiptsInserted'] - totals['receipts'],
                                        })
                                    if totals['paymentsInserted'] != totals['payments'] and totals['payments'] != 0:
                                        totalsDiffs.append({
                                            'Nome da Aba': sheet,
                                            'Dia Considerado': date,
                                            'Tipo': 'Pagamentos',
                                            'Total inserido': totals['paymentsInserted'],
                                            'Total': totals['payments'],
                                            'Diferença': totals['paymentsInserted'] - totals['payments'],
                                        })
                        #Se existem diferenças, converte as differenças para tabela html com o tabulate e printa na tela
                        if len(totalsDiffs) > 0:
                            cashiersWithDifferences += 1
                            print("Diferenças encontradas no caixa <strong>" + cashier.title() + ":</strong>")
                            print(tabulate(totalsDiffs, headers='keys', tablefmt='html'))
                        # Convert to dataframe
                        compact = pd.DataFrame(compact)

                        # Save to csv splited by ;
                        compact.to_csv(
                            mainPath +
                            file.replace('.xlsx', ' ' + config['config']['edited_name'] + '.csv')
                            , sep=';'
                            , index=False
                            , encoding='utf-8'
                        )

                    except Exception as e:
                        print("Erro ao processar arquivo " + mainPath + file + ": " + str(e))
                        print("")
                        continue

                    # break for file in folder
                    break
            else:
                print('Nenhum arquivo encontrado para: ' + cashier)

        #Se nao existir nenhuma diferença em nenhum caixa, printa na tela
        if cashiersWithDifferences == 0:
            print("Nenhuma diferença encontrada")
    else:
        print('Pasta de arquivos ' + mainPath + ' não encontrada')

    #Salva o log em um arquivo html na area de trabalho
    with open(mainPath + 'retorno robo caixas.html', 'w') as f:
        f.write(log)

#Protection against running the script twice
try:
    mes_teste = 5
    ano_teste = 2022

    #Se existir argumentos, define o call_id como o primeiro parametro, se não, define como None
    call_id = sys.argv[1] if len(sys.argv) > 1 else None

    #start robot with first argument
    robot = Robot(call_id)

    try:
        #from robot.parameters get 'mes' and 'ano' as int
        mes = int(robot.parameters['mes']) if call_id is not None else mes_teste
        ano = int(robot.parameters['ano']) if call_id is not None else ano_teste

        try:
            # call the main function passing the argument list
            Aluita_Caixas({'month': mes, 'year': ano})
            robot.setReturn(log)
        except Exception as e:
            robot.setReturn("Erro desconhecido: " + str(e))
    except Exception as e:
        robot.setReturn("Parametros passados invalidos: " + json.dumps(robot.parameters))
except Exception as e:
    print(e)
    pass

sys.exit(0)