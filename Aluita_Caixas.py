import configparser
import sys
import os
import pandas as pd

config = configparser.ConfigParser()

'''
    # find on all cells first date of the sheet
'''


def getDateFirstDateOfSheet(sheet):
    # For each row and column, verify if the cell is a date, if is a date, return the date
    for row in sheet._values:
        for column in row:
            # if column is not None and is not nan
            if pd.isnull(column) == False:
                try:
                    # if column is a date
                    if pd.to_datetime(column) is not None:
                        #convert date to sql date
                        return pd.to_datetime(column).strftime('%Y-%m-%d')
                except:
                    pass

    return None


def getParamInSectionWithFilters(section, text):
    try:
        for param in config[section].keys():
            if all(word.lower() in text.lower() for word in config[section][param].split(' ')):
                return str(param)
    except:
        pass
    return None


def Aluita_Caixas(args):
    # read config file aluita_caixas.ini
    config.read('aluita_caixas.ini')

    robot = {
        'month': 'Desktop',
        'year': 'Aluita Caixas',
    }

    # put 0 before the number
    month = str(robot['month']).zfill(2)

    # set on section 'config', month and year on config file
    config['paths']['month'] = str(month)
    config['paths']['year'] = str(robot['year'])

    # cashiers folder in config paths.boletinscaixa
    mainPath = config['paths']['boletinscaixa']

    # Get cashiers in cashier section
    cashiers = config['cashiers'].keys()

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
                # Open file and read all sheets without use headers, from column A to column F
                df = pd.read_excel(
                    mainPath + file, sheet_name=None, header=None, usecols='A:F')

                compact = []

                # For each sheet
                for sheet in df:
                    # If sheet is not empty
                    if df[sheet].empty == False:
                        # Get the first date of the sheet
                        date = getDateFirstDateOfSheet(df[sheet])
                        # If date is not None
                        if date is not None:
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
                                    if cashier == 'POA':
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

                                        totals[valueTitle] += f
                                        loat(row[-1])

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
                            print(totals)

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

                # break for file in folder
                break
        else:
            print('Nenhum arquivo encontrado para: ' + cashier)


# call the main function passing the argument list
Aluita_Caixas(sys.argv[1:])
