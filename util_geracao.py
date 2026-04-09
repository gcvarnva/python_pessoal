import pandas as pd
import sys
pasta_bibliotecas = r"Z:\Dados\Projetos_python\bibliotecas"
sys.path.append(str(pasta_bibliotecas))
import util_ressarcimento
from tqdm.auto import tqdm
nome_pasta = ""

# Dicionário: {'código IBGE': 'tamanho da IE'}
ie_tamanhos_por_municipio = {
    '12': '13',  # AC
    '27': '9',   # AL
    '13': '9',   # AM
    '16': '9',   # AP
    '29': '9',   # BA (na verdade pode ser 8 ou 9, mas colocamos 9 como regra geral)
    '23': '9',   # CE
    '53': '13',  # DF
    '32': '9',   # ES
    '52': '9',   # GO
    '21': '9',   # MA
    '31': '13',  # MG
    '50': '9',   # MS
    '51': '11',  # MT
    '15': '9',   # PA
    '25': '9',   # PB
    '26': '9',   # PE (há casos com 14 dígitos, mas 9 é o padrão principal)
    '22': '9',   # PI
    '41': '10',  # PR
    '33': '8',   # RJ
    '24': '9',   # RN
    '43': '10',  # RS
    '11': '9',   # RO
    '14': '9',   # RR
    '42': '9',   # SC
    '28': '9',   # SE
    '35': '12',  # SP
    '17': '9'    # TO
}

MODIFICA = False
TEM_0220 = True
UNIFORMIZA_FATORES = False
SO_GERAR_ENTRADAS = False
DIVIDIR = False
CARREGA_BLOCOH = True
ICMS_TOT_PELAS_SAIDAS = True
NOME_BLOCOH = "CNPJ  Estoque 31_10_2019_EMP0009-71.TXT"
fator_icms_tot = 1.0
#fator_conversao = 3.5
# fator_conversao = 2.5
#fator_conversao = 1.0
fator_conversao = 1.0
#fator_conversao = 1.0
lista_cfops_st = ['5405', '5409', '1411', '5403', '5411', '5401', '1403', '1409']
lista_cfops_st_teste = ['1101','1102','1111','1113','1116','1117','1118','1120','1121','1122','1124','1125','1126','1151','1152','1153','1154','1201','1202','1203','1204','1205','1206','1207','1208','1209','1251','1252','1253','1254','1255','1256','1257','1301','1302','1303','1304','1305','1306','1351','1352','1353','1354','1355','1356','1360','1401','1403','1406','1407','1408','1409','1410','1411','1414','1415','1451','1452','1501','1503','1504','1505','1506','1556','1557','1601','1602','1603','1604','1605','1651','1652','1653','1658','1659','1660','1661','1662','1663','1664','1901','1902','1903','1904','1905','1906','1907','1908','1909','1910','1911','1912','1913','1914','1915','1916','1917','1918','1919','1920','1921','1922','1923','1924','1925','1926','1931','1932','1933','1949','2101','2102','2111','2113','2116','2117','2118','2120','2121','2122','2124','2125','2126','2151','2152','2153','2154','2201','2202','2203','2204','2205','2206','2207','2208','2209','2251','2252','2253','2254','2255','2256','2257','2301','2302','2303','2304','2305','2306','2351','2352','2353','2354','2355','2356','2401','2403','2406','2407','2408','2409','2410','2411','2414','2415','2501','2503','2504','2505','2506','2556','2557','2603','2651','2652','2653','2658','2659','2660','2661','2662','2663','2664','2901','2902','2903','2904','2905','2906','2907','2908','2909','2910','2911','2912','2913','2914','2915','2916','2917','2918','2919','2920','2921','2922','2923','2924','2925','2931','2932','2933','2949','3101','3102','3126','3127','3201','3202','3205','3206','3207','3211','3217','3251','3301','3351','3352','3353','3354','3355','3356','3503','3556','3651','3652','3653','3930','3949','5101','5102','5103','5104','5105','5106','5109','5110','5111','5112','5113','5114','5115','5116','5117','5118','5119','5120','5122','5123','5124','5125','5151','5152','5153','5155','5156','5201','5202','5205','5206','5207','5208','5209','5210','5251','5252','5253','5254','5255','5256','5257','5258','5301','5302','5303','5304','5305','5306','5307','5351','5352','5353','5354','5355','5356','5357','5359','5360','5401','5402','5403','5405','5408','5409','5410','5411','5412','5413','5414','5415','5451','5501','5502','5503','5504','5505','5556','5557','5601','5602','5603','5605','5606','5651','5652','5653','5654','5655','5656','5657','5658','5659','5660','5661','5662','5663','5664','5665','5666','5667','5901','5902','5903','5904','5905','5906','5907','5908','5909','5910','5911','5912','5913','5914','5915','5916','5917','5918','5919','5920','5921','5922','5923','5924','5925','5926','5927','5928','5929','5931','5932','5933','5949','6101','6102','6103','6104','6105','6106','6107','6108','6109','6110','6111','6112','6113','6114','6115','6116','6117','6118','6119','6120','6122','6123','6124','6125','6151','6152','6153','6155','6156','6201','6202','6205','6206','6207','6208','6209','6210','6251','6252','6253','6254','6255','6256','6257','6258','6301','6302','6303','6304','6305','6306','6307','6351','6352','6353','6354','6355','6356','6357','6359','6360','6401','6402','6403','6404','6408','6409','6410','6411','6412','6413','6414','6415','6501','6502','6503','6504','6505','6556','6557','6603','6651','6652','6653','6654','6655','6656','6657','6658','6659','6660','6661','6662','6663','6664','6665','6666','6667','6901','6902','6903','6904','6905','6906','6907','6908','6909','6910','6911','6912','6913','6914','6915','6916','6917','6918','6919','6920','6921','6922','6923','6924','6925','6929','6931','6932','6933','6949','7101','7102','7105','7106','7127','7201','7202','7205','7206','7207','7210','7211','7251','7301','7358','7501','7556','7651','7654','7667','7930','7949','1662','1504','2504','5661','5662','6661','6662','1661','2661','2662']
dev_entrada = ['5201', '5202', '5205', '5206', '5207', '5208', '5209', '5210', '5410', '5411', '5412', '5413', '5414', '6201', '6202', '6205', '6206', '6207', '6208', '6209', '6210', '6410', '6411', '6412', '6413', '6414', '5661', '5662', '6661', '6662']
dev_saida = ['1411']
data_corte = "1/11/2019"
file_pos_val = r'Z:\Dados\Projetos_python\Tab_CFOPs -pos_validador.xlsx'
fat = 0.5

divisor = 1.0

def dividir_com_preservacao(v):
    if pd.isna(v):
        return v
    resultado = v / divisor

    if v == int(v):  # valor original era inteiro
        resultado = int(resultado)
        return max(1, resultado)
    # Em qualquer caso, garante mínimo de 0.001
    return max(0.001, resultado)

def remover_caracteres_nao_latin1(texto):
    if isinstance(texto, str):
        texto_latin1 = texto.encode('latin-1', errors='ignore').decode('latin-1', errors='ignore')
        return texto_latin1.replace('?', '')
    return texto

def calcular_estoque_inicial_por_produto(df):
    # Garante que a coluna DATA esteja em formato datetime
    df = df.copy()
    df['DATA'] = pd.to_datetime(df['DATA'])

    # Filtra apenas as datas a partir da data_corte
    df = df[df['DATA'] >= pd.to_datetime(data_corte, dayfirst=True)]

    # Ordena para cálculo retroativo
    df = df.sort_values(['COD_ITEM', 'DATA'], ascending=[True, False])

    resultados = []

    for cod, grupo in df.groupby('COD_ITEM'):
        grupo = grupo.copy()
        grupo['saldo_dia'] = grupo['QTD_ENTRADA'] - grupo['QTD_SAIDA']

        estoque = 0
        estoques_retroativos = []

        for saldo in grupo['saldo_dia']:
            estoque = max(estoque - saldo, 0)
            estoques_retroativos.append(estoque)

        estoques_retroativos.reverse()
        grupo['estoque_dia_inicio'] = estoques_retroativos
        grupo['estoque_inicial_minimo'] = estoques_retroativos[0]  # constante por produto a partir da data_corte

        resultados.append(grupo)

    resultado_final = pd.concat(resultados).sort_values(['COD_ITEM', 'DATA'])
    return resultado_final

def gera_0200(est_ini):
    reg0200 = est_ini[['COD_ITEM', 'descricao_efd', 'Código GTIN_nota','cod_unidade_efd', 'Código NCM_nota',\
                      'aliq', 'Código CEST_nota']].\
    drop_duplicates(subset='COD_ITEM').copy()
    reg0200 = reg0200.rename(columns={'descricao_efd':'DESCR_ITEM',\
                                     'Código GTIN_nota':'COD_BARRA',\
                                     'cod_unidade_efd':'UNID_INV',\
                                     'Código NCM_nota':'COD_NCM',\
                                     'aliq':'ALIQ_ICMS',\
                                     'Código CEST_nota':'CEST'})
    reg0200['COD_BARRA'] = reg0200['COD_BARRA'].fillna('')
    reg0200['COD_BARRA'] = reg0200['COD_BARRA'].apply(lambda x: str(x).strip())
    reg0200['COD_NCM'] = reg0200['COD_NCM'].fillna('00000000')
    reg0200['COD_NCM'] = reg0200['COD_NCM'].apply(lambda x: x.zfill(8))
    reg0200['ALIQ_ICMS'] =\
    reg0200['ALIQ_ICMS'].apply(lambda x: str(int(x)).zfill(2) if pd.notna(x) and float(x)!=0 else '00')
    reg0200['CEST'] =\
    reg0200['CEST'].apply(lambda x: str(int(float(x))).zfill(7) if pd.notna(x) and float(x) != 0 else '')
    reg0200['REG'] = '0200'
    reg0200 = reg0200[['REG', 'COD_ITEM', 'DESCR_ITEM', 'COD_BARRA', 'UNID_INV', 'COD_NCM', 'ALIQ_ICMS', 'CEST']]
    return reg0200

def gera_0150(nfe_auxs, est_ini):
    nfe_auxs = nfe_auxs.loc[nfe_auxs['Chave Acesso NFe'].isin(est_ini.drop_duplicates(subset='CHV_DOC')['CHV_DOC'])].\
    drop_duplicates(subset='Número CNPJ Emitente (char)')
    reg0150 = nfe_auxs[['Número CNPJ Emitente (char)', 'Nome Razão Social Emitente',\
                        'Código País Emitente', 'Número Inscrição Estadual Completa Emitente',\
                        'Código Município Fato Gerador']].\
    drop_duplicates(subset='Número CNPJ Emitente (char)').copy()
    reg0150 = reg0150.rename(columns={'Número CNPJ Emitente (char)':'COD_PART',\
                                     'Nome Razão Social Emitente':'NOME_0150',\
                                     'Código País Emitente':'COD_PAIS',\
                                     'Número Inscrição Estadual Completa Emitente':'IE_0150',\
                                     'Código Município Fato Gerador':'COD_MUN_0150'})
    reg0150['TAM_IE'] = reg0150['COD_MUN_0150'].str[:2].map(ie_tamanhos_por_municipio).astype(int)
    reg0150['NOME_0150'] = reg0150['NOME_0150'].apply(remover_caracteres_nao_latin1)
    #reg0150['IE_0150'] = reg0150['IE_0150'].apply(lambda x: x.strip().zfill(12))
    reg0150['IE_0150'] = reg0150.apply(lambda row: row['IE_0150'].strip().zfill(row['TAM_IE']), axis=1)
    reg0150['COD_PART'] = reg0150['COD_PART'].apply(lambda x: x.strip().zfill(14))
    reg0150['REG'] = '0150'
    reg0150['CNPJ_0150'] = reg0150['COD_PART']
    reg0150['CPF'] = ''
    reg0150 = reg0150[['REG', 'COD_PART', 'NOME_0150', 'COD_PAIS', 'CNPJ_0150', 'CPF', 'IE_0150', 'COD_MUN_0150']]
    reg0150 = reg0150.drop_duplicates(subset='COD_PART')
    return reg0150

def gera_1050(est_ini_mes, reg_1050_ant):

    dados_tmp = est_ini_mes.copy()
    dados_tmp['ref'] = dados_tmp['ano_mes'].apply(lambda x: str(x)[-2:] + str(x)[:4])
    dados_tmp['date_E_S'] = dados_tmp['DATA'].apply(lambda x: x.strftime('%d%m%Y'))
    mydata1100 = dados_tmp[['ref', 'date_E_S', 'IND_OPER', 'COD_ITEM', 'CFOP', 'QTD', 'ICMS_TOT',
                            'VL_CONFR', 'COD_LEGAL']].copy()
    mydata1100 = mydata1100.rename(columns={'IND_OPER': 'typeop', 'COD_ITEM': 'item.code',
                                            'CFOP': 'cfop', 'QTD': 'quant', 'ICMS_TOT': 'icms_sup',
                                            'VL_CONFR': 'confront_value_in_S', 'COD_LEGAL': 'legal_code'})
    mydata1200 = pd.DataFrame(columns=['ref', 'date_E_S', 'typeop', 'item.code', 'cfop', 'quant',
                           'icms_sup', 'confront_value_in_S', 'legal_code'])

    dados_tmp_sdupl = dados_tmp.drop_duplicates(subset='COD_ITEM')
    mydata1050 = \
        dados_tmp_sdupl[['ref', 'COD_ITEM',\
                         'estoque_inicial_minimo', 'ICMS_TOT_INI']].copy()
    mydata1050 = mydata1050.rename(columns={'COD_ITEM': 'itemcode'})
    if reg_1050_ant.empty:
        mydata1050['iq'] = mydata1050['estoque_inicial_minimo']
        mydata1050['iv'] = mydata1050['ICMS_TOT_INI']
    else:
        mydata1050 = pd.merge(mydata1050, reg_1050_ant[['itemcode', 'fq', 'fv']], on='itemcode', how='left',
                              indicator='_merge')
        mydata1050.loc[mydata1050['_merge'] == 'both', 'iq'] = mydata1050['fq']
        mydata1050.loc[mydata1050['_merge'] == 'both', 'iv'] = mydata1050['fv']
        mydata1050.loc[mydata1050['_merge'] == 'left_only', 'iq'] = mydata1050['estoque_inicial_minimo']
        mydata1050.loc[mydata1050['_merge'] == 'left_only', 'iv'] = mydata1050['ICMS_TOT_INI']
    mydata1050.drop(columns=['estoque_inicial_minimo', 'ICMS_TOT_INI'], inplace=True)
    mydata1050['fq'] = 0
    mydata1050['fv'] = 0

    ficha3 = util_ressarcimento.gerar_ficha3(mydata1050, mydata1100, mydata1200)
    # ##################################################################################################
    # ################################### Gerando excel da ficha3 ######################################
    # ##################################################################################################
    print("Gerando excel da ficha3...", end="", flush=True)
    ficha3_excel = ficha3.copy()
    nome_ficha3_completa = r"\ficha3_completa_" + ficha3_excel['ref'].iloc[0] + ".xlsx"
    ficha3_excel.to_excel(nome_pasta + nome_ficha3_completa, index=False)
    ficha3_excel['VLR_APURACAO'] = ficha3_excel['VLR_RESSARCIMENTO'] - ficha3_excel['VLR_COMPLEMENTO']

    ficha3_resumo = ficha3_excel.groupby(['ref', 'item.code']).agg( \
        {'QTD_SALDO': 'first', 'ICMS_TOT_SALDO': 'first', 'VLR_RESSARCIMENTO': 'sum', \
         'VLR_COMPLEMENTO': 'sum', 'VLR_APURACAO': 'sum'}).reset_index()

    pivot_table_excel = ficha3_excel.\
        pivot_table(values=['VLR_RESSARCIMENTO', 'VLR_COMPLEMENTO', 'VLR_APURACAO'],\
                    index='ref',\
                    aggfunc='sum',\
                    margins=True,\
                    margins_name='Total')
    nome_ficha3_excel = r"\ficha3_" + ficha3_excel['ref'].iloc[0] + ".xlsx"
    nome_ficha3_resumo = r"\ficha3_resumo_" + ficha3_excel['ref'].iloc[0] + ".xlsx"
    ficha3_resumo.to_excel(nome_pasta + nome_ficha3_resumo, index=False)
    pivot_table_excel.to_excel(nome_pasta + nome_ficha3_excel)
    print("Concluído!")

    ficha3 = ficha3.drop_duplicates(subset='item.code', keep='last')
    ficha3 = ficha3[['item.code', 'QTD_SALDO', 'ICMS_TOT_SALDO']]
    ficha3 = ficha3.rename(columns={'item.code': 'itemcode', 'QTD_SALDO': 'fq', 'ICMS_TOT_SALDO': 'fv'})
    mydata1050 = mydata1050.drop(columns=['fq', 'fv'])
    mydata1050 = pd.merge(mydata1050, ficha3, on='itemcode', how='left', indicator='erro')
    erros = len(mydata1050.loc[mydata1050['erro']=='left_only']) + len(mydata1050.loc[mydata1050['erro']=='right_only'])
    if erros != 0:
        print("⚠️⚠️⚠️⚠️⚠️ ERRO!!!!!!! ")
    mydata1050 = mydata1050.drop(columns=['erro'])
    #limite = 0.001
    # mydata1050[['iq', 'fq', 'iv', 'fv']] = \
    #     mydata1050[['iq', 'fq', 'iv', 'fv']].applymap(lambda x: 0 if abs(x) <= limite else x)
    mydata1050[['iq', 'fq', 'iv', 'fv']] = \
        mydata1050[['iq', 'fq', 'iv', 'fv']].applymap(lambda x: 0 if x < 0 else x)


    return mydata1050

def carrega_blocoh(nome=""):

    if nome == "":
        nome_arq = NOME_BLOCOH
    else:
        nome_arq = nome

    linhas_h010 = []

    with open(nome_arq, "r", encoding="latin1") as f:
        for linha in f:
            # tira quebras de linha e separa pelos pipes
            partes = linha.strip().split("|")

            # garantir que tem pelo menos 2 colunas antes de checar índice 1
            if len(partes) > 1 and partes[1] == "H010":
                linhas_h010.append(partes)

    # transforma em DataFrame
    df_h010 = pd.DataFrame(linhas_h010)
    df_h010 = df_h010.iloc[:, 1:-1]
    df_h010.columns = ["REG", "COD_ITEM", "UNID", "QTD", "VL_UNIT",\
                       "VL_ITEM", "IND_PROP", "COD_PART", "TXT_COMPL",\
                       "COD_CTA", "VL_ITEM_IR"]

    return df_h010

def set_nome_pasta(novo_nome):
    global nome_pasta
    nome_pasta = novo_nome

def gerar_ficha3_e_1050(est_ini, criar1050 = True):

    # #######################################################################################################
    # Preparação dos Dados ##################################################################################
    # #######################################################################################################
    est_ini = est_ini.copy()
    dados_tmp = est_ini.copy()
    dados_tmp = dados_tmp.loc[dados_tmp['ano_mes'] == dados_tmp['ano_mes'].min()]
    dados_tmp['ref'] = dados_tmp['ano_mes'].apply(lambda x: str(x)[-2:] + str(x)[:4])
    dados_tmp['date_E_S'] = dados_tmp['DATA'].apply(lambda x: x.strftime('%d%m%Y'))
    mydata1100 = dados_tmp[['ref', 'date_E_S', 'IND_OPER', 'COD_ITEM', 'CFOP', 'QTD', 'ICMS_TOT',
                            'VL_CONFR', 'COD_LEGAL']].copy()
    mydata1100 = mydata1100.rename(columns={'IND_OPER': 'typeop', 'COD_ITEM': 'item.code',
                                            'CFOP': 'cfop', 'QTD': 'quant', 'ICMS_TOT': 'icms_sup',
                                            'VL_CONFR': 'confront_value_in_S', 'COD_LEGAL': 'legal_code'})
    mydata1200 = pd.DataFrame(columns=['ref', 'date_E_S', 'typeop', 'item.code', 'cfop', 'quant',
                                       'icms_sup', 'confront_value_in_S', 'legal_code'])
    dados_tmp_sdupl = dados_tmp.drop_duplicates(subset='COD_ITEM')
    mydata1050 = \
        dados_tmp_sdupl[['ref', 'COD_ITEM', \
                         'estoque_inicial_minimo', 'ICMS_TOT_INI']].copy()
    mydata1050 = mydata1050.rename(columns={'COD_ITEM': 'itemcode', 'estoque_inicial_minimo': 'iq',\
                                            'ICMS_TOT_INI': 'iv'})
    # #######################################################################################################
    # FIM (Preparação dos Dados) ############################################################################
    # #######################################################################################################


    ficha3_1 = mydata1100[['ref', 'date_E_S', 'typeop', 'item.code', 'cfop', 'quant',
                           'icms_sup', 'confront_value_in_S', 'legal_code']].copy()
    ficha3_2 = mydata1200[['ref', 'date_E_S', 'typeop', 'item.code', 'cfop', 'quant',
                           'icms_sup', 'confront_value_in_S', 'legal_code']].copy()

    if ficha3_2.empty:
        ficha3 = ficha3_1
    else:
        ficha3 = pd.concat([ficha3_1, ficha3_2], ignore_index=True)

    # Substituir strings vazias por '0' e converter para int
    ficha3['legal_code'] = ficha3['legal_code'].str.strip().replace('', '0').astype(int)

    # Converter 'ref' para datetime (assumindo que o formato é 'mmaaaa')
    ficha3['ref'] = pd.to_datetime(ficha3['ref'], format='%m%Y')

    # Converter 'date_E_S' para datetime (assumindo que o formato é 'ddmmaaaa')
    ficha3['date_E_S'] = pd.to_datetime(ficha3['date_E_S'], format='%d%m%Y')

    # Ordenar o DataFrame pelas colunas especificadas
    # ficha3 = ficha3.sort_values(by=['ref', 'item.code', 'date_E_S', 'typeop'])
    ficha3 = ficha3.sort_values(by=['ref', 'item.code', 'date_E_S'])
    ficha3 = ficha3.reset_index(drop=True)

    # ######################################################################################################
    # Leitura do Excel com os CFOPs do pós-validador #######################################################
    # ######################################################################################################
    df_cfops_ress = pd.read_excel(file_pos_val)
    cfops_que_mudam = ['1662', '1504', '2504', '5661', '5662', '6661', '6662', '1661', '2661', '2662']
    # ######################################################################################################
    # FIM (Leitura do Excel com os CFOPs do pós-validador) #################################################
    # ######################################################################################################

    # ######################################################################################################
    # Verificação de CFOPs não válidos para ressarcimento ##################################################
    # ######################################################################################################
    # Selecionando os CFOPs não válidos para ressarcimento
    cfops_nao_validos = df_cfops_ress.loc[df_cfops_ress['Válido para Ressarcimento?'] == 'N', 'CFOP']
    # Verificando se algum dos CFOPs de 'ficha3' está na lista de CFOPs não válidos
    if ficha3['cfop'].isin(cfops_nao_validos).any():
        raise ValueError("Existem CFOPs não válidos para ressarcimento no conjunto de dados.")
    else:
        print(f"Sucesso: Todos os CFOPs presentes no arquivo do Contribuinte são aplicáveis ao Ressarcimento!")
    # ######################################################################################################
    # FIM (Verificação de CFOPs não válidos para ressarcimento) ############################################
    # ######################################################################################################

    # ######################################################################################################
    # Classificação das entradas e devoluções de compras ###################################################
    # ######################################################################################################
    ficha3['E_Dev_E'] = 0.0
    # Primeiro, vamos filtrar 'df_cfops_ress' para obter apenas as linhas válidas para ressarcimento
    condicoes_entradas = (df_cfops_ress['Entrada ou Saída'] == 'E') & \
                         (df_cfops_ress['Válido para Ressarcimento?'] == 'S')

    # Filtrar CFOPs válidos com SINAL '1' e '-1'
    cfops_entradas_1 = df_cfops_ress[condicoes_entradas & (df_cfops_ress['SINAL'] == 1)]
    cfops_entradas_neg1 = df_cfops_ress[condicoes_entradas & (df_cfops_ress['SINAL'] == -1)]

    # Criar listas dos CFOPs para comparação
    cfops_lista_ent_1 = cfops_entradas_1['CFOP'].tolist()
    cfops_lista_ent_1 = [str(cfop) for cfop in cfops_lista_ent_1]
    cfops_lista_ent_neg1 = cfops_entradas_neg1['CFOP'].tolist()
    cfops_lista_ent_neg1 = [str(cfop) for cfop in cfops_lista_ent_neg1]

    # Usar `loc` para definir os valores em 'E_Dev_E' baseado nos CFOPs filtrados
    ficha3.loc[ficha3['cfop'].isin(cfops_lista_ent_1), 'E_Dev_E'] = ficha3['quant']
    ficha3.loc[ficha3['cfop'].isin(cfops_lista_ent_neg1), 'E_Dev_E'] = -ficha3['quant']
    # ######################################################################################################
    # FIM (Classificação das entradas e devoluções de compras) #############################################
    # ######################################################################################################

    # ######################################################################################################
    # Classificação das saídas e devoluções de saídas ######################################################
    # ######################################################################################################
    ficha3['S_Dev_S'] = 0.0
    # Primeiro, vamos filtrar 'df_cfops_ress' para obter apenas as linhas válidas para ressarcimento
    condicoes_saidas = (df_cfops_ress['Entrada ou Saída'] == 'S') & \
                       (df_cfops_ress['Válido para Ressarcimento?'] == 'S')

    # Filtrar CFOPs válidos com SINAL '1' e '-1'
    cfops_saidas_1 = df_cfops_ress[condicoes_saidas & (df_cfops_ress['SINAL'] == 1)]
    cfops_saidas_neg1 = df_cfops_ress[condicoes_saidas & (df_cfops_ress['SINAL'] == -1)]

    # Criar listas dos CFOPs para comparação
    cfops_lista_sai_1 = cfops_saidas_1['CFOP'].tolist()
    cfops_lista_sai_1 = [str(cfop) for cfop in cfops_lista_sai_1]
    cfops_lista_sai_neg1 = cfops_saidas_neg1['CFOP'].tolist()
    cfops_lista_sai_neg1 = [str(cfop) for cfop in cfops_lista_sai_neg1]

    # Usar `loc` para definir os valores em 'E_Dev_E' baseado nos CFOPs filtrados
    ficha3.loc[ficha3['cfop'].isin(cfops_lista_sai_1), 'S_Dev_S'] = ficha3['quant']
    ficha3.loc[ficha3['cfop'].isin(cfops_lista_sai_neg1), 'S_Dev_S'] = -ficha3['quant']
    # ######################################################################################################
    # FIM (Classificação das saídas e devoluções de saídas) ################################################
    # ######################################################################################################

    ficha3['ref'] = ficha3['ref'].dt.strftime('%m%Y')

    #########################################################################################
    # Gera a Ficha 3 da PCAT ################################################################
    #########################################################################################
    # Criar as colunas abaixo, todas vazias em ficha3
    ficha3['QTD_SALDO'] = 0.0
    ficha3['ICMS_SAIDA_UNI'] = 0.0
    ficha3['ICMS_SAIDA'] = 0.0
    ficha3['ICMS_TOT_SALDO'] = 0.0
    ficha3['VLR_RESSARCIMENTO'] = 0.0
    ficha3['VLR_COMPLEMENTO'] = 0.0

    ref_ant = ficha3.loc[0, 'ref']
    primeira_linha = True

    ficha3_ref_filtrado = ficha3[ficha3['ref'] == ref_ant]
    # Criando o dicionário inicial com todos os valores de 'item.code' como False
    unique_item_codes = ficha3_ref_filtrado['item.code'].unique()
    item_status = {code: False for code in unique_item_codes}

    df_icms_tot_ini = est_ini[['COD_ITEM', 'estoque_inicial_minimo', 'ICMS_TOT_INI']].\
        drop_duplicates(subset='COD_ITEM').copy()
    df_icms_tot_ini = df_icms_tot_ini.rename(columns={'estoque_inicial_minimo': 'QTD_INI'})
    df_icms_tot_ini['calculado'] = False
    # Iterar sobre as linhas do DataFrame
    for index, row in tqdm(ficha3.iterrows(), total=ficha3.shape[0], desc="Processando"):
        itemcode_atual = row['item.code']
        ref_atual = row['ref']

        cfop_tmp = row['cfop']
        if cfop_tmp in cfops_que_mudam:
            zes = aplicar_regra_linha(cfop_tmp, row['ref'])
            if zes == 'E':
                ficha3.at[index, 'E_Dev_E'] = 0
            elif zes == 'S':
                ficha3.at[index, 'S_Dev_S'] = 0

        # Verifica se o 'ref' mudou, se sim, resete todos os valores do dicionário para False
        if ref_atual != ref_ant:
            ficha3_ref_filtrado = ficha3[ficha3['ref'] == ref_atual]
            # Criando o dicionário inicial com todos os valores de 'item.code' como False
            unique_item_codes = ficha3_ref_filtrado['item.code'].unique()
            item_status = {code: False for code in unique_item_codes}

        if (primeira_linha) or (ref_atual != ref_ant) or (not item_status[itemcode_atual]):
            primeira_linha = False
            ref_ant = ref_atual
            # Atualiza o status de 'item.code' para True
            item_status[itemcode_atual] = True

            try:
                # Obter os valores iniciais de quantidade e de ICMS de 'mydata1050'
                valor_inicial = mydata1050.loc[
                    (mydata1050['itemcode'] == row['item.code']) & (mydata1050['ref'] == row['ref']), 'iq'].item()
                valor_inicial_icms = mydata1050.loc[(mydata1050['itemcode'] == row['item.code']) & \
                    (mydata1050['ref'] == row['ref']), 'iv'].item()
                # if criar1050:
                #     valor_inicial_icms = (valor_inicial * row['confront_value_in_S'] * \
                #                          (1 + fat)) / row['S_Dev_S']
                #     # mydata1050.loc[(mydata1050['itemcode'] == row['item.code']) & (mydata1050['ref'] == row['ref']), 'iv'] = valor_inicial_icms
                #     est_ini.loc[(est_ini['COD_ITEM'] == row['item.code']), 'ICMS_TOT_INI'] = valor_inicial_icms
                # else:
            except:
                print("DEBUG")
                raise ValueError("Condição impossível: Há códigos de produto não encontrados no registro 1050!")

            # Inicializar o primeiro valor de 'QTD_SALDO'
            ficha3.at[index, 'QTD_SALDO'] = valor_inicial + row['E_Dev_E'] - row['S_Dev_S']
            # Cálculo do primeiro valor de 'ICMS_SAIDA_UNI'
            if (row['S_Dev_S'] > 0):
                valor_inicial_icms = (valor_inicial * row['confront_value_in_S']\
                                      * (1 + fat)) / row['S_Dev_S']
                df_icms_tot_ini.loc[df_icms_tot_ini['COD_ITEM'] == row['item.code'],\
                    'ICMS_TOT_INI'] = valor_inicial_icms
                df_icms_tot_ini.loc[df_icms_tot_ini['COD_ITEM'] == row['item.code'], \
                    'calculado'] = True
                if valor_inicial == 0:
                    set_trace()
                    raise ValueError("Não pode haver saída com estoque zerado!")
                else:
                    ficha3.at[index, 'ICMS_SAIDA_UNI'] = valor_inicial_icms / valor_inicial
                    ficha3.at[index, 'ICMS_SAIDA'] = row['S_Dev_S'] * ficha3.at[index, 'ICMS_SAIDA_UNI']
                    ficha3.at[index, 'ICMS_TOT_SALDO'] = valor_inicial_icms - ficha3.at[index, 'ICMS_SAIDA']
                    if (row['legal_code'] == 1 or row['legal_code'] == 2 or row['legal_code'] == 3 or row[
                        'legal_code'] == 4):
                        if ficha3.at[index, 'ICMS_SAIDA'] > row['confront_value_in_S']:
                            ficha3.at[index, 'VLR_RESSARCIMENTO'] = ficha3.at[index, 'ICMS_SAIDA'] - row[
                                'confront_value_in_S']
                        else:
                            if (row['legal_code'] == 1):
                                ficha3.at[index, 'VLR_COMPLEMENTO'] = row['confront_value_in_S'] - ficha3.at[
                                    index, 'ICMS_SAIDA']
            elif (row['S_Dev_S'] < 0):
                ficha3.at[index, 'ICMS_SAIDA'] = -row['icms_sup']
                ficha3.at[index, 'ICMS_TOT_SALDO'] = valor_inicial_icms + row['icms_sup']
                ficha3.at[index, 'ICMS_SAIDA_UNI'] = ficha3.at[index, 'ICMS_TOT_SALDO'] / ficha3.at[index, 'QTD_SALDO']
                if (row['legal_code'] == 1 or row['legal_code'] == 2 or row['legal_code'] == 3 or row[
                    'legal_code'] == 4):
                    if abs(ficha3.at[index, 'ICMS_SAIDA']) > abs(row['confront_value_in_S']):
                        ficha3.at[index, 'VLR_RESSARCIMENTO'] = ficha3.at[index, 'ICMS_SAIDA'] + row[
                            'confront_value_in_S']
                    else:
                        if (row['legal_code'] == 1):
                            ficha3.at[index, 'VLR_COMPLEMENTO'] = -row['confront_value_in_S'] - ficha3.at[
                                index, 'ICMS_SAIDA']
            elif (row['E_Dev_E'] > 0):
                ficha3.at[index, 'ICMS_SAIDA'] = 0
                ficha3.at[index, 'ICMS_TOT_SALDO'] = valor_inicial_icms + row['icms_sup']
                ficha3.at[index, 'ICMS_SAIDA_UNI'] = ficha3.at[index, 'ICMS_TOT_SALDO'] / ficha3.at[index, 'QTD_SALDO']
            elif (row['E_Dev_E'] < 0):
                ficha3.at[index, 'ICMS_SAIDA'] = 0
                ficha3.at[index, 'ICMS_TOT_SALDO'] = valor_inicial_icms - row['icms_sup']
                if ficha3.at[index, 'QTD_SALDO'] == 0:
                    ficha3.at[index, 'ICMS_TOT_SALDO'] = 0
                    ficha3.at[index, 'ICMS_SAIDA_UNI'] = 0
                else:
                    ficha3.at[index, 'ICMS_SAIDA_UNI'] = ficha3.at[index, 'ICMS_TOT_SALDO'] / ficha3.at[
                        index, 'QTD_SALDO']
            else:
                raise ValueError(
                    "Condição impossível: o arquivo PCAT está com CFOPs que não valem para o ressarcimento!")


        # ###########################################################################################
        # Segunda linha em diante de cada produto ###################################################
        # ###########################################################################################
        else:

            ficha3.at[index, 'QTD_SALDO'] = ficha3.at[index - 1, 'QTD_SALDO'] + row['E_Dev_E'] - row['S_Dev_S']

            # Cálculo do primeiro valor de 'ICMS_SAIDA_UNI'
            if (row['S_Dev_S'] > 0):
                if not (df_icms_tot_ini.loc[df_icms_tot_ini['COD_ITEM'] == row['item.code'],\
                        'calculado'].iloc[0]):
                    valor_inicial_icms = (valor_inicial * row['confront_value_in_S'] \
                                          * (1 + fat)) / row['S_Dev_S']
                    df_icms_tot_ini.loc[df_icms_tot_ini['COD_ITEM'] == row['item.code'], \
                        'ICMS_TOT_INI'] = valor_inicial_icms
                    df_icms_tot_ini.loc[df_icms_tot_ini['COD_ITEM'] == row['item.code'], \
                        'calculado'] = True
                ficha3.at[index, 'ICMS_SAIDA_UNI'] = ficha3.at[index - 1, 'ICMS_SAIDA_UNI']
                ficha3.at[index, 'ICMS_SAIDA'] = row['S_Dev_S'] * ficha3.at[index, 'ICMS_SAIDA_UNI']
                ficha3.at[index, 'ICMS_TOT_SALDO'] = ficha3.at[index - 1, 'ICMS_TOT_SALDO'] - ficha3.at[
                    index, 'ICMS_SAIDA']
                if (row['legal_code'] == 1 or row['legal_code'] == 2 or row['legal_code'] == 3 or row[
                    'legal_code'] == 4):
                    if ficha3.at[index, 'ICMS_SAIDA'] > row['confront_value_in_S']:
                        ficha3.at[index, 'VLR_RESSARCIMENTO'] = ficha3.at[index, 'ICMS_SAIDA'] - row[
                            'confront_value_in_S']
                    else:
                        if (row['legal_code'] == 1):
                            ficha3.at[index, 'VLR_COMPLEMENTO'] = row['confront_value_in_S'] - ficha3.at[
                                index, 'ICMS_SAIDA']
            elif (row['S_Dev_S'] < 0):
                ficha3.at[index, 'ICMS_SAIDA'] = -row['icms_sup']
                ficha3.at[index, 'ICMS_TOT_SALDO'] = ficha3.at[index - 1, 'ICMS_TOT_SALDO'] + row['icms_sup']
                ficha3.at[index, 'ICMS_SAIDA_UNI'] = ficha3.at[index, 'ICMS_TOT_SALDO'] / ficha3.at[index, 'QTD_SALDO']
                if (row['legal_code'] == 1 or row['legal_code'] == 2 or row['legal_code'] == 3 or row[
                    'legal_code'] == 4):
                    if abs(ficha3.at[index, 'ICMS_SAIDA']) > abs(row['confront_value_in_S']):
                        ficha3.at[index, 'VLR_RESSARCIMENTO'] = ficha3.at[index, 'ICMS_SAIDA'] + row[
                            'confront_value_in_S']
                    else:
                        if (row['legal_code'] == 1):
                            ficha3.at[index, 'VLR_COMPLEMENTO'] = -row['confront_value_in_S'] - ficha3.at[
                                index, 'ICMS_SAIDA']
            elif (row['E_Dev_E'] > 0):
                ficha3.at[index, 'ICMS_SAIDA'] = 0
                ficha3.at[index, 'ICMS_TOT_SALDO'] = ficha3.at[index - 1, 'ICMS_TOT_SALDO'] + row['icms_sup']
                ficha3.at[index, 'ICMS_SAIDA_UNI'] = ficha3.at[index, 'ICMS_TOT_SALDO'] / ficha3.at[index, 'QTD_SALDO']
            elif (row['E_Dev_E'] < 0):
                ficha3.at[index, 'ICMS_SAIDA'] = 0
                ficha3.at[index, 'ICMS_TOT_SALDO'] = ficha3.at[index - 1, 'ICMS_TOT_SALDO'] - row['icms_sup']
                if ficha3.at[index, 'QTD_SALDO'] == 0:
                    ficha3.at[index, 'ICMS_TOT_SALDO'] = 0
                    ficha3.at[index, 'ICMS_SAIDA_UNI'] = 0
                else:
                    ficha3.at[index, 'ICMS_SAIDA_UNI'] = ficha3.at[index, 'ICMS_TOT_SALDO'] / ficha3.at[
                        index, 'QTD_SALDO']
            else:
                print(f"index = {index}")
                linha = ficha3.iloc[index]
                print(linha)
                raise ValueError(
                    "Condição impossível: o arquivo PCAT está com CFOPs que não valem para o ressarcimento!")

    return ficha3, df_icms_tot_ini
    #########################################################################################
    # Gera a Ficha 3 da PCAT ################################################################
    #########################################################################################

def mediana_ponderada_por_grupo(df, group_cols, icms_col="icms_sup", qtd_col="qtd_efd",
                               out_col="mediana_pond_unit"):
    # unitário por entrada
    unit = df[icms_col] / df[qtd_col]

    # base mínima: chaves + unit + peso
    base = df.loc[:, list(group_cols)].copy()
    base["_unit_"] = unit.to_numpy()
    base["_w_"] = df[qtd_col].to_numpy()
    base["_icms_"] = df[icms_col].to_numpy()  # ← nova linha

    # 🔑 EXCLUI icms_sup == 0 DO CÁLCULO
    base = base.loc[base["_icms_"] > 0]  # ← novo filtro

    # ordena por grupo + unit
    base.sort_values(list(group_cols) + ["_unit_"], inplace=True, kind="mergesort")

    # pesos total e acumulado por grupo
    g = base.groupby(list(group_cols), sort=False)
    w_total = g["_w_"].transform("sum")
    w_cum = g["_w_"].cumsum()

    # ponto 50%: pega o primeiro unit onde cum >= total/2
    half = w_total / 2.0
    hit = base.loc[w_cum >= half, list(group_cols) + ["_unit_"]]
    med = hit.groupby(list(group_cols), sort=False)["_unit_"].first()

    # devolve alinhado ao df original (sem merge)
    keys = pd.MultiIndex.from_frame(df[list(group_cols)])
    out = pd.Series(keys.map(med), index=df.index, name=out_col)
    return out.fillna(0.0)

def relatorio_icms_suportado(entradas, nfe_sp_aliquotas, nome_pasta):
    print("Gerando relatório do ICMS suportado de entrada...", end="", flush=True)
    # entradas_resumidas = \
    #     entradas[['COD_ITEM', 'descricao_efd', 'cod_unidade_efd',
    #               'vlr_prod', 'QTD', 'icms', 'icms_st', 'bc_icms_st_ant',
    #               'DATA']]
    # teste = pd.merge(entradas_resumidas, nfe_sp_aliquotas,
    #                  left_on='Código Produto ou Serviço_nota', right_on='COD_ITEM', how='left',
    #                  indicator=True)
    # teste['aliq'] = teste['aliq'].fillna(0)
    # teste['icms_sup_1'] = (teste['Valor ICMS Operação_nota'] +
    #                        teste['Valor ICMS Substituição Tributária_nota'])
    # teste['icms_sup_1'] = (teste['icms_sup_1'] * 100).round().astype('int64') / 100
    # teste['icms_sup_2'] = (teste['Valor Base Cálculo ICMS ST Retido Operação Anterior_nota']
    #                        * teste['aliq'] / 100)
    # teste['icms_sup_2'] = (teste['icms_sup_2'] * 100).round().astype('int64') / 100
    # teste['icms_sup'] = np.maximum(teste['icms_sup_1'].fillna(0), teste['icms_sup_2'].fillna(0))


    teste = entradas[['COD_ITEM', 'descricao_efd', 'cod_unidade_efd', 'QTD',
                      'ICMS_TOT', 'DATA']].copy()
    teste['icms_sup_unit'] = teste['ICMS_TOT'] / teste['QTD']
    teste['icms_sup_unit'] = (teste['icms_sup_unit'] * 100).round().astype('int64') / 100
    teste = teste.sort_values(by=['COD_ITEM', 'cod_unidade_efd', 'icms_sup_unit'],
                              ascending=[True, True, False])
    caminho_completo = nome_pasta + r'\relatorio_icms_sup_entradas.xlsx'
    g = teste.groupby(['COD_ITEM', 'cod_unidade_efd'])['icms_sup_unit']
    teste['media_icms_sup'] = g.transform('mean')
    teste['media_icms_sup'] = (teste['media_icms_sup'] * 100).round().astype('int64') / 100
    teste['mediana_icms_sup'] = g.transform('median')
    teste['mediana_icms_sup'] = (teste['mediana_icms_sup'] * 100).round().astype('int64') / 100

    gcols = ["COD_ITEM", "cod_unidade_efd"]

    teste["mediana_pond_icms_sup_unit"] = mediana_ponderada_por_grupo(
        teste, gcols,
        icms_col="ICMS_TOT",
        qtd_col="QTD",
        out_col="mediana_pond_icms_sup_unit"
    )
    teste["mediana_pond_icms_sup_unit"] = (teste["mediana_pond_icms_sup_unit"] * 100).round().astype('int64') / 100
    # teste.to_excel(caminho_completo, index=False)
    teste_resumo = teste[['COD_ITEM', 'descricao_efd', 'cod_unidade_efd', 'QTD',
                          'DATA', 'ICMS_TOT', 'icms_sup_unit', 'media_icms_sup',
                          'mediana_icms_sup', 'mediana_pond_icms_sup_unit']]
    teste_resumo.to_excel(caminho_completo, index=False)

    caminho_completo = nome_pasta + r'\fat_conv_entradas.xlsx'
    teste_p_preencher = teste_resumo.copy()
    teste_p_preencher = teste_p_preencher.drop_duplicates(subset=['COD_ITEM', 'cod_unidade_efd'])
    teste_p_preencher['fat_conv'] = 1
    teste_p_preencher = teste_p_preencher[['COD_ITEM', 'cod_unidade_efd', 'fat_conv']]
    teste_p_preencher.to_excel(caminho_completo, index=False)

    print("relatório gerado!")

def relatorio_cfops_entrada(entradas_e_c170, nome_pasta):
    print("gerando CFOPs de entrada da EFD... ", end="", flush=True)
    cfops_c170 = entradas_e_c170.groupby(['cfop_efd']) \
        .agg({'qtd_efd': 'sum'}).reset_index().sort_values(by='qtd_efd', ascending=False)
    cfops_c170['cfop_efd'] = cfops_c170['cfop_efd'].astype(str).str.strip()
    cfops_c170['Desc_CFOPs'] = cfops_c170['cfop_efd'].map(desc_cfop)
    caminho_completo = nome_pasta + r'\cfops_entradas.xlsx'
    cfops_c170.to_excel(caminho_completo, index=False)
    print("concluído!")

def relatorio_ligacoes_entradas(entradas_e_c170, nome_pasta):

    print("gerando tabela de ligações das entradas... ", end="", flush=True)
    ligacoes_entradas = entradas_e_c170.groupby(['Código Produto ou Serviço_nota', \
                                                 'Descrição Produto_nota', \
                                                 'item_cod_efd_efd', \
                                                 'descricao_efd']) \
        .agg({'Valor Produto ou Serviço_nota': 'sum', 'Valor ICMS Operação_nota': 'sum', \
              'Valor ICMS Substituição Tributária_nota': 'sum', \
              'Valor Base Cálculo ICMS ST Retido Operação Anterior_nota': 'sum'}).reset_index()
    caminho_completo = nome_pasta + r'\ligacoes_entradas.xlsx'
    ligacoes_entradas.to_excel(caminho_completo, index=False)
    print("concluído!")

def uniformiza_fatores(entradas_e_c170, nome_pasta):
    entradas_e_c170['razao'] = entradas_e_c170['qtd_efd'].astype(float) \
                               / entradas_e_c170['Quantidade Comercial_nota'].astype(float)

    print("gerando tabela de conversão de unidades das entradas... ", end="", flush=True)
    conv_unid_efd = entradas_e_c170[
        ['Código Produto ou Serviço_nota', 'Unidade Comercial_nota', \
         'item_cod_efd_efd', 'cod_unidade_efd', 'razao']].drop_duplicates()
    caminho_completo = nome_pasta + r'\conversao_unidades_efd.xlsx'
    conv_unid_efd.to_excel(caminho_completo, index=False)
    print("Concluído!")

    mapeamento = entradas_e_c170[
        ['Código Produto ou Serviço_nota', 'Unidade Comercial_nota', 'item_cod_efd_efd', 'razao']]
    entradas_e_c170 = entradas_e_c170.drop('razao', axis=1)
    mapeamento_distinto = mapeamento.drop_duplicates()
    grupo_chave = ['Código Produto ou Serviço_nota', 'Unidade Comercial_nota', 'item_cod_efd_efd']
    # Para cada grupo, pegar a linha com o maior valor de 'razao'
    resultado = (
        mapeamento_distinto
        .sort_values('razao')  # organiza para que o menor venha primeiro
        .drop_duplicates(subset=grupo_chave, keep='first')  # PEGA O MAIOR VALOR DA RAZÃO
    )

    entradas_e_c170 = pd.merge(entradas_e_c170, resultado,
                               on=['Código Produto ou Serviço_nota', 'Unidade Comercial_nota', 'item_cod_efd_efd'],
                               how='left', indicator='entrada_e_razao')

    if MODIFICA:
        entradas_e_c170.loc[entradas_e_c170['razao'] < 1, 'razao'] = 1
        entradas_e_c170['razao'] = np.floor(entradas_e_c170['razao'] * fator_conversao)

    entradas_e_c170['qtd_efd'] = entradas_e_c170['Quantidade Comercial_nota'] * \
                                 entradas_e_c170['razao']

def converte_efd0220(entradas, nome_pasta):

    entradas['fat_0220'] = 1
    nome_efd0220 = f"{nome_pasta}/efd0220_completo.xlsx"
    # Caminho para o arquivo Excel já definido em nome_efd0220
    # Lendo apenas as colunas desejadas
    efd0220 = pd.read_excel(
        nome_efd0220,
        usecols=['Data Referência (AAAAMM)', 'Código Item', 'Descrição Unidade Conversão', 'Fator Conversão Unidade']
    )

    # Renomeando as colunas
    efd0220.rename(columns={
        'Data Referência (AAAAMM)': 'mes_ref_efd',
        'Código Item': 'COD_ITEM',
        'Descrição Unidade Conversão': 'cod_unidade_efd',
        'Fator Conversão Unidade': 'fat_novo'
    }, inplace=True)

    # Aplicando strip para evitar espaços em branco
    efd0220['mes_ref_efd'] = efd0220['mes_ref_efd'].astype(str).str.strip()
    efd0220['COD_ITEM'] = efd0220['COD_ITEM'].astype(str).str.strip()
    efd0220['cod_unidade_efd'] = efd0220['cod_unidade_efd'].astype(str).str.strip()

    # Merge temporário para obter o fat_novo correspondente
    df_merge = entradas.merge(
        efd0220[['mes_ref_efd', 'COD_ITEM', 'cod_unidade_efd', 'fat_novo']],
        on=['mes_ref_efd', 'COD_ITEM', 'cod_unidade_efd'],
        how='left'
    )
    # Atualizar fat_0220 apenas onde houver valor correspondente em fat_novo
    entradas['fat_0220'] = df_merge['fat_novo'].combine_first(entradas['fat_0220'])

    # entradas.loc[entradas['fat_0220']>=2, 'fat_0220'] = np.trunc(entradas['fat_0220']/2)
    # entradas.loc[entradas['fat_0220'] >= 4, 'fat_0220'] = np.trunc(entradas['fat_0220'] / 4)
    # entradas.loc[entradas['fat_0220'] >= 8, 'fat_0220'] = np.trunc(entradas['fat_0220'] / 8)
    # entradas.loc[entradas['fat_0220'] >= 12, 'fat_0220'] = np.trunc(entradas['fat_0220'] / 12)
    # entradas.loc[entradas['fat_0220'] >= 16, 'fat_0220'] = np.trunc(entradas['fat_0220'] / 16)
    # entradas.loc[entradas['fat_0220'] >= 2, 'fat_0220'] = 2

    entradas['QTD'] = entradas['QTD'] * entradas['fat_0220']
    if DIVIDIR:
        entradas['QTD'] = entradas['QTD'].apply(dividir_com_preservacao)
    entradas = entradas.drop(columns=['fat_0220'])

def relatorio_coditens_unids_diversas(entradas, nome_pasta):

    print("Gerando relatório dos códigos de item que têm unidades diferentes nas entradas...", end="", flush=True)
    # Passo 1: Contar unidades distintas por mes_ref_efd e COD_ITEM
    unidades_por_item_mes = entradas.groupby(['mes_ref_efd', 'COD_ITEM'])['cod_unidade_efd'].nunique().reset_index(
        name='qtd_unidades')

    # Passo 2: Filtrar apenas os que têm mais de uma unidade distinta
    itens_com_varias_unidades = unidades_por_item_mes[unidades_por_item_mes['qtd_unidades'] > 1]

    # Passo 3: Filtrar os dados originais com base nos que têm múltiplas unidades
    relatorio = entradas.merge(
        itens_com_varias_unidades[['mes_ref_efd', 'COD_ITEM']],
        on=['mes_ref_efd', 'COD_ITEM'],
        how='inner'
    )

    # Passo 4: Selecionar colunas relevantes e remover duplicatas
    relatorio = relatorio[['mes_ref_efd', 'COD_ITEM', 'cod_unidade_efd']].drop_duplicates()

    # Passo 5: Ordenar para visualização
    relatorio = relatorio.sort_values(by=['mes_ref_efd', 'COD_ITEM', 'cod_unidade_efd'])
    relatorio = relatorio.rename(columns={'cod_unidade_efd': 'cod_unidade_C170'})

    # Passo 6: Exportar para Excel
    nome_rel = f"{nome_pasta}/Relatorio-{CNPJ_ENTIDADE}_cod_item_entradas_unidades_diferentes.xlsx"
    relatorio.to_excel(nome_rel, index=False)