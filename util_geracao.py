import pandas as pd
import sys
pasta_bibliotecas = r"Z:\Dados\Projetos_python\bibliotecas"
sys.path.append(str(pasta_bibliotecas))
import util_ressarcimento

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
TEM_0220 = False
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

divisor = 3.1

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
