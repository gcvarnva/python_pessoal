import util_gui
import pandas as pd
import numpy as np
import os
import util_ressarcimento
from datetime import datetime
import matplotlib.pyplot as plt

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



dir_inicial = r"C:\Gustavo\Ressarcimento\Analises-DRT3\Simpatia\dados"
caminho_pasta = util_gui.selecionar_pasta(dir_inicial)
data_hora_atual = datetime.now().strftime('%d-%m-%Y_%Hh%Mm%Ss')
nome_pasta = f"{caminho_pasta}/pcat_{data_hora_atual}_ficms_{fator_icms_tot}_fconv_{fator_conversao}_div_{divisor}"
os.makedirs(nome_pasta, exist_ok=True)

# Leitura do Registro 0000 da EFD
caminho_completo = caminho_pasta + r'\EFD_Reg_0000.csv'
EFD_Reg_0000 = pd.read_csv(caminho_completo, sep=';', encoding='utf-8', quotechar='"', skiprows=3, dtype=str,\
                           usecols=['Nome Entidade', 'Mês Referência', 'Número CNPJ',\
                                    'Número Inscrição Estadual' ,'Código Município'])

# ############################################################################################################# #
# Leitura dos arquivos auxiliares da geração (registro 0150) ################################################## #
# ############################################################################################################# #
caminho_completo = caminho_pasta + r"\nfe_entrada_outras-p_geracao.csv"
nfe_ufs_auxs = pd.read_csv(caminho_completo, sep=';', encoding='utf-8', quotechar='"', skiprows=3, dtype=str,\
                              usecols=['Chave Acesso NFe', 'Número CNPJ Emitente (char)', 'Nome Razão Social Emitente',\
                                      'Código País Emitente' ,'Código Município Emitente',\
                                      'Número Inscrição Estadual Completa Emitente'])
nfe_ufs_auxs = nfe_ufs_auxs.drop_duplicates(subset=['Chave Acesso NFe'])
nfe_ufs_auxs = nfe_ufs_auxs.rename(columns={'Código Município Emitente': 'Código Município Fato Gerador'})

caminho_completo = caminho_pasta + r"\nfe_entrada_sp_aliquotas.csv"
nfe_sp_auxs = pd.read_csv(caminho_completo, sep=';', encoding='utf-8', quotechar='"', skiprows=3, dtype=str,\
                              usecols=['Chave Acesso NFe', 'Número CNPJ Emitente', 'Nome Razão Social Emitente',\
                                      'Código País Emitente' ,'Código Município Fato Gerador',\
                                      'Número Inscrição Estadual Completa Emitente', 'Número CPF Emitente'])
nfe_sp_auxs = nfe_sp_auxs.drop_duplicates(subset=['Chave Acesso NFe'])
nfe_sp_auxs['Número CNPJ Emitente'] = nfe_sp_auxs['Número CNPJ Emitente'].apply(lambda x: str(int(x)).zfill(14))
nfe_sp_auxs.drop(columns='Número CPF Emitente', inplace=True)
nfe_sp_auxs = nfe_sp_auxs.rename(columns={'Número CNPJ Emitente': 'Número CNPJ Emitente (char)'})

nfe_sp_aliquotas = pd.read_csv(caminho_completo, sep=';', encoding='utf-8', quotechar='"', skiprows=3, dtype=str,\
                              usecols=['Código Produto ou Serviço', 'Percentual Alíquota ICMS'])
nfe_sp_aliquotas = nfe_sp_aliquotas.rename(columns={'Código Produto ou Serviço': 'COD_ITEM',\
                                                    'Percentual Alíquota ICMS': 'aliq'})
# Converte a coluna 'aliq' para float (se necessário)
nfe_sp_aliquotas['aliq'] = nfe_sp_aliquotas['aliq'].str.replace(',', '.').astype(float)

# Seleciona apenas o maior valor de 'aliq' por 'COD_ITEM'
nfe_sp_aliquotas = (
    nfe_sp_aliquotas.groupby('COD_ITEM', as_index=False)['aliq']
    .max()
)
nfe_sp_aliquotas['aliq'] = nfe_sp_aliquotas['aliq'].fillna(0)

nfe_auxs = pd.concat([nfe_sp_auxs, nfe_ufs_auxs], ignore_index=True)
nfe_auxs = nfe_auxs.drop_duplicates(subset=['Chave Acesso NFe'])
nfe_auxs.loc[nfe_auxs['Código País Emitente'].isna(), 'Código País Emitente']='1058'
# ############################################################################################################# #
# FIM (Leitura dos arquivos auxiliares da geração (registro 0150)) ############################################ #
# ############################################################################################################# #

print("Lendo tabela de cnpjs...")
caminho_completo = caminho_pasta + r"\dados_cnpj.xlsx"
tab_cnpjs = pd.read_excel(caminho_completo, dtype=str)

# #####################################################################################################
# Lendo tabela_1 (notas de entrada e saída) ###########################################################
# #####################################################################################################
print("Lendo tabela das notas de entrada e saída...")
caminho_completo = caminho_pasta + r"\tabela_1.csv"
tabela_1 = pd.read_csv(caminho_completo, sep=';', encoding='utf-8', quotechar='"', dtype=str)

tabela_1['Data Emissão'] = pd.to_datetime(tabela_1['Data Emissão'],format='%Y-%m-%d')
tabela_1['Valor Produto ou Serviço'] = tabela_1['Valor Produto ou Serviço'].astype(float).fillna(0)
tabela_1['Valor ICMS Operação'] = tabela_1['Valor ICMS Operação'].astype(float).fillna(0)
tabela_1['Valor ICMS Substituição Tributária'] = tabela_1['Valor ICMS Substituição Tributária'].astype(float).fillna(0)

tabela_1['Valor Base Cálculo ICMS Substituição Tributária'] = \
tabela_1['Valor Base Cálculo ICMS Substituição Tributária'].astype(float).fillna(0)

tabela_1['Valor Base Cálculo ICMS ST Retido Operação Anterior'] = \
tabela_1['Valor Base Cálculo ICMS ST Retido Operação Anterior'].astype(float).fillna(0)

tabela_1['Valor Unitário Comercial'] = tabela_1['Valor Unitário Comercial'].astype(float).fillna(0)
tabela_1['Quantidade Comercial'] = tabela_1['Quantidade Comercial'].astype(float).fillna(0)
tabela_1['Número Item'] = tabela_1['Número Item'].astype(int)
tabela_1['IND_OPER'] = np.where((tabela_1['Tipo'] == 'dfe')
                                | (tabela_1['Tipo'] == 'saida_nfe')
                                | (tabela_1['Tipo'] == 'saida_nfce'), 1, 0)
tabela_1.columns = [col + '_nota' for col in tabela_1.columns]
# #####################################################################################################
# FIM (Leitura da tabela_1) ###########################################################################
# #####################################################################################################


# #####################################################################################################
# Entradas de Operação Própria ########################################################################
# #####################################################################################################
entradas_oppr = tabela_1.loc[tabela_1['Tipo_nota']=='entrada_oprprp',\
             ['Chave Acesso NFe_nota', 'Data Emissão_nota', 'Número Item_nota', 'IND_OPER_nota',\
              'Código Produto ou Serviço_nota', 'Código CFOP (04 Posições)_nota', 'Quantidade Comercial_nota',\
              'Valor Produto ou Serviço_nota',\
              'Valor ICMS Operação_nota', 'Valor ICMS Substituição Tributária_nota',\
              'Valor Base Cálculo ICMS ST Retido Operação Anterior_nota', 'Descrição Produto_nota',\
              'Número CNPJ Emitente_nota', 'Código GTIN_nota', 'Código NCM_nota',\
              'Código CEST_nota', 'Unidade Comercial_nota']]
entradas_oppr = \
entradas_oppr.loc[entradas_oppr['Código CFOP (04 Posições)_nota'].isin(lista_cfops_st)]
entradas_oppr = \
entradas_oppr.rename(columns={'Chave Acesso NFe_nota':'CHV_DOC',\
                              'Data Emissão_nota':'DATA',\
                              'Número Item_nota':'NUM_ITEM',\
                              'IND_OPER_nota':'IND_OPER',
                              'Código Produto ou Serviço_nota':'COD_ITEM',\
                              'Código CFOP (04 Posições)_nota': 'CFOP',\
                              'Quantidade Comercial_nota': 'QTD',\
                              'Valor Produto ou Serviço_nota': 'vlr_prod',\
                              'Valor ICMS Operação_nota': 'icms',\
                              'Valor ICMS Substituição Tributária_nota': 'icms_st',\
                              'Valor Base Cálculo ICMS ST Retido Operação Anterior_nota': 'bc_icms_st_ant',\
                              'Descrição Produto_nota': 'descricao_efd',\
                              'Unidade Comercial_nota': 'cod_unidade_efd'})
entradas_oppr['CFOP'] = entradas_oppr['CFOP'].astype(int)
# #####################################################################################################
# FIM (Entradas de Operação Própria) ##################################################################
# #####################################################################################################


#########################################################################################
# Lê o arquivo da EFD gerado pelo BO ####################################################
#########################################################################################
EFD_170_CSV = True
os.chdir(caminho_pasta)
# Supondo que você já tenha a lista de nomes de arquivos (fnamelist_2)
fnamelist_2 = [file for file in os.listdir() if 'efd_170' in file]
print("Arquivos encontrados - efd170:")
for arquivo in fnamelist_2:
    print(arquivo)
efdE_170sNAsdupl_df = util_ressarcimento.le_efd_c170(fnamelist_2, caminho_pasta,
                                                     EFD_170_CSV)
# ###########################################################################################
# #### FIM (Leitura do arquivo da EFD gerado pelo BO) #######################################
# ###########################################################################################


# ###########################################################################################
# Merge do c170 com a Tabela_1 ################################################################
# ###########################################################################################
efdE_170sNAsdupl_df.columns = [col + '_efd' for col in efdE_170sNAsdupl_df.columns]
entradas_e_c170 = pd.merge(tabela_1, efdE_170sNAsdupl_df,
                 left_on=['Chave Acesso NFe_nota', 'Número Item_nota'],
                 right_on=['chave_doc_efd_efd', 'item_num_efd'],
                 how='left', indicator='entrada_e_efd')
mapeamento = {
    'left_only': 'so_nota',
    'right_only': 'so_efd',
    'both': 'notas_e_efd'
}
# Aplicando o mapeamento na coluna desejada
entradas_e_c170['entrada_e_efd'] = entradas_e_c170['entrada_e_efd'].map(mapeamento)
entradas_e_c170 = entradas_e_c170.loc[entradas_e_c170['entrada_e_efd']=='notas_e_efd']
# entradas_e_c170 = \
# entradas_e_c170.loc[entradas_e_c170['Código CFOP (04 Posições)_nota'].isin(lista_cfops_st)]
tmp_entradas_e_c170 = \
entradas_e_c170.loc[entradas_e_c170['Código CFOP (04 Posições)_nota'].isin(lista_cfops_st)]
codigos_sem_repeticoes = tmp_entradas_e_c170['item_cod_efd_efd'].dropna().unique().tolist()
entradas_e_c170 = \
entradas_e_c170.loc[entradas_e_c170['item_cod_efd_efd'].isin(codigos_sem_repeticoes)]
# ###########################################################################################
# FIM(Merge do c170 com a Tabela_1) #########################################################
# ###########################################################################################


# ###############################################################################################
# Descobrindo os fatores de conversão ###########################################################
# ###############################################################################################
entradas_e_c170['razao'] = entradas_e_c170['qtd_efd'].astype(float)\
/ entradas_e_c170['Quantidade Comercial_nota'].astype(float)
mapeamento = entradas_e_c170[['Código Produto ou Serviço_nota', 'Unidade Comercial_nota', 'item_cod_efd_efd', 'razao']]
entradas_e_c170 = entradas_e_c170.drop('razao', axis=1)
mapeamento_distinto = mapeamento.drop_duplicates()
grupo_chave = ['Código Produto ou Serviço_nota', 'Unidade Comercial_nota', 'item_cod_efd_efd']
# Para cada grupo, pegar a linha com o maior valor de 'razao'
resultado = (
    mapeamento_distinto
    .sort_values('razao')  # organiza para que o menor venha primeiro
    .drop_duplicates(subset=grupo_chave, keep='first')  # PEGA O MAIOR VALOR DA RAZÃO
)
print(entradas_e_c170.shape[0])
print(resultado.shape[0])
entradas_e_c170 = pd.merge(entradas_e_c170, resultado,
                 on=['Código Produto ou Serviço_nota', 'item_cod_efd_efd'],
                 how='left', indicator='entrada_e_razao')
print(entradas_e_c170.shape[0])
print(entradas_e_c170['entrada_e_razao'].value_counts())

entradas_e_c170.loc[entradas_e_c170['razao']<1, 'razao'] = 1
entradas_e_c170['razao'] = np.floor(entradas_e_c170['razao'] * fator_conversao)

# entradas_e_c170['Quantidade Comercial_nota'] = entradas_e_c170['Quantidade Comercial_nota'] * \
# entradas_e_c170['razao']
# ###############################################################################################
# FIM (Descobrindo os fatores de conversão) #####################################################
# ###############################################################################################


# ###############################################################################################
# Manipulando o entradas_e_c170 #################################################################
# ###############################################################################################
teste = entradas_e_c170
entradas_e_c170 = entradas_e_c170[['Chave Acesso NFe_nota', 'mes_ref_efd', 'data_e_s_efd', 'Número Item_nota',\
                                   'IND_OPER_nota', 'item_cod_efd_efd', 'cfop_efd', 'qtd_efd',\
                                   'Valor Produto ou Serviço_nota',\
                                   'Valor ICMS Operação_nota', 'Valor ICMS Substituição Tributária_nota',\
                                   'Valor Base Cálculo ICMS ST Retido Operação Anterior_nota',\
                                   'descricao_efd', 'Número CNPJ Emitente_nota', 'Código GTIN_nota',\
                                   'Código NCM_nota', 'Código CEST_nota', 'cod_unidade_efd']]
entradas_e_c170['cfop_efd'] = entradas_e_c170['cfop_efd'].astype(int).astype(str)
entradas_e_c170 = \
entradas_e_c170.loc[entradas_e_c170['cfop_efd'].isin(lista_cfops_st_teste)]
entradas_e_c170 = \
entradas_e_c170.rename(columns={'Chave Acesso NFe_nota':'CHV_DOC',\
                                'data_e_s_efd':'DATA',\
                                'Número Item_nota':'NUM_ITEM',\
                                'IND_OPER_nota':'IND_OPER',
                                'item_cod_efd_efd':'COD_ITEM',\
                                'cfop_efd': 'CFOP',\
                                'qtd_efd': 'QTD',\
                                'Valor Produto ou Serviço_nota': 'vlr_prod',\
                                'Valor ICMS Operação_nota': 'icms',\
                                'Valor ICMS Substituição Tributária_nota': 'icms_st',\
                                'Valor Base Cálculo ICMS ST Retido Operação Anterior_nota': 'bc_icms_st_ant'})
entradas_e_c170['DATA'] = pd.to_datetime(entradas_e_c170['DATA'],dayfirst=True)
#entradas_e_c170['CFOP'] = entradas_e_c170['CFOP'].astype(int)
# ###############################################################################################
# FIM (Manipulando o entradas_e_c170) ###########################################################
# ###############################################################################################


# ###############################################################################################
# Gera um DataFrame com todas as notas de entrada (inclusive as de operação própria) ############
# ###############################################################################################
# Todas as entradas
entradas = pd.concat([entradas_e_c170, entradas_oppr], ignore_index=True)
# Gera a referência para as notas de emissão própria, uma vez que elas não são escrituradas no EFD170
entradas['mes_ref_efd'] = entradas['mes_ref_efd'].fillna(0).astype(int).astype(str)
entradas['mes_aaaamm'] = entradas['DATA'].dt.strftime('%Y%m')
entradas.loc[entradas['mes_ref_efd'] == '0', 'mes_ref_efd'] = entradas['mes_aaaamm']
entradas.drop(columns='mes_aaaamm', inplace=True)

# Aplicar strip em algumas colunas
entradas['mes_ref_efd'] = entradas['mes_ref_efd'].astype(str).str.strip()
entradas['COD_ITEM'] = entradas['COD_ITEM'].astype(str).str.strip()
entradas['cod_unidade_efd'] = entradas['cod_unidade_efd'].astype(str).str.strip()
# ###############################################################################################
# FIM (Gera um DataFrame com todas as notas de entrada (inclusive as de operação própria)) ######
# ###############################################################################################

# ###############################################################################################
# Lê o EFD0220 ##################################################################################
# ###############################################################################################
entradas['fat_0220'] = 1
nome_efd0220 = f"{caminho_pasta}/efd0220_completo.xlsx"
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
entradas['QTD'] = entradas['QTD'] * entradas['fat_0220']
entradas['QTD'] = entradas['QTD'].apply(dividir_com_preservacao)
entradas = entradas.drop(columns=['fat_0220'])
# ###############################################################################################
# FIM (Lê o EFD0220) ############################################################################
# ###############################################################################################

# ###############################################################################################
# ################## Obtendo os dados do registro 0000 ##########################################
# ###############################################################################################
CNPJ_ENTIDADE = tabela_1.loc[tabela_1['Tipo_nota']=='saida_nfe','Número CNPJ Emitente_nota'].iloc[0]
NOME_ENTIDADE = tab_cnpjs.loc[tab_cnpjs['cnpj']==CNPJ_ENTIDADE, 'nome'].iloc[0]
IE_ENTIDADE = tab_cnpjs.loc[tab_cnpjs['cnpj']==CNPJ_ENTIDADE, 'ie'].iloc[0]
COD_MUN_ENTIDADE = tab_cnpjs.loc[tab_cnpjs['cnpj']==CNPJ_ENTIDADE, 'codmun'].iloc[0]
# Esse código é fixo
COD_VER = "01"
# 00 - Remessa Regular; 02 - Remessa de arquivo para substituição de arquivo remetido anteriormente
# 01 - Remessa de arquivo requerido por intimação específica
COD_FIN = "00"
# ###############################################################################################
# ################## FIM (Obtendo os dados do registro 0000) ####################################
# ###############################################################################################


# ############################################################################################################# #
# Finalização da geração dos arquivos auxiliares (Registro 0150) ############################################## #
# ############################################################################################################# #
print("Finalizando geração dos arquivos auxiliares...")
nfe_auxs_oppr = entradas_oppr[['CHV_DOC']].drop_duplicates(subset='CHV_DOC').copy()
nfe_auxs_oppr = nfe_auxs_oppr.rename(columns={'CHV_DOC': 'Chave Acesso NFe'})
nfe_auxs_oppr['Número CNPJ Emitente (char)'] = EFD_Reg_0000['Número CNPJ'].iloc[0]
nfe_auxs_oppr['Nome Razão Social Emitente'] = EFD_Reg_0000['Nome Entidade'].iloc[0]
nfe_auxs_oppr['Número Inscrição Estadual Completa Emitente'] = EFD_Reg_0000['Número Inscrição Estadual'].iloc[0]
nfe_auxs_oppr['Código Município Fato Gerador'] = EFD_Reg_0000['Código Município'].iloc[0]
nfe_auxs_oppr['Código País Emitente'] = '1058'

nfe_auxs = pd.concat([nfe_auxs, nfe_auxs_oppr], ignore_index=True)
nfe_auxs = nfe_auxs.drop_duplicates(subset='Chave Acesso NFe')
# ############################################################################################################# #
# FIM (Finalização da geração dos arquivos auxiliares (Registro 0150)) ######################################## #
# ############################################################################################################# #

# ############################################################################################################# #
# 'COD_ITEM' (entradas) com unidades distintas ################################################################ #
# ############################################################################################################# #
print("Gerando relatório dos códigos de item que têm unidades diferentes nas entradas...")
# Passo 1: Contar unidades distintas por mes_ref_efd e COD_ITEM
unidades_por_item_mes = entradas.groupby(['mes_ref_efd', 'COD_ITEM'])['cod_unidade_efd'].nunique().reset_index(name='qtd_unidades')

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
# ############################################################################################################# #
# FIM ('COD_ITEM' (entradas) com unidades distintas) ########################################################## #
# ############################################################################################################# #

# ############################################################################################################# #
# Análise da Razão ############################################################################################ #
# ############################################################################################################# #
# IMPLEMENTAR AQUI
# ############################################################################################################# #
# FIM (Análise da Razão) ###################################################################################### #
# ############################################################################################################# #


# ############################################################################################################# #
# Montagem das Entradas ####################################################################################### #
# ############################################################################################################# #
print("Montando as entradas...")
entradas = pd.merge(entradas, nfe_sp_aliquotas,
                 on=['COD_ITEM'],
                 how='left', indicator='entrada_e_aliquotas')
entradas['aliq'] = entradas['aliq'].fillna(18)

# Calculando a primeira opção: 'icms' + 'icms_st'
primeira_opcao = entradas['icms'] + entradas['icms_st']
segunda_opcao = entradas['bc_icms_st_ant'] * entradas['aliq']/100
# Definindo o 'icms_sup_devido' como o maior valor entre as duas opções calculadas
entradas['ICMS_TOT'] = primeira_opcao.combine(segunda_opcao, max)

entradas['VL_CONFR'] = 0
entradas['COD_LEGAL'] = ""
entradas['CFOP'] = entradas['CFOP'].astype(str)
# As devoluções de saída têm que ter código legal preenchidas
entradas.loc[entradas['CFOP'].isin(dev_saida),'COD_LEGAL'] = "1"
entradas.loc[entradas['CFOP'].isin(dev_saida),'VL_CONFR'] = entradas['vlr_prod'] * entradas['aliq'] / 100

entradas.loc[(entradas['CFOP'].isin(dev_saida)) & (entradas['VL_CONFR']==0),'COD_LEGAL'] = "0"

entradas['SUBTIPO'] = 0
entradas.loc[entradas['CFOP'].isin(dev_saida),'SUBTIPO'] = 3
# ############################################################################################################# #
# FIM (Montagem das Entradas) ################################################################################# #
# ############################################################################################################# #


# ############################################################################################################# #
# Montagem das Saídas ######################################################################################### #
# ############################################################################################################# #
print("Montando as saídas...")
saidas = tabela_1.loc[tabela_1['IND_OPER_nota']==1]
saidas = saidas.loc[saidas['Código CFOP (04 Posições)_nota'].isin(lista_cfops_st_teste)]
saidas = saidas.rename(columns={'Chave Acesso NFe_nota':'CHV_DOC',\
                                'Data Emissão_nota':'DATA',\
                                'Número Item_nota':'NUM_ITEM',\
                                'IND_OPER_nota':'IND_OPER',\
                                'Código Produto ou Serviço_nota':'COD_ITEM',\
                                'Código CFOP (04 Posições)_nota': 'CFOP',\
                                'Quantidade Comercial_nota': 'QTD',\
                                'Valor Produto ou Serviço_nota': 'vlr_prod',\
                                'Valor ICMS Operação_nota': 'icms',\
                                'Valor ICMS Substituição Tributária_nota': 'icms_st',\
                                'Valor Base Cálculo ICMS ST Retido Operação Anterior_nota': 'bc_icms_st_ant',\
                                'Descrição Produto_nota': 'descricao_efd',\
                                'Unidade Comercial_nota': 'cod_unidade_efd'
                                })
saidas['mes_ref_efd'] = saidas['DATA'].dt.strftime('%Y%m')

aliqs_saidas = entradas[['mes_ref_efd', 'COD_ITEM', 'aliq']].drop_duplicates()
saidas = pd.merge(saidas, aliqs_saidas,
                 on=['mes_ref_efd', 'COD_ITEM'],
                 how='left', indicator='saida_e_0200')
saidas = saidas.loc[saidas['saida_e_0200']=='both']
saidas = saidas[saidas['QTD'] != 0]
saidas['VL_CONFR'] = saidas['vlr_prod'] * saidas['aliq']/100
saidas['ICMS_TOT'] = 0
saidas['COD_LEGAL'] = "1"

saidas.loc[saidas['VL_CONFR']<0.01,'VL_CONFR'] = 0
saidas.loc[saidas['VL_CONFR']==0,'COD_LEGAL'] = "0"

saidas['SUBTIPO'] = 2
saidas.loc[saidas['CFOP'].isin(dev_entrada),'SUBTIPO'] = 1
saidas.loc[saidas['CFOP'].isin(dev_entrada),'COD_LEGAL'] = ""
saidas.loc[saidas['CFOP'].isin(dev_entrada),'VL_CONFR'] = 0
primeira_op = saidas['icms'] + saidas['icms_st']
segunda_op = saidas['bc_icms_st_ant'] * saidas['aliq']/100
saidas.loc[saidas['CFOP'].isin(dev_entrada),'ICMS_TOT'] = primeira_op.combine(segunda_op, max)
# ############################################################################################################# #
# FIM (Montagem das Saídas) ################################################################################### #
# ############################################################################################################# #


# ############################################################################################################# #
# Montagem dos dados (União das entradas e das saídas)######################################################### #
# ############################################################################################################# #
print("Concatenando entradas e saídas...")
# 1. Unir os dois DataFrames mantendo apenas as colunas relevantes
dados = pd.concat([
    entradas[['COD_ITEM', 'DATA', 'CHV_DOC', 'NUM_ITEM', 'CFOP', 'QTD', 'IND_OPER', 'ICMS_TOT',\
             'VL_CONFR', 'COD_LEGAL', 'SUBTIPO', 'Número CNPJ Emitente_nota', 'Código GTIN_nota',\
             'Código NCM_nota', 'Código CEST_nota', 'descricao_efd', 'cod_unidade_efd', 'aliq']],
    saidas[['COD_ITEM', 'DATA', 'CHV_DOC', 'NUM_ITEM', 'CFOP', 'QTD', 'IND_OPER', 'ICMS_TOT',\
           'VL_CONFR', 'COD_LEGAL', 'SUBTIPO', 'Número CNPJ Emitente_nota', 'Código GTIN_nota',\
           'Código NCM_nota', 'Código CEST_nota', 'descricao_efd', 'cod_unidade_efd', 'aliq']]
], ignore_index=True)

# 2. Criar duas colunas separadas para entrada e saída
dados['QTD_ENTRADA'] = dados.apply(lambda x: x['QTD'] if x['IND_OPER'] == 0 else 0, axis=1)
dados['QTD_SAIDA'] = dados.apply(lambda x: x['QTD'] if x['IND_OPER'] == 1 else 0, axis=1)

# Filtra apenas as entradas (IND_OPER == 0)
filtro = dados['IND_OPER'] == 0
dados_filtrados = dados.loc[filtro]

# Separar as linhas com ICMS_TOT > 0
com_icms = dados_filtrados[dados_filtrados['ICMS_TOT'] != 0]

# Agrupamento considerando apenas ICMS_TOT > 0
agrupado_com_icms = com_icms.groupby('COD_ITEM').agg({
    'ICMS_TOT': 'sum',
    'QTD': 'sum'
})
agrupado_com_icms['ICMS_TOT_UNI'] = (agrupado_com_icms['ICMS_TOT'] / agrupado_com_icms['QTD']) * fator_icms_tot

# Criar dataframe com todos os COD_ITEM únicos de entradas
cod_items_unicos = dados_filtrados[['COD_ITEM']].drop_duplicates().set_index('COD_ITEM')

# Junta com os resultados calculados, preenchendo ICMS_TOT_UNI=0 onde não havia ICMS_TOT > 0
agrupado_final = cod_items_unicos.join(agrupado_com_icms['ICMS_TOT_UNI'], how='left').fillna(0)

# Teste - relatório dos valores de icms_tot_uni
nome_rel_icms_tot_uni = f"{nome_pasta}/Relatorio-{CNPJ_ENTIDADE}_icms_tot_uni.xlsx"
agrupado_final.to_excel(nome_rel_icms_tot_uni, index=False)

# Junta com o dataframe original
dados = dados.merge(
    agrupado_final,
    how='left',
    left_on='COD_ITEM',
    right_index=True
)
# ############################################################################################################# #
# FIM (Montagem dos dados (União das entradas e das saídas)) ################################################## #
# ############################################################################################################# #


# ############################################################################################################# #
# Cálculo do estoque inicial ################################################################################## #
# ############################################################################################################# #
print("Calculando o estoque inicial...")
est_ini = calcular_estoque_inicial_por_produto(dados)
est_ini = est_ini.sort_values(by=['DATA', 'SUBTIPO'])
est_ini['ICMS_TOT_INI'] = est_ini['estoque_inicial_minimo'] * est_ini['ICMS_TOT_UNI']
est_ini['ano_mes'] = est_ini['DATA'].dt.to_period('M')

print("Gerando os estoques iniciais, para confronto com o Bloco H...")
est_ini_nov2019 = est_ini[est_ini['ano_mes'].astype(str) == '2019-11'].copy()
est_ini_nov2019 = est_ini_nov2019[['ano_mes', 'COD_ITEM', 'estoque_inicial_minimo']].drop_duplicates()
nome_est_nov19 = f"{nome_pasta}/estoques_iniciais-{CNPJ_ENTIDADE}_31_10_2019.xlsx"
est_ini_nov2019.to_excel(nome_est_nov19, index=False)

# est = est_ini[['COD_ITEM', 'ICMS_TOT_INI']]
# est = est.drop_duplicates()
# est = est[est['ICMS_TOT_INI']!=0].sort_values(by='ICMS_TOT_INI', ascending=True)
#
# # 2. Calcula o percentil (menores valores = percentis menores)
# est['percentil'] = est['ICMS_TOT_INI'].rank(pct=True)
#
# # 3. Define o multiplicador decrescente (valores menores recebem mais aumento)
# # Exemplo: de 1.5 (valores menores) até 1.0 (valores maiores)
# est['multiplicador'] = 1.5 - 0.5 * est['percentil']
#
# # 4. Aplica o multiplicador no est_ini
# # Junta os multiplicadores de volta ao est_ini com base em COD_ITEM
# est_ini = est_ini.merge(est[['COD_ITEM', 'multiplicador']], on='COD_ITEM', how='left')
#
#
# est_ini['multiplicador'] = est_ini['multiplicador'].fillna(1)
# #original['ICMS_TOT_INI'] = est_ini['ICMS_TOT_INI'].copy()
# est_ini['ICMS_TOT_INI'] = est_ini['ICMS_TOT_INI'] * est_ini['multiplicador']

# ############################################################################################################# #
# FIM (Cálculo do estoque inicial) ############################################################################ #
# ############################################################################################################# #


# ############################################################################################################# #
# Montagem e geração dos arquivos da PCAT 42/2018 ############################################################# #
# ############################################################################################################# #
print("Iniciando Montagem e geração dos arquivos da PCAT...")
# Cria coluna auxiliar com o ano e mês
est_ini['ano_mes'] = est_ini['DATA'].dt.to_period('M')
cods_ress = ['1', '2', '3', '4']
# Garante que não haja duplicações por COD_ITEM no mesmo mês
grupo_mes = est_ini.groupby(['ano_mes', 'COD_ITEM']).agg({
    'estoque_inicial_minimo': 'first',
    'ICMS_TOT_INI': 'first'
}).reset_index()

# Para cada mês, gerar um arquivo .txt
reg_1050_ant = pd.DataFrame(columns=['ref', 'itemcode', 'iq', 'iv', 'fq', 'fv'])
for periodo, grupo in grupo_mes.groupby('ano_mes'):
    nome_arquivo = f"{nome_pasta}/RessarcimentoST-{CNPJ_ENTIDADE}_{periodo.strftime('%Y-%m')}.txt"
    #nome_arquivo = f"C:/temp_calculo/RessarcimentoST-{CNPJ_ENTIDADE}_{periodo.strftime('%Y-%m')}.txt"

    periodo_0000 = periodo.strftime('%m%Y')
    linha_0000 = f"0000|{periodo_0000}|{NOME_ENTIDADE}|{CNPJ_ENTIDADE}|{IE_ENTIDADE}|{COD_MUN_ENTIDADE}|{COD_VER}|{COD_FIN}\n"
    linhas_0150 = []
    linhas_0200 = []
    linhas_1050 = []
    linhas_1100 = []

    # Filtra o est_ini correspondente ao mês atual
    est_ini_mes = est_ini[est_ini['ano_mes'] == periodo]
    reg0150 = gera_0150(nfe_auxs, est_ini_mes)
    reg0200 = gera_0200(est_ini_mes)
    reg1050 = gera_1050(est_ini_mes, reg_1050_ant)

    reg_1050_ant = pd.merge(reg1050, reg_1050_ant, on='itemcode', how='outer', indicator='total')
    reg_1050_ant.loc[reg_1050_ant['total'] == 'right_only', ['iq_x', 'fq_x', 'iv_x', 'fv_x']] = \
        reg_1050_ant.loc[reg_1050_ant['total'] == 'right_only', ['iq_y', 'fq_y', 'iv_y', 'fv_y']].values
    reg_1050_ant = reg_1050_ant.drop(columns=['iq_y', 'fq_y', 'iv_y', 'fv_y', 'total'])
    reg_1050_ant = reg_1050_ant.rename(columns={'iq_x': 'iq', 'fq_x':'fq' , 'iv_x': 'iv', 'fv_x': 'fv'})


    # raise SystemExit("Parando aqui para DEBUG!")

    # Percorre cada linha do reg0150 e escreve no arquivo
    for _, row in reg0150.iterrows():
        linha_0150 = f"{row['REG']}|{row['COD_PART']}|{row['NOME_0150']}|{row['COD_PAIS']}|{row['CNPJ_0150']}|{row['CPF']}|{row['IE_0150']}|{row['COD_MUN_0150']}\n"
        linhas_0150.append(linha_0150)

    # Percorre cada linha do reg0200 e escreve no arquivo
    for _, row in reg0200.iterrows():
        linha_0200 = f"{row['REG']}|{row['COD_ITEM']}|{row['DESCR_ITEM']}|{row['COD_BARRA']}|{row['UNID_INV']}|{row['COD_NCM']}|{row['ALIQ_ICMS']}|{row['CEST']}\n"
        linhas_0200.append(linha_0200)

    # Primeiro cria as linhas 1050
    # for _, row in grupo.iterrows():
    #     linha_1050 = f"1050|{row['COD_ITEM']}|{row['estoque_inicial_minimo']:.3f}|{row['ICMS_TOT_INI']:.2f}|{row['estoque_inicial_minimo']:.3f}|{row['ICMS_TOT_INI']:.2f}\n"
    #     linha_1050 = linha_1050.replace(".", ",")
    #     linhas_1050.append(linha_1050)
    for _, row in reg1050.iterrows():
        linha_1050 = f"1050|{row['itemcode']}|{row['iq']:.3f}|{row['iv']:.2f}|{row['fq']:.3f}|{row['fv']:.2f}\n"
        linha_1050 = linha_1050.replace(".", ",")
        linhas_1050.append(linha_1050)

    # Depois cria as linhas 1100
    for _, detalhe in est_ini_mes.iterrows():
        if (detalhe['IND_OPER'] == 0) and (detalhe['CFOP'] not in dev_saida):
            linha_1100 = (
                f"1100|{detalhe['CHV_DOC']}|{detalhe['DATA'].strftime('%d%m%Y')}|{detalhe['NUM_ITEM']}|"
                f"{detalhe['IND_OPER']}|{detalhe['COD_ITEM']}|{detalhe['CFOP']}|"
                f"{detalhe['QTD']:.3f}|{detalhe['ICMS_TOT']:.2f}||\n"
            )
        elif (detalhe['IND_OPER'] == 1) and (detalhe['CFOP'] in dev_entrada):
            linha_1100 = (
                f"1100|{detalhe['CHV_DOC']}|{detalhe['DATA'].strftime('%d%m%Y')}|{detalhe['NUM_ITEM']}|"
                f"{detalhe['IND_OPER']}|{detalhe['COD_ITEM']}|{detalhe['CFOP']}|"
                f"{detalhe['QTD']:.3f}|{detalhe['ICMS_TOT']:.2f}||\n"
            )
        elif (detalhe['IND_OPER'] == 1) and (detalhe['CFOP'] not in dev_entrada) and \
                (detalhe['COD_LEGAL'] in cods_ress):
            linha_1100 = (
                f"1100|{detalhe['CHV_DOC']}|{detalhe['DATA'].strftime('%d%m%Y')}|{detalhe['NUM_ITEM']}|"
                f"{detalhe['IND_OPER']}|{detalhe['COD_ITEM']}|{detalhe['CFOP']}|"
                f"{detalhe['QTD']:.3f}||{detalhe['VL_CONFR']:.2f}|{detalhe['COD_LEGAL']}\n"
            )
        elif (detalhe['IND_OPER'] == 1) and (detalhe['CFOP'] not in dev_entrada) and \
                (detalhe['COD_LEGAL'] not in cods_ress):
            linha_1100 = (
                f"1100|{detalhe['CHV_DOC']}|{detalhe['DATA'].strftime('%d%m%Y')}|{detalhe['NUM_ITEM']}|"
                f"{detalhe['IND_OPER']}|{detalhe['COD_ITEM']}|{detalhe['CFOP']}|"
                f"{detalhe['QTD']:.3f}|||{detalhe['COD_LEGAL']}\n"
            )
        elif (detalhe['IND_OPER'] == 0) and (detalhe['CFOP'] in dev_saida) and \
                (detalhe['COD_LEGAL'] in cods_ress):
            linha_1100 = (
                f"1100|{detalhe['CHV_DOC']}|{detalhe['DATA'].strftime('%d%m%Y')}|{detalhe['NUM_ITEM']}|"
                f"{detalhe['IND_OPER']}|{detalhe['COD_ITEM']}|{detalhe['CFOP']}|"
                f"{detalhe['QTD']:.3f}|{detalhe['ICMS_TOT']:.2f}|{detalhe['VL_CONFR']:.2f}|{detalhe['COD_LEGAL']}\n"
            )
        elif (detalhe['IND_OPER'] == 0) and (detalhe['CFOP'] in dev_saida) and \
                (detalhe['COD_LEGAL'] not in cods_ress):
            linha_1100 = (
                f"1100|{detalhe['CHV_DOC']}|{detalhe['DATA'].strftime('%d%m%Y')}|{detalhe['NUM_ITEM']}|"
                f"{detalhe['IND_OPER']}|{detalhe['COD_ITEM']}|{detalhe['CFOP']}|"
                f"{detalhe['QTD']:.3f}|{detalhe['ICMS_TOT']:.2f}||{detalhe['COD_LEGAL']}\n"
            )
        else:
            raise SystemExit("Erro!!!!!!!!!!!!!!!!")

        linha_1100 = linha_1100.replace(".", ",")
        linhas_1100.append(linha_1100)

    # Escreve o arquivo: primeiro 1050, depois 1100
    with open(nome_arquivo, 'w', encoding='utf-8') as f:
        f.writelines([linha_0000] + linhas_0150 + linhas_0200 + linhas_1050 + linhas_1100)

    print(f"Arquivo gerado: {nome_arquivo}")
# ############################################################################################################# #
# FIM (Montagem e geração dos arquivos da PCAT 42/2018) ####################################################### #
# ############################################################################################################# #