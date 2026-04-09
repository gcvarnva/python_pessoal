import sys
from sys import exception

import util_geracao

pasta_bibliotecas = r"Z:\Dados\Projetos_python\bibliotecas"
sys.path.append(str(pasta_bibliotecas))
import util_gui
import pandas as pd
import numpy as np
import os
import util_ressarcimento
from datetime import datetime
import matplotlib.pyplot as plt
from util_geracao import *
from pathlib import Path

GERAR_RELATORIOS = False
LE_FATORES = True
CONVERTE_EFD0220 = False


dir_inicial = r"Z:"
caminho_pasta = util_gui.selecionar_pasta(dir_inicial)
data_hora_atual = datetime.now().strftime('%d-%m-%Y_%Hh%Mm%Ss')
nome_pasta = f"{caminho_pasta}/pcat_{data_hora_atual}_ficms_{fator_icms_tot}_fconv_{fator_conversao}_div_{divisor}"
os.makedirs(nome_pasta, exist_ok=True)
set_nome_pasta(nome_pasta)

os.chdir(caminho_pasta)

caminho = Path(caminho_pasta)
tem_excel_fat_conv = (Path(caminho_pasta) / "fat_conv_entradas.xlsx").exists()
tem_excel_fat_conv_saidas = (Path(caminho_pasta) / "conversao_unidades_saidas.xlsx").exists()

#Carregando bloco H do contribuinte (início do período)
blocoh = carrega_blocoh()
cod_itens_blocoh = blocoh['COD_ITEM'].str.strip().dropna().unique().tolist()
# Leitura do Registro 0000 da EFD
caminho_completo = caminho_pasta + r'\EFD_Reg_0000.csv'
EFD_Reg_0000 = pd.read_csv(caminho_completo, sep=';', encoding='utf-8', quotechar='"', skiprows=3, dtype=str,\
                           usecols=['Nome Entidade', 'Mês Referência', 'Número CNPJ',\
                                    'Número Inscrição Estadual' ,'Código Município'])

# Lê o arquivo Excel com a descrição dos CFOPs
pasta_atual = Path(__file__).parent  # pega a pasta onde o script está
caminho_excel = pasta_atual / "CFOPs Sumarizados 2021_06_10.xlsx"
desc_cfop, sugest_cfop = util_ressarcimento.le_desc_CFOPs(caminho_excel)

# ############################################################################################################# #
# Leitura dos arquivos auxiliares da geração (registro 0150) ################################################## #
# ############################################################################################################# #
caminho_completo = caminho_pasta + r"\nfe_entrada_outras-p_geracao.csv"
nfe_ufs_auxs = pd.read_csv(caminho_completo, sep=';', encoding='utf-8', quotechar='"', skiprows=3, dtype=str,\
                              usecols=['Chave Acesso NFe', 'Número CNPJ Emitente (char)',\
                                       'Nome Razão Social Emitente',\
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
                              usecols=['Código Produto ou Serviço',\
                                       'Percentual Alíquota ICMS'])
nfe_sp_aliquotas = nfe_sp_aliquotas\
    .rename(columns={'Código Produto ou Serviço': 'COD_ITEM',\
                     'Percentual Alíquota ICMS': 'aliq'})

# Converte a coluna 'aliq' para float (se necessário)
nfe_sp_aliquotas['aliq'] = nfe_sp_aliquotas['aliq'].str.replace(',', '.').astype(float)

# Seleciona apenas o maior valor de 'aliq' por 'COD_ITEM'
nfe_sp_aliquotas = (
    nfe_sp_aliquotas.groupby('COD_ITEM', as_index=False)['aliq']
    .max()
)
nfe_sp_aliquotas['aliq'] = nfe_sp_aliquotas['aliq'].fillna(18)

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
print("Lendo tabela das notas de entrada e saída...", end="", flush=True)
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
tabela_1 = tabela_1.drop_duplicates(subset=['Chave Acesso NFe_nota', 'Número Item_nota'])
print("leitura concluída!")
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
print("Merge do C170 com a tabela 1...", end="", flush=True)
efdE_170sNAsdupl_df.columns = [col + '_efd' for col in efdE_170sNAsdupl_df.columns]
efdE_170sNAsdupl_df = efdE_170sNAsdupl_df.rename(\
    columns={'chave_doc_efd_efd': 'chave', 'item_num_efd': 'numero_item'})
tabela_1 = tabela_1.rename(\
    columns={'Chave Acesso NFe_nota': 'chave', 'Número Item_nota': 'numero_item'})
entradas_e_c170 = pd.merge(tabela_1,efdE_170sNAsdupl_df,how='inner',on=['chave', 'numero_item'],\
                    indicator=True, validate="one_to_one" )

entradas_e_c170_com_st = entradas_e_c170.loc[\
    (entradas_e_c170['Valor Base Cálculo ICMS ST Retido Operação Anterior_nota']>0)\
        | (entradas_e_c170['Valor ICMS Substituição Tributária_nota']>0)]

codigos_sem_repeticoes = entradas_e_c170_com_st['item_cod_efd_efd'].dropna().unique().tolist()
entradas_e_c170 = \
entradas_e_c170.loc[entradas_e_c170['item_cod_efd_efd'].isin(codigos_sem_repeticoes)]
# entradas_e_c170 = \
# entradas_e_c170.loc[entradas_e_c170['item_cod_efd_efd'].isin(cod_itens_blocoh)]
print("merge concluído!")
# ###########################################################################################
# FIM(Merge do c170 com a Tabela_1) #########################################################
# ###########################################################################################

# ###########################################################################################
# Tabela de CFOPs das entradas ##############################################################
# ###########################################################################################
if GERAR_RELATORIOS:
    util_geracao.relatorio_cfops_entrada(entradas_e_c170, nome_pasta)
# ###########################################################################################
# FIM (Tabela de CFOPs das entradas) ########################################################
# ###########################################################################################

# ###########################################################################################
# Tabela de ligações das entradas ###########################################################
# ###########################################################################################
if GERAR_RELATORIOS:
    util_geracao.relatorio_ligacoes_entradas(entradas_e_c170, nome_pasta)
# ###############################################################################################
# Uniformizando os fatores de conversão ###########################################################
# ###############################################################################################
if not LE_FATORES:
    if UNIFORMIZA_FATORES:
        util_geracao.uniformiza_fatores(entradas_e_c170, nome_pasta)
# ###############################################################################################
# Manipulando o entradas_e_c170 #################################################################
# ###############################################################################################
entradas_e_c170 = entradas_e_c170[['chave', 'mes_ref_efd', 'data_e_s_efd', 'numero_item',\
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
entradas_e_c170.rename(columns={'chave':'CHV_DOC',\
                                'data_e_s_efd':'DATA',\
                                'numero_item':'NUM_ITEM',\
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
print("Finalizando dataframe das notas de entrada (inclusive as de operação própria)...", end="", flush=True)
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

entradas = entradas.loc[entradas['COD_ITEM'].isin(codigos_sem_repeticoes)]
#entradas = entradas.loc[entradas['COD_ITEM'].isin(cod_itens_blocoh)]
print("finalizado!")
# ###############################################################################################
# FIM (Gera um DataFrame com todas as notas de entrada (inclusive as de operação própria)) ######
# ###############################################################################################

# ###############################################################################################
# Lê o EFD0220 ##################################################################################
# ###############################################################################################
if not LE_FATORES:
    if CONVERTE_EFD0220:
        util_geracao.converte_efd0220(entradas, nome_pasta)

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
print("Finalizando geração dos arquivos auxiliares... ", end="", flush=True)
nfe_auxs_oppr = entradas_oppr[['CHV_DOC']].drop_duplicates(subset='CHV_DOC').copy()
nfe_auxs_oppr = nfe_auxs_oppr.rename(columns={'CHV_DOC': 'Chave Acesso NFe'})
nfe_auxs_oppr['Número CNPJ Emitente (char)'] = EFD_Reg_0000['Número CNPJ'].iloc[0]
nfe_auxs_oppr['Nome Razão Social Emitente'] = EFD_Reg_0000['Nome Entidade'].iloc[0]
nfe_auxs_oppr['Número Inscrição Estadual Completa Emitente'] = EFD_Reg_0000['Número Inscrição Estadual'].iloc[0]
nfe_auxs_oppr['Código Município Fato Gerador'] = EFD_Reg_0000['Código Município'].iloc[0]
nfe_auxs_oppr['Código País Emitente'] = '1058'

nfe_auxs = pd.concat([nfe_auxs, nfe_auxs_oppr], ignore_index=True)
nfe_auxs = nfe_auxs.drop_duplicates(subset='Chave Acesso NFe')
print("Concluído!")
# ############################################################################################################# #
# FIM (Finalização da geração dos arquivos auxiliares (Registro 0150)) ######################################## #
# ############################################################################################################# #

# 'COD_ITEM' (entradas) com unidades distintas
if GERAR_RELATORIOS:
    util_geracao.relatorio_coditens_unids_diversas(entradas, nome_pasta)

# ############################################################################################################# #
# Montagem das Entradas ####################################################################################### #
# ############################################################################################################# #
print("Finalizando a montagem das entradas...", end="", flush=True)
entradas = pd.merge(entradas, nfe_sp_aliquotas,
                 on=['COD_ITEM'],
                 how='left', indicator='entrada_e_aliquotas')
entradas['aliq'] = entradas['aliq'].fillna(18)
aliquotas_entradas = entradas.drop_duplicates(
    subset=['COD_ITEM', 'descricao_efd', 'Código NCM_nota', 'aliq', 'mes_ref_efd']
)[['COD_ITEM', 'descricao_efd', 'Código NCM_nota', 'aliq', 'mes_ref_efd']]
caminho_completo = nome_pasta + r'\ligacoes_aliquotas_entradas.xlsx'
aliquotas_entradas.to_excel(caminho_completo, index=False)


# Calculando a primeira opção: 'icms' + 'icms_st'
primeira_opcao = entradas['icms'] + entradas['icms_st']
segunda_opcao = entradas['bc_icms_st_ant'] * entradas['aliq']/100
# Definindo o 'icms_sup_devido' como o maior valor entre as duas opções calculadas
entradas['ICMS_TOT'] = primeira_opcao.combine(segunda_opcao, max)

entradas['VL_CONFR'] = 0.0
entradas['COD_LEGAL'] = ""
entradas['CFOP'] = entradas['CFOP'].astype(str)
# As devoluções de saída têm que ter código legal preenchidas
entradas.loc[entradas['CFOP'].isin(dev_saida),'COD_LEGAL'] = "1"
entradas.loc[entradas['CFOP'].isin(dev_saida),'VL_CONFR'] = entradas['vlr_prod'] * entradas['aliq'] / 100

entradas.loc[(entradas['CFOP'].isin(dev_saida)) & (entradas['VL_CONFR']==0),'COD_LEGAL'] = "0"

entradas['SUBTIPO'] = 0
entradas.loc[entradas['CFOP'].isin(dev_saida),'SUBTIPO'] = 3
print("montagem concluída!")
# ############################################################################################################# #
# FIM (Montagem das Entradas) ################################################################################# #
# ############################################################################################################# #

# ###########################################################################################
# ######## Relatório - ICMS Suportado de entrada ############################################
# ###########################################################################################
if LE_FATORES and (not tem_excel_fat_conv):
    util_geracao.relatorio_icms_suportado(entradas, nfe_sp_aliquotas, nome_pasta)
    print("Ajuste os fatores de conversão e rode novamente.")
    print("FIM DA EXECUÇÃO!.")
    sys.exit(0)
elif LE_FATORES:
    caminho_completo = caminho_pasta + r"\fat_conv_entradas.xlsx"
    df_fat_conv = pd.read_excel(caminho_completo)
    entradas = pd.merge(entradas, df_fat_conv, on=['COD_ITEM', 'cod_unidade_efd'], how='left',
                        validate='many_to_one', indicator=True)
    qtd_diferentes_both = entradas['_merge'].ne('both').sum()
    if qtd_diferentes_both >0:
        raise SystemExit("Erro de Merge nos fatores de conversão de entradas!")
    entradas = entradas.drop(columns='_merge')
    entradas['QTD'] = entradas['QTD'] * entradas['fat_conv']
# ###########################################################################################
# ######## Fim - Relatório - ICMS Suportado de entrada ######################################
# ###########################################################################################

# ############################################################################################################# #
# Montagem das Saídas ######################################################################################### #
# ############################################################################################################# #
print("Montando as saídas...", end="", flush=True)
saidas = tabela_1.loc[tabela_1['IND_OPER_nota']==1]
saidas = saidas.loc[saidas['Código CFOP (04 Posições)_nota'].isin(lista_cfops_st_teste)]
saidas = \
saidas.loc[saidas['Código Produto ou Serviço_nota'].isin(codigos_sem_repeticoes)]
# saidas = \
# saidas.loc[saidas['Código Produto ou Serviço_nota'].isin(cod_itens_blocoh)]
saidas = saidas.rename(columns={'chave':'CHV_DOC',\
                                'Data Emissão_nota':'DATA',\
                                'numero_item':'NUM_ITEM',\
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

saidas = pd.merge(saidas, nfe_sp_aliquotas,
                 on=['COD_ITEM'],
                 how='left', indicator='saida_e_aliquotas')
saidas['aliq'] = saidas['aliq'].fillna(18)
#saidas = saidas.loc[saidas['saida_e_0200']=='both']

# ###############################################################################################
# # Conversão de unidades das saídas ############################################################
# ###############################################################################################
saidas['cod_unidade_efd'] = saidas['cod_unidade_efd'].str.strip()
efd0220['cod_unidade_efd'] = efd0220['cod_unidade_efd'].str.strip()

if LE_FATORES and (not tem_excel_fat_conv_saidas):
    print("gerando tabela de conversão de unidades das saídas... ", end="", flush=True)
    conv_unid_saidas = saidas[\
        ['COD_ITEM', 'cod_unidade_efd']].drop_duplicates().copy()
    conv_unid_saidas['fat_conv'] = 1
    caminho_completo = nome_pasta + r'\conversao_unidades_saidas.xlsx'
    conv_unid_saidas.to_excel(caminho_completo, index=False)
    print("Concluído!")
    print("Ajuste os fatores de conversão de saída e rode novamente.")
    print("FIM DA EXECUÇÃO!.")
    sys.exit(0)
elif LE_FATORES:
    # PAREI AQUI (08/04/2026)
    caminho_completo = caminho_pasta + r"\fat_conv_entradas.xlsx"
    df_fat_conv = pd.read_excel(caminho_completo)
    entradas = pd.merge(entradas, df_fat_conv, on=['COD_ITEM', 'cod_unidade_efd'], how='left',
                        validate='many_to_one', indicator=True)
    qtd_diferentes_both = entradas['_merge'].ne('both').sum()
    if qtd_diferentes_both > 0:
        raise SystemExit("Erro de Merge nos fatores de conversão de entradas!")
    entradas = entradas.drop(columns='_merge')
    entradas['QTD'] = entradas['QTD'] * entradas['fat_conv']

saidas['fator'] = 1
saidas = pd.merge(saidas, efd0220[['cod_unidade_efd', 'fat_novo']], on=['cod_unidade_efd'],\
                  how='left', indicator='_tem_fator')
saidas.loc[saidas['_tem_fator']=='both', 'fator'] = saidas['fat_novo']

saidas.loc[(saidas['_tem_fator']=='left_only') & (saidas['cod_unidade_efd']=='CX'), 'fator'] = 12
saidas.loc[(saidas['_tem_fator']=='left_only') & (saidas['cod_unidade_efd']=='EV'), 'fator'] = 12
saidas.loc[(saidas['_tem_fator']=='left_only') & (saidas['cod_unidade_efd']=='FD'), 'fator'] = 12
saidas.loc[(saidas['_tem_fator']=='left_only') & (saidas['cod_unidade_efd']=='CJ'), 'fator'] = 30
saidas.loc[(saidas['_tem_fator']=='left_only') & (saidas['cod_unidade_efd']=='PC'), 'fator'] = 4

saidas['QTD'] = saidas['QTD']*saidas['fator']
saidas.drop(columns=['fator', 'fat_novo', '_tem_fator'], inplace=True)

saidas = saidas[saidas['QTD'] != 0]
saidas['VL_CONFR'] = saidas['vlr_prod'] * saidas['aliq']/100
saidas['ICMS_TOT'] = 0.0

saidas['COD_LEGAL'] = "0"
saidas.loc[saidas['CFOP']=='5405', 'COD_LEGAL'] = "1"
saidas.loc[saidas['COD_LEGAL']=='0', 'VL_CONFR'] = 0.0

saidas.loc[saidas['VL_CONFR']<0.01,'VL_CONFR'] = 0.0
saidas.loc[saidas['VL_CONFR']==0,'COD_LEGAL'] = "0"

saidas['SUBTIPO'] = 2
saidas.loc[saidas['CFOP'].isin(dev_entrada),'SUBTIPO'] = 1
saidas.loc[saidas['CFOP'].isin(dev_entrada),'COD_LEGAL'] = ""
saidas.loc[saidas['CFOP'].isin(dev_entrada),'VL_CONFR'] = 0.0
primeira_op = saidas['icms'] + saidas['icms_st']
segunda_op = saidas['bc_icms_st_ant'] * saidas['aliq']/100
saidas.loc[saidas['CFOP'].isin(dev_entrada),'ICMS_TOT'] = primeira_op.combine(segunda_op, max)

df_icms_tot = saidas[['COD_ITEM', 'QTD', 'VL_CONFR']].copy()
df_icms_tot['ICMS_TOT_UNI'] = df_icms_tot['VL_CONFR'] / df_icms_tot['QTD'] * 1.5
df_icms_tot = df_icms_tot.drop_duplicates(subset=['COD_ITEM'])
df_icms_tot = df_icms_tot.rename(columns={'ICMS_TOT_UNI': 'ICMS_TOT_UNI_NOVO'})

print("Concluído!")
# ############################################################################################################# #
# FIM (Montagem das Saídas) ################################################################################### #
# ############################################################################################################# #

# ############################################################################################################# #
# Montagem dos dados (União das entradas e das saídas)######################################################### #
# ############################################################################################################# #
print("Concatenando entradas e saídas...", end="", flush=True)
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
print("Concluído!")
# ############################################################################################################# #
# FIM (Montagem dos dados (União das entradas e das saídas)) ################################################## #
# ############################################################################################################# #


# ############################################################################################################# #
# Cálculo do estoque inicial ################################################################################## #
# ############################################################################################################# #
print("Calculando o estoque inicial...", end="", flush=True)
est_ini = calcular_estoque_inicial_por_produto(dados)
est_ini = est_ini.sort_values(by=['DATA', 'SUBTIPO'])
est_ini['ano_mes'] = est_ini['DATA'].dt.to_period('M')
blocoh = blocoh.rename(columns={'QTD': 'QTD_blocoh'})
blocoh['QTD_blocoh'] = blocoh['QTD_blocoh'].str.replace(",", ".", regex=False).astype(float)
blocoh_util = blocoh[['COD_ITEM', 'QTD_blocoh']].copy()

est_ini = est_ini.loc[(est_ini['estoque_inicial_minimo']==0) | \
    ((est_ini['estoque_inicial_minimo']>0) &\
        (est_ini['COD_ITEM'].isin(blocoh_util['COD_ITEM'])))]
est_ini = pd.merge(est_ini, blocoh[['COD_ITEM', 'QTD_blocoh']],\
                           on=['COD_ITEM'], how='left', indicator = True)
# if not est_ini.loc[est_ini['_merge'] != 'both'].empty:
#     print("ERRO - existem registros na PCAT que não estão no Bloco H!!!!!")
#     raise exception()
#est_ini = est_ini.loc[est_ini['QTD_blocoh']>=est_ini['estoque_inicial_minimo']]
#est_ini['estoque_inicial_minimo'] = est_ini['QTD_blocoh']
# est_ini = est_ini.loc[(est_ini['estoque_inicial_minimo']>=0) & \
#                       (est_ini['estoque_inicial_minimo']<=4*est_ini['QTD_blocoh'])]

est_ini['ICMS_TOT_INI'] = est_ini['estoque_inicial_minimo'] * est_ini['ICMS_TOT_UNI']
est_ini.drop(columns='QTD_blocoh', inplace=True)

if ICMS_TOT_PELAS_SAIDAS:
    est_ini = pd.merge(est_ini, df_icms_tot[['COD_ITEM', 'ICMS_TOT_UNI_NOVO']],\
                             on='COD_ITEM', how='left', indicator='_teste')
    est_ini.loc[est_ini['_teste']=='both', 'ICMS_TOT_UNI'] = est_ini['ICMS_TOT_UNI_NOVO']
    est_ini['ICMS_TOT_INI'] = est_ini['estoque_inicial_minimo'] * est_ini['ICMS_TOT_UNI']
    est_ini.drop(columns=['_teste', 'ICMS_TOT_UNI_NOVO'], inplace=True)


print("Concluído!")
print("Analisando confronto entre os estoques da PCAT e o Bloco H...", end="", flush=True)
estoque_pcat = est_ini[['COD_ITEM', 'descricao_efd', 'estoque_inicial_minimo', 'ICMS_TOT_INI']]\
    .drop_duplicates(subset='COD_ITEM')
estoque_pcat['COD_ITEM'] = estoque_pcat['COD_ITEM'].apply(lambda x: x.strip())
blocoh['COD_ITEM'] = blocoh['COD_ITEM'].apply(lambda x: x.strip())
df_anal_estoque = pd.merge(estoque_pcat, blocoh[['COD_ITEM', 'QTD_blocoh', 'VL_UNIT']],\
                           on=['COD_ITEM'], how='left', indicator = 'True')
df_anal_estoque.rename(columns={'estoque_inicial_minimo': 'Estoque_PCAT',\
                                'QTD_blocoh': 'Estoque_BlocoH'}, inplace=True)
caminho_completo = nome_pasta + r"\teste_pcat_e_blocoh.xlsx"
df_anal_estoque.to_excel(caminho_completo, index=False)
print("Concluído!")
print("gerando resumo dos códigos de item...", end="", flush=True)
caminho_completo = nome_pasta + r"\resumo_cod_item.xlsx"
# Agrupar por COD_ITEM e somar os campos desejados
resumo = (
    est_ini.groupby("COD_ITEM", as_index=False)
         .agg({
        "descricao_efd": "first",
        "ICMS_TOT_UNI": "first",
        "estoque_inicial_minimo": "first",
        "ICMS_TOT_INI": "first",
        "ICMS_TOT": "sum",
        "VL_CONFR": "sum",
        "QTD_SAIDA": "sum"
        })
)
resumo.rename(columns={'estoque_inicial_minimo': 'Estoque_Ini'}, inplace=True)
# Exportar para Excel
resumo.to_excel(caminho_completo, index=False)
print("Concluído!")

# ############################################################################################################# #
# FIM (Cálculo do estoque inicial) ############################################################################ #
# ############################################################################################################# #


# ############################################################################################################# #
# Montagem e geração dos arquivos da PCAT 42/2018 ############################################################# #
# ############################################################################################################# #
print("Iniciando Montagem e geração dos arquivos da PCAT...")
# Cria coluna auxiliar com o ano e mês
est_ini['ano_mes'] = est_ini['DATA'].dt.to_period('M')

ficha3_primeira, df_icms_tot_ini = gerar_ficha3_e_1050(est_ini)

cods_ress = ['1', '2', '3', '4']
# Garante que não haja duplicações por COD_ITEM no mesmo mês
grupo_mes = est_ini.groupby(['ano_mes', 'COD_ITEM']).agg({
    'estoque_inicial_minimo': 'first',
    'ICMS_TOT_INI': 'first'
}).reset_index()

# Para cada mês, gerar um arquivo .txt
reg_1050_ant = pd.DataFrame(columns=['itemcode', 'iq', 'iv', 'fq', 'fv'])
for periodo, grupo in grupo_mes.groupby('ano_mes'):
    nome_arquivo = f"{nome_pasta}/RessarcimentoST-{CNPJ_ENTIDADE}_{periodo.strftime('%Y-%m')}.txt"

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
    if reg_1050_ant.empty:
        reg_1050_ant = reg1050[['itemcode', 'iq', 'iv', 'fq', 'fv']].copy()
    else:
        reg_1050_ant = pd.merge(reg1050[['itemcode', 'iq', 'iv', 'fq', 'fv']], reg_1050_ant, on='itemcode',\
                                how='outer', indicator='total')
        reg_1050_ant.loc[reg_1050_ant['total'] == 'right_only', ['iq_x', 'fq_x', 'iv_x', 'fv_x']] = \
            reg_1050_ant.loc[reg_1050_ant['total'] == 'right_only', ['iq_y', 'fq_y', 'iv_y', 'fv_y']].values
        reg_1050_ant = reg_1050_ant.drop(columns=['iq_y', 'fq_y', 'iv_y', 'fv_y', 'total'])
        reg_1050_ant = reg_1050_ant.rename(columns={'iq_x': 'iq', 'fq_x':'fq' , 'iv_x': 'iv', 'fv_x': 'fv'})



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
print("SUCESSO - EXECUÇÃO ENCERRADA SEM ERROS!")