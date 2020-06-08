import xlwings as xw
import numpy  as np
import pandas as pd
import datetime as dt

def liquidate(btc,consolidated):
    print("liquidando")
    btc_l = pd.DataFrame.copy(btc)
    btc_consilidated = pd.DataFrame.copy(consolidated)

    l_table = btc_consilidated[ (btc_consilidated['Cover'] == 'Liquidate') ]
    print(l_table)
    l_list = l_table.values.tolist()

    boleta_l = []
    for i in range(len(l_list)):  #todas as liquidaçoes
        fund = l_list[i][0]
        instrument = l_list[i][1]
        subl = btc_l[ (btc_l['Fund'] == fund) & (btc_l['Instrument'] == instrument) ] #df com todos os contratos de um papel
        #listas com infos dos contratos de um papel
        qty_list = subl['QTY'].values.tolist()
        contract_list = subl['Contract'].values.tolist()
        rate_list = subl['Rate'].values.tolist()
        exp_list = pd.to_datetime(subl['Expiration'].values).tolist()

        single_instrument_l = [] #lista para as liquidaçoes de um papel
        for j in range(len(qty_list)): #montando as liquidações de um papel
            liqlist_individual = [fund,instrument,qty_list[j],'TOMADOR',rate_list[j],exp_list[j],contract_list[j],'Liquidar']
            boleta_l.append(liqlist_individual)

    boleta_l_df = pd.DataFrame(boleta_l,columns=['Fund','Ticker','Qty','Side','Rate','Exp','Contrato','Tipo'])
    print(boleta_l_df)
    return boleta_l_df






    return 0


