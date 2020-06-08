import xlwings as xw
import numpy  as np
import pandas as pd

def create_table (book):

    wb = xw.Book(book)
    sht = wb.sheets['SQL']
    pos = sht.range('B3').options(pd.DataFrame, expand='table').value
    btc = sht.range('G3').options(pd.DataFrame, expand='table').value

#----Creating New SQL, without blank fund names
#----Might be unnecessary after BPREV starts showing up correctly in SQL
    sht2 = wb.sheets['New SQL']
    pos = pos.replace(np.nan,'BPREV',regex=True)
    sht2.range('A1').value = pos
    btc = btc.replace(np.nan,'BPREV',regex=True)
    sht2.range('G1').value = btc
#------------------------------------------------------
    btc_matrix = []
    group = btc.groupby(['Fund', 'Instrument'])
    for group_name, group_df in group:
        tomador = group_df.query("Side == 'T'")['QTY'].sum()
        doador = group_df.query("Side == 'D'")['QTY'].sum()
        gross = pos[ (pos['Fund']==group_name[0]) & (pos['Instrument']==group_name[1]) ]['Qty'].sum()
        net = gross+tomador-doador
        #calculate cover
        if(gross ==0):
            cover = "Liquidate"
        elif (-(doador+tomador)/gross < 0 ):
            cover = "Liquidate"
        else:
            cover = '{:.1f}%'.format(((-doador-tomador)/gross)*100)

        #Consolidate pos and btc
        consolidation = [group_name[0],group_name[1],tomador,doador,gross,net,cover]
        btc_matrix.append(consolidation)

    btc_consolidated = pd.DataFrame(btc_matrix,columns = ['Fund','Instrument','Tomador','Doador','Gross Shares','Net Shares','Cover'])
    sht = wb.sheets['New SQL']
    sht.range('S1').value = btc_consolidated

    #print(sum)
    return  pos,btc,btc_consolidated