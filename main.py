from create_table import *
from liquidate import *
from create_boleta import *

def main():
    print("Creating table")
    wb = 'Aluguel.xlsb'
    tables = create_table(wb)
    #tables = [pos,btc,btc_consolidated]
    liquidacoes = liquidate(tables[1],tables[2])
    create_boleta(wb,liquidacoes)


main()