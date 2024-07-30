import openpyxl as pyxl
import pyautogui as pygui
workbook = pyxl.load_workbook ('vendas_de_produtos.xlsx')

vendas_sheet = workbook['vendas']

for linha in vendas_sheet.iter_rows(min_row= 2) :
    #campo nome
    pygui.click(1171, 331, duration=1.5)
    pygui.write(linha[0].value)
    #campo produto
    pygui.click(1184, 352, duration=1.5)
    pygui.write(linha[1].value)
    #campo quantidade
    pygui.click(1183, 378, duration=1.5)
    pygui.write(str(linha[2].value))
    #campo categoria
    pygui.click(1248, 406, duration=1.5)
    pygui.write(linha[3].value)
    #salvar
    pygui.click(1117, 437, duration=1.5)
    #ok
    pygui.click(654, 428, duration=1.5)
