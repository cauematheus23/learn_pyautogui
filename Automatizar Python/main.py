import openpyxl
import pyperclip
import pyautogui
#entrar na planilha
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']
#Copiar informação de um campo e colar no seu campo correspondente
for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1449,341,duration=1)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('tab')
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('tab')
    categoria = linha[2].value
    codigo_produto = linha[3].value
    peso_produto = linha[4].value
    dimensoes = linha[5].value
    preco = linha[6].value
    quantidade_estoque = linha[7].value
    data_validade = linha[8].value
    cor_produto = linha[9].value
    tamanho = linha[10].value
    material = linha[11].value
    codigo_barras = linha[12].value
    localizacao = linha[13].value