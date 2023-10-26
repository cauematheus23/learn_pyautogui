import openpyxl
import pyperclip
import pyautogui
from time import sleep
#entrar na planilha
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']
#Copiar informação de um campo e colar no seu campo correspondente
for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1468,345,duration=1)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('tab')
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('tab')
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('tab')
    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('tab')
    peso_produto = linha[4].value
    pyperclip.copy(peso_produto)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('tab')
    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.hotkey('ctrl','v')
    pyautogui.click(1419,881,duration =1)
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(1406,369,duration=1)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('tab')
    quantidade_estoque = linha[7].value
    pyperclip.copy(quantidade_estoque)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('tab')
    data_validade = linha[8].value
    pyperclip.copy(data_validade)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('tab')
    cor_produto = linha[9].value
    pyperclip.copy(cor_produto)
    pyautogui.hotkey('ctrl','v')
    tamanho = linha[10].value
    pyautogui.click(1433,716,duration=1)
    sleep(1)
    if tamanho == 'Pequeno':
        pyautogui.hotkey('tab')
    elif tamanho == 'Médio':
        pyautogui.click(1408,764,duration=1)
    elif tamanho == 'Grande':
        pyautogui.click(1420,785,duration=1)
    pyautogui.hotkey('tab')
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.hotkey('ctrl','v')
    pyautogui.click(1416,868,duration=1)
    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(1434,392,duration=1)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('tab')
    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('tab')
    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('tab')
    codigo_barras = linha[15].value
    pyperclip.copy(codigo_barras)
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('tab')
    localizacao = linha[16].value
    pyperclip.copy(pais_origem)
    pyautogui.hotkey('ctrl','v')
    pyautogui.click(1414,837,duration=1)
    pyautogui.click(2115,172,duration=1)
    pyautogui.click(1918,616,duration=1)