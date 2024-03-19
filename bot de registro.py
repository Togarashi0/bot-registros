import openpyxl
import pyperclip
import pyautogui
from time import sleep

workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(2328,347,duration=1)
    pyautogui.hotkey('ctrl', 'v')
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(2327,442,duration=1)
    pyautogui.hotkey('ctrl', 'v')
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(2379,565,duration=1)
    pyautogui.hotkey('ctrl', 'v')
    cod_prod = linha[3].value
    pyperclip.copy(cod_prod)
    pyautogui.click(2370,651,duration=1)
    pyautogui.hotkey('ctrl', 'v')
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(2371,739,duration=1)
    pyautogui.hotkey('ctrl', 'v')
    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(2363,831,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(2272,878,duration=1)
    sleep(3)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(2299,372,duration=1)
    pyautogui.hotkey('ctrl', 'v')
    quantidade = linha[7].value
    pyperclip.copy(quantidade)
    pyautogui.click(2302,462,duration=1)
    pyautogui.hotkey('ctrl', 'v')
    data_val = linha[8].value
    pyperclip.copy(data_val)
    pyautogui.click(2300,536,duration=1)
    pyautogui.hotkey('ctrl', 'v')
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(2294,630,duration=1)
    pyautogui.hotkey('ctrl', 'v')
    tamanho = linha[10].value
    pyautogui.click(2301,713,duration=1)
    if tamanho == 'Pequeno':
        pyautogui.click(2288,742,duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(2300,769,duration=1)
    else:
        pyautogui.click(2349,792,duration=1)

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(2325,803,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(2269,867,duration=1)
    sleep(2)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(2277,397,duration=1)
    pyautogui.hotkey('ctrl', 'v')
    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(2280,472,duration=1)
    pyautogui.hotkey('ctrl', 'v')
    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(2299,574,duration=1)
    pyautogui.hotkey('ctrl', 'v')
    cod_barras = linha[15].value
    pyperclip.copy(cod_barras)
    pyautogui.click(2299,693,duration=1)
    pyautogui.hotkey('ctrl', 'v')
    localizacao = linha[16].value
    pyperclip.copy(localizacao)
    pyautogui.click(2281,779,duration=1)
    pyautogui.hotkey('ctrl', 'v')
    #botao concluir
    pyautogui.click(2266,843,duration=1)
    #botao de confirmação
    pyautogui.click(3039,174,duration=1)
    #botao de confirmação2
    pyautogui.click(2857,626,duration=1)
   
