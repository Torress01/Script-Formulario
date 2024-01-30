import pyautogui
import openpyxl
import pyperclip
from time import sleep

workbook = openpyxl.load_workbook('produtos_ficticios.xlsx');
sheet_produtos = workbook['Produtos']

for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1342,293,duration=1)
    pyautogui.hotkey('ctrl','v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(1257,394,duration=1)
    pyautogui.hotkey('ctrl','v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1260,556,duration=1)
    pyautogui.hotkey('ctrl','v')

    codigo = linha[3].value
    pyperclip.copy(codigo)
    pyautogui.click(1261,662,duration=1)
    pyautogui.hotkey('ctrl','v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1321,769,duration=1)
    pyautogui.hotkey('ctrl','v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(1288,873,duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(1291,964,duration=1)
    sleep(3)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(1280,319,duration=1)
    pyautogui.hotkey('ctrl','v')

    quantidade = linha[7].value
    pyperclip.copy(quantidade)
    pyautogui.click(1262,423,duration=1)
    pyautogui.hotkey('ctrl','v')

    data = linha[8].value
    pyperclip.copy(data)
    pyautogui.click(1259,540,duration=1)
    pyautogui.hotkey('ctrl','v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(1270,652,duration=1)
    pyautogui.hotkey('ctrl','v')

    tamanho = linha[10].value
    pyperclip.copy(tamanho)
    pyautogui.click(1273,747,duration=1)
    #ler informacao da planilha e clicar de acordo com a posicao
    if tamanho == 'Pequeno':
        pyautogui.click(1283,786,duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(1292,818,duration=1)
    else:
        pyautogui.click(1290,845,duration=1)

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(1283,858,duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(1291,932,duration=1)
    sleep(3)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(1261,343,duration=1)
    pyautogui.hotkey('ctrl','v')

    pais = linha[13].value
    pyperclip.copy(pais)
    pyautogui.click(1269,454,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(1265,565,duration=1)
    pyautogui.hotkey('ctrl','v')

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(1278,723,duration=1)
    pyautogui.hotkey('ctrl','v')

    localizacao = linha[16].value
    pyperclip.copy(localizacao)
    pyautogui.click(1278,840,duration=1)
    pyautogui.hotkey('ctrl','v')    

    #Concluir
    pyautogui.click(1291,911,duration=1)
    #Confirmar
    pyautogui.click(1803,205,duration=1)
    #2ºConfirmar
    pyautogui.click(1803,205,duration=1)
    #Adicionar mais um 
    pyautogui.click(1554,635,duration=1)


