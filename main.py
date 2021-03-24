# -*- coding: utf-8 -*-
"""
- O item 1 tem que ser administracao Local
- Salvar como valor tudo dentro da aba Orçamento e Calculo e apagar as outras abas
- 
@author: felipebrandao
"""
linhaInicial = 62
linhaUltima = 72
colunaInicial = 13
#cposName = 'CPOS 178'

time0 = 1
time1 = 2#float ( input("obs-siconv em primeiro plano, excel em segundo, celula do codigo marcado, digite o tempo lento") )
time2 = 5#float( input("digite o tempo rapido") )
n_frentes = 4 #number of frentes de obra
item_times = 1#int( input("digite a quantidade de itens sinapi"))
bdi1 = "24,22" #16,79 #24,23
bdi2 = "16,78"



import pandas as pd
import pyautogui
#import win32clipboard
import matplotlib.pyplot as plt
from PIL import Image
from PIL import ImageChops
import time
import xerox
#import win32clipboard
#import locale

#locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

orc = pd.read_excel('planilha.xlsx', sheet_name='ORÇAMENTO', encoding='latin-1')
calc = pd.read_excel('planilha.xlsx', sheet_name='CÁLCULO')

n_evento = orc.iloc[linhaInicial-2, colunaInicial+1 ].split('.')
n_evento = int(n_evento[1])

def equals(first_img, second_img):
    """ 
    Identifies if first_img and second_img are identical to each other. 
    Args:
     first_img: a PIL image object of the original image 
    second_img: a PIL image object of the modified image 
    Returns: A boolean stating whether or not the two images are identical, True means they are identical. 
    """ 

    if first_img is None or second_img is None: 
        return False 
    diff = ImageChops.difference(first_img, second_img)
    if diff.getbbox() != None: 
        return False
    else: 
        diff = None
        return True 
    


def cadastroItem(is_SINAPI,primeiraVez,codigo,descricao, unidade, precoUnitario, bdi, quantidades):  
        
        time.sleep(2)
        pyautogui.PAUSE = time0
        if primeiraVez == 1:
            pyautogui.keyDown('alt')
            pyautogui.hotkey('tab')
            pyautogui.keyUp('alt')
        #pyautogui.confirm('Prosseguir0')
        
        pyautogui.hotkey('tab')
        pyautogui.hotkey('enter')
        pyautogui.PAUSE = 2
        
        #image = pyautogui.locateOnScreen('detect_my_job1.png')
        #while image != None:
        #    print('prosseguir 1')  
        #   image = pyautogui.locateOnScreen('detect_my_job.png')
        time.sleep(2)
        #im = pyautogui.screenshot('test1-1.png', region=(660,560, 80, 40))
        #im = pyautogui.screenshot('test1-1-trabalho.png', region=(670,490, 80, 30))
        im = pyautogui.screenshot('test1-1-trabalho.png', region=(670,460, 80, 30))
        imgplot = plt.imshow(im)
        plt.show()
        im2 = Image.open("aguarde_trabalho.png")
        equalImages = equals(im, im2)
        print(equalImages)
        
        while equalImages == True:
            im = pyautogui.screenshot('test1-2-trabalho.png', region=(670,460, 80, 30))
            equalImages = equals(im, im2)
            print('cadastro item stop 1')
            
        #pyautogui.confirm('Prosseguir1')
        #pyautogui.PAUSE = time2
        pyautogui.hotkey('right')
        pyautogui.PAUSE = 2
    
        im = pyautogui.screenshot('test2-1-trabalho.png', region=(670,490, 80, 30))    
        equalImages = equals(im, im2)
        while equalImages == True:
            im = pyautogui.screenshot('test.png', region=(670,490, 80, 30))
            equalImages = equals(im, im2)
            print('cadastro item stop 2')
        
        #pyautogui.confirm('Prosseguir2')
        pyautogui.hotkey('tab')
        pyautogui.PAUSE = time0
        #pyautogui.hotkey('down')
        pyautogui.typewrite(str(n_evento), interval=0.001)
        
        if n_evento >= 10:
            time.sleep(3)
            pyautogui.typewrite(str(n_evento), interval=0.001)
        
        pyautogui.PAUSE = time1
        pyautogui.hotkey('tab')
        pyautogui.PAUSE = 0.25
        pyautogui.hotkey('tab')
        pyautogui.PAUSE = 2
        
        if is_SINAPI == 1: # eh SINAPI
            pyautogui.hotkey('s')
            pyautogui.PAUSE = 1
            pyautogui.hotkey('tab')
            pyautogui.hotkey('enter')
            
            im = pyautogui.screenshot('test2.png', region=(970,495, 90, 25))
            imgplot = plt.imshow(im)
            plt.show()
            imlupapesquisar = Image.open("lupaPesquisar.png")
            equalImages = equals(im, imlupapesquisar)
            #while equalImages == False:
            im = pyautogui.screenshot('test2.png', region=(970,495, 200, 100))
             #   equalImages = equals(im, imlupapesquisar)
             #   print(equalImages)
            
            #pyautogui.confirm('Prosseguir - clicou na lupa?')
            pyautogui.hotkey('shift','tab')
            pyautogui.hotkey('left')
            pyautogui.hotkey('tab')
            ##pyautogui.keyDown('alt')
            ##if primeiraVez == 1:
            ##    pyautogui.hotkey('tab')
            ##pyautogui.hotkey('tab')
            ##pyautogui.keyUp('alt')
            #copia e cola o CODIGO
            ##pyautogui.hotkey('ctrl','c') #copia e cola o CODIGO
            ##pyautogui.hotkey('right')
            ##pyautogui.hotkey('right')
            ##pyautogui.hotkey('right')
            ##pyautogui.hotkey('right')
            
            ##pyautogui.keyDown('alt')
            ##pyautogui.hotkey('tab')
            ##pyautogui.keyUp('alt')
            pyautogui.PAUSE = 0.5
            
            pyautogui.typewrite( codigo, interval=0.1)
            
            #pyautogui.hotkey('ctrl','v')
            pyautogui.hotkey('tab')
            pyautogui.hotkey('enter')
            pyautogui.PAUSE = 2
            time.sleep(3)
            #pyautogui.confirm('Prosseguir - PRONTO PARA APERTAR ENTER?')
            pyautogui.hotkey('enter')
            pyautogui.hotkey('tab')
            pyautogui.hotkey('tab')
            pyautogui.hotkey('tab')
            pyautogui.hotkey('enter')
            pyautogui.PAUSE = 2
            #copia e cola o CUSTO UNITÁRIO
            ##pyautogui.keyDown('alt')
            ##pyautogui.hotkey('tab')
            ##pyautogui.keyUp('alt')
            
            time.sleep(3)
            #pyautogui.confirm('Prosseguir - NO EXCEL PRONTO COPIAR CUSTO?')
            ##pyautogui.hotkey('ctrl','c')
            #pyautogui.confirm('Prosseguir - COPIOU O CUSTO?')
            ##pyautogui.keyDown('alt')
            ##pyautogui.hotkey('tab')
            ##pyautogui.keyUp('alt')
            pyautogui.typewrite( precoUnitario, interval=0.1)
            ##pyautogui.hotkey('ctrl','v')
            
            
        if is_SINAPI == 0: #nao eh SINAPI, eh CPOS, SIURB,...
            pyautogui.hotkey('o')
            pyautogui.PAUSE = time2
            pyautogui.hotkey('tab')
            pyautogui.PAUSE = 0.25
    
            #copy and paste CODIGO
            ##pyautogui.keyDown('alt')
            ##if primeiraVez == 1:
            ##    pyautogui.hotkey('tab')
            ##pyautogui.hotkey('tab')
            ##pyautogui.keyUp('alt')
            ##pyautogui.hotkey('ctrl','c')
            ##pyautogui.keyDown('alt')
            ##pyautogui.hotkey('tab')
            ##pyautogui.keyUp('alt')
            pyautogui.PAUSE = 0.5
            #pyautogui.hotkey('tab')
            #pyautogui.keyDown('alt')
            #pyautogui.confirm('Prosseguir3 - pronto para colar o codigo')
            pyautogui.typewrite( codigo, interval=0.1)
            ##pyautogui.hotkey('ctrl','v')
            #fim do copia e cola
            pyautogui.hotkey('tab')
        
            #copy and paste
            #descricao
            #pyautogui.confirm('Prosseguir3-1 - o codigo deveria ir para o excel e copiar a descricao')
            pyautogui.PAUSE = 0.25
            ##pyautogui.keyDown('alt')
            ##pyautogui.hotkey('tab')
            #pyautogui.hotkey('tab')
            ##pyautogui.keyUp('alt')
            ##pyautogui.hotkey('right')
            ##pyautogui.hotkey('ctrl','c')
            ##pyautogui.keyDown('alt')
            ##pyautogui.hotkey('tab')
            ##pyautogui.keyUp('alt')
            #pyautogui.hotkey('tab')
            #pyautogui.typewrite( descricao, interval=0.1)
            xerox.copy(descricao)
            xerox.paste()
            pyautogui.hotkey('ctrl','v')
            #fim do copia e cola
            #pyautogui.hotkey('tab')
            #copia e cola a UNIDADE
            ##pyautogui.keyDown('alt')
            ##pyautogui.hotkey('tab')
            ##pyautogui.keyUp('alt')
            ##pyautogui.hotkey('right')
            ##pyautogui.hotkey('ctrl','c')
            #OBS: criar mecanismo para analisar termo copiado
            ##win32clipboard.OpenClipboard()
            ##unidade = win32clipboard.GetClipboardData()
            ##print(unidade)
            ##win32clipboard.CloseClipboard()
            #pyautogui.confirm('Prosseguir3-1-1 - copiou a unidade?')   
            ##pyautogui.keyDown('alt')
            ##pyautogui.hotkey('tab')
            ##pyautogui.keyUp('alt')
            pyautogui.PAUSE = 2
            pyautogui.hotkey('tab')
            pyautogui.hotkey('space')
            
            pyautogui.typewrite(unidade, interval=0.1)   
            #pyautogui.confirm('Prosseguir3-2 - tudo certo com a unidade?')
            pyautogui.PAUSE = 1
            #UNIDADE
            #pyautogui.confirm('Prosseguir3-2 COLOCAR UNIDADE')
            #pyautogui.typewrite(unidade, interval=0.1)
            pyautogui.hotkey('tab')
            #copy and paste
            #custo referencia
            pyautogui.PAUSE = 0.25
            ##pyautogui.keyDown('alt')
            ##pyautogui.hotkey('tab')
            ##pyautogui.keyUp('alt')
            pyautogui.PAUSE = 0.25
            #pyautogui.hotkey('right')
            ##pyautogui.hotkey('right')
            ##pyautogui.hotkey('right')
            ##pyautogui.hotkey('ctrl','c')
            ##pyautogui.keyDown('alt')
            ##pyautogui.hotkey('tab')
            ##pyautogui.keyUp('alt')
            pyautogui.PAUSE = time1
            #pyautogui.hotkey('tab')
            pyautogui.typewrite(precoUnitario, interval=0.1)   
            ##pyautogui.hotkey('ctrl','v')
            #fim do copia e cola
            pyautogui.hotkey('tab')
            
            #custo unitario
            ##pyautogui.hotkey('ctrl','v')
            pyautogui.typewrite(precoUnitario, interval=0.1)   
        
        #COMUM PARA TANTO SINAPI COMO PARA AS OUTRAS TABELAS
        pyautogui.PAUSE = 1
        pyautogui.hotkey('tab')
        #pyperclip.copy(bdi)
        #pyperclip.paste()
        xerox.copy(bdi)
        pyautogui.hotkey('ctrl','v')
        #pyautogui.typewrite( bdi, interval=1)
        pyautogui.PAUSE = 0.25
        
        #com eventos
        pyautogui.hotkey('tab')
        #pyautogui.hotkey('down')
        pyautogui.typewrite(str(n_evento), interval=0.1)
        #fim com eventos
        
        #copy and paste
        #quantidade (COM EVENTOS)
        pyautogui.PAUSE = 0.25
        pyautogui.hotkey('tab')
        ##pyautogui.keyDown('alt')
        ##pyautogui.hotkey('tab')
        ##pyautogui.keyUp('alt')
        pyautogui.PAUSE = 0.25
        #pyautogui.hotkey('left') antes pegava da propria PO
        ##pyautogui.hotkey('ctrl','pagedown') #muda para a aba CALCULO
        ##pyautogui.hotkey('ctrl','c')
        
        pyautogui.PAUSE = 1
        
        for x in range (1,n_frentes+1):
            
            #pyautogui.typewrite(quantidades[x-1], interval=0.1)
            time.sleep(2)
            xerox.copy(quantidades[x-1])
            pyautogui.hotkey('ctrl','v')
            
            
            pyautogui.hotkey('tab')
            ##pyautogui.keyDown('alt')
            ##pyautogui.hotkey('tab')
            ##pyautogui.keyUp('alt') #muda para o excel
            ##if x != n_frentes+1:
            ##    pyautogui.hotkey('right')
            ##    pyautogui.hotkey('ctrl','c')
            ##    print(x)
                
            #if x == (n_frentes):
        print('aqui')
        ##pyautogui.hotkey('down')
        ##for y in range (1, n_frentes+1): #deixa o cursor na primeira frente do proximo item
        ##    pyautogui.hotkey('left')
        ##pyautogui.hotkey('ctrl', 'pageup') #volta para a aba PO
        ##pyautogui.hotkey('down')
        ##pyautogui.hotkey('left')
        ##pyautogui.hotkey('left')
        ##pyautogui.hotkey('left')
        ##pyautogui.hotkey('left') #deixa o cursor no codigo do proximo item na PO
        ##pyautogui.keyDown('alt')
        ##pyautogui.hotkey('tab')
        ##pyautogui.keyUp('alt') #muda para o siconv
            
        
        pyautogui.hotkey('tab')
        pyautogui.hotkey('tab')
        pyautogui.hotkey('enter') #aperta salvar no siconv
        
        im = pyautogui.screenshot('test1-1-trabalho.png', region=(670,490, 80, 30))
        imgplot = plt.imshow(im)
        plt.show()
        im2 = Image.open("aguarde_trabalho.png")
        equalImages = equals(im, im2)
        print(equalImages)
        
        while equalImages == True:
            im = pyautogui.screenshot('test1-2-trabalho.png', region=(670,490, 80, 30))
            equalImages = equals(im, im2)
            print('cadastro item stop 1')


def cadastroMacroItem(primeiraVez,descricaoMacro):

        time.sleep(2)
        pyautogui.PAUSE = 1
        if primeiraVez == 1:
            pyautogui.keyDown('alt')
            pyautogui.hotkey('tab')
            pyautogui.keyUp('alt')
    
        pyautogui.hotkey('tab')
        pyautogui.hotkey('enter')
        
        time.sleep(2)
        im = pyautogui.screenshot('test1-1-trabalho.png', region=(670,490, 80, 30))
        imgplot = plt.imshow(im)
        plt.show()
        im2 = Image.open("aguarde_trabalho.png")
        equalImages = equals(im, im2)
        print(equalImages)
        
        while equalImages == True:
            im = pyautogui.screenshot('test1-2-trabalho.png', region=(670,490, 80, 30))
            equalImages = equals(im, im2)
            print('cadastro item stop 1')
            
        #pyautogui.confirm('Prosseguir1')
        time.sleep(2)
        pyautogui.hotkey('space') #seleciona macroservico
        pyautogui.hotkey('tab')
        pyautogui.hotkey('tab')
        
        ##pyautogui.keyDown('alt') #muda para o excel para copiar
        #e colar nome do macroservico
        ##pyautogui.hotkey('tab')
        ##pyautogui.hotkey('tab')
        ##pyautogui.keyUp('alt')
        
        ##pyautogui.hotkey('right')
        ##pyautogui.hotkey('ctrl', 'c')
        ##pyautogui.hotkey('down')
        ##pyautogui.hotkey('left')
        ##pyautogui.hotkey('ctrl', 'pagedown')
        ##pyautogui.hotkey('down')
        ##pyautogui.hotkey('ctrl', 'pageup')
        
        ##pyautogui.keyDown('alt')
        ##pyautogui.hotkey('tab')
        ##pyautogui.keyUp('alt')
        
        ##pyautogui.hotkey('ctrl', 'v')
        #######
        ##pyautogui.typewrite(descricaoMacro, interval=0.1)
        xerox.copy(descricaoMacro)
        xerox.paste()
        pyautogui.hotkey('ctrl','v')
        
        pyautogui.hotkey('tab')
        pyautogui.hotkey('tab')
        pyautogui.hotkey('enter')
        
        im = pyautogui.screenshot('test1-1-trabalho.png', region=(670,490, 80, 30))
        imgplot = plt.imshow(im)
        plt.show()
        im2 = Image.open("aguarde_trabalho.png")
        equalImages = equals(im, im2)
        print(equalImages)
        
        while equalImages == True:
            im = pyautogui.screenshot('test1-2-trabalho.png', region=(670,490, 80, 30))
            equalImages = equals(im, im2)
            print('cadastro macroitem stop 1')
    


        
#### MAIN
primeiraVez = 1
while linhaInicial <= linhaUltima:
    
    
    tipoNivel = orc.iloc[linhaInicial-2, colunaInicial ]
    
    if tipoNivel == "Serviço":
        tablePrice = orc.iloc[linhaInicial-2, colunaInicial+2 ]
        if tablePrice == 'SINAPI' or tablePrice == 'SINAPI-I':
            is_SINAPI = 1 # eh SINAPI = 1 ; eh outra tabela = 0;
        else:
            is_SINAPI = 0 # eh SINAPI = 1 ; eh outra tabela = 0;
            
        codigo = orc.iloc[linhaInicial-2, colunaInicial+3 ]
        descricao = orc.iloc[linhaInicial-2, colunaInicial+4 ]
        unidade = orc.iloc[linhaInicial-2, colunaInicial+5 ]
        precoUnitario = str(orc.iloc[linhaInicial-2, colunaInicial+7 ])
        precoUnitario = precoUnitario.replace('.',',')
        bdiTipo = orc.iloc[linhaInicial-2, colunaInicial+8 ]
        
        if bdiTipo == "BDI 2":
            bdi = bdi2
        elif bdiTipo == "BDI 1":
            bdi = bdi1
        quantidades = []
        
        for i in range(n_frentes):
            quantidades.append( str(calc.iloc[linhaInicial-2, colunaInicial+3+i ]) )
            quantidades[i] = quantidades[i].replace('.',',')
            
            
        cadastroItem(is_SINAPI,primeiraVez,codigo,descricao, unidade, precoUnitario, bdi, quantidades)
        
    elif tipoNivel == "Nível 2":
        descricaoMacro = orc.iloc[linhaInicial-2, colunaInicial+4 ]
        cadastroMacroItem(primeiraVez,descricaoMacro)
        
    #elif tipoNivel == "Nível 3" or tipoNivel == "Nível 4":
        
        
    linhaInicial = linhaInicial + 1
    primeiraVez = 0
    n_evento = orc.iloc[linhaInicial-2, colunaInicial+1 ].split('.')
    n_evento = int(n_evento[1])
    print(linhaInicial)
