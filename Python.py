#UNIR PDfs e trasnformar em EXCEL e colocar VBA em ponto de bala (2 macros juntas)
# UNIR PDFS
import pyautogui
import time

pyautogui.alert("Irá começar")
pyautogui.PAUSE=1

#acessar site de unir pdfs
pyautogui.press('winleft')
pyautogui.write('Chrome')
pyautogui.press('enter')
pyautogui.write('https://smallpdf.com/pt/juntar-pdf')
time.sleep(1)
pyautogui.press('enter')
pyautogui.hotkey('winleft','up')

#área de trabalho
pyautogui.hotkey('winleft','d')
#abrir pasta de pdfs a unir
pyautogui.moveTo(1318, 234)
pyautogui.doubleClick(1318, 234)
#maximizar janela
pyautogui.hotkey('winleft','up')
#clicar dentro da pasta e maximizar tela
pyautogui.moveTo(350, 136,duration=1)
pyautogui.click(350, 136)
pyautogui.hotkey('ctrl','a')
#arrastar arquivos
pyautogui.mouseDown(477, 380)
pyautogui.moveTo(669, 566)
pyautogui.hotkey('alt','tab')
pyautogui.mouseUp(669, 566)
#esperar fazer conversão
time.sleep(80)
#clicar nas opções
pyautogui.moveTo(830,436)
pyautogui.click(830,436)
pyautogui.moveTo(863,629)
pyautogui.click(863,629)
#juntar pdfs
pyautogui.moveTo(1079,682)
pyautogui.click(1079,682)
#esperar
time.sleep(10)
#baixar
pyautogui.moveTo(1146, 329)
pyautogui.click(1146, 329)
pyautogui.moveTo(999, 382)
pyautogui.click(999, 382)
time.sleep(5)
pyautogui.write('renomear')
pyautogui.press('enter')
#arrastar pdf unido
time.sleep(5)
pyautogui.moveTo(211, 701)
pyautogui.click(211,701)
pyautogui.moveTo(275,631)
pyautogui.click(275,631)
pyautogui.hotkey('ctrl','x')
pyautogui.hotkey('winleft','d')
pyautogui.moveTo(1308, 335)
pyautogui.doubleClick(1308, 335)
pyautogui.hotkey('ctrl','v')
#conectar com próxima macro
pyautogui.hotkey('alt','f4')
pyautogui.moveTo(1311, 232)
pyautogui.doubleClick(1311, 232)
pyautogui.hotkey('alt','f4')
pyautogui.moveTo(711, 748,duration=1)
pyautogui.moveTo(711, 748,duration=1)
pyautogui.moveTo(756, 666,duration=1)
pyautogui.click(756, 666,duration=1)
pyautogui.hotkey('alt','f4')
pyautogui.moveTo(513, 757,duration=1)
pyautogui.moveTo(514, 666,duration=1)
pyautogui.click(514, 666,duration=1)
pyautogui.hotkey('alt','f4')

#PDFS PARA EXCEL

#dar pausa entre comandos
#pyautogui.PAUSE=1

#fechar pasta PDFs a unir
#pyautogui.hotkey('alt','tab')
#pyautogui.hotkey('alt','f4')

#abrir site
pyautogui.press('winleft')
pyautogui.write('Chrome')
pyautogui.press('enter')
pyautogui.write('https://easypdf.com/pt/pdf-em-excel')
time.sleep(1)
pyautogui.press('enter')
pyautogui.hotkey('winleft','up')
pyautogui.moveTo(1132,245,duration=1)
time.sleep(13)
pyautogui.click(1132,245)

#ir para área de trabalho
pyautogui.hotkey('winleft','d')
#entrar na pasta do PDF
pyautogui.moveTo(1313, 336,duration=1)
pyautogui.doubleClick(1313, 336)
#maximizar janela
pyautogui.hotkey('winleft','Up')
#selecionar PDF
pyautogui.moveTo(377, 130,duration=1)
pyautogui.click(377, 130)
pyautogui.hotkey('ctrl','a')
pyautogui.mouseDown(1101,120)
pyautogui.moveTo(220,694,duration=1)
pyautogui.mouseUp(220,694)
#arrastar arquivo
pyautogui.moveTo(390,135,duration=1)
pyautogui.mouseDown(390,135)
pyautogui.moveTo(688,564,duration=1)
pyautogui.hotkey('alt','tab')
pyautogui.mouseUp()
#esperar fazer conversão
time.sleep(80)
#clicar fora
pyautogui.moveTo(1132,245,duration=1)
pyautogui.click(1132,245)
#clicar em download
pyautogui.moveTo(649,505,duration=1)
pyautogui.click(649,505)
#ir para área de trabalho
pyautogui.hotkey('winleft','d')
#clicar na pasta unir 500...
pyautogui.moveTo(1313, 430,duration=1)
pyautogui.doubleClick(1313, 430)
#maximizar janela
pyautogui.hotkey('winleft','up')
#clicar dentro da pasta 
pyautogui.moveTo(845,494,duration=1)
pyautogui.click(845,494)
#clicar no arquivo planilhas a unir
pyautogui.moveTo(254,134,duration=1)
pyautogui.doubleClick(254,134)
#clicar na planilha
pyautogui.moveTo(254,134,duration=1)
pyautogui.doubleClick(254,134)
#habiilitar conteúdo
time.sleep(5)
pyautogui.moveTo(435,163,duration=1)
pyautogui.doubleClick(435,163)
#mouse em cima da internet
pyautogui.moveTo(711,754,duration=1)
# clicar no site de conversão
pyautogui.moveTo(808,668,duration=1)
pyautogui.click(808,668)
#clicar na planilha baixada para abrir
pyautogui.moveTo(105,698,duration=1)
pyautogui.doubleClick(105,698)
time.sleep(5)
pyautogui.hotkey('winleft','up')
#selecionar células para copiar e copiar
pyautogui.moveTo(82, 164,duration=1)
pyautogui.click(82, 164)
pyautogui.hotkey('ctrl','t')
pyautogui.hotkey('ctrl','c')
#abrir arquivo planilhas a unir para colar
pyautogui.moveTo(749,743,duration=1)
pyautogui.moveTo(659,674,duration=1)
pyautogui.click(659,674)
#colar células no arquivo planilhas a unir
pyautogui.hotkey('winleft','up')
pyautogui.hotkey('ctrl','up')
pyautogui.moveTo(54,222,duration=1)
pyautogui.click(54,222)
#colar 
pyautogui.hotkey('ctrl','v')
#selecionar célula de fora
pyautogui.moveTo(1006,304,duration=1)
pyautogui.click(1006,304)
#colocar mouse sobre instruções
##pyautogui.moveTo(820,241,duration=1)
##pyautogui.click(820,241)
#preencher nome no formulário
##pyautogui.write('Supermen')
##pyautogui.moveTo(817,238,duration=1)
##pyautogui.click(817,238)
##time.sleep(35)
#colocar mouse sobre executar
pyautogui.moveTo(847,269,duration=1) 

pyautogui.alert('Finalizado!')

----------------------------------------------------------------------------------------
import pyautogui
pyautogui.position()
----------------------------------------------------------------------------------------
#planilhar empenhos

import pyautogui
import time

pyautogui.alert("Irá começar")
pyautogui.PAUSE=1

pyautogui.hotkey('winleft','m')
pyautogui.moveTo(760, 753,duration=1)
pyautogui.moveTo(618, 640,duration=1)
pyautogui.click(618, 640)
pyautogui.hotkey('ctrl','b')
pyautogui.hotkey('alt','f4')
pyautogui.hotkey('winleft','d')

#clicar pasta UNIR 500-550 teste
pyautogui.moveTo(1304,430)
pyautogui.doubleClick(1304,430)

#maximizar janela
pyautogui.hotkey('winleft','up')
#clicar na planilha "Planilha teste unificar"
pyautogui.moveTo(254,151)
pyautogui.doubleClick(254,151)
time.sleep(8)
#clicar na macro do VBA
pyautogui.moveTo(188,441)
pyautogui.click(188,441)
pyautogui.hotkey('ctrl','up')
pyautogui.moveTo(179,222)
pyautogui.click(179,222)
pyautogui.alert("Macro realizada com sucesso!")
