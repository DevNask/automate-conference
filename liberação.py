import pyautogui as bot
from pyautogui import FAILSAFE
import pyperclip
import pandas as pd 
from time import sleep
import re

bot.FAILSAFE = True

arquivo = r"C:\Users\Detran\Desktop\gbnf\robo liberacao (meu)\pagamento.xlsx"

df = pd.read_excel(arquivo)
resultado = []

for renavam in df['renavam']:
    
    renavam_formatado = '000' + str(renavam)
    sleep(1)

    # Passo 1 : Pesquisa tela do PEPM com Chassi
    bot.moveTo(37,76, duration=0.5)
    bot.click()
    sleep(1)
    bot.write('pren',interval=0.7)
    pyperclip.copy(str(renavam))
    bot.hotkey('ctrl','v', interval=0.7)
    bot.press('enter')

    sleep(1.5)

    # Passo 2 : Copiar a Área de Bloqueio na Tela do PEPM
    bot.moveTo(101,411,duration=0.7)
    bot.mouseDown()
    bot.moveTo(765,444,duration=0.7)
    bot.mouseUp()
    
    sleep(0.5)

    bot.hotkey('ctrl','c', interval=0.7)
    bloqueio = pyperclip.paste().upper()

    bloqueios_indesejados = ['JUDICIAL','RENAJUD-CIRCULACAO','RENAJUD-PENHORA','RENAJUD-TRANSFERENCIA','QUEIXA DE FURTO','QUEIXA DE ROUBO']
    bloqueios_encontrados = [b for b in bloqueios_indesejados if b in bloqueio]

    if bloqueios_encontrados:
        print(f'Veículo com restrição : {bloqueios_encontrados}')
        continue

    sleep(1.5)

    # Passo 3 : Conferir Resultado da Tela de Pagamento
    bot.moveTo(37,76, duration=0.5)
    bot.click()
    bot.write('txut', interval=0.7)
    pyperclip.copy(str(renavam))
    bot.hotkey('ctrl','v', interval=0.7)
    bot.press('enter')

    sleep(1.5)

    # Copia conteúdo da tela
    bot.moveTo(17,782, duration=0.7)
    bot.mouseDown()
    bot.moveTo(610,784, duration=0.7)
    bot.mouseUp()
    bot.hotkey('ctrl', 'C', interval=0.7)

    # Transformar o texto em formato limpo
    registro = pyperclip.paste().strip().upper()

    # Listas de condições
    registros_indesejados = [
        'NENHUM REGISTRO ENCONTRADO PARA O PARAMETR',
        'PARAM.NAO INF.-INFORME RENAVAM OU CPFCNPJ OU PLACA OU '
    ]
    registros_possiveis = [
        'EXISTEM MAIS REGISTROS PARA SEREM EXIBIDOS'
    ]
    registros_desejados = [
        'PESQUISA CONCLUIDA COM SUCESSO'
    ]

    if any(texto in registro for texto in registros_indesejados):
        print(f'Renavam {renavam}: Nenhum registro encontrado.')
        bot.moveTo(17,782, duration=0.7)
        bot.mouseDown()
        bot.moveTo(610,784, duration=0.7)
        bot.mouseUp()
        bot.hotkey('ctrl','c')
        sleep(0.5)
    
        texto_copiado = pyperclip.paste().strip().upper()

        # Adiciona o resultado na planilha com a mensagem copiada como "taxa"
        resultado.append({
            'renavam': renavam,
            'taxa': texto_copiado,  # Agora salvando o texto real da tela
            'status': 'Sem Registro'
        })
    
    else:
        if any(texto in registro for texto in registros_possiveis):
            print("Registros Possíveis Detectado, copiando mensagem")
            bot.press('enter')
            bot.moveTo(21,780, duration=0.7)
            bot.mouseDown()
            bot.moveTo(435,785, duration=0.7)
            bot.mouseUp()
            bot.hotkey('ctrl', 'c')
            registro = pyperclip.paste().strip().upper()

        if any(texto in registro for texto in registros_desejados):
            print("Pesquisa concluída com sucesso.")
            bot.moveTo(318,470, duration=0.7)
            bot.mouseDown()
            bot.moveTo(365,472, duration=0.7)
            bot.mouseUp()
            bot.hotkey('ctrl', 'c')

            taxa = pyperclip.paste().strip()
            print(f'Valor da taxa copiado: {taxa}')
            try:
                taxa_num = int(taxa)
            except ValueError:
                taxa_num = None

            if taxa_num == 2006:
                print(f'Taxa de R$2006 paga para renavam {renavam}')
                resultado.append({'renavam': renavam, 'taxa': taxa, 'status': 'PAGO'})
            else:
                print(f'Valor diferente para renavam {renavam}: {taxa}')
                resultado.append({'renavam': renavam, 'taxa': taxa, 'status': 'VALOR DIFERENTE'})

    # SALVAR RESULTADO A CADA LAÇO
    df_resultado = pd.DataFrame(resultado)
    df_resultado.to_excel(r'C:\Users\Detran\Desktop\gbnf\robo liberacao (meu)\resultado_geral.xlsx', index=False)