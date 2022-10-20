import speech_recognition as sr
import pyttsx3
import datetime
import pywhatkit
import pyautogui
import pyperclip
import pandas as pd
import openpyxl
import numpy
import time
import tweepy
import requests
import json


audio = sr.Recognizer()
maquina = pyttsx3.init()
pyautogui.PAUSE = 2



def automation_email():  

    link = "###"
   
    pyautogui.press("winleft")
    pyautogui.write("opera")
    pyautogui.press("enter")

    pyperclip.copy(link)

    pyautogui.hotkey('ctrl','v')
    pyautogui.press("enter")
    time.sleep(5)
    pyautogui.click(x=334, y=298, clicks=2, button='left')
    pyautogui.click(x=334, y=298, clicks=2, button='right')
    pyautogui.click(x=416, y=745)
    time.sleep(10)

    df = pd.read_excel(r'###')

    faturamento = df['Valor Final'].sum()
    quantidade = df['Quantidade'].sum()

    pyautogui.press("winleft")
    pyautogui.write("email")
    pyautogui.press("enter")
    pyautogui.click(x=104, y=100, button='left')
    pyautogui.write("###")
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')

    assunto = "Relatório de Vendas de Ontem"
    pyperclip.copy(assunto)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.press('tab')
    texto = f"""
    Bom dia, tudo bom?

    O faturamento de ontem foi de: R$ {faturamento:,.2f}
    A quantidade de produtos foi de: {quantidade:,} 

    Obrigado, até mais.
    """

    pyperclip.copy(texto)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.hotkey('ctrl', 'enter')

def twitar_dolar():

    api = tweepy.Client(
    consumer_key='###',
    consumer_secret='###',
    access_token='###',
    access_token_secret='###')

    running = True
    seconds = 1
    end = 0
    while(running):
        seconds -=1
        if(seconds <= end):
            try:
                req = requests.get ('https://economia.awesomeapi.com.br/all/USD-BRL')
                cotacao = req.json()
                cotacao_print = ("#### Cotação do Dolar ####" + "\n" + "Moeda:"  + cotacao["USD"]["name"] + "\n" + "Data: " + cotacao["USD"]["create_date"] + "\n" + "Valor atual: R$" + cotacao["USD"]["bid"] + "\n" + "Maior Valor: R$" + cotacao["USD"]["high"] + "\n" + "Menor Valor: R$" + cotacao["USD"]["low"])

                tweet = api.create_tweet(text=cotacao_print)
                print(tweet)
                time.sleep(10)
                
            except Exception as error:
                print("Erro:", error)
    
   
def executa_comando():
    try:
        with sr.Microphone() as source:
            print('Ouvindo...')
            voz = audio.listen(source)
            comando = audio.recognize_google(voz, language='pt-BR')
            comando = comando.lower()
            if 'computador' in comando:
                comando = comando.replace('computador', '')
                maquina.say(comando)
                maquina.runAndWait()

    except:
        print('Microfone não está ok')

    return comando

def comando_voz_usuario():
    comando = executa_comando()

    if 'que horas são' in comando:
        hora = datetime.datetime.now().strftime('%H:%M')
        maquina.say('Agora são' + hora)
        maquina.runAndWait()

    elif 'toque' in comando:
        musica = comando.replace('toque','')
        resultado = pywhatkit.playonyt(musica)
        maquina.say('Tocando música')
        maquina.runAndWait()

    elif 'abra o navegador' in comando:
        abra_navegador = comando.replace('abra o navegador', '')
        resultado = pyautogui.press("winleft"), pyautogui.write("opera"), pyautogui.press("enter")
        maquina.say('Abrindo o navegador')
        maquina.runAndWait()
    
    elif 'abra o youtube' in comando:
        abra_navegador_youtube = comando.replace('abra o youtube', '')
        resultado = pyautogui.press("winleft"), pyautogui.write("opera"), pyautogui.press("enter"), pyautogui.write("youtube"), pyautogui.press("enter")
        maquina.say('Abrindo o navegador e o youtube')
        maquina.runAndWait()
    
    elif 'enviar' in comando:
        enviar_email = comando.replace('enviar', '')
        resultado = automation_email()
        maquina.say('Automatizando email')
        maquina.runAndWait()
    
    elif 'publicar' in comando:
        enviar_tweet = comando.replace('publicar', '')
        resultado = twitar_dolar()
        maquina.say('Automatizando tweets')
        maquina.runAndWait()
        

comando_voz_usuario() 