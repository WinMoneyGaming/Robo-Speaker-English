import win32com.client as wincom

print('Welcome to Robo Speaker v1.1 Created By Ravinder Verma.')
speak = wincom.Dispatch("SAPI.SpVoice")
while True:
    x = input('Enter what you want to speak: ')
    if x == 'q':
        speak.Speak('Bye Bye My Friend!')
        print('Bye Bye My Friend!')
        break
    speak.Speak(x)
