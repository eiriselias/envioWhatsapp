import time
import pyautogui as pg
import sys
import os
from tkinter import messagebox

lista_errores = []

def ruta_recurso(relative_path):
    try:
        # Si es un .exe, busca en la carpeta temporal _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        # Si es el script .py, busca en la carpeta actual
        base_path = os.path.abspath(".")
    
    # Unimos la base con el nombre del archivo
    return os.path.join(base_path, relative_path)


def envio(numero, mensaje):
    encontrado = False  
    inicio = time.time()

    ruta_imagen = ruta_recurso('enviar.png')

    print(f'Buscando boton para el {numero}...')

    while not encontrado and (time.time() - inicio < 35):
        try:
            btn = pg.locateOnScreen(ruta_imagen, confidence=0.8, grayscale=True)
            
            if btn is not None:
                pg.click(btn)
                time.sleep(1.0) 
                pg.press('enter')
                print(f'✅ Mensaje enviado a {numero}')
                
                encontrado = True
                
                time.sleep(2) 
                
                pg.hotkey('ctrl', 'w') 
                time.sleep(1.0)
                pg.press('enter')
                
        except (pg.ImageNotFoundException, Exception):
            messagebox.showwarning("Error", "No se encontró el botón de enviar mensaje.")
            time.sleep(2) 
            
    if not encontrado:
        print(f'❌ Tiempo agotado para {numero}. ¿WhatsApp cargó correctamente?')
        lista_errores.append(numero)