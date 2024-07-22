from flask import Flask, request, render_template, redirect, url_for
from werkzeug.utils import secure_filename
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import urllib.parse
import time
import pyautogui as tempoEspera
from openpyxl import load_workbook

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return redirect(request.url)
    file = request.files['file']
    if file.filename == '':
        return redirect(request.url)
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        return redirect(url_for('send_messages', filepath=filepath))
    return 'Invalid file format'

@app.route('/send_messages')
def send_messages():
    filepath = request.args.get('filepath')
    planilhaDadosContato = load_workbook(filepath)
    sheet_selecionada = planilhaDadosContato['Dados']
    navegadorChrome = webdriver.Chrome()

    navegadorChrome.get("https://web.whatsapp.com/")
    WebDriverWait(navegadorChrome, 60).until(EC.presence_of_element_located((By.ID, 'side')))

    for linha in range(2, len(sheet_selecionada['A']) + 1):
        numeroContato = sheet_selecionada['A%s' % linha].value
        mensagemContato = sheet_selecionada['B%s' % linha].value

        texto = urllib.parse.quote(f"{mensagemContato}")
        link = f"https://web.whatsapp.com/send?phone={numeroContato}&text={texto}"
        navegadorChrome.get(link)

        WebDriverWait(navegadorChrome, 60).until(EC.presence_of_element_located((By.ID, 'side')))
        tempoEspera.sleep(2)

        alerta_presente = len(navegadorChrome.find_elements(By.XPATH, '//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[1]')) > 0
        tempoEspera.sleep(2)

        if not alerta_presente:
            tempoEspera.press('enter')
            sheet_selecionada['C%s' % linha].value = 'Ok'
            tempoEspera.sleep(2)
        else:
            sheet_selecionada['C%s' % linha].value = 'Número inválido'
            tempoEspera.sleep(2)

    planilhaDadosContato.save(filepath)
    navegadorChrome.quit()
    return 'Messages sent and planilha updated!'

if __name__ == '__main__':
    app.run(debug=True)
