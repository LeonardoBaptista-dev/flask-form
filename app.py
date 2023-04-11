import os
import csv
from flask import Flask, render_template, request, redirect, url_for, send_from_directory
from openpyxl import load_workbook

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        name = request.form['name']
        email = request.form['email']
        message = request.form['message']

        # Abre o arquivo CSV, cria se não existir, e adiciona a nova linha
        filename_csv = 'data.csv'
        if os.path.isfile(os.path.join(app.config['UPLOAD_FOLDER'], filename_csv)):
            mode = 'a'
        else:
            mode = 'w'
        with open(os.path.join(app.config['UPLOAD_FOLDER'], filename_csv), mode, newline='') as file:
            writer = csv.writer(file)
            if mode == 'w':
                writer.writerow(['Name', 'Email', 'Message'])
            writer.writerow([name, email, message])

        # Redireciona para a página de sucesso
        return redirect(url_for('success'))

    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    name = request.form['name']
    email = request.form['email']
    message = request.form['message']

    # Abre o arquivo CSV, cria se não existir, e adiciona a nova linha
    filename_csv = 'data.csv'
    if os.path.isfile(os.path.join(app.config['UPLOAD_FOLDER'], filename_csv)):
        mode = 'a'
    else:
        mode = 'w'
    with open(os.path.join(app.config['UPLOAD_FOLDER'], filename_csv), mode, newline='') as file:
        writer = csv.writer(file)
        if mode == 'w':
            writer.writerow(['Name', 'Email', 'Message'])
        writer.writerow([name, email, message])

    # Redireciona para a página de sucesso
    return redirect(url_for('success'))

@app.route('/success')
def success():
    return "Formulário enviado com sucesso!"

@app.route('/download')
def download():
    filename_csv = 'data.csv'
    return send_from_directory(directory=app.config['UPLOAD_FOLDER'], filename=filename_csv, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
