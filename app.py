# -*- coding: utf-8 -*-

import os
from flask import Flask, render_template, request, send_from_directory

# Cria a pasta de uploads se ela não existir
if not os.path.exists("uploads"):
    os.makedirs("uploads")

app = Flask(__name__)

@app.route("/")
def index():
    """Renderiza a página inicial com o formulário de upload."""
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload():
    """Recebe os arquivos, salva e retorna uma mensagem de sucesso."""
    cdr_file = request.files["cdr_file"]
    product_list = request.form["product_list"]

    # Salva o arquivo .cdr
    cdr_filename = cdr_file.filename
    cdr_file.save(os.path.join("uploads", cdr_filename))

    # Salva a lista de produtos em um arquivo de texto
    list_filename = f"{os.path.splitext(cdr_filename)[0]}_lista.txt"
    with open(os.path.join("uploads", list_filename), "w", encoding="utf-8") as f:
        f.write(product_list)

    return f"""
    <h1>Arquivos Recebidos com Sucesso!</h1>
    <p>Arquivo CDR: {cdr_filename}</p>
    <p>Lista de Produtos salva como: {list_filename}</p>
    <p>O processamento começará em breve.</p>
    """

if __name__ == "__main__":
    from waitress import serve
    serve(app, host="0.0.0.0", port=8080)
