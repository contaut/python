from flask import Flask
from flask import request
import subprocess
# -*- coding: utf-8 -*-
   

app = Flask(__name__)


@app.route("/")

def init():
    login = request.args.get('login')
    senha = request.args.get('senha') 
    nome = request.args.get('nome') 
    f= open("dados.txt","w+")
    f.write(login)
    f.write('\n')
    f.write(senha)
    f.write('\n')
    f.write(nome)
    f.close()
    params = 'contaut_worker.py'
    command = ('python', params)
    subprocess.Popen(command)
    return "Hello World! <strong>I am learning Flask</strong>"

@app.route("/documents")
def retrieve():
    
    return "documents"
app.run()
