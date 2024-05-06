import os
from flask import Flask, render_template, request, jsonify, session
import subprocess
import json
from flask import request
from flask import jsonify
import re
import logging

app = Flask(__name__, template_folder='templates')
app.secret_key = os.urandom(24)

# Definição do filtro parse_json
def parse_json(data):
    return json.loads(data)

# Adição do filtro parse_json ao ambiente Jinja2
app.jinja_env.filters['parse_json'] = parse_json

# Rota para servir o index.html
@app.route('/')
def main():
    return render_template('index.html')

# Função para formatar o resultado e inserir no Flask
def format_results(results):
    formatted_results = []
    header_mapping = {
        'PSComputerName': 'Servidor',
        'Name': 'Nome do Serviço',
        'State': 'Estado do Serviço'
    }

    formatted_results.append("{:<15} | {:<40} | {:<15}".format(header_mapping['PSComputerName'], header_mapping['Name'], header_mapping['State']))
    formatted_results.append("-" * 75)

    for result in results:
        lines = result.strip().split('\n')
        for line in lines[2:]:
            line = line.strip()
            if line:
                parts = line.split()
                formatted_line = "{:<15} | {:<40} | {:<15}".format(parts[0], ' '.join(parts[1:-1]), parts[-1])
                formatted_results.append(formatted_line)

    return formatted_results

# Função para realizar a primeira consulta
@app.route('/consultar', methods=['POST'])
def consultar():
    session['nome_cliente'] = request.form['nome_cliente']

    resultados = []
    for computador in servidores:
        comando = f'Get-WmiObject -Class Win32_Service -ComputerName {computador} | Where-Object {{ $_.PathName -like "*{session["nome_cliente"]}*" }} | Where-Object {{ $_.Name -like "*motor*" }} | Select-Object PSComputerName, Name, State'

        resultado_json = subprocess.check_output(['powershell', comando], text=True)
        
        resultados.append(resultado_json)

    # Formatar os resultados em linhas
    formatted_results = format_results(resultados)

    return jsonify(resultados=formatted_results)

# Função para realizar a segunda consulta
@app.route('/consultar_cliente_ambiente', methods=['POST'])
def consultar_cliente_ambiente():
    session['codigo_cliente'] = request.form['codigo_cliente']  # Adicionando à sessão
    tipo_ambiente = request.form['tipo_ambiente']

    if tipo_ambiente == "_p":  
        ambiente = "p"
    else:
        ambiente = "h"
    
    session['ambiente'] = ambiente  # Adicionando à sessão

    nome_cliente_formatado = f"{session['nome_cliente']}_{session['codigo_cliente']}_{ambiente}"

    resultados = []
    servicos = set()  # Usando um conjunto para evitar serviços duplicados
    for computador in servidores:
        comando = f'Get-WmiObject -Class Win32_Service -ComputerName {computador} | Where-Object {{ $_.PathName -like "*{nome_cliente_formatado}*" }} | Select-Object PSComputerName, Name, State'

        resultado_json = subprocess.check_output(['powershell', comando], text=True)
        resultados.append(resultado_json)

        servicos_computador = [line.strip() for line in resultado_json.strip().split('\n')[1:] if line.strip() and not line.startswith(('PSComputerName', '--------------'))]
        servicos.update(servicos_computador)
    
    # Formatar os resultados em linhas
    formatted_results = format_results(resultados)

    # Enviar apenas os nomes dos serviços para o frontend
    return jsonify(resultados=formatted_results, servicos=[{'Name': servico} for servico in servicos])

# Função para iniciar o(s) serviço(s)
@app.route('/iniciar_servico', methods=['POST'])
def iniciar_servico():
    tipo_ambiente = request.form.get('tipo_ambiente')

    if tipo_ambiente == "_p":  
        ambiente = "p"
    else:
        ambiente = "h"

    nome_cliente_formatado = f"{session['nome_cliente']}_{session['codigo_cliente']}_{session['ambiente']}"
    servicos_selecionados = request.form.getlist('servico_selecionado[]')

    for servico_selecionado in servicos_selecionados:
        if servico_selecionado.lower() == "todos":  # Verifica se "Todos" foi selecionado
            for computador in servidores:
                # Comando para parar todos os serviços
                comando_iniciar = f'(Get-WMIObject win32_service -ComputerName {computador} | Where-Object {{ $_.PathName -like "*{nome_cliente_formatado}*" }}).StartService()'
                # Executar o comando para parar todos os serviços
                subprocess.run(['powershell', comando_iniciar], text=True)

        else:
            # Expressão regular para extrair o nome do serviço com base nos critérios fornecidos
            servico_regex = r'(Motor|SeniorInst|CSM|HCMIntegrator|Wiipo|Concentradora|SAM|IntegrationBack|Middleware)\S*'
            # Extrair o nome do serviço usando a expressão regular
            match = re.search(servico_regex, servico_selecionado)
            if match:
                nome_servico = match.group()
            
            for computador in servidores:
                # Comando para parar o serviço específico
                comando_iniciar = f'(Get-WMIObject win32_service -ComputerName {computador} | Where-Object {{ $_.PathName -like "*{nome_cliente_formatado}*" -and $_.Name -eq "{nome_servico}" }}).StartService()'
                # Executar o comando para parar o serviço específico
                subprocess.run(['powershell', comando_iniciar], text=True)

    return 'Serviço(s) parado(s) com sucesso', 200  # OK

# Função para parar o(s) serviço(s)
@app.route('/parar_servico', methods=['POST'])
def parar_servico():
    tipo_ambiente = request.form.get('tipo_ambiente')

    if tipo_ambiente == "_p":  
        ambiente = "p"
    else:
        ambiente = "h"

    nome_cliente_formatado = f"{session['nome_cliente']}_{session['codigo_cliente']}_{session['ambiente']}"
    servicos_selecionados = request.form.getlist('servico_selecionado[]')

    for servico_selecionado in servicos_selecionados:
        if servico_selecionado.lower() == "todos":  # Verifica se "Todos" foi selecionado
            for computador in servidores:
                # Comando para parar todos os serviços
                comando_parar = f'(Get-WMIObject win32_service -ComputerName {computador} | Where-Object {{ $_.PathName -like "*{nome_cliente_formatado}*" }}).StopService()'
                # Executar o comando para parar todos os serviços
                subprocess.run(['powershell', comando_parar], text=True)
        else:
            # Expressão regular para extrair o nome do serviço com base nos critérios fornecidos
            servico_regex = r'(Motor|SeniorInst|CSM|HCMIntegrator|Wiipo|Concentradora|SAM|IntegrationBack|Middleware)\S*'
            # Extrair o nome do serviço usando a expressão regular
            match = re.search(servico_regex, servico_selecionado)
            if match:
                nome_servico = match.group()
            
            for computador in servidores:
                # Comando para parar o serviço específico
                comando_parar = f'(Get-WMIObject win32_service -ComputerName {computador} | Where-Object {{ $_.PathName -like "*{nome_cliente_formatado}*" -and $_.Name -eq "{nome_servico}" }}).StopService()'
                # Executar o comando para parar o serviço específico
                subprocess.run(['powershell', comando_parar], text=True)

    return 'Serviço(s) parado(s) com sucesso', 200  # OK

servidores = ["OCSENAPLH01", "OCSENAPL01", "OCSENAPL02", "OCSENAPL03", "OCSENAPL04", "OCSENINT01", "OCSENMDW01"] # Adicione mais se necessário

if __name__ == '__main__':
    logging.basicConfig(filename='app.log', level=logging.DEBUG)
    app.run(host='OCSENAPLH01', port=5000, debug=True)
