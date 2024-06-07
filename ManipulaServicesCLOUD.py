import subprocess
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk

# Função para exibir mensagem de instrução
def exibir_mensagem():
    messagebox.showinfo("Instrução", 'Selecione o ambiente desejado (Produção ou Homologação) e clique em "Consultar".\n\nNa barra inferior, selecione o serviço desejado ou selecione "Todos" para executar o comando em todos os serviços apresentados.')

# Função para realizar a consulta e exibir resultados
def primeira_consulta():
    # Definir o nome do cliente fixo
    nome_cliente = "CLOUD_99999"

    # Obter os dados do formulário
    tipo_ambiente = var_tipo_ambiente.get()

    # Formatando os dados
    if tipo_ambiente == "Produção":
        ambiente = "p"
    else:
        ambiente = "h"

    nome_cliente_formatado = f"{nome_cliente}_{ambiente}".lower()

    # Limpar a tabela anterior
    limpar_tabela()

    # Lista para armazenar todos os serviços
    all_services = []

    for computador in computadores:
        # Comando para verificar o serviço
        comando = f'Get-WmiObject -Class Win32_Service -ComputerName {computador} | Where-Object {{ $_.PathName -like "*{nome_cliente_formatado}*" }} | Format-Table Name, State'

        # Executa o comando no PowerShell e captura a saída
        resultado = subprocess.check_output(['powershell', comando], text=True)

        # Inserir resultado na tabela
        inserir_na_tabela(resultado)

        # Adicionar os serviços da consulta atual à lista de todos os serviços
        lines = resultado.strip().split('\n')
        services = [' '.join(line.split()[:-1]) for line in lines if len(line.split()) >= 2]  # Capturar o nome completo do serviço
        all_services.extend(services)

    # Preencher o Combobox com todos os serviços
    preencher_combobox(all_services)

    # Atualizar a janela para exibir o Combobox e os botões
    combobox_servicos.grid(row=4, column=1, padx=5, pady=5)
    botao_iniciar.grid(row=4, column=0, padx=5, pady=5)
    botao_parar.grid(row=4, column=2, padx=5, pady=5)

# Função para preencher o Combobox com os nomes reais dos serviços
def preencher_combobox(services):
    services = [service for service in services if service != "Name"]  # Excluir a linha "Name" se presente
    combobox_servicos['values'] = ["Todos"] + services
    combobox_servicos.config(width=40)  # Definir a largura do combobox

# Função para limpar a tabela
def limpar_tabela():
    for item in treeview.get_children():
        treeview.delete(item)

# Função para inserir resultado na tabela
def inserir_na_tabela(resultado):
    lines = resultado.strip().split('\n')
    for line in lines[2:]:  # Começa a partir da 3ª linha para ignorar os cabeçalhos
        data = line.split()
        # Verificar se a linha contém os valores esperados
        if len(data) >= 2:
            # Capturar Name e State
            name = ' '.join(data[:-1])  # Juntar todos os elementos, exceto o último
            state = data[-1]  # Último elemento é o estado
            # Substituir "Running" por "Em execução" e "Stopped" por "Parado"
            state = "Em Execução" if state == "Running" else "Parado" if state == "Stopped" else state
            treeview.insert('', 'end', values=(name, state))

# Função para obter o serviço selecionado no Combobox
def obter_servico_selecionado():
    selected_text = combobox_servicos.get()
    return selected_text

# Função para INICIAR os serviços
def iniciar_servicos():
    # Definir o nome do cliente fixo
    nome_cliente = "CLOUD_99999"

    # Obter os dados do formulário
    tipo_ambiente = var_tipo_ambiente.get()

    servico_selecionado = obter_servico_selecionado()

    # Formatando os dados
    if tipo_ambiente == "Produção":
        ambiente = "p"
    else:
        ambiente = "h"

    nome_cliente_formatado = f"{nome_cliente}_{ambiente}".lower()

    # Loop sobre os computadores
    for computador in computadores:
        # Comando para iniciar o serviço
        if servico_selecionado == "Todos":
            comando_iniciar = f'Get-WMIObject win32_service -ComputerName {computador} | Where-Object {{ $_.PathName -like "*{nome_cliente_formatado}*" }} | ForEach-Object {{ $_.StartService() }}'
        else:
            comando_iniciar = f'(Get-WMIObject win32_service -ComputerName {computador} | Where-Object {{ $_.PathName -like "*{nome_cliente_formatado}*" -and $_.Name -like "*{servico_selecionado}*" }}).StartService()'

        print(comando_iniciar)

        # Executa o comando no PowerShell para iniciar o serviço
        resultado_iniciar = subprocess.run(['powershell', '-Command', comando_iniciar])

    # Atualizar a tabela com o novo estado do serviço
    primeira_consulta()

# Função para PARAR os serviços
def parar_servicos():
    # Definir o nome do cliente fixo
    nome_cliente = "CLOUD_99999"

    # Obter os dados do formulário
    tipo_ambiente = var_tipo_ambiente.get()

    servico_selecionado = obter_servico_selecionado()

    # Formatando os dados
    if tipo_ambiente == "Produção":
        ambiente = "p"
    else:
        ambiente = "h"

    nome_cliente_formatado = f"{nome_cliente}_{ambiente}".lower()

    # Loop sobre os computadores
    for computador in computadores:
        # Comando para iniciar o serviço
        if servico_selecionado == "Todos":
            comando_parar= f'Get-WMIObject win32_service -ComputerName {computador} | Where-Object {{ $_.PathName -like "*{nome_cliente_formatado}*" }} | ForEach-Object {{ $_.StopService() }}'
        else:
            comando_parar = f'(Get-WMIObject win32_service -ComputerName {computador} | Where-Object {{ $_.PathName -like "*{nome_cliente_formatado}*" -and $_.Name -like "*{servico_selecionado}*" }}).StopService()'

        print(comando_parar)

        # Executa o comando no PowerShell para iniciar o serviço
        resultado_iniciar = subprocess.run(['powershell', '-Command', comando_parar])

    # Atualizar a tabela com o novo estado do serviço
    primeira_consulta()

# Criar a janela
janela = tk.Tk()
janela.title("Manipula Serviços HCM SaaS Orion")

# Criar os rótulos e campos de entrada
label_tipo_ambiente = tk.Label(janela, text='Selecione o Ambiente:')
label_tipo_ambiente.grid(row=1, column=0, padx=5, pady=5)

var_tipo_ambiente = tk.StringVar(janela)
var_tipo_ambiente.set("Homologação")  # Valor padrão

opcao_producao = tk.Radiobutton(janela, text="Produção", variable=var_tipo_ambiente, value="Produção")
opcao_producao.grid(row=1, column=1, padx=5, pady=5)

opcao_homologacao = tk.Radiobutton(janela, text="Homologação", variable=var_tipo_ambiente, value="Homologação")
opcao_homologacao.grid(row=1, column=2, padx=5, pady=5)

# Criar estilo para a tabela
style = ttk.Style()
style.configure("Treeview", rowheight=40)  # Ajuste o valor de rowheight conforme necessário

# Adicionar widget Treeview para mostrar os resultados
treeview = ttk.Treeview(janela, columns=('Name', 'State'), show='headings')
treeview.heading('Name', text='Nome do Serviço')
treeview.heading('State', text='Estado')
treeview.grid(row=20, column=0, columnspan=3, padx=100, pady=100)  # Ajuste para ocupar mais colunas

# Definir a largura das colunas
treeview.column('Name', width=400)  # Ajuste o valor de width conforme necessário
treeview.column('State', width=200)  # Ajuste o valor de width conforme necessário

treeview.grid(row=5, column=0, columnspan=3, padx=5, pady=5)

# Botão de mensagem de instrução
botao_instrucao = tk.Button(janela, text="Instrução", command=exibir_mensagem)
botao_instrucao.grid(row=3, column=2, padx=5, pady=5)

# Botão para consultar os serviços
botao_consultar = tk.Button(janela, text="Consultar Serviços", command=primeira_consulta)
botao_consultar.grid(row=3, column=0, padx=5, pady=5)

# Botão para iniciar os serviços
botao_iniciar = tk.Button(janela, text="Iniciar Serviço(s)", command=iniciar_servicos)

# Botão para parar os serviços
botao_parar = tk.Button(janela, text="Parar Serviço(s)", command=parar_servicos)

# Combobox para selecionar o serviço
combobox_servicos = ttk.Combobox(janela, state="readonly")

# Computadores para consulta
computadores = ["nb021505"]  # Adicione mais se necessário

janela.mainloop()
