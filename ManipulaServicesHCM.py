import subprocess
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk

# Função para exibir mensagem de instrução
def exibir_mensagem():
    messagebox.showinfo("Instrução", 'Digite o nome do cliente e clique em "Consultar", será consultado o Serviço do Motor e-Social.\n\nConfirme o código HCM do cliente, digite o código no campo que será habilitado e escolha o tipo de ambiente, Produção ou Homologação, em seguida, clique em "Consultar" novamente.\n\nNa barra inferior, selecione o serviço desejado ou selecione "Todos" para executar o comando em todos os serviços apresentados.')

# Função para realizar a primeira consulta e exibir resultados
def primeira_consulta():
    # Obter os dados do formulário
    nome_cliente = entry_nome_cliente.get()

    # Limpar a tabela anterior
    limpar_tabela()

    for computador in computadores:
        # Comando para verificar o serviço
        comando = f'Get-WmiObject -Class Win32_Service -ComputerName {computador} | Where-Object {{ $_.PathName -like "*{nome_cliente}*" }} | Where-Object {{ $_.Name -like "*motor*" }} | Format-Table PSComputerName, Name, State'

        # Executa o comando no PowerShell e captura a saída
        resultado = subprocess.check_output(['powershell', comando], text=True)

        # Inserir resultado na tabela
        inserir_na_tabela(resultado)

    # Habilitar o campo para digitar o código HCM
    label_codigo_hcm.grid(row=2, column=0, padx=5, pady=5)
    entry_codigo_hcm.grid(row=2, column=1, padx=5, pady=5)
    entry_codigo_hcm.config(state=tk.NORMAL)  # Habilitar o campo
    botao_consultar.config(command=segunda_consulta)  # Configurar o botão para executar a segunda consulta

# Função para limpar a tabela
def limpar_tabela():
    for item in treeview.get_children():
        treeview.delete(item)

# Função para inserir resultado na tabela
def inserir_na_tabela(resultado):
    lines = resultado.strip().split('\n')
    for line in lines:
        data = line.split()
        # Verificar se a linha contém os valores esperados
        if len(data) >= 3:
            # Capturar PSComputerName, Name e State
            ps_computer_name = data[0]
            name = data[1]
            # Encontrar o campo 'State'
            state = data[-1]
            treeview.insert('', 'end', values=(ps_computer_name, name, state))

# Função para realizar a segunda consulta e exibir resultados
def segunda_consulta():
    # Obter o código HCM
    codigo_hcm = entry_codigo_hcm.get()

    # Obter os dados do formulário
    nome_cliente = entry_nome_cliente.get()
    tipo_ambiente = var_tipo_ambiente.get()

    # Formatando os dados
    if tipo_ambiente == "Produção":
        ambiente = "p"
    else:
        ambiente = "h"

    nome_cliente_formatado = f"{nome_cliente}_{codigo_hcm}_{ambiente}".lower()

    # Limpar a tabela anterior
    limpar_tabela()

    # Lista para armazenar todos os serviços
    all_services = []

    for computador in computadores:
        # Comando para verificar o serviço
        comando = f'Get-WmiObject -Class Win32_Service -ComputerName {computador} | Where-Object {{ $_.PathName -like "*{nome_cliente_formatado}*" }} | Format-Table PSComputerName, Name, State'

        # Executa o comando no PowerShell e captura a saída
        resultado = subprocess.check_output(['powershell', comando], text=True)

        # Inserir resultado na tabela
        inserir_na_tabela_segunda(resultado)

        # Adicionar os serviços da consulta atual à lista de todos os serviços
        lines = resultado.strip().split('\n')
        services = [line.split()[1] for line in lines if len(line.split()) >= 3]  # Capturar apenas o nome do serviço
        all_services.extend(services)

    # Preencher o Combobox com todos os serviços
    preencher_combobox(all_services)

    # Atualizar a janela para exibir o Combobox e os botões
    combobox_servicos.grid(row=7, columnspan=3, padx=5, pady=5)
    botao_iniciar.grid(row=7, column=0, padx=5, pady=5)
    botao_parar.grid(row=7, column=2, padx=5, pady=5)

def inserir_na_tabela_segunda(resultado):
    lines = resultado.strip().split('\n')
    for line in lines:
        data = line.split()
        # Verificar se a linha contém os valores esperados
        if len(data) >= 3:
            # Capturar PSComputerName, Name e State
            ps_computer_name = data[0]
            name = data[1]
            # Encontrar o campo 'State'
            state = data[-1]
            treeview.insert('', 'end', values=(ps_computer_name, name, state))

# Função para preencher o Combobox com os nomes dos serviços
def preencher_combobox(services):
    services = [service for service in services if service != "Name"]  # Excluir a linha "Name" se presente
    combobox_servicos['values'] = ["Todos"] + services

# Função para INICIAR os serviços
def iniciar_servicos():
    # Obter o código HCM, nome do cliente e serviço selecionado
    codigo_hcm = entry_codigo_hcm.get()
    nome_cliente = entry_nome_cliente.get()
    servico_selecionado = combobox_servicos.get()

    # Definir o ambiente com base na seleção
    tipo_ambiente = var_tipo_ambiente.get()
    if tipo_ambiente == "Produção":
        ambiente = "p"
    else:
        ambiente = "h"

    # Formatar o nome do cliente para correspondência com os serviços
    nome_cliente_formatado = f"{nome_cliente}_{codigo_hcm}_{ambiente}"

    # Loop sobre os computadores
    for computador in computadores:
        # Comando para iniciar o serviço
        if servico_selecionado == "Todos":
            comando_iniciar = f'Get-WMIObject win32_service -ComputerName {computador} | Where-Object {{ $_.PathName -like "*{nome_cliente_formatado}*" }} | ForEach-Object {{ $_.StartService() }}'
        else:
            comando_iniciar = f'(Get-WMIObject win32_service -ComputerName {computador} | Where-Object {{ $_.PathName -like "*{nome_cliente_formatado}*" -and $_.Name -eq "{servico_selecionado}" }}).StartService()'

        # Executa o comando no PowerShell para iniciar o serviço
        resultado_iniciar = subprocess.run(['powershell', '-Command', comando_iniciar])

        # Atualizar a tabela com o novo estado do serviço
        segunda_consulta()

# Função para PARAR os serviços
def parar_servicos():
    # Obter o código HCM, nome do cliente e serviço selecionado
    codigo_hcm = entry_codigo_hcm.get()
    nome_cliente = entry_nome_cliente.get()
    servico_selecionado = combobox_servicos.get()

    # Definir o ambiente com base na seleção
    tipo_ambiente = var_tipo_ambiente.get()
    if tipo_ambiente == "Produção":
        ambiente = "p"
    else:
        ambiente = "h"

    # Formatar o nome do cliente para correspondência com os serviços
    nome_cliente_formatado = f"{nome_cliente}_{codigo_hcm}_{ambiente}"

    # Loop sobre os computadores
    for computador in computadores:
        # Comando para parar o serviço
        if servico_selecionado == "Todos":
            comando_parar = f'Get-WMIObject win32_service -ComputerName {computador} | Where-Object {{ $_.PathName -like "*{nome_cliente_formatado}*" }} | ForEach-Object {{ $_.StopService() }}'
        else:
            comando_parar = f'(Get-WMIObject win32_service -ComputerName {computador} | Where-Object {{ $_.PathName -like "*{nome_cliente_formatado}*" -and $_.Name -eq "{servico_selecionado}" }}).StopService()'

        # Executa o comando no PowerShell para parar o serviço
        resultado_parar = subprocess.run(['powershell', '-Command', comando_parar])

        # Atualizar a tabela com o novo estado do serviço
        segunda_consulta()

# Criar a janela
janela = tk.Tk()
janela.title("Manipulador de Serviços HCM")

# Criar os rótulos e campos de entrada
label_nome_cliente = tk.Label(janela, text="NOME DO CLIENTE:")
label_nome_cliente.grid(row=0, column=0, padx=5, pady=5)
entry_nome_cliente = tk.Entry(janela)
entry_nome_cliente.grid(row=0, column=1, padx=5, pady=5)

label_codigo_hcm = tk.Label(janela, text="CÓDIGO HCM:")

label_tipo_ambiente = tk.Label(janela, text="TIPO DE AMBIENTE:")
label_tipo_ambiente.grid(row=1, column=0, padx=5, pady=5)

var_tipo_ambiente = tk.StringVar(janela)
var_tipo_ambiente.set("Homologação")  # Valor padrão

opcao_producao = tk.Radiobutton(janela, text="PRODUÇÃO", variable=var_tipo_ambiente, value="Produção")
opcao_producao.grid(row=1, column=1, padx=5, pady=5)

opcao_homologacao = tk.Radiobutton(janela, text="HOMOLOGAÇÃO", variable=var_tipo_ambiente, value="Homologação")
opcao_homologacao.grid(row=1, column=2, padx=5, pady=5)

# Adicionar widget Treeview para mostrar os resultados
treeview = ttk.Treeview(janela, columns=('PSComputerName', 'Name', 'State'), show='headings')
treeview.heading('PSComputerName', text='SERVIDOR')
treeview.heading('Name', text='NOME DO SERVIÇO')
treeview.heading('State', text='STATUS')
treeview.grid(row=5, columnspan=3, padx=5, pady=5)

# Botão para consultar os serviços
botao_consultar = tk.Button(janela, text="CONSULTAR", command=primeira_consulta)
botao_consultar.grid(row=4, columnspan=3, padx=5, pady=5)

# Botão de mensagem de instrução
botao_instrucao = tk.Button(janela, text="INSTRUÇÃO", command=exibir_mensagem)
botao_instrucao.grid(row=3, columnspan=3, padx=5, pady=5)

# Botão para iniciar os serviços
botao_iniciar = tk.Button(janela, text="INICIAR SERVIÇO(S)", command=iniciar_servicos)

# Botão para parar os serviços
botao_parar = tk.Button(janela, text="PARAR SERVIÇO(S)", command=parar_servicos)

# Campo de entrada para o código HCM
entry_codigo_hcm = tk.Entry(janela)

# Combobox para selecionar o serviço
combobox_servicos = ttk.Combobox(janela, state="readonly")

# Computadores para consulta
computadores = ["nb021505"]  # Adicione mais se necessário

janela.mainloop()
