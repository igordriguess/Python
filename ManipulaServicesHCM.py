import subprocess
import tkinter as tk
from tkinter import ttk

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

    # Habilitar o campo para digitar o código HCM e selecionar o ambiente
    label_tipo_ambiente.grid(row=2, column=0, padx=5, pady=5)
    opcao_producao = tk.Radiobutton(janela, text="Produção", variable=var_tipo_ambiente, value="Produção")
    opcao_producao.grid(row=2, column=1, padx=5, pady=5)
    opcao_homologacao = tk.Radiobutton(janela, text="Homologação", variable=var_tipo_ambiente, value="Homologação")
    opcao_homologacao.grid(row=2, column=2, padx=5, pady=5)

    label_codigo_hcm.grid(row=1, column=0, padx=5, pady=5)
    entry_codigo_hcm.grid(row=1, column=1, padx=5, pady=5)
    entry_codigo_hcm.config(state=tk.NORMAL)  # Habilitar o campo

    # Mostrar o botão para consultar cliente
    botao_consultar_cliente.grid(row=1, column=2, padx=5, pady=5)

    limpar_combobox()

# Função para limpar a tabela
def limpar_tabela():
    for item in treeview.get_children():
        treeview.delete(item)

# Função para limpar as informações do Combobox
def limpar_combobox():
    combobox_servicos['values'] = []

# Função para inserir resultado na tabela
def inserir_na_tabela(resultado):
    lines = resultado.strip().split('\n')
    for line in lines:
        # Ignorar linhas que começam com "PSComputerName" ou "-----------"
        if not line.startswith("PSComputerName") and not line.startswith("-----------"):
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
        # Ignorar linhas que começam com "PSComputerName" ou "-----------"
        if not line.startswith("PSComputerName") and not line.startswith("-----------"):
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
    combobox_servicos.config(width=40)  # Definir a largura do combobox

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
janela.title("Manipulador de Serviços HCM SaaS Orion")

# Criar os rótulos e campos de entrada
label_nome_cliente = tk.Label(janela, text="Nome do Cliente:")
label_nome_cliente.grid(row=0, column=0, padx=5, pady=5)
entry_nome_cliente = tk.Entry(janela)
entry_nome_cliente.grid(row=0, column=1, padx=5, pady=5)

label_codigo_hcm = tk.Label(janela, text="Código HCM:")

label_tipo_ambiente = tk.Label(janela, text="Tipo de Ambiente:")

var_tipo_ambiente = tk.StringVar(janela)
var_tipo_ambiente.set("Homologação")  # Valor padrão

# Criar estilo para a tabela
style = ttk.Style()
style.configure("Treeview", rowheight=40)  # Ajuste o valor de rowheight conforme necessário

# Adicionar widget Treeview para mostrar os resultados
treeview = ttk.Treeview(janela, columns=('PSComputerName', 'Name', 'State'), show='headings')
treeview.heading('PSComputerName', text='Servidor')
treeview.heading('Name', text='Nome do Serviço')
treeview.heading('State', text='Status')
treeview.grid(row=20, column=0, columnspan=3, padx=5, pady=5)

# Definir a largura das colunas
treeview.column('PSComputerName', width=200)  # Ajuste o valor de width conforme necessário
treeview.column('Name', width=400)  # Ajuste o valor de width conforme necessário
treeview.column('State', width=200)  # Ajuste o valor de width conforme necessário

# Botão para consultar os serviços na primeira consulta
botao_consultar = tk.Button(janela, text="Consultar Código HCM", command=primeira_consulta)
botao_consultar.grid(row=0, column=2, padx=5, pady=5)

# Botão para consultar os serviços na segunda consulta
botao_consultar_cliente = tk.Button(janela, text="Consultar Cliente", command=segunda_consulta)
botao_consultar_cliente.grid(row=1, column=2, padx=5, pady=5)
botao_consultar_cliente.grid_remove()  # Ocultar o botão inicialmente

# Botão para iniciar os serviços
botao_iniciar = tk.Button(janela, text="Iniciar Serviço(s)", command=iniciar_servicos)

# Botão para parar os serviços
botao_parar = tk.Button(janela, text="Parar Serviço(s)", command=parar_servicos)

# Campo de entrada para o código HCM
entry_codigo_hcm = tk.Entry(janela)

# Combobox para selecionar o serviço
combobox_servicos = ttk.Combobox(janela, state="readonly")

# Computadores para consulta
computadores = ["OCSENAPLH01", "OCSENAPL01", "OCSENAPL02", "OCSENAPL03", "OCSENAPL04", "OCSENINT01", "OCSENMDW01"] # Adicione mais se necessário

janela.mainloop()
