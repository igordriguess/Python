import subprocess
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk

# Função para exibir mensagem de instrução
def exibir_mensagem():
    messagebox.showinfo("Instrução", 'Digite o nome do cliente e clique em "Consultar", será consultado o Serviço de Informação da Instalação. Valide o código HCM do cliente, digite o código no campo que será habilitado e escolha o tipo de ambiente, em seguida, clique em "Consultar" novamente.')

# Função para realizar a primeira consulta e exibir resultados
def primeira_consulta():
    # Obter os dados do formulário
    nome_cliente = entry_nome_cliente.get()

    # Limpar a tabela anterior
    limpar_tabela()

    for computador in computadores:
        # Comando para verificar o serviço
        comando = f'Get-WmiObject -Class Win32_Service -ComputerName {computador} | Where-Object {{ $_.PathName -like "*{nome_cliente}*" }} | Where-Object {{ $_.Name -like "*SeniorInstInfo*" }} | Format-Table PSComputerName, Name, PathName, State'

        # Executa o comando no PowerShell e captura a saída
        resultado = subprocess.check_output(['powershell', comando], text=True)

        # Inserir resultado na tabela
        inserir_na_tabela(resultado)

    # Habilitar o campo para digitar o código HCM
    label_codigo_hcm.grid(row=2, column=0, padx=5, pady=5)
    entry_codigo_hcm.grid(row=2, column=1, padx=5, pady=5)
    entry_codigo_hcm.config(state=tk.NORMAL)  # Habilitar o campo
    botao_consultar.config(command=segunda_consulta)  # Configurar o botão para executar a segunda consulta

# Função para converter o resultado para formato JSON
def convert_to_json(resultado):
    lines = resultado.strip().split('\n')
    resultados_json = []
    for line in lines:
        data = line.split()
        # Verificar se a linha contém os valores esperados
        if len(data) >= 4:
            # Capturar PSComputerName, Name, PathName e State
            ps_computer_name = data[0]
            name = data[1]
            path_name = data[2]
            # O campo State pode estar na última posição ou na posição 3
            state = data[-1] if data[-1].startswith("Running") or data[-1].startswith("Stopped") else data[3]
            resultado_dict = {
                "PSComputerName": ps_computer_name,
                "Name": name,
                "PathName": path_name,
                "State": state
            }
            resultados_json.append(resultado_dict)
    return resultados_json

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
        if len(data) >= 4:
            # Capturar PSComputerName, Name, PathName e State
            ps_computer_name = data[0]
            name = data[1]
            path_name = data[2]
            # O campo State pode estar na última posição ou na posição 3
            state = data[-1] if data[-1].startswith("Running") or data[-1].startswith("Stopped") else data[3]
            treeview.insert('', 'end', values=(ps_computer_name, name, path_name, state))
            # Adicionar uma linha em branco após cada inserção
            treeview.insert('', 'end', values=('', '', '', ''))

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

    nome_cliente_formatado = f"{nome_cliente}_{codigo_hcm}_{ambiente}"

    # Limpar a tabela anterior
    limpar_tabela()

    for computador in computadores:
        # Comando para verificar o serviço
        comando = f'Get-WmiObject -Class Win32_Service -ComputerName {computador} | Where-Object {{ $_.PathName -like "*{nome_cliente_formatado}*" }} | Format-Table PSComputerName, Name, PathName, State'

        # Executa o comando no PowerShell e captura a saída
        resultado = subprocess.check_output(['powershell', comando], text=True)

        # Inserir resultado na tabela
        inserir_na_tabela(resultado)

# Função para INICIAR os serviços
def iniciar_servicos():
    # Obter o código HCM e o nome do cliente
    codigo_hcm = entry_codigo_hcm.get()
    nome_cliente = entry_nome_cliente.get()

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
        comando_iniciar = f'(Get-WMIObject win32_service -ComputerName {computador} | Where-Object {{ $_.PathName -like "*{nome_cliente_formatado}*" }}).StartService()'

        # Executa o comando no PowerShell para parar o serviço
        resultado_iniciar = subprocess.run(['powershell', '-Command', comando_iniciar])

        # Comando para verificar novamente o serviço
        comando_consultar_um = f'Get-WmiObject -Class Win32_Service -ComputerName {computador} | Where-Object {{ $_.PathName -like "*{nome_cliente_formatado}*" }} | Format-Table PSComputerName, Name, PathName, State'

        # Executa o comando no PowerShell e captura a saída do comando_dois
        resultado_comando_um = subprocess.check_output(['powershell', comando_consultar_um], text=True)

        # Inserir resultado na tabela
        inserir_na_tabela(resultado_comando_um)

# Função para PARAR os serviços
def parar_servicos():
    # Obter o código HCM e o nome do cliente
    codigo_hcm = entry_codigo_hcm.get()
    nome_cliente = entry_nome_cliente.get()

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
        comando_parar = f'(Get-WMIObject win32_service -ComputerName {computador} | Where-Object {{ $_.PathName -like "*{nome_cliente_formatado}*" }}).StopService()'

        # Executa o comando no PowerShell para parar o serviço
        resultado_parar = subprocess.run(['powershell', '-Command', comando_parar])

        # Comando para verificar novamente o serviço
        comando_consultar_dois = f'Get-WmiObject -Class Win32_Service -ComputerName {computador} | Where-Object {{ $_.PathName -like "*{nome_cliente_formatado}*" }} | Format-Table PSComputerName, Name, PathName, State'

        # Executa o comando no PowerShell e captura a saída do comando_dois
        resultado_comando_dois = subprocess.check_output(['powershell', comando_consultar_dois], text=True)

        # Inserir resultado na tabela
        inserir_na_tabela(resultado_comando_dois)

# Criar a janela
janela = tk.Tk()
janela.title("Consulta de Serviços HCM")

# Criar os rótulos e campos de entrada
label_nome_cliente = tk.Label(janela, text="Nome do Cliente:")
label_nome_cliente.grid(row=0, column=0, padx=5, pady=5)
entry_nome_cliente = tk.Entry(janela)
entry_nome_cliente.grid(row=0, column=1, padx=5, pady=5)

label_codigo_hcm = tk.Label(janela, text="Código HCM:")

label_tipo_ambiente = tk.Label(janela, text="Tipo de Ambiente:")
label_tipo_ambiente.grid(row=1, column=0, padx=5, pady=5)

var_tipo_ambiente = tk.StringVar(janela)
var_tipo_ambiente.set("Produção")  # Valor padrão

opcao_producao = tk.Radiobutton(janela, text="Produção", variable=var_tipo_ambiente, value="Produção")
opcao_producao.grid(row=1, column=1, padx=5, pady=5)

opcao_homologacao = tk.Radiobutton(janela, text="Homologação", variable=var_tipo_ambiente, value="Homologação")
opcao_homologacao.grid(row=1, column=0, padx=5, pady=5)

# Adicionar widget Treeview para mostrar os resultados
treeview = ttk.Treeview(janela, columns=('PSComputerName', 'Name', 'PathName', 'State'), show='headings')
treeview.heading('PSComputerName', text='PSComputerName')
treeview.heading('Name', text='Name')
treeview.heading('PathName', text='PathName')
treeview.heading('State', text='State')
treeview.grid(row=5, columnspan=2, padx=5, pady=5)

# Botão para consultar os serviços
botao_consultar = tk.Button(janela, text="Consultar", command=primeira_consulta)
botao_consultar.grid(row=4, columnspan=2, padx=5, pady=5)

# Botão de mensagem de instrução
botao_instrucao = tk.Button(janela, text="Instrução", command=exibir_mensagem)
botao_instrucao.grid(row=3, columnspan=2, padx=5, pady=5)

# Botão para iniciar os serviços
botao_iniciar = tk.Button(janela, text="Iniciar Serviços", command=iniciar_servicos)
botao_iniciar.grid(row=6, column=0, padx=5, pady=5)

# Botão para parar os serviços
botao_parar = tk.Button(janela, text="Parar Serviços", command=parar_servicos)
botao_parar.grid(row=6, column=1, padx=5, pady=5)

# Campo de entrada para o código HCM
entry_codigo_hcm = tk.Entry(janela)

# Computadores para consulta
computadores = ["OCSENAPLH01", "OCSENAPL01", "OCSENAPL02", "OCSENAPL03", "OCSENAPL04", "OCSENINT01", "OCSENMDW01"]  # Adicione mais se necessário

janela.mainloop()
