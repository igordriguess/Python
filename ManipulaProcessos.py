import subprocess
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk

# Função para exibir mensagem de instrução
def exibir_mensagem():
    messagebox.showinfo("INSTRUÇÃO", 'Digite o SERVIDOR e o nome do CLIENTE e clique em "Consultar", serão exibidos todos os processos em execução do cliente.\n\nDigite o ID do processo ou o NOME DO PROCESSO caso deseje filtrar por um processo específico e clique em "Filtrar"\n\nClique no botão "Encerrar Processo(s) Filtrado(s)" para encerrar o processo filtrado anteriormente pelo ID ou pelo NOME.\n\nClique em "Encerrar Todos" para finalizar TODOS os processos em execução do cliente.')

# Função para realizar a primeira consulta e exibir resultados
def primeira_consulta():
    # Obter os dados do formulário
    nome_cliente = entry_nome_cliente.get()
    nome_servidor = entry_nome_servidor.get()

    # Limpar a tabela anterior
    limpar_tabela()

    # Executar comando PowerShell Remoting em cada servidor na lista
    comando = f"Invoke-Command -ComputerName {nome_servidor} -ScriptBlock {{ Get-Process | Where-Object {{ $_.Path -like '*{nome_cliente}*' }} | Select-Object PSComputerName, ProcessName, Path, Id}}"
    comando = ["powershell", "-Command", comando]  # Adiciona o prefixo powershell para execução do comando

    # Capturar a saída do comando
    processo = subprocess.run(comando, capture_output=True, text=True)

    # Inserir resultado na tabela
    inserir_na_tabela(processo.stdout)

# Função para limpar a tabela
def limpar_tabela():
    for item in treeview.get_children():
        treeview.delete(item)

# Função para inserir resultado na tabela
def inserir_na_tabela(output):
    # Dividir a saída em linhas
    linhas = output.strip().split('\n')
    for linha in linhas:
        # Dividir cada linha em colunas
        dados = linha.strip().split(': ')
        # Inserir dados na tabela
        treeview.insert('', 'end', values=dados)

# Função para realizar a segunda consulta e exibir resultados
def segunda_consulta():
    # Obter os dados do formulário
    nome_cliente = entry_nome_cliente.get()
    nome_servidor = entry_nome_servidor.get()
    id_nome_processo = entry_id_nome_processo.get()

    # Limpar a tabela anterior
    limpar_tabela()

    # Verificar se o campo é um ID ou um nome de processo
    if id_nome_processo.isdigit():  # Se for um ID, encerra pelo ID
        comando = f"Invoke-Command -ComputerName {nome_servidor} -ScriptBlock {{ Get-Process | Where-Object {{ $_.Path -like '*{nome_cliente}*' -and $_.Id -eq '{id_nome_processo}'}} | Select-Object PSComputerName, ProcessName, Path, Id}}"
    else:  # Se for um nome, encerra pelo nome
        comando = f"Invoke-Command -ComputerName {nome_servidor} -ScriptBlock {{ Get-Process | Where-Object {{ $_.Path -like '*{nome_cliente}*' -and $_.Name -like '*{id_nome_processo}*'}} | Select-Object PSComputerName, ProcessName, Path, Id}}"
    comando = ["powershell", "-Command", comando]  # Adiciona o prefixo powershell para execução do comando

    # Capturar a saída do comando
    processo = subprocess.run(comando, capture_output=True, text=True)

    # Inserir resultado na tabela
    inserir_na_tabela(processo.stdout)

# Função para finalizar o processo pelo ID ou pelo Nome
def encerrar_processo():
    # Obter os dados do formulário
    nome_cliente = entry_nome_cliente.get()
    nome_servidor = entry_nome_servidor.get()
    id_nome_processo = entry_id_nome_processo.get()

    # Limpar a tabela anterior
    limpar_tabela()

    # Executar comando PowerShell Remoting em cada servidor na lista
    if id_nome_processo.isdigit():  # Se for um ID, encerra pelo ID
        comando = f"Invoke-Command -ComputerName {nome_servidor} -ScriptBlock {{ Get-Process | Where-Object {{ $_.Path -like '*{nome_cliente}*' -and $_.Id -eq '{id_nome_processo}'}} | Stop-Process -Force}}"
    else:  # Se for um nome, encerra pelo nome
        comando = f"Invoke-Command -ComputerName {nome_servidor} -ScriptBlock {{ Get-Process | Where-Object {{ $_.Path -like '*{nome_cliente}*' -and $_.Name -like '*{id_nome_processo}*'}} | Stop-Process -Force}}"
    #comando = f"Invoke-Command -ComputerName {nome_servidor} -ScriptBlock {{ Get-Process | Where-Object {{ $_.Name -like '*{id_nome_processo}*' -and $_.Id -eq '{id_nome_processo}'}} | Stop-Process -Force}}"
    comando = ["powershell", "-Command", comando]  # Adiciona o prefixo powershell para execução do comando

    # Capturar a saída do comando
    processo = subprocess.run(comando, capture_output=True, text=True)

    # Inserir resultado na tabela
    inserir_na_tabela(processo.stdout)

# Função para finalizar TODOS os processos
def encerrar_todos_processos():
    # Obter os dados do formulário
    nome_cliente = entry_nome_cliente.get()
    nome_servidor = entry_nome_servidor.get()

    # Limpar a tabela anterior
    limpar_tabela()

    # Executar comando PowerShell Remoting em cada servidor na lista
    comando = f"Invoke-Command -ComputerName {nome_servidor} -ScriptBlock {{ Get-Process | Where-Object {{ $_.Path -like '*{nome_cliente}*' }} | Stop-Process -Force}}"
    comando = ["powershell", "-Command", comando]  # Adiciona o prefixo powershell para execução do comando

    # Capturar a saída do comando
    processo = subprocess.run(comando, capture_output=True, text=True)

    # Inserir resultado na tabela
    inserir_na_tabela(processo.stdout)

# Criar a janela
janela = tk.Tk()
janela.title("Manipulador de Processos")

# Criar os rótulos e campos de entrada
label_nome_servidor = tk.Label(janela, text="Servidor:")
label_nome_servidor.grid(row=0, column=0, padx=5, pady=5)
entry_nome_servidor = tk.Entry(janela)
entry_nome_servidor.grid(row=0, column=1, padx=5, pady=5)

label_nome_cliente = tk.Label(janela, text="Nome do Cliente:")
label_nome_cliente.grid(row=1, column=0, padx=5, pady=5)
entry_nome_cliente = tk.Entry(janela)
entry_nome_cliente.grid(row=1, column=1, padx=5, pady=5)

label_id_ou_nome_processo = tk.Label(janela, text="Filtrar pelo ID ou Nome do Processo:")
label_id_ou_nome_processo.grid(row=2, column=0, padx=5, pady=5)
entry_id_nome_processo = tk.Entry(janela)
entry_id_nome_processo.grid(row=2, column=1, padx=5, pady=5)

label_encerrar_processo = tk.Label(janela, text="ID do Processo:")

# Adicionar widget Treeview para mostrar os resultados
treeview = ttk.Treeview(janela, columns=('PSComputerName', 'ProcessName'), show='headings')
treeview.heading('PSComputerName', text='Campos')
treeview.heading('ProcessName', text='Detalhes do Processo')
treeview.grid(row=10, columnspan=2, padx=5, pady=5)

# Definir a altura da Treeview (em número de linhas)
treeview['height'] = 20  # Ajuste o valor conforme necessário

# Definindo a largura da Treeview
treeview.column("PSComputerName", width=200)
treeview.column("ProcessName", width=400)

# Botão para consultar os serviços
botao_consultar = tk.Button(janela, text="Consultar", command=primeira_consulta)
botao_consultar.grid(row=1, columnspan=2, padx=5, pady=5)

# Botão para consultar os serviços
botao_consultar = tk.Button(janela, text="Filtrar", command=segunda_consulta)
botao_consultar.grid(row=2, columnspan=2, padx=5, pady=5)

# Botão de mensagem de instrução
botao_instrucao = tk.Button(janela, text="Instrução", command=exibir_mensagem)
botao_instrucao.grid(row=0, columnspan=2, padx=5, pady=5)

# Botão para encerrar o processo pelo ID
botao_encerrar_processo = tk.Button(janela, text="Encerrar Processo(s) Filtrado(s)", command=encerrar_processo)
botao_encerrar_processo.grid(row=4, columnspan=2, padx=5, pady=5)

# Botão para encerrar TODOS os processos
botao_encerrar_todos_processos = tk.Button(janela, text="Encerrar Todos", command=encerrar_todos_processos)
botao_encerrar_todos_processos.grid(row=15, columnspan=2, padx=5, pady=5)

# Campo de entrada para o código HCM
entry_encerrar_processo = tk.Entry(janela)

janela.mainloop()
