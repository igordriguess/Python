import subprocess
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk

# Função para exibir mensagem de instrução
def exibir_mensagem():
    messagebox.showinfo("Instrução", 'Digite o SERVIDOR e o nome do CLIENTE, clique em "Consultar", serão exibidos todos os processos em execução.\n\nCaso deseje encerrar algum processo, digite o ID correspondente e clique em "Encerrar Processo".\n\nClique em "Encerrar Todos" para finalizar TODOS os processos apresentados na tela.')

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

    # Habilitar o campo para digitar o ID do Processo
    label_encerrar_processo.grid(row=4, column=0, padx=5, pady=5)
    entry_encerrar_processo.grid(row=4, column=1, padx=5, pady=5)
    entry_encerrar_processo.config(state=tk.NORMAL)  # Habilitar o campo
    botao_encerrar_processo.config(command=encerrar_processo)  # Configurar o encerrar o processo

# Função para finalizar o processo pelo ID
def encerrar_processo():
    # Obter os dados do formulário
    nome_cliente = entry_nome_cliente.get()
    nome_servidor = entry_nome_servidor.get()
    encerrar_processo = entry_encerrar_processo.get()

    # Limpar a tabela anterior
    limpar_tabela()

    # Executar comando PowerShell Remoting em cada servidor na lista
    comando = f"Invoke-Command -ComputerName {nome_servidor} -ScriptBlock {{ Get-Process | Where-Object {{ $_.Id -like '*{encerrar_processo}*' }} | Stop-Process -Force}}"
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

# Botão de mensagem de instrução
botao_instrucao = tk.Button(janela, text="Instrução", command=exibir_mensagem)
botao_instrucao.grid(row=0, columnspan=2, padx=5, pady=5)

# Botão para encerrar o processo pelo ID
botao_encerrar_processo = tk.Button(janela, text="Encerrar o Processo ID", command=encerrar_processo)
botao_encerrar_processo.grid(row=4, columnspan=2, padx=5, pady=5)

# Botão para encerrar TODOS os processos
botao_encerrar_todos_processos = tk.Button(janela, text="Encerrar Todos", command=encerrar_todos_processos)
botao_encerrar_todos_processos.grid(row=15, columnspan=2, padx=5, pady=5)

# Campo de entrada para o código HCM
entry_encerrar_processo = tk.Entry(janela)

janela.mainloop()
