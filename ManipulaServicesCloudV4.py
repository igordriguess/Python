import subprocess
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import messagebox, Tk, ttk as tk_tt

# Dicionário para mapear nomes reais e mascarados
nome_mascarado_para_real = {}

# Lista de computadores para executar os comandos
computadores = ["localhost"]  # Substitua com os nomes reais dos computadores

# Configurações globais da interface
root = Tk()
root.title("Gerenciador de Serviços Statum")

# Variável para o tipo de ambiente
var_tipo_ambiente = ttk.StringVar(value="Produção")

# Função para exibir mensagem de instrução
def exibir_mensagem():
    messagebox.showinfo(
        "Instrução",
        'Selecione o ambiente desejado (Produção ou Homologação) e clique em "Consultar".\n\n'
        'Na barra inferior, selecione o serviço desejado ou selecione "Todos" para executar o comando em todos os serviços apresentados.',
    )

# Função para realizar a consulta e exibir resultados
def primeira_consulta():
    # Definir o nome do cliente fixo
    nome_cliente = "senior\\"

    # Obter os dados do formulário
    tipo_ambiente = var_tipo_ambiente.get()

    # Definir o filtro do ambiente e as variáveis de controle
    if tipo_ambiente == "Produção":
        ambiente = nome_cliente.lower()  # Produção busca pelo nome do cliente
        comando_wildfly = True           # WildFly será incluído
        comando_glassfish_teste = False  # Glassfish Teste não será incluído
    elif tipo_ambiente == "Homologação":
        ambiente = "teste"               # Homologação busca apenas por "Teste"
        comando_wildfly = False          # WildFly não será incluído
        comando_glassfish_teste = True   # Glassfish Teste será incluído

    # Limpar a tabela anterior
    limpar_tabela()

    # Lista para armazenar todos os serviços
    all_services = []

    if ambiente:  # Só realiza a consulta se um ambiente válido foi definido
        for computador in computadores:
            # Inicializar o comando básico
            comando = f'Get-WmiObject -Class Win32_Service -ComputerName {computador}'

            # Adicionar os filtros dependendo do ambiente
            if comando_wildfly:
                comando += f' | Where-Object {{ $_.PathName -like "*{ambiente}*" -or $_.PathName -like "*wildfly*" -or $_.PathName -like "*domain\\*" -or $_.PathName -like "*gestaoponto\\*" -or $_.PathName -like "*csmcenter\\*" -or $_.PathName -like "*spoolsv*"}}'
            elif comando_glassfish_teste:
                comando += f' | Where-Object {{ $_.PathName -like "*{ambiente}*" -or $_.PathName -like "*domainteste\\*" -or $_.PathName -like "*gestaopontoteste\\*" }}'

            comando += ' | Format-Table Name, State'

            #print(f"Comando Glassfish Produção: {comando}")
            #print(f"Comando Glassfish Teste: {comando}")

            # Executa o comando no PowerShell e captura a saída
            try:
                resultado = subprocess.check_output(['powershell', comando], text=True)
            except subprocess.CalledProcessError as e:
                resultado = f"Erro ao consultar {computador}: {e.output}"

            # Inserir resultado na tabela
            inserir_na_tabela(resultado)

            # Adicionar os serviços da consulta atual à lista de todos os serviços
            lines = resultado.strip().split('\n')
            services = [' '.join(line.split()[:-1]) for line in lines if len(line.split()) >= 2]  # Capturar o nome completo do serviço
            all_services.extend(services)

    # Preencher o Combobox com todos os serviços
    preencher_combobox(all_services)

# Função para preencher o Combobox com os nomes reais dos serviços
def preencher_combobox(services):
    services = [service for service in services if service != "Name"]  # Excluir a linha "Name" se presente
    masked_services = []
    for service in services:
        masked_name = "Senior Motor e-Social" if "motor" in service.lower() else \
                      "Senior Serviço de Informação da Instalação" if "seniorinst" in service.lower() else \
                      "Senior Integrador Wiipo" if "wiipo" in service.lower() else \
                      "Senior Middleware" if "middleware" in service.lower() else \
                      "Senior SAM Integrador" if "integrador" in service.lower() else \
                      "Senior Integrador G7" if "integration" in service.lower() else \
                      "Senior Concentradora" if "concentradora" in service.lower() else \
                      "Senior CSM Center" if "center" in service.lower() else \
                      "Senior SDE" if "sde_sde" in service.lower() else \
                      "Senior SDE Print Service" if "sde_print_sde" in service.lower() else \
                      "Senior SDE Teste" if "sde_sdeteste" in service.lower() else \
                      "Senior SDE Print Service Teste" if "sde_print_sdeteste" in service.lower() else \
                      "Senior WildFly" if "wildfly" in service.lower() else \
                      "Senior Glassfish 4 domain" if "domain" in service.lower() and "domainteste" not in service.lower() else \
                      "Senior Glassfish 4 domainteste" if "domainteste" in service.lower() else \
                      "Senior Glassfish 4 gestaoponto" if "gestaoponto" in service.lower() and "gestaopontoteste" not in service.lower() else \
                      "Senior Glassfish 4 gestaopontoteste" if "gestaopontoteste" in service.lower() else \
                      "Senior Glassfish 4 csmcenter" if "csmcenter" in service.lower() else \
                      "Spooler de Impressão" if "spooler" in service.lower() else \
                      "Senior Integrador HCM" if "integrator" in service.lower() else service
        masked_services.append(masked_name)
        nome_mascarado_para_real[masked_name] = service

        # Verifica se o nome do serviço contém '-----' antes de exibir
        #if '----' not in service.lower():
        #    print(f"Nome real do serviço: {service.lower()}")
            
        #if '----' not in masked_name:
        #    print(f"Nome do serviço com a máscara: {masked_name}")

    combobox_servicos['values'] = ["Todos"] + masked_services
    combobox_servicos.config(width=40)  # Definir a largura do combobox

# Função para limpar a tabela
def limpar_tabela():
    for item in treeview.get_children():
        treeview.delete(item)

# Função para inserir resultado na tabela
def inserir_na_tabela(resultado):
    # Contador para alternar entre as cores das linhas
    index = 0

    # Processar as linhas do resultado
    lines = resultado.strip().split('\n')
    for line in lines[2:]:  # Ignorar as duas primeiras linhas (cabeçalhos)
        data = line.split()
        
        # Verificar se a linha contém os valores esperados
        if len(data) >= 2:
            # Capturar Name e State
            name = ' '.join(data[:-1])  # Juntar todos os elementos, exceto o último
            state = data[-1]  # Último elemento é o estado
            
            # Substituir "Running" por "Em execução" e "Stopped" por "Parado"
            state = "Em Execução" if state == "Running" else "Parado" if state == "Stopped" else state
            
            # Mascarar o nome do serviço conforme as regras
            masked_name = (
                "Senior Motor e-Social" if "motor" in name.lower() else
                "Senior Serviço de Informação da Instalação" if "seniorinst" in name.lower() else
                "Senior Integrador Wiipo" if "wiipo" in name.lower() else
                "Senior Integrador HCM" if "integrator" in name.lower() else
                "Senior SAM Integrador" if "integrador" in name.lower() else
                "Senior Integrador G7" if "integration" in name.lower() else
                "Senior Concentradora" if "concentradora" in name.lower() else
                "Senior CSM Center" if "center" in name.lower() else
                "Senior SDE" if "sde_sde" in name.lower() else
                "Senior SDE Print Service" if "sde_print_sde" in name.lower() else
                "Senior SDE Teste" if "sde_sdeteste" in name.lower() else
                "Senior SDE Print Service Teste" if "sde_print_sdeteste" in name.lower() else
                "Senior WildFly" if "wildfly" in name.lower() else
                "Senior Glassfish 4 domain" if "domain" in name.lower() and "domainteste" not in name.lower() else
                "Senior Glassfish 4 domainteste" if "domainteste" in name.lower() else
                "Senior Glassfish 4 gestaoponto" if "gestaoponto" in name.lower() and "gestaopontoteste" not in name.lower() else
                "Senior Glassfish 4 gestaopontoteste" if "gestaopontoteste" in name.lower() else
                "Senior Glassfish 4 csmcenter" if "csmcenter\\" in name.lower() else
                "Spooler de Impressão" if "spooler" in name.lower() else \
                "Senior Middleware" if "middleware" in name.lower() else name
            )
            
            # Mapeamento do nome mascarado para o nome real
            nome_mascarado_para_real[masked_name] = name
            
            # Definir a tag de cor alternada (oddrow/evenrow)
            tag = "oddrow" if index % 2 == 0 else "evenrow"
            index += 1  # Incrementar o índice para alternar as tags
            
            # Inserir no Treeview
            treeview.insert("", "end", values=(masked_name, state), tags=(tag,))

            # Verifica se o nome do serviço contém '-----' antes de exibir
            #if '----' not in name.lower():
            #    print(f"Nome real do serviço: {name.lower()}")
                
            #if '----' not in masked_name:
            #    print(f"Nome do serviço com a máscara: {masked_name}")

# Função para obter o serviço selecionado no Combobox
def obter_servico_selecionado():
    selected_text = combobox_servicos.get()
    return selected_text

# Função para INICIAR os serviços
def iniciar_servicos():
    # Definir o nome do cliente fixo
    nome_cliente = "senior\\"

    # Obter os dados do formulário
    tipo_ambiente = var_tipo_ambiente.get()

    # Obter o serviço selecionado pelo usuário
    servico_selecionado = obter_servico_selecionado()

    # Configurar o ambiente
    if tipo_ambiente == "Produção":
        ambiente = nome_cliente.lower()  # Produção busca pelo nome do cliente
    elif tipo_ambiente == "Homologação":
        ambiente = "teste"  # Homologação busca apenas por "Teste"
    else:
        ambiente = None  # Caso não seja Produção ou Homologação

    # Mapear o nome mascarado para o nome real
    servico_real = nome_mascarado_para_real.get(servico_selecionado, servico_selecionado)

    # Loop sobre os computadores
    if ambiente:  # Só realiza a execução se o ambiente for válido
        for computador in computadores:
            # Construir o comando para iniciar o serviço
            if servico_real == "Todos":
                # Iniciar todos os serviços do ambiente, incluindo WildFly apenas em Produção
                comando_iniciar = (
                    f'Get-WMIObject win32_service -ComputerName {computador} | '
                    f'Where-Object {{ $_.PathName -like "*{ambiente}*" '
                    f'{"-or $_.PathName -like \"*wildfly*\" -or $_.PathName -like \"*domain\\*\" -or $_.PathName -like \"*gestaoponto\\*\" -or $_.PathName -like \"*csmcenter\\*\" -or $_.PathName -like \"*spoolsv*\" " if tipo_ambiente == "Produção" else ""} }} | '
                    f'ForEach-Object {{ if ($_) {{ $_.StartService() }} }}'
                )
            else:
                # Iniciar apenas o serviço selecionado
                comando_iniciar = (
                    f'$service = Get-WMIObject win32_service -ComputerName {computador} | '
                    f'Where-Object {{ $_.Name -eq "{servico_real}" }}; '
                    f'if ($service) {{ $service.StartService() }};'
                )

            # Exibir o comando para debug
            print(f"Executando comando: {comando_iniciar}")

            # Executa o comando no PowerShell para iniciar o serviço
            try:
                resultado_iniciar = subprocess.run(['powershell', '-Command', comando_iniciar], text=True, check=True)
            except subprocess.CalledProcessError as e:
                print(f"Erro ao iniciar o serviço no computador {computador}: {e}")

    # Atualizar a tabela com o novo estado do serviço
    primeira_consulta()

# Função para PARAR os serviços
def parar_servicos():
    # Definir o nome do cliente fixo
    nome_cliente = "senior\\"

    # Obter os dados do formulário
    tipo_ambiente = var_tipo_ambiente.get()

    # Obter o serviço selecionado pelo usuário
    servico_selecionado = obter_servico_selecionado()

    # Configurar o ambiente
    if tipo_ambiente == "Produção":
        ambiente = nome_cliente.lower()  # Produção busca pelo nome do cliente
    elif tipo_ambiente == "Homologação":
        ambiente = "teste"  # Homologação busca apenas por "Teste"
    else:
        ambiente = None  # Caso não seja Produção ou Homologação

    # Mapear o nome mascarado para o nome real
    servico_real = nome_mascarado_para_real.get(servico_selecionado, servico_selecionado)

    # Loop sobre os computadores
    if ambiente:  # Só realiza a execução se o ambiente for válido
        for computador in computadores:
            # Construir o comando para iniciar o serviço
            if servico_real == "Todos":
                # Iniciar todos os serviços do ambiente, incluindo WildFly apenas em Produção
                comando_parar = (
                    f'Get-WMIObject win32_service -ComputerName {computador} | '
                    f'Where-Object {{ $_.PathName -like "*{ambiente}*" '
                    f'{"-or $_.PathName -like \"*wildfly*\" -or $_.PathName -like \"*domain\\*\" -or $_.PathName -like \"*gestaoponto\\*\" -or $_.PathName -like \"*csmcenter\\*\" -or $_.PathName -like \"*spoolsv*\" " if tipo_ambiente == "Produção" else ""} }} | '
                    f'ForEach-Object {{ if ($_) {{ $_.StopService() }} }}'
                )
            else:
                # Iniciar apenas o serviço selecionado
                comando_parar = (
                    f'$service = Get-WMIObject win32_service -ComputerName {computador} | '
                    f'Where-Object {{ $_.Name -eq "{servico_real}" }}; '
                    f'if ($service) {{ $service.StopService() }};'
                )

            # Exibir o comando para debug
            print(f"Executando comando: {comando_parar}")

            # Executa o comando no PowerShell para iniciar o serviço
            try:
                resultado_iniciar = subprocess.run(['powershell', '-Command', comando_parar], text=True, check=True)
            except subprocess.CalledProcessError as e:
                print(f"Erro ao iniciar o serviço no computador {computador}: {e}")

    # Atualizar a tabela com o novo estado do serviço
    primeira_consulta()

# Configurar interface gráfica
frame_top = ttk.Frame(root, padding=10)
frame_top.grid(row=0, column=0, columnspan=3)

##################### Cabeçalho #####################
ttk.Label(frame_top, text="Ambiente:", font=("Roboto", 12)).grid(row=0, column=0, padx=5, pady=5)
combo_ambiente = ttk.Combobox(frame_top, textvariable=var_tipo_ambiente, values=["Produção", "Homologação"], state="readonly", font=("Roboto", 12))
combo_ambiente.grid(row=0, column=1, padx=5, pady=5)
botao_consultar = ttk.Button(frame_top, text="Consultar", command=primeira_consulta, bootstyle="primary")
botao_consultar.grid(row=0, column=2, padx=5, pady=5)
botao_instrucoes = ttk.Button(frame_top, text="Instruções", command=exibir_mensagem, bootstyle="info")
botao_instrucoes.grid(row=0, column=3, padx=5, pady=5)

##################### Treeview #####################
# Criação do Treeview com estilo
treeview = ttk.Treeview(root, columns=("Serviço", "Estado"), show="headings", height=20)
treeview.grid(row=1, column=0, columnspan=3, padx=10, pady=10)
# Configurando os cabeçalhos
treeview.heading("Serviço", text="Serviço", anchor="center")
treeview.heading("Estado", text="Estado", anchor="center")
# Configurando as colunas
treeview.column("Serviço", anchor="center", width=400)
treeview.column("Estado", anchor="center", width=200)
# Linhas alternadas (efeito zebra)
treeview.tag_configure("oddrow", background="#F9F9F9")
treeview.tag_configure("evenrow", background="#FFFFFF")
# Estilo do Treeview
style = ttk.Style()
style.configure("Treeview", font=("Roboto", 12), rowheight=16)  # Fonte e altura das linhas
style.configure("Treeview.Heading", font=("Roboto", 12, "bold"))  # Fonte dos cabeçalhos
style.map("Treeview", background=[("selected", "#DCE4F2")], foreground=[("selected", "#000000")]) # Efeito de seleção da linha

##################### Barra de seleção "Combobox" #####################
combobox_servicos = ttk.Combobox(root, state="readonly", font=("Roboto", 12))
combobox_servicos.grid(row=2, column=0, padx=5, pady=5, columnspan=3)
combobox_servicos.config(width=40)

##################### Botões Iniciar e Parar #####################
botao_iniciar = ttk.Button(root, text="Iniciar", command=iniciar_servicos, bootstyle="success", width=20)
botao_iniciar.grid(row=3, column=0, padx=5, pady=5)
botao_parar = ttk.Button(root, text="Parar", command=parar_servicos, bootstyle="danger", width=20)
botao_parar.grid(row=3, column=2, padx=5, pady=5)

# Exibir a janela principal
root.mainloop()
