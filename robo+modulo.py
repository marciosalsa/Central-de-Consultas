import pyperclip
import pyautogui
import time
import threading
import pygetwindow as gw
import tkinter as tk
import keyboard
import pandas as pd
from tkinter import filedialog, Menu
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import re
from tkinter import messagebox
import processador_excel


# Variáveis globais
parar_robo = False
pausado = False  
iniciado = False  
index_atual = 0  # Para rastrear o índice atual do número
numeros = []
tempo_espera = 2  # Tempo de espera inicial em segundos

def abrir_processador_excel():
    # Selecionar o arquivo e processar com o módulo
    file_path = processador_excel.selecionar_arquivo()  # Usa a função para selecionar o arquivo
    if file_path:
        processador_excel.analisar_primeira_coluna(file_path)  # Chama a função do módulo

def atualizar_display(mensagem):
    """Adiciona uma mensagem ao widget de texto e rola para baixo."""
    status_text.config(state=tk.NORMAL)  # Habilita a edição
    status_text.insert(tk.END, mensagem + "\n")
    status_text.see(tk.END)  # Rolar para o final
    status_text.config(state=tk.DISABLED)  # Desabilita a edição novamente

def atualizar_tempo_label():
    """Atualiza o texto do rótulo com o tempo de espera atual."""
    tempo_label.config(text=str(tempo_espera) + "s")  # Atualiza o rótulo com a unidade "s"

def incrementar_tempo():
    """Aumenta o tempo de espera em 1 segundo e atualiza o rótulo."""
    global tempo_espera
    if tempo_espera < 10:  # Verifica se o tempo atual é menor que 15
        tempo_espera += 1
        atualizar_display(f"Tempo de espera: {tempo_espera} segundos.")
        atualizar_tempo_label()  # Atualiza o rótulo
    else:
        atualizar_display("O timer máximo é 10 segundos")  # Mensagem informativa

def decrementar_tempo():
    """Diminui o tempo de espera em 1 segundo, mínimo 1 segundo, e atualiza o rótulo."""
    global tempo_espera
    if tempo_espera > 1:
        tempo_espera -= 1
        atualizar_display(f"Tempo de espera: {tempo_espera} segundos.")
        atualizar_tempo_label()  # Atualiza o rótulo    
    else:
        atualizar_display("O timer mínimo é 1 segundo")  # Mensagem informativa

def reiniciar_robo():
    global iniciado, index_atual, pausado, parar_robo, numeros

    # Reseta as variáveis globais
    iniciado = False
    index_atual = 0
    pausado = False
    parar_robo = False
    
    atualizar_display("Reiniciando o robô.")

    # Lê os números novamente
    numeros = ler_numeros_do_arquivo()  # Recarrega todos os números do arquivo
    

# Função para ler os números do arquivo
def ler_numeros_do_arquivo():
    with open('numeros.txt', 'r') as file:
        return [num.strip() for num in file.readlines()]

# Função para copiar o número
def copiar_numero(numero):
    pyperclip.copy(numero)

# Função para focar na janela "agendamento"
def trocar_janela():
    nomenclatura = "agendamento"
    janelas = gw.getAllTitles()
    
    for janela in janelas:
        if janela.startswith(nomenclatura):
            janela_encontrada = gw.getWindowsWithTitle(janela)[0]
            janela_encontrada.activate()
            return
    atualizar_display(f"Nenhuma janela encontrada que comece com '{nomenclatura}'.")

# Função para clicar em uma posição específica (coordenadas x, y)
def clicar_em_posicao(x, y):
    pyautogui.click(x=240, y=89)

def apagar_num_anterior():
    for _ in range(10):  # Aperte 'backspace' 10 vezes
        pyautogui.press('backspace')  

# Função para colar o número
def colar_numero():
    pyautogui.hotkey('ctrl', 'v')

# Função para apertar Enter
def apertar_enter():
    pyautogui.press('enter')

# Função principal para executar o processo
def executar_processo():
    global index_atual, pausado, parar_robo
    
    # Inicia a thread para monitorar a tecla Esc
    monitorar_thread = threading.Thread(target=monitorar_teclas, daemon=True)
    monitorar_thread.start()

    trocar_janela()

    while not parar_robo:
        # Enquanto estiver pausado, espera
        if pausado:
            time.sleep(1)  # Aguarda 1 segundo antes de verificar novamente
            continue

        # Se todos os números foram processados
        if index_atual >= len(numeros):
            atualizar_display("Todos os números processados.")
            return

        numero = numeros[index_atual]  # Pega o número atual

        # Verifica se o número é '0', 'null' ou está vazio
        if numero == '0' or numero.lower() == 'null' or numero == '':
            atualizar_display("Número inválido encontrado.\nRobô pausado.")
            pausado = True
            index_atual += 1  # Move para o próximo número
            continue

        atualizar_display(f"Número lido: {numero}")

        # Copia o número e processa
        copiar_numero(numero)              
        time.sleep(0.1)
        clicar_em_posicao(240, 89)
        time.sleep(0.1)
        apagar_num_anterior()
        time.sleep(0.1)
        colar_numero()
        time.sleep(0.1)
        apertar_enter()
        time.sleep(2)
        apertar_enter()
        time.sleep(0.1)
        apagar_num_anterior()
        

        # Incrementa o índice
        index_atual += 1
        time.sleep(tempo_espera)

def iniciar_robo():
    global iniciado, index_atual
    if not iniciado:  # Verifica se o robô ainda não foi iniciado
        iniciado = True
        btn_iniciar.config(state=tk.DISABLED)  # Desabilita o botão após a primeira execução
        atualizar_display("Robô iniciado.")  # Atualiza o display
        global numeros
        numeros = ler_numeros_do_arquivo()  # Lê os números
        index_atual = 0  # Reseta o índice
        threading.Thread(target=executar_processo, daemon=True).start()  # Executa o processo em uma thread separada
    else:
        atualizar_display("Robô já está em execução.")  # Mensagem informativa se o robô já estiver rodando

def pausar_robo():
    global pausado
    pausado = not pausado
    estado = "pausado" if pausado else "despausado"
    atualizar_display(f"Robô {estado}.")

def monitorar_teclas():
    global parar_robo, root
    while True:
        if keyboard.is_pressed('esc'):
            parar_robo = True
            atualizar_display("Tecla ESC pressionada. \nParando o robô e fechando o programa...")
            root.quit()  # Encerra o loop principal do Tkinter
            root.destroy()  # Fecha a janela
            break
        if keyboard.is_pressed('p'):
            pausar_robo()
            time.sleep(0.1)  # Aguarda um pouco para evitar múltiplas chamadas
        time.sleep(0.1)  # Pequeno delay para não sobrecarregar o loop

# Função para iniciar a interface Tkinter
def iniciar_interface():
    global root, status_text, btn_iniciar, tempo_label
    root = tk.Tk()
    root.geometry("250x300")
    root.title("Controle do Robô")

    # Criar um menu
    menu_bar = Menu(root)
    root.config(menu=menu_bar)

    # Adicionar um menu "Arquivo"
    arquivo_menu = Menu(menu_bar, tearoff=0)
    menu_bar.add_cascade(label="Arquivo", menu=arquivo_menu)

    # Adicionar um item ao menu "Arquivo"
    arquivo_menu.add_command(label="Processar Excel", command=abrir_processador_excel)

    

     # Definir a janela como sempre no topo
    root.attributes("-topmost", True)

    btn_iniciar = tk.Button(root, text="Iniciar", command=iniciar_robo)
    btn_iniciar.pack(pady=10)

    btn_pausar = tk.Button(root, text="Pausar/Despausar", command=pausar_robo)
    btn_pausar.pack(pady=10)

    btn_reiniciar = tk.Button(root, text="Reiniciar", command=reiniciar_robo)
    btn_reiniciar.pack(pady=10)

    # Frame para os botões de ajuste de tempo e o rótulo
    frame_ajuste = tk.Frame(root)
    frame_ajuste.pack(pady=5)

    # Rótulo indicando "Timer:"
    tempo_texto_label = tk.Label(frame_ajuste, text="Timer:", font=("Arial", 9))
    tempo_texto_label.pack(side=tk.LEFT, padx=1)

    # Rótulo para mostrar o tempo de espera atual
    tempo_label = tk.Label(frame_ajuste, font=("Arial", 9), width=5)
    tempo_label.pack(side=tk.LEFT, padx=1)
    atualizar_tempo_label()  # Inicializa o rótulo com o valor atual

    # Botão de diminuir tempo (seta para baixo)
    btn_decrementar = tk.Button(frame_ajuste, text="↓", command=decrementar_tempo, font=("Arial", 7), width=3)
    btn_decrementar.pack(side=tk.LEFT, padx=3)

    # Botão de aumentar tempo (seta para cima)
    btn_incrementar = tk.Button(frame_ajuste, text="↑", command=incrementar_tempo, font=("Arial", 7), width=3)
    btn_incrementar.pack(side=tk.LEFT, padx=3)

    # Adiciona um widget Text para exibir informações do robô
    status_text = tk.Text(root, height=10, width=40)
    status_text.pack(pady=10)

    # Adiciona o atalho para pausar o robô
    root.bind('<p>', pausar_robo)

    # Torna o widget de texto não editável
    status_text.config(state=tk.DISABLED)  # Desativa a edição

    root.protocol("WM_DELETE_WINDOW", root.quit)  # Fecha o aplicativo corretamente
    root.mainloop()

    # Inicia a thread para monitorar a tecla Esc
    monitorar_thread = threading.Thread(target=monitorar_teclas, daemon=True)
    monitorar_thread.start()

# Inicia a interface
iniciar_interface()

# Propriedade de Marcio Ferreira Salsa (04/10/2024)

