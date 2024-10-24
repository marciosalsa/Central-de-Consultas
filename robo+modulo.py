import pyperclip
import pyautogui
import time
import threading
import pygetwindow as gw
import tkinter as tk
import keyboard
from tkinter import Menu
import processador_excel
import comparador_excel
from comparador_excel import iniciar_interface

parar_robo = False
pausado = False  
iniciado = False  
index_atual = 0 
numeros = []
tempo_espera = 2 
lock = threading.Lock() 

def monitorar_teclas():
    """Função para monitorar as teclas Esc e P."""
    global parar_robo, root
    while True:
        if keyboard.is_pressed('esc'):
            parar_robo = True
            atualizar_display("Tecla ESC pressionada. \nParando o robô e fechando o programa...")
            root.quit()
            break
        if keyboard.is_pressed('p'):
            pausar_robo()
            time.sleep(0.5)  # Pequena pausa para evitar múltiplas ativações
        time.sleep(0.1)

# Inicia o monitoramento das teclas Esc e P em uma thread separada
threading.Thread(target=monitorar_teclas, daemon=True).start()

def abrir_comparador_excel():
    """Abre a interface do comparador de Excel."""
    comparador_excel.iniciar_interface()  


def abrir_processador_excel():
    file_path = processador_excel.selecionar_arquivo()
    if file_path:
        processador_excel.analisar_primeira_coluna(file_path, exibir_mensagem=atualizar_display)
                
    else:
        atualizar_display("Nenhum arquivo foi selecionado.")

def atualizar_display(mensagem):
    """Adiciona uma mensagem ao widget de texto e rola para baixo."""
    if status_text: 
        status_text.config(state=tk.NORMAL)
        status_text.insert(tk.END, mensagem + "\n")
        status_text.see(tk.END)
        status_text.config(state=tk.DISABLED)
    else:
        print("Erro: Widget status_text não encontrado.")

def atualizar_tempo_label():
    """Atualiza o texto do rótulo com o tempo de espera atual."""
    tempo_label.config(text=str(tempo_espera) + "s")  

def incrementar_tempo():
    """Aumenta o tempo de espera em 1 segundo e atualiza o rótulo."""
    global tempo_espera
    with lock:
        if tempo_espera < 10:  
            tempo_espera += 1
            atualizar_display(f"Tempo de espera: {tempo_espera} segundos.")
            atualizar_tempo_label()  
        else:
            atualizar_display("O timer máximo é 10 segundos")  

def decrementar_tempo():
    """Diminui o tempo de espera em 1 segundo, mínimo 1 segundo, e atualiza o rótulo.""" 
    global tempo_espera
    with lock:
        if tempo_espera > 1:
            tempo_espera -= 1
            atualizar_display(f"Tempo de espera: {tempo_espera} segundos.")
            atualizar_tempo_label()  
        else:
            atualizar_display("O timer mínimo é 1 segundo")  

def reiniciar_robo():
    global iniciado, index_atual, pausado, parar_robo, numeros
    iniciado = False
    index_atual = 0
    pausado = False
    parar_robo = False
    atualizar_display("Reiniciando o robô.")
    numeros = ler_numeros_do_arquivo()  # Recarrega todos os números do arquivo

def ler_numeros_do_arquivo():
    with open('numeros.txt', 'r') as file:
        return [num.strip() for num in file.readlines()]

def copiar_numero(numero):
    pyperclip.copy(numero)

def trocar_janela():
    nomenclatura = "agendamento"
    janelas = gw.getAllTitles()
    
    for janela in janelas:
        if janela.startswith(nomenclatura):
            janela_encontrada = gw.getWindowsWithTitle(janela)[0]
            janela_encontrada.activate()
            return
    atualizar_display(f"Nenhuma janela encontrada que comece com '{nomenclatura}'.")

def clicar_em_posicao(x, y):
    pyautogui.click(x=x, y=y)

def apagar_num_anterior():
    for _ in range(10):
        pyautogui.press('backspace')  

def colar_numero():
    pyautogui.hotkey('ctrl', 'v')

def apertar_enter():
    pyautogui.press('enter')

def executar_processo():
    global index_atual, pausado, parar_robo

    trocar_janela()

    while not parar_robo:
        if pausado:
            time.sleep(1)
            continue

        if index_atual >= len(numeros):
            atualizar_display("Todos os números processados.")
            return

        numero = numeros[index_atual]

        if numero == '0' or numero.lower() == 'null' or numero == '':
            atualizar_display("Número inválido encontrado.\nRobô pausado.")
            pausado = True
            index_atual += 1
            continue

        atualizar_display(f"Número lido: {numero}")

        copiar_numero(numero)        
        clicar_em_posicao(240, 89)        
        apagar_num_anterior()        
        colar_numero()        
        apertar_enter()
        time.sleep(tempo_espera)
        apertar_enter()        
        apagar_num_anterior()        
        apertar_enter()        
        apagar_num_anterior()
        
        with lock:
            index_atual += 1
        

def iniciar_robo():
    global iniciado, index_atual
    if not iniciado:  
        iniciado = True
        btn_iniciar.config(state=tk.DISABLED)
        atualizar_display("Robô iniciado.")  
        global numeros
        numeros = ler_numeros_do_arquivo()
        index_atual = 0
        threading.Thread(target=executar_processo, daemon=True).start()  
    else:
        atualizar_display("Robô já está em execução.")  

def pausar_robo(event=None):
    global pausado
    pausado = not pausado
    estado = "pausado" if pausado else "despausado"
    atualizar_display(f"Robô {estado}.")

def iniciar_interface():
    global root, status_text, btn_iniciar, tempo_label
    root = tk.Tk()
    root.geometry("290x352")
    root.title("Controle do Robô")

    menu_bar = Menu(root)
    root.config(menu=menu_bar)

    arquivo_menu = Menu(menu_bar, tearoff=0)
    menu_bar.add_cascade(label="Arquivo", menu=arquivo_menu)
    arquivo_menu.add_command(label="Processar Excel", command=abrir_processador_excel)
    arquivo_menu.add_command(label="Comparar Excel", command=abrir_comparador_excel)
     
    root.attributes("-topmost", True)

    btn_iniciar = tk.Button(root, text="Iniciar", command=iniciar_robo)
    btn_iniciar.pack(pady=10)

    btn_pausar = tk.Button(root, text="Pausar/Despausar", command=pausar_robo)
    btn_pausar.pack(pady=10)

    btn_reiniciar = tk.Button(root, text="Reiniciar", command=reiniciar_robo)
    btn_reiniciar.pack(pady=10)

    frame_ajuste = tk.Frame(root)
    frame_ajuste.pack(pady=5)

    tempo_texto_label = tk.Label(frame_ajuste, text="Timer:", font=("Arial", 9))
    tempo_texto_label.pack(side=tk.LEFT, padx=1)

    tempo_label = tk.Label(frame_ajuste, font=("Arial", 9), width=5)
    tempo_label.pack(side=tk.LEFT, padx=1)
    atualizar_tempo_label()

    btn_decrementar = tk.Button(frame_ajuste, text="↓", command=decrementar_tempo, font=("Arial", 7), width=3)
    btn_decrementar.pack(side=tk.LEFT, padx=3)

    btn_incrementar = tk.Button(frame_ajuste, text="↑", command=incrementar_tempo, font=("Arial", 7), width=3)
    btn_incrementar.pack(side=tk.LEFT, padx=3)

    status_text = tk.Text(root, height=10, width=40)
    status_text.pack(pady=10)    
   
    root.protocol("WM_DELETE_WINDOW", root.quit)
    root.mainloop()

iniciar_interface()
