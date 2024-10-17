import pandas as pd
import tkinter as tk
from tkinter import filedialog

def selecionar_arquivo():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=[("Arquivo Excel", "*.xlsx *.xls")]
    )
    return file_path

def localizar_coluna_paciente(file_path):
    df = pd.read_excel(file_path)
    coluna_paciente = None

    for col in df.columns:
        if "paciente" in str(col).lower():
            coluna_paciente = col
            break

    if coluna_paciente:
        print(f"A coluna 'paciente' foi encontrada: {coluna_paciente}")
        return coluna_paciente
    else:
        print("A coluna 'paciente' não foi encontrada.")
        return None

def somase_por_paciente(file_path):
    df = pd.read_excel(file_path)
    coluna_paciente = localizar_coluna_paciente(file_path)
    
    if not coluna_paciente:
        print("A coluna de 'paciente' não foi encontrada no arquivo.")
        return
    
    # Define a coluna de valores com base na coluna "paciente" identificada
    if coluna_paciente == 'E':
        coluna_valor = 'K'
    elif coluna_paciente == 'B':
        coluna_valor = 'F'
    else:
        print(f"A coluna de valores não foi definida para a coluna 'paciente' encontrada em '{coluna_paciente}'.")
        return

    if coluna_valor in df.columns:
        # Realiza o somatório dos valores na coluna correspondente para cada paciente
        resultado_soma = df.groupby(coluna_paciente)[coluna_valor].sum().reset_index()
        print("\nSomatório por paciente:")
        print(resultado_soma)
        return resultado_soma
    else:
        print(f"A coluna de valores '{coluna_valor}' não foi encontrada no arquivo.")

# Execução
file_path = selecionar_arquivo()
if file_path:
    somase_por_paciente(file_path)
else:
    print("Nenhum arquivo foi selecionado.")
