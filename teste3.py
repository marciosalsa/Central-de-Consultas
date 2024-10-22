import pandas as pd
from tkinter import Tk, Button, Label, filedialog

def selecionar_arquivo1():
    global arquivo1
    arquivo1 = filedialog.askopenfilename(title="Selecione a primeira planilha", filetypes=[("Excel files", "*.xlsx;*.xls")])
    label_arquivo1.config(text=arquivo1)

def selecionar_arquivo2():
    global arquivo2
    arquivo2 = filedialog.askopenfilename(title="Selecione a segunda planilha", filetypes=[("Excel files", "*.xlsx;*.xls")])
    label_arquivo2.config(text=arquivo2)

def processar_planilhas():
    # Carregar os DataFrames das planilhas selecionadas
    df1 = pd.read_excel(arquivo1, usecols=[0, 4, 10])  
    df2 = pd.read_excel(arquivo2, usecols=[0, 1, 5], skiprows=4)  

    # Renomear colunas
    df1.columns = ['Guia', 'Paciente', 'Valor']
    df2.columns = ['Guia', 'Paciente', 'Valor']

    # Remover espaços extras dos nomes dos pacientes
    df1['Paciente'] = df1['Paciente'].str.strip()
    df2['Paciente'] = df2['Paciente'].str.strip()

    # Agrupar pacientes por nome e somar os valores
    df1 = df1.groupby(['Paciente'], as_index=False).agg({'Valor': 'sum'})  
    df2 = df2.groupby(['Paciente'], as_index=False).agg({'Valor': 'sum'})

    # Merge dos dois DataFrames com base nos pacientes
    df_merge = pd.merge(df1, df2, how='outer', on='Paciente', suffixes=('_df1', '_df2'))
    df_merge['diferenca'] = df_merge['Valor_df1'].fillna(0) - df_merge['Valor_df2'].fillna(0)

    # Definir uma tolerância para as diferenças
    tolerancia = 0.02
    df_diferencas = df_merge[(df_merge['diferenca'].abs() > tolerancia) | (df_merge['diferenca'].isna())]
 
    df_diferencas.to_excel('df1_teste.xlsx', index=False)
    
         
    
    label_resultado.config(text="Planilhas processadas e diferenças salvas em 'df1_teste.xlsx'.")



# Configuração da interface Tkinter
root = Tk()
root.title("Processador de Planilhas")

# Botões para selecionar arquivos
botao_arquivo1 = Button(root, text="Selecionar Planilha Fontana", command=selecionar_arquivo1)
botao_arquivo1.pack(pady=10)

label_arquivo1 = Label(root, text="Nenhum arquivo selecionado")
label_arquivo1.pack()

botao_arquivo2 = Button(root, text="Selecionar Planilha Agendamento", command=selecionar_arquivo2)
botao_arquivo2.pack(pady=10)

label_arquivo2 = Label(root, text="Nenhum arquivo selecionado")
label_arquivo2.pack()

# Botão para processar as planilhas
botao_processar = Button(root, text="Processar Planilhas", command=processar_planilhas)
botao_processar.pack(pady=20)

# Label para exibir resultados
label_resultado = Label(root, text="")
label_resultado.pack(pady=10)

root.mainloop()
