import pandas as pd
from tkinter import Tk, Button, Label, filedialog
from fuzzywuzzy import process

# Variáveis globais
arquivo1 = None
arquivo2 = None
df_final = None  # Inicializar df_final globalmente
label_arquivo1 = None
label_arquivo2 = None
label_resultado = None

# Funções para seleção de arquivos
def selecionar_arquivo1():
    global arquivo1
    arquivo1 = filedialog.askopenfilename(title="Selecione a primeira planilha", filetypes=[("Excel files", "*.xlsx;*.xls")])
    label_arquivo1.config(text=arquivo1)

def selecionar_arquivo2():
    global arquivo2
    arquivo2 = filedialog.askopenfilename(title="Selecione a segunda planilha", filetypes=[("Excel files", "*.xlsx;*.xls")])
    label_arquivo2.config(text=arquivo2)

# Função para localizar similaridades
def localizar_similaridades(df):
    global df_final
    df_copy = df.copy()

    for index, row in df.iterrows():
        paciente = row['Paciente']
        valor_df1 = row['Valor_df1']
        valor_df2 = row['Valor_df2']

        if pd.isna(valor_df1) or pd.isna(valor_df2):
            if pd.isna(valor_df2):
                match = process.extractOne(paciente, df_copy[df_copy['Valor_df1'].isna()]['Paciente'])
            elif pd.isna(valor_df1):
                match = process.extractOne(paciente, df_copy[df_copy['Valor_df2'].isna()]['Paciente'])

            if match and match[1] >= 90:
                matched_row = df_copy[df_copy['Paciente'] == match[0]]
                if not matched_row.empty:
                    if pd.isna(valor_df1):
                        df.at[index, 'Valor_df1'] = matched_row['Valor_df1'].values[0]
                    if pd.isna(valor_df2):
                        df.at[index, 'Valor_df2'] = matched_row['Valor_df2'].values[0]
                    df.at[index, 'diferenca'] = df.at[index, 'Valor_df1'] - df.at[index, 'Valor_df2']

    df_final = df.groupby('Paciente', as_index=False).agg({
        'Valor_df1': 'sum',
        'Valor_df2': 'sum',
        'diferenca': 'sum'
    })

    df_final = df_final[df_final['diferenca'] != 0]
    return df_final.reset_index(drop=True)

# Funções para filtrar DataFrames
def filtrar_df1_por_df_final():
    global df_final  
    df1 = pd.read_excel(arquivo1, usecols=[0, 4, 9, 10])  
    df1.columns = ['Guia1', 'Paciente', 'Procedimento1', 'Valor1']
    pacientes_final = df_final['Paciente'].astype(str).tolist()  
    indices_para_excluir = []

    for index, row in df1.iterrows():
        paciente_str = str(row['Paciente']).strip()  
        if pd.isna(paciente_str) or paciente_str == "":  
            indices_para_excluir.append(index)
            continue

        match = process.extractOne(paciente_str, pacientes_final, score_cutoff=90)
        if match is None:  
            indices_para_excluir.append(index)

    df1_filtrado = df1.drop(indices_para_excluir).reset_index(drop=True)
    df1_filtrado['Paciente'] = df1_filtrado['Paciente'].str.strip()
    df1_filtrado = df1_filtrado.sort_values(by=['Paciente', 'Procedimento1']).reset_index(drop=True)
    return df1_filtrado

def filtrar_df2_por_df_final():
    global df_final  
    df2 = pd.read_excel(arquivo2, usecols=[0, 1, 2, 5], skiprows=4)
    df2.columns = ['Guia2', 'Paciente', 'Procedimento2', 'Valor2']
    pacientes_final = df_final['Paciente'].astype(str).tolist()
    indices_para_excluir = []

    for index, row in df2.iterrows():
        paciente_str = str(row['Paciente']).strip()
        if pd.isna(paciente_str) or paciente_str == "":
            indices_para_excluir.append(index)
            continue

        match = process.extractOne(paciente_str, pacientes_final, score_cutoff=90)
        if match is None:
            indices_para_excluir.append(index)

    df2_filtrado = df2.drop(indices_para_excluir).reset_index(drop=True)
    df2_filtrado['Paciente'] = df2_filtrado['Paciente'].str.strip()
    df2_filtrado = df2_filtrado.sort_values(by=['Paciente', 'Procedimento2']).reset_index(drop=True)
    return df2_filtrado

# Função principal de processamento
def processar_planilhas():
    global df_final  
    df1 = pd.read_excel(arquivo1, usecols=[0, 4, 10])
    df2 = pd.read_excel(arquivo2, usecols=[0, 1, 5], skiprows=4)
    df1.columns = ['Guia', 'Paciente', 'Valor']
    df2.columns = ['Guia', 'Paciente', 'Valor']
    df1['Paciente'] = df1['Paciente'].str.strip()
    df2['Paciente'] = df2['Paciente'].str.strip()

    df1 = df1.groupby(['Paciente'], as_index=False).agg({'Valor': 'sum'})
    df2 = df2.groupby(['Paciente'], as_index=False).agg({'Valor': 'sum'})

    df_merge = pd.merge(df1, df2, how='outer', on='Paciente', suffixes=('_df1', '_df2'))
    df_merge['diferenca'] = df_merge['Valor_df1'].fillna(0) - df_merge['Valor_df2'].fillna(0)

    tolerancia = 0.02
    df_diferencas = df_merge[(df_merge['diferenca'].abs() > tolerancia) | (df_merge['diferenca'].isna())]
    df_diferencas = localizar_similaridades(df_diferencas)

    df1_novo = filtrar_df1_por_df_final()
    df2_novo = filtrar_df2_por_df_final()

    df_combinado = pd.concat([df1_novo, df2_novo], axis=1)

    with pd.ExcelWriter('resultado_planilhas.xlsx', engine='openpyxl') as writer:
        df_combinado.to_excel(writer, sheet_name='Dados Combinados', index=False)
        df_final.to_excel(writer, sheet_name='df_final', index=False)

    label_resultado.config(text="Planilhas processadas e diferenças salvas em 'resultado_planilhas.xlsx'.")
    print("Dados salvos com sucesso em 'resultado_planilhas.xlsx'.")

# Se for rodado diretamente, inicia o Tkinter
def iniciar_interface():    
    global label_arquivo1, label_arquivo2, label_resultado  # Declare as variáveis globais

    root = Tk()
    root.title("Comparador de Planilhas")

    botao_arquivo1 = Button(root, text="Selecionar Planilha Fontana", command=selecionar_arquivo1)
    botao_arquivo1.pack(pady=10)

    label_arquivo1 = Label(root, text="Nenhum arquivo selecionado")
    label_arquivo1.pack()

    botao_arquivo2 = Button(root, text="Selecionar Planilha Agendamento", command=selecionar_arquivo2)
    botao_arquivo2.pack(pady=10)

    label_arquivo2 = Label(root, text="Nenhum arquivo selecionado")
    label_arquivo2.pack()

    botao_processar = Button(root, text="Processar Planilhas", command=processar_planilhas)
    botao_processar.pack(pady=20)

    label_resultado = Label(root, text="")
    label_resultado.pack(pady=10)

    root.mainloop()        

if __name__ == "__main__":
    iniciar_interface()
