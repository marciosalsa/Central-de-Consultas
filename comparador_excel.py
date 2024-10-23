import pandas as pd
from tkinter import Tk, Button, Label, filedialog
from fuzzywuzzy import process

def selecionar_arquivo1():
    global arquivo1
    arquivo1 = filedialog.askopenfilename(title="Selecione a primeira planilha", filetypes=[("Excel files", "*.xlsx;*.xls")])
    label_arquivo1.config(text=arquivo1)

def selecionar_arquivo2():
    global arquivo2
    arquivo2 = filedialog.askopenfilename(title="Selecione a segunda planilha", filetypes=[("Excel files", "*.xlsx;*.xls")])
    label_arquivo2.config(text=arquivo2)


def localizar_similaridades(df):
    # Cria uma cópia do DataFrame para não alterar o original
    df_copy = df.copy()

    # Itera pelas linhas do DataFrame
    for index, row in df.iterrows():
        paciente = row['Paciente']
        valor_df1 = row['Valor_df1']
        valor_df2 = row['Valor_df2']

        # Verifica se Valor_df1 ou Valor_df2 está vazio
        if pd.isna(valor_df1) or pd.isna(valor_df2):
            # Se Valor_df2 está vazio, procura por correspondente com Valor_df1 vazio
            if pd.isna(valor_df2):
                match = process.extractOne(paciente, df_copy[df_copy['Valor_df1'].isna()]['Paciente'])

            # Se Valor_df1 está vazio, procura por correspondente com Valor_df2 vazio
            elif pd.isna(valor_df1):
                match = process.extractOne(paciente, df_copy[df_copy['Valor_df2'].isna()]['Paciente'])

            # Se um match for encontrado e a similaridade for maior ou igual a 80%
            if match and match[1] >= 90:
                matched_row = df_copy[df_copy['Paciente'] == match[0]]

                if not matched_row.empty:
                    # Atualiza os valores
                    if pd.isna(valor_df1):
                        df.at[index, 'Valor_df1'] = matched_row['Valor_df1'].values[0]
                    if pd.isna(valor_df2):
                        df.at[index, 'Valor_df2'] = matched_row['Valor_df2'].values[0]

                    # Calcula a diferença
                    df.at[index, 'diferenca'] = df.at[index, 'Valor_df1'] - df.at[index, 'Valor_df2']

    # Agrupar por paciente e somar os valores
    df_final = df.groupby('Paciente', as_index=False).agg({
        'Valor_df1': 'sum',
        'Valor_df2': 'sum',
        'diferenca': 'sum'
    })

    # Remover linhas onde a diferença é zero
    

    # Remover linhas duplicadas de baixo para cima
    to_drop = []
    
    # Percorre o DataFrame de baixo para cima
    for i in range(len(df_final) - 1, 0, -1):  # Começa do final e vai até a primeira linha
        if (df_final.iloc[i]['Valor_df1'] == df_final.iloc[i - 1]['Valor_df1'] and
            df_final.iloc[i]['Valor_df2'] == df_final.iloc[i - 1]['Valor_df2'] and
            df_final.iloc[i]['diferenca'] == df_final.iloc[i - 1]['diferenca']):
            # Adiciona o índice da linha atual para remoção
            to_drop.append(i)

    # Remove as linhas identificadas
    df_final = df_final.drop(to_drop, errors='ignore')

    df_final = df_final[df_final['diferenca'] != 0]

    return df_final.reset_index(drop=True)



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

    # Chama a função para localizar similaridades, agora com df_diferencas
    df_diferencas = localizar_similaridades(df_diferencas)

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
