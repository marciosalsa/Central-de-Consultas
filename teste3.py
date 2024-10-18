import pandas as pd
import tkinter as tk
from tkinter import filedialog
from fuzzywuzzy import fuzz
from fuzzywuzzy import process


def selecionar_arquivo():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=[("Arquivo Excel", "*.xlsx *.xls")]
    )
    return file_path

def localizar_coluna_paciente(file_path, max_attempts=5):
    for skiprows in range(max_attempts): 
        df = pd.read_excel(file_path, skiprows=skiprows)
        coluna_paciente_idx = None
        
        # Verificar as colunas do DataFrame atual
        for col in df.columns:
            if "paciente" in str(col).lower():
                coluna_paciente_idx = df.columns.get_loc(col)  # Obter o índice da coluna
                break

        if coluna_paciente_idx is not None:
            # Condições para definir coluna_valor_idx
            if coluna_paciente_idx == 4:
                coluna_valor_idx = 10
            elif coluna_paciente_idx == 1:
                coluna_valor_idx = 5
            else:
                print("Planilha inválida.")
                return None
            
            print(f"A coluna 'paciente' foi encontrada na tentativa {skiprows + 1}: {coluna_paciente_idx}")
            print(f"Índice da coluna de valor: {coluna_valor_idx}")
            return coluna_paciente_idx, coluna_valor_idx  # Retorna ambos os índices

    print("A coluna 'paciente' não foi encontrada em nenhuma das tentativas.")
    return None

def somar_linhas_por_nome(file_path, coluna_paciente_idx, coluna_valor_idx):
    
    if coluna_paciente_idx == 1:
        df = pd.read_excel(file_path, skiprows=4)  
    else:
        df = pd.read_excel(file_path)  
    
    # Obter o nome das colunas com base nos índices
    coluna_paciente = df.columns[coluna_paciente_idx]
    coluna_valor = df.columns[coluna_valor_idx]
    
    # Remover espaços em branco dos nomes dos pacientes
    df[coluna_paciente] = df[coluna_paciente].str.strip()

    # Agrupar pela coluna "paciente" e somar os valores correspondentes
    soma_por_paciente = df.groupby(coluna_paciente)[coluna_valor].sum().reset_index()

    # Exibir o resultado
    print("Soma por paciente:")
    print(soma_por_paciente)
    
    # Exibir os nomes das colunas do DataFrame resultante
    print("Colunas do DataFrame resultante:", soma_por_paciente.columns.tolist())
    
    # Renomear a coluna "paciente" para garantir consistência
    soma_por_paciente.rename(columns={coluna_paciente: 'paciente'}, inplace=True)

    return soma_por_paciente

def comparar_dataframes(df1, df2, tolerancia=0.02):
    # Realiza uma mesclagem (join) dos DataFrames para comparar
    comparacao = pd.merge(df1, df2, on='paciente', how='outer', suffixes=('_1', '_2'))
    comparacao['diferenca'] = comparacao.iloc[:, 1] - comparacao.iloc[:, 2]  # Usando índices para referenciar as colunas
    
    # Filtrar as diferenças que não são iguais a zero ou que não estão dentro da tolerância
    diferencas = comparacao[(comparacao['diferenca'].isna()) | (comparacao['diferenca'] > tolerancia)]
    
    return diferencas
# Execução
file_path1 = selecionar_arquivo()
if file_path1:
    resultado1 = localizar_coluna_paciente(file_path1)
    if resultado1:
        coluna_paciente_idx_1, coluna_valor_idx_1 = resultado1
        
        # Executar a soma por paciente na primeira planilha
        resultado_soma1 = somar_linhas_por_nome(file_path1, coluna_paciente_idx_1, coluna_valor_idx_1)

        # Selecionar a segunda planilha
        file_path2 = selecionar_arquivo()
        if file_path2:
            resultado2 = localizar_coluna_paciente(file_path2)
            if resultado2:
                coluna_paciente_idx_2, coluna_valor_idx_2 = resultado2
                
                # Executar a soma por paciente na segunda planilha
                resultado_soma2 = somar_linhas_por_nome(file_path2, coluna_paciente_idx_2, coluna_valor_idx_2)
                
                # Comparar os dois DataFrames
                comparacao_resultado = comparar_dataframes(resultado_soma1, resultado_soma2)

                # Filtrar apenas as diferenças diferentes de zero
                diferencas = comparacao_resultado[comparacao_resultado['diferenca'] != 0]

                # Exibir os resultados da comparação
                print("Comparação entre as duas planilhas (diferenças):")
                print(diferencas)

                # Salvar as diferenças em um novo arquivo Excel
                output_path_diff = filedialog.asksaveasfilename(
                    title="Salvar resultados das diferenças",
                    defaultextension=".xlsx",
                    filetypes=[("Arquivo Excel", "*.xlsx")]
                )
                if output_path_diff:
                    diferencas.to_excel(output_path_diff, index=False)
                    print(f"Arquivo Excel das diferenças salvo com sucesso em {output_path_diff}")
else:
    print("Nenhum arquivo foi selecionado.")
    