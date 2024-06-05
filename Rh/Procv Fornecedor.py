import pandas as pd
import tkinter as tk
from tkinter import filedialog

def atualizar_folha(folha_path, cadastro_path, output_path):
    # Carregar os arquivos Excel
    folha_df = pd.read_excel(folha_path)
    cadastro_df = pd.read_excel(cadastro_path)

    # Renomear as colunas para facilitar a manipulação
    cadastro_df.columns = ['Cod_Fornecedores', 'Cod_Loja', 'Descricao_Fornecedores', 'Values']
    cadastro_df = cadastro_df[['Cod_Fornecedores', 'Descricao_Fornecedores']]

    # Remover linhas que não possuem valores válidos e duplicatas
    cadastro_df = cadastro_df.dropna(subset=['Cod_Fornecedores', 'Descricao_Fornecedores'])
    cadastro_df = cadastro_df.drop_duplicates(subset='Descricao_Fornecedores', keep='first')

    # Preencher a coluna 'E2_FORNECE' na folha com os códigos correspondentes
    folha_df['E2_FORNECE'] = folha_df['E2_NOMFOR'].map(cadastro_df.set_index('Descricao_Fornecedores')['Cod_Fornecedores'])

    # Salvar o DataFrame atualizado de volta em um novo arquivo Excel
    folha_df.to_excel(output_path, index=False)
    print(f"Arquivo atualizado salvo em: {output_path}")

def main():
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal

    # Solicitar que o usuário selecione o arquivo da Folha
    folha_path = filedialog.askopenfilename(title="Selecione o arquivo da Folha", filetypes=[("Excel files", "*.xlsx")])
    if not folha_path:
        print("Nenhum arquivo de Folha selecionado. Saindo...")
        return

    # Solicitar que o usuário selecione o arquivo de Cadastro de Fornecedores
    cadastro_path = filedialog.askopenfilename(title="Selecione o arquivo de Cadastro de Fornecedores", filetypes=[("Excel files", "*.xlsx")])
    if not cadastro_path:
        print("Nenhum arquivo de Cadastro selecionado. Saindo...")
        return

    # Solicitar que o usuário selecione o diretório para salvar a Folha atualizada
    output_dir = filedialog.askdirectory(title="Selecione o diretório para salvar a Folha atualizada")
    if not output_dir:
        print("Nenhum diretório selecionado. Saindo...")
        return

    output_path = f"{output_dir}/Folha_Atualizada.xlsx"

    # Executar a função
    atualizar_folha(folha_path, cadastro_path, output_path)

if __name__ == "__main__":
    main()
