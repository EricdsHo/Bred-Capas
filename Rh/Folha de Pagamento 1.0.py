import pandas as pd
import re
from tkinter import Tk, filedialog
import datetime

def convert_excel_format():
    # Esconder a janela principal do Tk
    root = Tk()
    root.withdraw()

    # Solicitando ao usuário para selecionar o arquivo Excel de origem
    file_path = filedialog.askopenfilename(
        title='Selecione o arquivo Excel',
        filetypes=[('Arquivos Excel', '*.xlsx')]
    )
    
    # Se o usuário cancelar a seleção do arquivo, o script é encerrado
    if not file_path:
        print("Seleção de arquivo cancelada.")
        return

    # Solicitando ao usuário para selecionar o diretório de destino
    destination_path = filedialog.asksaveasfilename(
        title='Salvar arquivo convertido como',
        filetypes=[('Arquivos Excel', '*.xlsx')],
        defaultextension='.xlsx'
    )
    
    # Se o usuário cancelar a seleção do destino, o script é encerrado
    if not destination_path:
        print("Seleção de destino cancelada.")
        return

    # Solicitar ao usuário os dados para os campos adicionais
    e2_num = input("Digite o valor para E2_NUM: ")
    e2_num = int(e2_num)
    e2_naturez = input("Digite o valor para E2_NATUREZ: ")
    e2_vencto = input("Digite o valor para E2_VENCTO: ")
    e2_hist = input("Digite o valor para E2_HIST: ")

    # Definindo E2_EMISSAO para a data atual
    e2_emissao = datetime.datetime.now().strftime("%d/%m/%Y")

    # Lendo o arquivo Excel
    df = pd.read_excel(file_path)

    # Preenchendo a coluna 'Loja' para baixo, para casos de células mescladas
    df['Loja'] = df['Loja'].ffill()

    # Substituindo os nomes das lojas de "BRED XX" para "01XX"
    df['Loja'] = df['Loja'].apply(lambda x: re.sub(r'BRED (\d+)', r'01\1', str(x)))

    # Substituições específicas nas colunas
    df['Loja'] = df['Loja'].replace({
        'BRED DISTRIBUIDORA': '0198',
        'ADM': '0199',
        'PRÓ - LABORE': '01ZZ'
    })

    # Removendo linhas em branco, títulos e entradas sem valor a receber
    df = df[df['Nome'].notna() & ~df['Nome'].isin(['', 'Sum', 'Nome', 'GRUPO BRED', 'SÓCIOS']) & df['Valor TOTAL (5º)'].notna() & (df['Valor TOTAL (5º)'] != 0)]

    # Criando um novo DataFrame com o formato final
    new_columns = [
        'E2_FILIAL', 'E2_PREFIXO', 'E2_NUM', 'E2_TIPO', 'E2_NATUREZ', 'E2_FORNECE',
        'E2_LOJA', 'E2_NOMFOR', 'E2_EMISSAO', 'E2_VENCTO', 'E2_VENCREA', 'E2_VALOR', 'E2_HIST'
    ]
    new_df = pd.DataFrame(columns=new_columns)
    new_df['E2_FILIAL'] = df['Loja']
    new_df['E2_PREFIXO'] = 'RHI'
    new_df['E2_NUM'] = e2_num
    new_df['E2_TIPO'] = 'TF'
    new_df['E2_NATUREZ'] = e2_naturez
    new_df['E2_LOJA'] = '0001'
    new_df['E2_NOMFOR'] = df['Nome']
    new_df['E2_VALOR'] = df['Valor TOTAL (5º)']
    new_df['E2_HIST'] = e2_hist
    new_df['E2_EMISSAO'] = e2_emissao
    new_df['E2_VENCTO'] = e2_vencto
    new_df['E2_VENCREA'] = e2_vencto

    # Atualizando o campo E2_NUM de forma progressiva para cada linha
    for index, _ in new_df.iterrows():
        new_df.at[index, 'E2_NUM'] = int(str(e2_num).zfill(9))  # Formatando para 9 dígitos
        e2_num += 1  # Incrementando para a próxima linha
    
    # Salvando o novo DataFrame em um novo arquivo Excel
    new_df.to_excel(destination_path, index=False)
    print(f'Arquivo convertido com sucesso: {destination_path}')

# Executando a função
if __name__ == '__main__':
    convert_excel_format()
    input("Pressione Enter para sair...")  # Manter o terminal aberto até que o usuário pressione Enter
