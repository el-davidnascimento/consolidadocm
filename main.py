import os
import pandas as pd

pastas = [
    r'G:\.shortcut-targets-by-id\1kArAZwgCxrjbQwQOPEzeJLtMUll3VVJ7\13. Dados\13.2. RMR\13.2.2. Escolas\13.2.2.1. UNI\13.2.2.1.2. Quantitativo\13.2.2.1.2.1. Base\13.2.2.1.2.1.2. Base Escolas\13.2.2.1.2.1.2.5. Campanha de Matrícula 2024', ### UNIVERSO
#    r'G:\.shortcut-targets-by-id\1kArAZwgCxrjbQwQOPEzeJLtMUll3VVJ7\13. Dados\13.2. RMR\13.2.2. Escolas\13.2.2.2. MVM\13.2.2.2.2. Quantitativo\13.2.2.2.2.1. Base\13.2.2.2.2.1.2. Base Escolas\13.2.2.2.2.1.2.6. Campanha de Matrícula 2024', ### MULTIVERSO MESSEJANA
#    r'G:\.shortcut-targets-by-id\1kArAZwgCxrjbQwQOPEzeJLtMUll3VVJ7\13. Dados\13.2. RMR\13.2.2. Escolas\13.2.2.3. MVDA\13.2.2.3.2. Quantitativo\13.2.2.3.2.1. Base\13.2.2.3.2.1.1. Base Escolas\13.2.2.3.2.1.1.5. Campanha de Matrícula 2024' ### MULTIVERSO DAMAS
]
dados_concatenados = pd.DataFrame()
# Loop para percorrer os arquivos na pasta
for pasta in pastas:
    for arquivo in os.listdir(pasta):
        caminho_arquivo = os.path.join(pasta, arquivo)
        if os.path.isfile(caminho_arquivo) and arquivo.endswith('.xlsx'):
            dados_planilha = pd.read_excel(caminho_arquivo)
            dados_concatenados = pd.concat([dados_concatenados, dados_planilha])

caminho_destino = r'G:\.shortcut-targets-by-id\1kArAZwgCxrjbQwQOPEzeJLtMUll3VVJ7\13. Dados\13.8. Campanha de Matrícula\13.8.1. Auxiliar\BASE MANUAL.xlsx'  # Substitua pelo caminho desejado
# Salvar os dados concatenados em uma única planilha no arquivo Excel no caminho especificado
dados_concatenados.to_excel(caminho_destino, sheet_name='MATREMAT SERIE', index=False)
