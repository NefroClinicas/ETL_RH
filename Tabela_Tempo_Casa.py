import pandas as pd
from datetime import date
import numpy as np

# =======================================================================
# 1. CARREGUE SEU DATAFRAME PRINCIPAL
# =======================================================================
# ATENÇÃO: Altere o caminho do arquivo para o seu caso
caminho_arquivo = 'Base_MODELADA_PowerBI_V4.xlsx' 
try:
    # Tenta ler o arquivo
    df_fatos = pd.read_excel(caminho_arquivo)
    print(f"Dados carregados com sucesso de: {caminho_arquivo}")
except FileNotFoundError:
    print(f"ERRO: Arquivo não encontrado no caminho: {caminho_arquivo}")
    print("Certifique-se de que o nome do arquivo e o caminho estão corretos.")
    # Cria um DataFrame vazio para evitar quebra do script, caso ocorra erro
    df_fatos = pd.DataFrame(columns=['Pessoa_ID', 'Admissão'])


# 2. TRATAMENTO INICIAL DE COLUNAS (garantir tipos)
# Garante que a coluna 'Admissão' esteja em formato datetime. 
# O errors='coerce' transforma datas inválidas/faltantes em NaT.
df_fatos['Admissão'] = pd.to_datetime(df_fatos['Admissão'], errors='coerce').dt.normalize()

# 3. CÁLCULO REVISADO: Meses na Empresa (Mais Preciso)
# Esta lógica usa a média de dias no mês (30.4375) para calcular meses COMPLETOs,
# resolvendo o problema de precisão do dia que ocorre no cálculo ano/mês simples.
data_atual = pd.to_datetime(date.today())

dias_na_empresa = (data_atual - df_fatos['Admissão']).dt.days

# Arredonda para baixo (floor) para obter o número de meses COMPLETOS
df_fatos['MesesNaEmpresa'] = np.floor(dias_na_empresa / 30.4375).astype('Int64')

# 4. CRIAÇÃO DA COLUNA DE CATEGORIA (Lógica EXATA do Usuário)
def classificar_tempo_de_casa(meses):
    # 4.1. Trata valores NaN ou NaT (necessário para estabilidade)
    if pd.isna(meses):
        # Retorna uma etiqueta clara para dados inválidos
        return "8. Data de Admissão Inválida"

    # Conversão explícita para inteiro (Int64)
    meses = int(meses)

    # 4.2. Lógica de Classificação com Etiquetas do Usuário
    if meses < 6:
        return "1. Menos de 6 meses"
    elif 6 <= meses < 12:
        return "2. Mais de 6 meses" # 6 a 11 meses
    elif 12 <= meses < 24:
        return "3. 1 Ano"          # 12 a 23 meses
    elif 24 <= meses < 36:
        return "4. 2 Anos"          # 24 a 35 meses
    elif 36 <= meses < 48:
        return "5. 3 Anos"          # 36 a 47 meses
    elif 48 <= meses < 60:
        return "6. 4 Anos"          # 48 a 59 meses
    elif meses >= 60:
        return "7. 5 Anos ou Mais" # 60 meses ou mais
    else:
        # Caso catch-all
        return "9. Não Classificado"

df_fatos['Tempo de Casa Categoria Python'] = df_fatos['MesesNaEmpresa'].apply(classificar_tempo_de_casa)

# 5. RESULTADO FINAL (DataFrame de Saída)
# O Power BI só precisa da CHAVE e da NOVA COLUNA
if 'Pessoa_ID' in df_fatos.columns:
    df_resultado = df_fatos[['Pessoa_ID', 'Tempo de Casa Categoria Python']].copy()
    
    # Salva em CSV para que o Power BI acesse o dado
    df_resultado.to_csv('TempoDeCasa_Calculado.csv', index=False, encoding='utf-8')
    print("\nCálculo finalizado!")
    print("Arquivo 'TempoDeCasa_Calculado.csv' gerado com as categorias atualizadas conforme sua solicitação.")
else:
    print("\nAVISO: Coluna 'Pessoa_ID' não encontrada. Verifique o nome da coluna no seu DataFrame.")
