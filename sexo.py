import pandas as pd
from gender_guesser_br import Genero # Importa a biblioteca brasileira
import re

# --- CONFIGURAÇÃO ---
nome_arquivo_excel = 'Nomes_e_Sexo_Inferido.xlsx'
coluna_nome = 'Nome'
coluna_sexo = 'Sexo'
# --------------------

# 1. Carregar o arquivo Excel
try:
    df = pd.read_excel(nome_arquivo_excel)
except FileNotFoundError:
    print(f"Erro: O arquivo '{nome_arquivo_excel}' não foi encontrado.")
    exit()

# 2. Definir a condição de re-tratamento (Desconhecido ou Ambos)
condicao_tratar = (df[coluna_sexo] == 'Desconhecido') | (df[coluna_sexo] == 'Ambos')
df_a_tratar = df[condicao_tratar].copy()

if df_a_tratar.empty:
    print("Nenhum nome 'Desconhecido' ou 'Ambos' restante. Processo finalizado.")
    exit()

total_inicial_problemas = len(df_a_tratar)
print(f"Iniciando tentativa final com gender-guesser-br para {total_inicial_problemas} nomes.")
print("ATENÇÃO: Este processo pode ser lento ou travar devido à dependência da API do IBGE.")

# 3. Funções de Limpeza de Nome (Reutilizadas da última melhoria)
def extrair_primeiro_nome(nome_completo):
    """Extrai e limpa o primeiro nome da string completa."""
    if pd.isna(nome_completo):
        return None
    nome = str(nome_completo).strip()
    # Limpeza agressiva: remove tudo que não for letra ou espaço
    nome_limpo_completo = re.sub(r'[^a-zA-Z\s]', ' ', nome) 
    palavras = [p for p in nome_limpo_completo.split() if p]
    if not palavras:
        return None
    return palavras[0].capitalize()

# 4. Função de Inferência com 'gender-guesser-br'
def inferir_sexo_br(primeiro_nome):
    """Inferir o sexo usando a biblioteca gender-guesser-br (IBGE)."""
    if not primeiro_nome:
        return None 
    
    # O pacote brasileiro requer instanciar a classe
    try:
        # A API do IBGE pode falhar ou demorar. O try/except é crucial aqui.
        resultado = Genero(primeiro_nome)() 
        
        # Traduz os resultados: 'masculino', 'feminino' ou 'Não Encontrado'
        if resultado == 'masculino':
            return 'Masculino'
        elif resultado == 'feminino':
            return 'Feminino'
        else: # 'Não Encontrado'
            return None # Retorna None para indicar que não houve sucesso
            
    except Exception as e:
        # Captura erros de API, timeout, etc.
        # print(f"Erro ao consultar nome {primeiro_nome}: {e}") # Descomente para debug
        return None 

# 5. Aplicação e Atualização
# Cria a coluna com o primeiro nome limpo SÓ nos nomes a serem tratados
df_a_tratar['Primeiro_Nome'] = df_a_tratar[coluna_nome].apply(extrair_primeiro_nome)

# Aplica a função de inferência (a parte lenta/instável)
novos_sexos = df_a_tratar['Primeiro_Nome'].apply(inferir_sexo_br)

# 6. ATUALIZAÇÃO DO DATAFRAME ORIGINAL
# Filtrar apenas os resultados que não são None (ou seja, onde houve sucesso na API)
sucesso_br = novos_sexos.dropna()

# Atualiza a coluna 'Sexo' do DataFrame original (df) com os novos valores,
# usando o índice da série 'sucesso_br' para garantir que SÓ os sucessos sejam atualizados.
df.loc[sucesso_br.index, coluna_sexo] = sucesso_br

# 7. Salvar o arquivo atualizado
df.to_excel(nome_arquivo_excel, index=False)
print("\nArquivo atualizado com a tentativa final de tratamento (gender-guesser-br) salvo.")

# 8. Imprimir os novos totais
novos_totais = df[coluna_sexo].value_counts()
print("\n--- Totais Finais por Categoria de Sexo ---")
print(novos_totais)

# Cálculo da melhoria
total_final_problemas = novos_totais.get('Desconhecido', 0) + novos_totais.get('Ambos', 0)
melhorados = total_inicial_problemas - total_final_problemas

print(f"\nTotal inicial de problemas ('Desconhecido' + 'Ambos'): {total_inicial_problemas}")
print(f"Total final de problemas ('Desconhecido' + 'Ambos'): {total_final_problemas}")
print(f"Registros classificados com sucesso por gender-guesser-br: {melhorados}")