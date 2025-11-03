import pandas as pd
import os
import time
from datetime import date

# --- CONFIGURAÇÃO DE ARQUIVOS ---
ARQUIVO_ENTRADA = 'Base_BI_Consolidada2.csv'
# Alterado o nome do arquivo para V5 para refletir a nova versão corrigida
ARQUIVO_FINAL_MODELADO = 'Base_MODELADA_PowerBI_V4.xlsx' 

# LISTA DE COLUNAS DESEJADAS (Adicionado 'Sexo')
COLUNAS_DESEJADAS = [
    'Nome', 'Empresa', 'Cadastro', 'Admissão', 'Cargo', 'C.Custo', 'Descrição (C.Custo)',
    'Data Afastamento', 'Título Reduzido (Cargo)', 'Descrição (Raça/Etnia)', 'Sexo', # <-- NOVO: Adicionado 'Sexo' para a Dim_Pessoa
    'Descrição (Cat. eSocial)', 'Causa', 'Descrição (Causa)', 'Escala', 'Descrição (Escala)',
    'Filial', 'Apelido (Filial)', 'Código Fornecedor', 'Descrição (Motivo Alt. Salário)',
    'Data Adicionais', 'Data Aposentadoria', 'Data Cargo', 'Data C.Custo', 'Data Ult. Alt. Cat.',
    'Data de Chegada', 'Data Escala', 'Data Estabilidade', 'Data Escala VTR', 'Data Filial',
    'Data Histórico de Contrato', 'Data Inclusão', 'Data Local', 'Nascimento', 'Opção FGTS',
    'Data Posto', 'Data Ass. PPR', 'Data de Reintegração', 'Data Salário', 'Data Cat. SEFIP',
    'Última Simulação', 'Data Sindicato', 'Data FGTS', 'Data Vínculo', 'Cadastramento PIS',
    'Dependentes IR', 'Dependentes Saf', 'Dep. Saldo FGTS', 'Estado Civil', 'Descrição (Estado Civil)',
    'Instrução', 'Descrição (Instrução)', 'Nome (Empresa)', 'Nome (Cadastro O. Contrato)',
    'Nome (Empresa O. Contrato)', 'Descrição (Tipo O. Contrato)', '% Desempenho',
    '% Insalubridade', '% Base IR Transportista', '% ISS', '% FGTS', 'Período Pagto',
    'Descrição (Período Pagto)', '% Periculosidade', '% Reajuste', '% Base INSS Transportista',
    'Raça/Etnia', 'Recebe 13° Salário', 'Situação', 'Descrição (Situação)',
    'Descrição (T. Adm)', 'Descrição (T. Contrato)'
]

print(f"Iniciando tratamento do arquivo: {ARQUIVO_ENTRADA}")

# 1. Leitura da Base de Dados
try:
    df = pd.read_csv(ARQUIVO_ENTRADA)
    df.columns = df.columns.str.strip()
    print(f"Base de dados lida com sucesso. Total de linhas: {len(df)}")
except FileNotFoundError:
    print(f"ERRO: Arquivo '{ARQUIVO_ENTRADA}' não encontrado. Verifique o caminho.")
    exit()
except Exception as e:
    print(f"ERRO durante a leitura do arquivo: {e}")
    exit()

# --- FUNÇÃO PARA TRATAR O NOME DA COLUNA SALÁRIO (Mantida para robustez) --


# 2. SELEÇÃO DE COLUNAS (Tornada flexível)
colunas_para_selecao = [col for col in COLUNAS_DESEJADAS if col in df.columns]
df = df[colunas_para_selecao]
print(f"DataFrame filtrado para {len(df.columns)} colunas desejadas e existentes.")

# --- FUNÇÃO DE LIMPEZA DE CHAVE (Para ser usada em todo ID) ---
def limpar_chave(df, coluna_id):
    """Garante que IDs/Chaves sejam strings, sem espaços e em caixa alta."""
    if coluna_id in df.columns:
        # 1. Tenta remover o '.0' de números lidos como float/int antes de converter para string
        if df[coluna_id].dtype in ['int64', 'float64']:
             df[coluna_id] = df[coluna_id].astype(str).str.replace(r'\.0$', '', regex=True)
        # 2. Converte para string, remove espaços e coloca em caixa alta
        df[coluna_id] = df[coluna_id].astype(str).str.strip().str.upper()
    return df

# ===================================================================
# 3. LIMPEZA E TRATAMENTO DE DADOS (APLICANDO LIMPEZA ROBUSTA)
# ===================================================================

# A. Aplica Limpeza às Colunas que Serão Chaves (IDs)
print("\nIniciando limpeza das chaves de relacionamento...")
colunas_para_limpar = ['Nome', 'Empresa', 'Cadastro', 'Cargo', 'C.Custo', 'Filial']
for col in colunas_para_limpar:
    df = limpar_chave(df, col)
print("Limpeza de chaves de relacionamento (strip e upper) concluída.")

# B. Limpeza de Linhas VAZIAS (Remove linhas onde 'Nome' está nulo ou ficou 'NAN')
df.dropna(subset=['Nome', 'Cadastro'], inplace=True)
df = df[df['Nome'] != 'NAN']
print(f"Linhas com 'Nome' e/ou 'Cadastro' vazios removidas. Linhas restantes: {len(df)}")

# C. Preenchimento de nulos em colunas específicas
if 'Descrição (T. Adm)' in df.columns:
    df['Descrição (T. Adm)'] = df['Descrição (T. Adm)'].fillna('NÃO INFORMADO')
    print("Valores nulos em 'Descrição (T. Adm)' preenchidos.")

# 4. CONVERSÃO DE DATAS
COLUNAS_DE_DATA = [col for col in COLUNAS_DESEJADAS if 'Data' in col or 'Admissão' in col or 'Nascimento' in col or 'Última Simulação' in col]
print("\nIniciando conversão de colunas de data...")
for col in COLUNAS_DE_DATA:
    if col in df.columns:
        df[col] = pd.to_datetime(df[col], errors='coerce', dayfirst=True)
print("Conversão de datas concluída.")

# ===================================================================
# 5. GERAÇÃO DAS TABELAS DIMENSÃO (HIERARQUIA EMPRESA/FILIAL MODIFICADA)
# ===================================================================
print("\nGerando Tabelas Dimensão com NOVO ID SEQUENCIAL (incluindo hierarquia Empresa/Filial)...")

# --- DIMENSÃO 5: Dim_Empresa (PAI) ---
dim_empresa = df[['Empresa', 'Nome (Empresa)']].copy().dropna(subset=['Empresa'])
dim_empresa.drop_duplicates(subset=['Empresa'], inplace=True)
dim_empresa.insert(0, 'Empresa_ID', range(1, len(dim_empresa) + 1))
print(f"Dimensão Empresa (Pai) criada com {len(dim_empresa)} registros únicos (Chave: Empresa).")

# --- DIMENSÃO 4: Dim_Filial (FILHO) ---
colunas_filial_origem = ['Filial', 'Apelido (Filial)', 'Empresa']
dim_filial = df[[col for col in colunas_filial_origem if col in df.columns]].copy().dropna(subset=['Filial'])
dim_filial.drop_duplicates(subset=['Filial', 'Empresa'], inplace=True)

if 'Apelido (Filial)' in dim_filial.columns:
    dim_filial['Apelido (Filial)'] = dim_filial['Apelido (Filial)'].fillna('NÃO INFORMADO')

dim_filial.insert(0, 'Filial_ID', range(1, len(dim_filial) + 1))

# JUNTA COM Dim_Empresa para obter a Chave Estrangeira (FK)
dim_filial = pd.merge(dim_filial, dim_empresa[['Empresa', 'Empresa_ID']], on='Empresa', how='left')

# Remove a chave original 'Empresa' do Dim_Filial
dim_filial.drop(columns=['Empresa'], inplace=True)
print(f"Dimensão Filial (Filho) criada com {len(dim_filial)} registros únicos e ligada à Empresa (FK Empresa_ID).")


# --- DIMENSÃO 1: Dim_Pessoa ---
# Coluna 'Nascimento' é mantida aqui para o cálculo de Idade/Faixa Etária
colunas_pessoa = [
    'Nome', 'Cadastro', 'Nascimento', 'Sexo', # <-- NOVO: Coluna 'Sexo' adicionada
    'Estado Civil', 'Descrição (Estado Civil)',
    'Instrução', 'Descrição (Instrução)', 'Raça/Etnia', 'Descrição (Raça/Etnia)',
    'Dependentes IR', 'Dependentes Saf', 'Dep. Saldo FGTS', 'Cadastramento PIS',
    'Nome (Cadastro O. Contrato)'
]
dim_pessoa = df[[col for col in colunas_pessoa if col in df.columns]].copy()
dim_pessoa.drop_duplicates(subset=['Nome', 'Cadastro'], inplace=True)
dim_pessoa.insert(0, 'Pessoa_ID', range(1, len(dim_pessoa) + 1))
print(f"Dimensão Pessoa criada com {len(dim_pessoa)} registros ÚNICOS, incluindo 'Sexo'.")

# ********************************************************************
# ** CÁLCULO DE IDADE E FAIXA ETÁRIA **
# ********************************************************************
if 'Nascimento' in dim_pessoa.columns:
    print("Calculando 'Idade Atual' e 'Faixa Etária'...")
    hoje = pd.to_datetime(date.today())
    dim_pessoa['Idade Atual'] = (hoje - dim_pessoa['Nascimento']).dt.days / 365.25
    dim_pessoa['Idade Atual'] = dim_pessoa['Idade Atual'].apply(lambda x: int(x) if pd.notna(x) else pd.NA).astype('Int64')

    bins = [0, 20, 25, 30, 35, 40, 50, 60, 100]
    labels = [
        "1. Abaixo de 20", "2. 20 - 24 Anos", "3. 25 - 29 Anos", "4. 30 - 34 Anos",
        "5. 35 - 39 Anos", "6. 40 - 49 Anos", "7. 50 - 59 Anos", "8. 60 ou Mais"
    ]
    dim_pessoa['Faixa Etária (Pirâmide)'] = pd.cut(
        dim_pessoa['Idade Atual'],
        bins=bins,
        labels=labels,
        right=False
    ).astype(str).str.replace('nan', '9. Idade Inválida')
    print("Colunas 'Idade Atual' e 'Faixa Etária' adicionadas à Dim_Pessoa.")
else:
    print("AVISO: Coluna 'Nascimento' não encontrada na Dim_Pessoa. Idade não calculada.")

# --- DIMENSÃO 2: Dim_Cargo ---
dim_cargo = df[['Cargo', 'Título Reduzido (Cargo)']].copy().dropna(subset=['Cargo'])
dim_cargo.drop_duplicates(subset=['Cargo'], inplace=True)
dim_cargo.insert(0, 'Cargo_ID', range(1, len(dim_cargo) + 1))
print(f"Dimensão Cargo criada com {len(dim_cargo)} registros únicos.")

# --- DIMENSÃO 3: Dim_CCusto ---
dim_ccusto = df[['C.Custo', 'Descrição (C.Custo)']].copy().dropna(subset=['C.Custo'])
dim_ccusto.drop_duplicates(subset=['C.Custo'], inplace=True)
dim_ccusto.insert(0, 'CCusto_ID', range(1, len(dim_ccusto) + 1))
print(f"Dimensão C.Custo criada com {len(dim_ccusto)} registros únicos.")


# ===================================================================
# 6. CRIAÇÃO DA TABELA FATO (Fato_Contratos)
# ===================================================================
print("\nPreparando a Tabela Fato com todos os novos IDs...")
df_fato = df.copy()

# 1. Merges Padrão (Pessoa, Cargo, CCusto)
# NOTA: O 'Sexo' NÃO precisa ser mesclado de volta, pois é uma coluna de dimensão.
df_fato = pd.merge(df_fato, dim_pessoa[['Nome', 'Cadastro', 'Pessoa_ID']], on=['Nome', 'Cadastro'], how='left')
df_fato = pd.merge(df_fato, dim_cargo[['Cargo', 'Cargo_ID']], on='Cargo', how='left')
df_fato = pd.merge(df_fato, dim_ccusto[['C.Custo', 'CCusto_ID']], on='C.Custo', how='left')

# 2. Merge Empresa (Obtém Empresa_ID)
df_fato = pd.merge(df_fato, dim_empresa[['Empresa', 'Empresa_ID']], on='Empresa', how='left')

# 3. Merge Filial (Obtém Filial_ID)
# Cria o DataFrame de chaves composto para o merge
dim_filial_chaves = dim_filial[['Filial_ID', 'Filial', 'Empresa_ID']]

df_fato = pd.merge(
    df_fato, 
    dim_filial_chaves, 
    on=['Filial', 'Empresa_ID'], 
    how='left',
    suffixes=('_drop', '') 
)

print("Todos os novos IDs (Surrogate Keys) foram adicionados à Tabela Fato.")

# 4. Seleção e Limpeza da Tabela Fato Final
# Remoção das colunas originais que viraram IDs
colunas_para_descartar = ['Nome', 'Cargo', 'C.Custo', 'Empresa', 'Filial', 'Sexo']
df_fato.drop(columns=colunas_para_descartar, inplace=True, errors='ignore')


# B. Selecionar colunas para a Tabela Fato
colunas_fato = [
    # NOVOS IDs (Chaves Estrangeiras)
    'Pessoa_ID', 'Cargo_ID', 'CCusto_ID', 'Filial_ID', 'Empresa_ID', 
    # CHAVE DE AUDITORIA: Adicionado 'Cadastro' conforme solicitado
    'Cadastro', 
    # Dados de Contrato/Fato
    'Admissão', 'Data Afastamento', 'Data Salário', 'Data Cargo', 'Data C.Custo',
    'Data de Reintegração', 'Data Vínculo', 'Última Simulação', 'Data Inclusão', '% Desempenho', '% Insalubridade', '% Periculosidade',
    '% Reajuste', '% FGTS', '% ISS', 'Dependentes IR', 'Dependentes Saf',
    'Situação', 'Descrição (Situação)', 'Causa', 'Descrição (Causa)', 'Escala',
    'Descrição (Escala)', 'Opção FGTS', 'Período Pagto', 'Descrição (Período Pagto)',
    'Descrição (T. Adm)', 'Descrição (T. Contrato)', 'Descrição (Cat. eSocial)',
    'Descrição (Motivo Alt. Salário)', 'Recebe 13° Salário', 'Código Fornecedor'
]

df_fato_final = df_fato[[col for col in colunas_fato if col in df_fato.columns]]
print("Tabela Fato final criada com as colunas de IDs e medidas.")

# ===================================================================
# 7. SALVAMENTO EM MÚLTIPLAS ABAS
# ===================================================================
print(f"\nSalvando Modelo Estrela/Snowflake em EXCEL: {ARQUIVO_FINAL_MODELADO} (Múltiplas abas)")
try:
    with pd.ExcelWriter(ARQUIVO_FINAL_MODELADO, engine='xlsxwriter') as writer:
        df_fato_final.to_excel(writer, sheet_name='Fato_Contratos', index=False)
        
        # Dimensões
        # 'Nome' não é incluído na dimensão Pessoa final, mas 'Sexo' sim.
        colunas_dim_pessoa_final = [col for col in dim_pessoa.columns if col not in ['Nome']]
        dim_pessoa[colunas_dim_pessoa_final].to_excel(writer, sheet_name='Dim_Pessoa', index=False)

        dim_cargo.drop(columns=['Cargo']).to_excel(writer, sheet_name='Dim_Cargo', index=False)
        dim_ccusto.drop(columns=['C.Custo']).to_excel(writer, sheet_name='Dim_CCusto', index=False)

        # Dimensões Empresa e Filial (Hierárquicas)
        dim_filial.to_excel(writer, sheet_name='Dim_Filial', index=False) 
        dim_empresa.drop(columns=['Empresa']).to_excel(writer, sheet_name='Dim_Empresa', index=False) 

    print("\n---------------------------------------------------")
    print("Modelagem e salvamento CONCLUÍDOS com sucesso!")
    print(f"Total de Pessoas Únicas (IDs): {len(dim_pessoa)}")
    print(f"O arquivo EXCEL '{ARQUIVO_FINAL_MODELADO}' (6 abas) está pronto.")
    print("No Power BI, as relações devem ser: Dim_Empresa[Empresa_ID] (1) -> Dim_Filial[Empresa_ID] (*)")
    print("E a principal: Dim_Filial[Filial_ID] (1) -> Fato_Contratos[Filial_ID] (*)")
    print("---------------------------------------------------")

except Exception as e:
    print(f"\nERRO ao salvar o arquivo Excel: {e}")
