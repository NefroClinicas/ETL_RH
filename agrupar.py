import pandas as pd
import os
import time
from datetime import date
import re
from gender_guesser_br import Genero


PASTA_ORIGEM_ONEDRIVE = r"C:\Users\LuisGuilhermeMoraesd\Nefroclinicas Serviço de Nefrologia e Dialise Ltda\Nefroclinicas - 07 - DADOS (1)"
PASTA_TRABALHO = r"C:\Users\LuisGuilhermeMoraesd\OneDrive - Nefroclinicas Serviço de Nefrologia e Dialise Ltda\Área de Trabalho\Limpar_Dados"

# Nomes dos arquivos de saída e auxiliares
ARQUIVO_SEXO = os.path.join(PASTA_TRABALHO, "Nomes_e_Sexo_Inferido.xlsx")
ARQUIVO_INTERMEDIARIO_CSV = os.path.join(PASTA_TRABALHO, "Base_BI_Consolidada2.csv")
ARQUIVO_FINAL_MODELADO = os.path.join(PASTA_TRABALHO, "Base_MODELADA_PowerBI_V4.xlsx")
# Arquivos Auxiliares (Faltas e Absenteísmo)
ARQUIVO_FALTAS = os.path.join(PASTA_TRABALHO, "faltas.csv")
ARQUIVO_ABS = os.path.join(PASTA_TRABALHO, "ABS.csv")


# Lista EXATA das colunas a serem extraídas e usadas no modelo
COLUNAS_DESEJADAS = [
'Nome', 'Sexo','CPF', 'Empresa', 'Cadastro', 'Admissão', 'Cargo', 'C.Custo', 
'Descrição (C.Custo)', 'Data Afastamento', 'Título Reduzido (Cargo)', 
'Descrição (Raça/Etnia)', 'Descrição (Cat. eSocial)', 'Causa', 
'Descrição (Causa)', 'Escala', 'Descrição (Escala)', 'Filial', 
'Apelido (Filial)', 'Código Fornecedor', 'Descrição (Motivo Alt. Salário)',
'Data Adicionais', 'Data Aposentadoria', 'Data Cargo', 'Data C.Custo', 
'Data Ult. Alt. Cat.', 'Data de Chegada', 'Data Escala', 'Data Estabilidade', 
'Data Escala VTR', 'Data Filial', 'Data Histórico de Contrato', 'Data Inclusão', 
'Data Local', 'Nascimento', 'Opção FGTS', 'Data Posto', 'Data Ass. PPR', 
'Data de Reintegração', 'Data Salário', 'Data Cat. SEFIP', 'Última Simulação', 
'Data Sindicato', 'Data FGTS', 'Data Vínculo', 'Cadastramento PIS', 
'Dependentes IR', 'Dependentes Saf', 'Dep. Saldo FGTS', 'Estado Civil', 
'Descrição (Estado Civil)', 'Instrução', 'Descrição (Instrução)', 
'Nome (Empresa)', 'Nome (Cadastro O. Contrato)', 'Nome (Empresa O. Contrato)',
'Descrição (Tipo O. Contrato)', '% Desempenho', '% Insalubridade', 
'% Base IR Transportista', '% ISS', '% FGTS', 'Período Pagto', 
'Descrição (Período Pagto)', '% Periculosidade', '% Reajuste', 
'% Base INSS Transportista', 'Raça/Etnia', 'Recebe 13° Salário', 'Situação', 
'Descrição (Situação)', 'Descrição (T. Adm)', 'Descrição (T. Contrato)'
]


# =======================================================================
# 2. FUNÇÕES DE SUPORTE
# =======================================================================

# -----------------------------------------------------------------------
# ** NOVAS FUNÇÕES DE INFERÊNCIA COPIADAS DO SEXO.PY **
# -----------------------------------------------------------------------
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

def inferir_sexo_br(primeiro_nome):
    """Inferir o sexo usando a biblioteca gender-guesser-br (IBGE)."""
    if not primeiro_nome:
        return None 
    
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
        # print(f"Erro ao consultar nome {primeiro_nome}: {e}") # Descomente para debug
        return None 
# -----------------------------------------------------------------------


def carregar_dados_sexo(caminho_arquivo):
    """Carrega o arquivo de sexo (Excel) e prepara o DataFrame para merge."""
    print(f"\n-> Carregando dados de Sexo de: {os.path.basename(caminho_arquivo)}")
    try:
        # Tenta carregar o arquivo Excel
        df_sexo = pd.read_excel(caminho_arquivo)
        
        # Garante que as colunas 'Nome' e 'Sexo' existam e limpa nomes de colunas
        df_sexo.columns = df_sexo.columns.str.strip()

        if 'Nome' not in df_sexo.columns or 'Sexo' not in df_sexo.columns:
            print("  ERRO: O arquivo de sexo não contém as colunas 'Nome' e/ou 'Sexo'.")
            return None

        # Normaliza valores de Sexo (e.g., Desconhecido/Ambos para NaN, ou usa somente Masculino/Feminino)
        # O melhor é deixar o valor como está para ter a referência manual
        df_sexo = df_sexo[['Nome', 'Sexo']].copy()
        
        # Limpa espaços em branco nos Nomes do arquivo de Sexo para garantir o merge
        df_sexo['Nome'] = df_sexo['Nome'].astype(str).str.strip().str.upper() # Adiciona .upper()
        df_sexo.drop_duplicates(subset=['Nome'], keep='first', inplace=True)
        
        # Tratamento: Converte 'Desconhecido', 'Ambos' ou string vazia para NA para ser preenchido pela inferência
        valores_a_anular = ['DESCONHECIDO', 'AMBOS', '']
        df_sexo['Sexo'] = df_sexo['Sexo'].astype(str).str.upper()
        df_sexo.loc[df_sexo['Sexo'].isin(valores_a_anular), 'Sexo'] = pd.NA
        
        # Converte o resto de volta para Capitalize/Normal
        df_sexo['Sexo'] = df_sexo['Sexo'].str.capitalize()

        print(f"  Sucesso: {len(df_sexo)} nomes únicos carregados com dados de Sexo.")
        return df_sexo
        
    except FileNotFoundError:
        print(f"  AVISO: Arquivo de sexo não encontrado no caminho: {caminho_arquivo}")
        return None
    except Exception as e:
        print(f"  ERRO ao processar o arquivo de sexo (EXCEL): {e}")
        return None


def transformar_e_selecionar(caminho_arquivo_local, colunas_desejadas):
    """Carrega o arquivo (CSV ou XLS/XLSX), trata, renomeia e seleciona colunas."""
    nome_arquivo = os.path.basename(caminho_arquivo_local)
    
    # Verifica a extensão para decidir como ler (prioriza CSV/separadores, senão tenta Excel)
    if nome_arquivo.lower().endswith('.xls') or nome_arquivo.lower().endswith('.xlsx'):
        print(f"  -> Processando com Pandas: {nome_arquivo} (Lendo como Excel)")
        try:
            df = pd.read_excel(caminho_arquivo_local)
        except Exception as e:
            print(f"  ERRO CRÍTICO ao ler {nome_arquivo} como Excel: {e}")
            return None
    else:
        print(f"  -> Processando com Pandas: {nome_arquivo} (Lendo como CSV, tentando delimitadores)")
        # 1. Tenta carregar o arquivo CSV, testando delimitadores
        try:
            df = pd.read_csv(caminho_arquivo_local, sep=';', encoding='latin1') 
        except Exception:
            try:
                df = pd.read_csv(caminho_arquivo_local, sep=',', encoding='latin1')
            except Exception as e:
                print(f"  ERRO CRÍTICO ao ler {nome_arquivo} como CSV. Falha com ';' e ','.")
                print(f"  Detalhes: {e}")
                return None

    # NOVO PASSO: LIMPEZA E PADRONIZAÇÃO DOS NOMES DAS COLUNAS DO ARQUIVO
    df.columns = df.columns.str.strip()
    colunas_reais = df.columns.tolist()
    
    # --- BLOCO DE MUDANÇA DE NOME E TRATAMENTO DE AUSENTES ---
    
    # 2. Mapeamento de Colunas Ausentes (Salário)
    # Tenta renomear 'Salário Simulado' para 'Valor Salário'
    if 'Valor Salário' not in colunas_reais and 'Salário Simulado' in colunas_reais:
        df.rename(columns={'Salário Simulado': 'Valor Salário'}, inplace=True)
        print(f"  AVISO: 'Salário Simulado' renomeado para 'Valor Salário' em {nome_arquivo}.")
    
    # 3. Lista de Colunas Disponíveis
    # Filtra COLUNAS_DESEJADAS (agora mantendo 'Sexo', pois será a coluna de destino)
    
    colunas_presentes = [col.strip() for col in colunas_desejadas if col.strip() in df.columns]

    colunas_ausentes = [col.strip() for col in colunas_desejadas if col.strip() not in df.columns]
    
    if colunas_ausentes:
        print(f"  AVISO: Colunas não encontradas e IGNORADAS em {nome_arquivo}: {colunas_ausentes}")


    # 4. Seleção de Colunas
    try:
        df_selecionado = df[colunas_presentes].copy()
        
        # Limpa espaços em branco nos Nomes do DataFrame principal e coloca em caixa alta
        if 'Nome' in df_selecionado.columns:
            df_selecionado['Nome'] = df_selecionado['Nome'].astype(str).str.strip().str.upper()
        return df_selecionado
    except KeyError as e:
        print(f"  ERRO: Seleção final falhou após o mapeamento. Detalhes: {e}")
        return None

def limpar_chave(df, coluna_id):
    """Garante que IDs/Chaves sejam strings, sem espaços e em caixa alta."""
    if coluna_id in df.columns:
        # 1. Tenta remover o '.0' de números lidos como float/int antes de converter para string
        if df[coluna_id].dtype in ['int64', 'float64']:
            # Use uma expressão regular mais segura para remover apenas .0 no final
            df[coluna_id] = df[coluna_id].astype(str).str.replace(r'\.0$', '', regex=True)
        # 2. Converte para string, remove espaços e coloca em caixa alta
        df[coluna_id] = df[coluna_id].astype(str).str.strip().str.upper().str.replace('NAN', '') # Remove 'NAN' residual
    return df

def calcular_idade_faixa_etaria(df):
    """Calcula Idade Atual e Faixa Etária no DataFrame Pessoa."""
    if 'Nascimento' in df.columns:
        print("Calculando 'Idade Atual' e 'Faixa Etária'...")
        hoje = pd.to_datetime(date.today())
        
        # Garante que 'Nascimento' esteja como datetime antes de calcular
        df['Nascimento'] = pd.to_datetime(df['Nascimento'], errors='coerce', dayfirst=True)
        
        df['Idade Atual'] = (hoje - df['Nascimento']).dt.days / 365.25
        df['Idade Atual'] = df['Idade Atual'].apply(lambda x: int(x) if pd.notna(x) else pd.NA).astype('Int64')

        bins = [0, 20, 25, 30, 35, 40, 50, 60, 100]
        labels = [
            "1. Abaixo de 20", "2. 20 - 24 Anos", "3. 25 - 29 Anos", "4. 30 - 34 Anos",
            "5. 35 - 39 Anos", "6. 40 - 49 Anos", "7. 50 - 59 Anos", "8. 60 ou Mais"
        ]
        
        # Usa .fillna() para Faixa Etária inválida
        df['Faixa Etária (Pirâmide)'] = pd.cut(
            df['Idade Atual'],
            bins=bins,
            labels=labels,
            right=False
        ).astype(str).str.replace('nan', '9. Idade Inválida')
        print("Colunas 'Idade Atual' e 'Faixa Etária' adicionadas à Dim_Pessoa.")
    else:
        print("AVISO: Coluna 'Nascimento' não encontrada. Idade não calculada.")
    return df

# ***********************************************************************
# NOVA FUNÇÃO PARA PROCESSAR CSVs AUXILIARES (FALTAS/ABSENTEÍSMO)
# ***********************************************************************
def etl_processa_csv_auxiliar(caminho_arquivo, nome_tabela):
    """
    Função dedicada para ler e processar arquivos CSV auxiliares (como Faltas ou Absenteísmo).
    Faz a limpeza básica, padronização do Nome, limpeza de Horários, conversão de ABS e de Data.
    """
    print(f"\n-> Processando arquivo auxiliar: {os.path.basename(caminho_arquivo)} para a tabela {nome_tabela}")
    try:
        # Tenta ler o CSV, usando latin1 por ser comum em dados brasileiros
        # Mantendo o sep=';' da sua última tentativa.
        df = pd.read_csv(caminho_arquivo, sep=';', encoding='latin1')
    except FileNotFoundError:
        print(f"  AVISO: Arquivo '{os.path.basename(caminho_arquivo)}' não encontrado. Pulando.")
        return None
    except Exception as e:
        try:
            # Tenta com vírgula como separador (fallback)
            df = pd.read_csv(caminho_arquivo, sep=',', encoding='latin1')
        except Exception as e:
            print(f"  ERRO CRÍTICO ao ler {os.path.basename(caminho_arquivo)} como CSV. Detalhes: {e}")
            return None

    # Limpeza básica e padronização dos cabeçalhos das colunas
    df.columns = df.columns.str.strip().str.replace(r'[^a-zA-Z0-9\s\(\)%]', '', regex=True)
    
    # Padronização da coluna Nome
    if 'Nome' in df.columns:
        df['Nome'] = df['Nome'].astype(str).str.strip().str.upper()
        print("  Coluna 'Nome' padronizada (Upper, Strip).")
        
    # --------------------------------------------------------------------------------
    # 🌟 1. NOVO: LIMPEZA DAS COLUNAS DE HORÁRIO (Removendo o :00 final)
    # --------------------------------------------------------------------------------
    colunas_horario = ['Previsto', 'Ausencia', 'Presenca']
    
    for col in colunas_horario:
        # Verifica se a coluna existe
        if col in df.columns:
            # Remove o padrão ":00" APENAS se estiver no final da string (usando regex '$')
            df[col] = df[col].astype(str).str.replace(r':00$', '', regex=True)
            # Remove ":00:00" em colunas de Absenteísmo.
            df[col] = df[col].astype(str).str.replace(r'00:00$', '00', regex=True)
            print(f"  Coluna '{col}' limpa (removido o sufixo :00).")
    # --------------------------------------------------------------------------------
    
    # --------------------------------------------------------------------------------
    # 🌟 2. NOVO: CONVERSÃO DA COLUNA 'ABS' PARA NÚMERO DECIMAL (%)
    # --------------------------------------------------------------------------------
    if 'ABS' in df.columns:
        try:
            # 1. Limpa o '%' e espaços
            df['ABS'] = df['ABS'].astype(str).str.replace('%', '', regex=False).str.strip()
            
            # 2. Substitui a vírgula (,) por ponto (.) como separador decimal
            df['ABS'] = df['ABS'].str.replace(',', '.', regex=False)
            
            # 3. Converte para número e divide por 100 (para decimal: 5.88% -> 0.0588)
            df['ABS'] = pd.to_numeric(df['ABS'], errors='coerce') / 100
            print("  Coluna 'ABS' convertida para número decimal (%).")
        except Exception as e:
            print(f"  AVISO: Falha ao converter coluna 'ABS' para numérico. Detalhes: {e}")
    # --------------------------------------------------------------------------------
    
    # Conversão da coluna Data (mantida a correção anterior)
    if 'Data' in df.columns:
        
        # 1. REMOVER O DIA DA SEMANA
        # O padrão '.*,' busca qualquer coisa (.*) até a primeira vírgula (,) e a própria vírgula.
        df['Data'] = df['Data'].astype(str).str.replace(r'.*,', '', regex=True).str.strip()
        
        # 2. CONVERSÃO FINAL PARA DATETIME
        df['Data'] = pd.to_datetime(df['Data'], errors='coerce', dayfirst=True)
        print("  Coluna 'Data' limpa (dia da semana removido) e convertida para DateTime.")
        
    print(f"  Tabela '{nome_tabela}' processada com {len(df)} linhas.")
    return df


# =======================================================================
# 3. LÓGICA PRINCIPAL: CONSOLIDAÇÃO (ETL)
# =======================================================================

def etl_consolida_e_salva_csv(colunas_desejadas):
    """Busca arquivos, processa, consolida, faz o merge de Sexo e salva o CSV intermediário."""
    
    lista_dataframes = []
    
    print("\n" + "="*70)
    print(f"ETAPA 1/2: CONSOLIDAÇÃO E INFERÊNCIA DE SEXO | Início: {time.ctime()}")
    print("="*70)

    # PASSO 0: Carregar Dados Auxiliares (Sexo)
    df_sexo = carregar_dados_sexo(ARQUIVO_SEXO) 
    
    # 1. BUSCAR TODOS OS ARQUIVOS DA PASTA
    try:
        todos_arquivos = [
            os.path.join(PASTA_ORIGEM_ONEDRIVE, f)
            for f in os.listdir(PASTA_ORIGEM_ONEDRIVE)
            if os.path.isfile(os.path.join(PASTA_ORIGEM_ONEDRIVE, f)) and not f.startswith('.') 
        ]
    except FileNotFoundError:
        print(f"\nERRO: Pasta de origem não encontrada: {PASTA_ORIGEM_ONEDRIVE}. Execute 'ambiente_de_teste.py' primeiro.")
        return None

    # 2. FILTRA OS ARQUIVOS QUE PROVAVELMENTE PODEM SER CSVs/XLSs
    arquivos_compativeis = [
        f for f in todos_arquivos 
        if f.lower().endswith(('.xls', '.xlsx', '.csv'))
    ]

    if not arquivos_compativeis:
        print("\n" + "="*70)
        print("NENHUM ARQUIVO COMPATÍVEL (.xls, .xlsx ou .csv) ENCONTRADO na pasta especificada.")
        print(f"Caminho verificado: {PASTA_ORIGEM_ONEDRIVE}")
        print("="*70)
        return None

    print(f"-> {len(arquivos_compativeis)} arquivos compatíveis encontrados. Iniciando consolidação...")
    
    # 3. Processamento e Consolidação
    for arquivo in arquivos_compativeis:
        # A coluna 'Sexo' pode vir aqui, mas será substituída pelo merge/inferência
        df_processado = transformar_e_selecionar(arquivo, colunas_desejadas) 
        
        if df_processado is not None and not df_processado.empty:
            df_processado['Origem_Arquivo'] = os.path.basename(arquivo)
            lista_dataframes.append(df_processado)

    # 4. Combinação Final (Concatenar)
    if lista_dataframes:
        CONTEUDO_FINAL = pd.concat(lista_dataframes, ignore_index=True, sort=False)
        
        # --- PASSO 5: MERGE E INFERÊNCIA CONDICIONAL DE SEXO (LÓGICA UNIFICADA) ---
        
        # 5.1 PREPARAÇÃO: Remove qualquer coluna 'Sexo' que tenha vindo dos arquivos brutos
        CONTEUDO_FINAL.drop(columns=['Sexo'], inplace=True, errors='ignore')
        
        # Adiciona uma coluna 'Sexo' vazia para ser preenchida
        if 'Sexo' not in CONTEUDO_FINAL.columns:
            CONTEUDO_FINAL['Sexo'] = pd.NA

        # 5.2 MERGE COM DADOS DE SEXO (Excel) - Prioridade Manual
        if df_sexo is not None and 'Nome' in CONTEUDO_FINAL.columns:
            print("\n-> Realizando Merge com a informação de Sexo do Excel (Prioridade Manual)...")
            
            # Merge Left Join
            df_merged = pd.merge(
                CONTEUDO_FINAL.drop(columns=['Sexo'], errors='ignore'), 
                df_sexo,
                on='Nome',
                how='left'
            )
            # Define o resultado do merge como o CONTEUDO_FINAL temporário
            CONTEUDO_FINAL = df_merged
            
            print(f"  Merge concluído. {CONTEUDO_FINAL['Sexo'].count()} valores de Sexo manual/existente carregados.")
        else:
            print("  AVISO: Arquivo de Sexo manual não carregado. Pulando merge.")

        # 5.3 INFERÊNCIA CONDICIONAL: Aplica gender-guesser-br APENAS nos valores nulos.
        
        # 5.3.1 Identifica os registros SEM Sexo (Nulos)
        # O .isna() trata valores NA do Pandas e None
        condicao_inferir = CONTEUDO_FINAL['Sexo'].isna()
        df_a_inferir = CONTEUDO_FINAL[condicao_inferir].copy()
        
        if not df_a_inferir.empty:
            total_a_inferir = len(df_a_inferir)
            print(f"\n  -> Aplicando Inferência do IBGE (Lenta) a {total_a_inferir} registros NA/Nulos...")

            # 5.3.2 Aplica a função de limpeza de nome
            df_a_inferir['Primeiro_Nome'] = df_a_inferir['Nome'].apply(extrair_primeiro_nome)
            
            # 5.3.3 Aplica a função de inferência
            novos_sexos = df_a_inferir['Primeiro_Nome'].apply(inferir_sexo_br)
            
            # 5.3.4 Atualiza o DataFrame principal (CONTEUDO_FINAL)
            sucesso_br = novos_sexos.dropna()
            
            # O .loc garante que os novos valores sejam colocados DE VOLTA no DataFrame principal, APENAS nos nulos
            CONTEUDO_FINAL.loc[sucesso_br.index, 'Sexo'] = sucesso_br
            
            print(f"  {sucesso_br.count()} valores de Sexo preenchidos por Inferência.")
        else:
            print("  Nenhum valor nulo (NA) de Sexo restante para inferência. Passo ignorado.")
        
        # --- FIM DO PASSO 5 ---

        # 6. Adiciona as colunas ausentes na consolidação final (se alguma faltou em todos)
        # Garante que todas as colunas desejadas estejam presentes
        for col in [str(coluna).strip() for coluna in colunas_desejadas if coluna is not None]:
            if col not in CONTEUDO_FINAL.columns:
                CONTEUDO_FINAL[col] = pd.NA
        
        # 7. Salvar o Resultado Intermediário (CSV)
        colunas_finais_ordenadas = [c.strip() for c in colunas_desejadas] + ['Origem_Arquivo']
        colunas_para_salvar = [col for col in colunas_finais_ordenadas if col in CONTEUDO_FINAL.columns]

        CONTEUDO_FINAL[colunas_para_salvar].to_csv(ARQUIVO_INTERMEDIARIO_CSV, index=False, encoding='utf-8')
        
        print("\n" + "="*70)
        print(f"ETAPA 1/2 COMPLETA! CSV Intermediário salvo em: {ARQUIVO_INTERMEDIARIO_CSV}")
        print(f"Total de linhas consolidadas (incluindo duplicadas): {len(CONTEUDO_FINAL)}")
        print("="*70)
        return CONTEUDO_FINAL

    else:
        print("\nNenhum arquivo pôde ser processado com sucesso.")
        return None

# =======================================================================
# 4. LÓGICA PRINCIPAL: MODELAGEM (ELT)
# =======================================================================

def etl_modela_e_salva_excel(df_input, colunas_desejadas):
    """
    Recebe o DataFrame consolidado, limpa, cria o modelo estrela/snowflake 
    e salva no arquivo final Excel de múltiplas abas, incluindo tabelas auxiliares.
    """
    if df_input is None or df_input.empty:
        print("ERRO: O DataFrame consolidado está vazio ou não foi gerado.")
        return
        
    print("\n" + "="*70)
    print("ETAPA 2/2: MODELAGEM E CRIAÇÃO DE MODELO ESTRELA/SNOWFLAKE")
    print("="*70)

    df = df_input.copy()
    
    # ********************************************************************
    # ** NOVO PASSO: PROCESSAMENTO DAS TABELAS AUXILIARES **
    # ********************************************************************
    df_faltas = etl_processa_csv_auxiliar(ARQUIVO_FALTAS, 'Fato_Faltas')
    df_abs = etl_processa_csv_auxiliar(ARQUIVO_ABS, 'Fato_Absenteismo')


    # 1. SELEÇÃO DE COLUNAS
    colunas_para_selecao = [col for col in colunas_desejadas if col in df.columns]
    df = df[colunas_para_selecao + ['Origem_Arquivo']].copy()
    print(f"DataFrame filtrado para {len(df.columns)} colunas desejadas e existentes.")

    # 2. LIMPEZA E TRATAMENTO DE DADOS

    # A. Aplica Limpeza às Colunas que Serão Chaves (IDs)
    print("\nIniciando limpeza e padronização das chaves de relacionamento...")
    # ** ALTERAÇÃO: Adiciona 'CPF' à lista de chaves a serem limpas
    colunas_para_limpar = ['Nome', 'CPF', 'Empresa', 'Cadastro', 'Cargo', 'C.Custo', 'Filial']
    for col in colunas_para_limpar:
        df = limpar_chave(df, col)
    print("Limpeza de chaves de relacionamento (strip e upper) concluída.")

    # B. Limpeza de Linhas VAZIAS (Remove linhas onde a chave única de pessoa ('CPF') está nula ou vazia)
    df.replace('', pd.NA, inplace=True) # Substitui strings vazias por NA para o dropna
    
    # ** ALTERAÇÃO: Garante que CPF (nova chave única) e Nome (descrição essencial) existam
    df.dropna(subset=['CPF', 'Nome'], inplace=True)
    
    # Remove valores string 'NAN' nas chaves
    df = df[df['CPF'] != 'NAN'] 
    df = df[df['Nome'] != 'NAN'] 
    print(f"Linhas com CPF e/ou Nome vazios removidas. Linhas restantes: {len(df)}")


    # C. Preenchimento de nulos em colunas específicas
    if 'Descrição (T. Adm)' in df.columns:
        df['Descrição (T. Adm)'] = df['Descrição (T. Adm)'].fillna('NÃO INFORMADO')
        print("Valores nulos em 'Descrição (T. Adm)' preenchidos.")
        
    # NOVO: Preenche Sexo nulo com valor 'INFERÊNCIA FALHOU' (para fins de visualização)
    if 'Sexo' in df.columns:
        df['Sexo'] = df['Sexo'].fillna('Não Definido/Inferido')
        print("Valores nulos em 'Sexo' preenchidos com 'Não Definido/Inferido'.")


    # 3. CONVERSÃO DE DATAS
    COLUNAS_DE_DATA = [col for col in colunas_desejadas if 'Data' in col or 'Admissão' in col or 'Nascimento' in col or 'Última Simulação' in col]
    print("\nIniciando conversão de colunas de data...")
    for col in COLUNAS_DE_DATA:
        if col in df.columns:
            # Força o tratamento como data no formato brasileiro dia/mês/ano
            df[col] = pd.to_datetime(df[col], errors='coerce', dayfirst=True)
    print("Conversão de datas concluída.")

    # ===================================================================
    # 4. GERAÇÃO DAS TABELAS DIMENSÃO
    # ===================================================================
    print("\nGerando Tabelas Dimensão com ID Sequencial (Chaves Suplementares)...")

    # --- DIMENSÃO 5: Dim_Empresa (PAI) ---
    dim_empresa = df[['Empresa', 'Nome (Empresa)']].copy().dropna(subset=['Empresa'])
    dim_empresa.drop_duplicates(subset=['Empresa'], inplace=True)
    dim_empresa.insert(0, 'Empresa_ID', range(1, len(dim_empresa) + 1))
    print(f"Dimensão Empresa (Pai) criada com {len(dim_empresa)} registros únicos.")

    # --- DIMENSÃO 4: Dim_Filial (FILHO) ---
    colunas_filial_origem = ['Filial', 'Apelido (Filial)', 'Empresa']
    dim_filial = df[[col for col in colunas_filial_origem if col in df.columns]].copy().dropna(subset=['Filial'])
    dim_filial.drop_duplicates(subset=['Filial', 'Empresa'], inplace=True)

    if 'Apelido (Filial)' in dim_filial.columns:
        dim_filial['Apelido (Filial)'] = dim_filial['Apelido (Filial)'].fillna('NÃO INFORMADO')

    dim_filial.insert(0, 'Filial_ID', range(1, len(dim_filial) + 1))

    # JUNTA COM Dim_Empresa para obter a Chave Estrangeira (FK)
    dim_filial = pd.merge(dim_filial, dim_empresa[['Empresa', 'Empresa_ID']], on='Empresa', how='left')
    dim_filial.drop(columns=['Empresa'], inplace=True, errors='ignore')
    print(f"Dimensão Filial (Filho) criada com {len(dim_filial)} registros únicos e ligada à Empresa (FK Empresa_ID).")


    # --- DIMENSÃO 1: Dim_Pessoa ---
    colunas_pessoa = [
        'Nome', 'CPF', 'Cadastro', 'Nascimento', 'Sexo', # ** CPF adicionado aqui **
        'Estado Civil', 'Descrição (Estado Civil)',
        'Instrução', 'Descrição (Instrução)', 'Raça/Etnia', 'Descrição (Raça/Etnia)',
        'Dependentes IR', 'Dependentes Saf', 'Dep. Saldo FGTS', 'Cadastramento PIS',
        'Nome (Cadastro O. Contrato)'
    ]
    # Seleciona apenas as colunas que existem no df
    cols_existentes_pessoa = [col for col in colunas_pessoa if col in df.columns]
    
    dim_pessoa = df[cols_existentes_pessoa].copy()
    
    # ** ALTERAÇÃO: Usa CPF como chave única para drop_duplicates **
    if 'CPF' in dim_pessoa.columns:
        dim_pessoa.drop_duplicates(subset=['CPF'], inplace=True)
        print("Chave única para Dim_Pessoa definida como CPF.")
    else:
        # Fallback
        dim_pessoa.drop_duplicates(subset=['Nome', 'Cadastro'], inplace=True)
        print("ALERTA: Coluna 'CPF' não encontrada. Usando ['Nome', 'Cadastro'] como chave de Pessoa.")
    
    dim_pessoa.insert(0, 'Pessoa_ID', range(1, len(dim_pessoa) + 1))
    
    # ********************************************************************
    # ** CÁLCULO DE IDADE E FAIXA ETÁRIA **
    # ********************************************************************
    dim_pessoa = calcular_idade_faixa_etaria(dim_pessoa)
    
    print(f"Dimensão Pessoa criada com {len(dim_pessoa)} registros ÚNICOS, incluindo 'Sexo'.")

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
    # 5. CRIAÇÃO DA TABELA FATO (Fato_Contratos)
    # ===================================================================
    print("\nPreparando a Tabela Fato com todos os novos IDs...")
    df_fato = df.copy()

    # 1. Merges Padrão (Pessoa, Cargo, CCusto, Empresa, Filial)
    
    # ** ALTERAÇÃO: Merge Pessoa (Chave Única: CPF) **
    chave_merge_pessoa = ['CPF']
    cols_merge_pessoa = chave_merge_pessoa + ['Pessoa_ID']

    df_fato = pd.merge(df_fato, dim_pessoa[cols_merge_pessoa], 
                             on=chave_merge_pessoa, how='left', suffixes=('', '_Pessoa_ID_drop'))
    df_fato.drop(columns=[c for c in df_fato.columns if '_Pessoa_ID_drop' in c], inplace=True, errors='ignore')
    
    # Merge Cargo
    df_fato = pd.merge(df_fato, dim_cargo[['Cargo', 'Cargo_ID']], on='Cargo', how='left')
    
    # Merge C.Custo
    df_fato = pd.merge(df_fato, dim_ccusto[['C.Custo', 'CCusto_ID']], on='C.Custo', how='left')

    # Merge Empresa (Obtém Empresa_ID)
    df_fato = pd.merge(df_fato, dim_empresa[['Empresa', 'Empresa_ID']], on='Empresa', how='left')

    # Merge Filial (Obtém Filial_ID) - Chave Composta: Filial + Empresa_ID (Snowflake)
    dim_filial_chaves = dim_filial[['Filial_ID', 'Filial', 'Empresa_ID']]

    df_fato = pd.merge(
        df_fato, 
        dim_filial_chaves, 
        on=['Filial', 'Empresa_ID'], 
        how='left',
        suffixes=('_drop', '') 
    )

    print("Todos os novos IDs (Surrogate Keys) foram adicionados à Tabela Fato.")

    # 2. Seleção e Limpeza da Tabela Fato Final
    # Remoção das colunas originais que viraram IDs/Chaves naturais
    # ** ALTERAÇÃO: Adiciona 'CPF' à lista de colunas a serem descartadas **
    colunas_para_descartar = ['Nome', 'CPF', 'Cargo', 'C.Custo', 'Empresa', 'Filial', 'Sexo']
    df_fato.drop(columns=[col for col in colunas_para_descartar if col in df_fato.columns], 
                 inplace=True, errors='ignore')


    # B. Selecionar colunas para a Tabela Fato
    colunas_fato_desejadas = [
        # NOVOS IDs (Chaves Estrangeiras)
        'Pessoa_ID', 'Cargo_ID', 'CCusto_ID', 'Filial_ID', 'Empresa_ID', 
        # CHAVE DE AUDITORIA
        'Cadastro', 
        # Dados de Contrato/Fato
        'Admissão', 'Data Afastamento', 'Data Salário', 'Data Cargo', 'Data C.Custo',
        'Data de Reintegração', 'Data Vínculo', 'Última Simulação', 'Data Inclusão', 
        '% Desempenho', '% Insalubridade', '% Periculosidade', '% Reajuste', 
        '% FGTS', '% ISS', 'Dependentes IR', 'Dependentes Saf',
        'Situação', 'Descrição (Situação)', 'Causa', 'Descrição (Causa)', 'Escala',
        'Descrição (Escala)', 'Opção FGTS', 'Período Pagto', 'Descrição (Período Pagto)',
        'Descrição (T. Adm)', 'Descrição (T. Contrato)', 'Descrição (Cat. eSocial)',
        'Descrição (Motivo Alt. Salário)', 'Recebe 13° Salário', 'Código Fornecedor',
        'Origem_Arquivo' # Mantém a coluna de auditoria do arquivo original
    ]

    df_fato_final = df_fato[[col for col in colunas_fato_desejadas if col in df_fato.columns]].copy()
    print("Tabela Fato final criada com as colunas de IDs e medidas.")

    # ===================================================================
    # 6. SALVAMENTO EM MÚLTIPLAS ABAS (EXCEL)
    # ===================================================================
    
    # --- ADIÇÃO DE MÉTRICAS FINAIS SOLICITADAS ---
    total_colaboradores_unicos = len(dim_pessoa)
    total_contratos = len(df_fato_final)
    
    print("\n" + "#"*70)
    print("MÉTRICAS CHAVE DO MODELO DE DADOS")
    print(f"1. Total de Contratos/Registros (Tabela Fato): {total_contratos}")
    print(f"2. Total de Colaboradores Únicos (Baseado em CPF): {total_colaboradores_unicos}")
    print("#"*70)
    
    print(f"\nSalvando Modelo Estrela/Snowflake em EXCEL: {ARQUIVO_FINAL_MODELADO} (Múltiplas abas)")
    try:
        with pd.ExcelWriter(ARQUIVO_FINAL_MODELADO, engine='xlsxwriter') as writer:
            df_fato_final.to_excel(writer, sheet_name='Fato_Contratos', index=False)
            
            # Dimensões
            # ** ALTERAÇÃO: Dim_Pessoa agora mantém Nome e Cadastro como atributos, CPF é a chave natural **
            dim_pessoa.to_excel(writer, sheet_name='Dim_Pessoa', index=False)

            dim_cargo.drop(columns=['Cargo']).to_excel(writer, sheet_name='Dim_Cargo', index=False)
            dim_ccusto.drop(columns=['C.Custo']).to_excel(writer, sheet_name='Dim_CCusto', index=False)

            # Dimensões Empresa e Filial (Hierárquicas)
            dim_filial.to_excel(writer, sheet_name='Dim_Filial', index=False) 
            dim_empresa.drop(columns=['Empresa']).to_excel(writer, sheet_name='Dim_Empresa', index=False) 
            
            # ********************************************************************
            # ** SALVANDO AS NOVAS TABELAS AUXILIARES **
            # ********************************************************************
            if df_faltas is not None:
                df_faltas.to_excel(writer, sheet_name='Fato_Faltas', index=False)
                print("Tabela auxiliar 'Fato_Faltas' salva.")
            if df_abs is not None:
                df_abs.to_excel(writer, sheet_name='Fato_Absenteismo', index=False)
                print("Tabela auxiliar 'Fato_Absenteismo' salva.")
            
            # ********************************************************************

        print("\n" + "="*70)
        print("Modelagem e salvamento CONCLUÍDOS com sucesso!")
        print(f"O arquivo EXCEL '{ARQUIVO_FINAL_MODELADO}' (6 + Auxiliares abas) está pronto.")
        print("Instrução para Power BI: As chaves 'Empresa_ID' e 'Filial_ID' ligam as dimensões à Fato, formando o Snowflake/Estrela.")
        print("="*70)

    except Exception as e:
        print(f"\nERRO ao salvar o arquivo Excel: {e}")


# =======================================================================
# 5. EXECUÇÃO DO FLUXO COMPLETO
# =======================================================================
def run_full_etl():
    """Executa a consolidação seguida da modelagem."""
    
    # 1. Executa a Etapa de Consolidação e salva o CSV intermediário
    df_consolidado = etl_consolida_e_salva_csv(COLUNAS_DESEJADAS)
    
    # 2. Executa a Etapa de Modelagem
    if df_consolidado is not None:
        etl_modela_e_salva_excel(df_consolidado, COLUNAS_DESEJADAS)

if __name__ == "__main__":
    # Verificação simples para lembrar de rodar o setup, caso a pasta de origem não exista
    if not os.path.exists(PASTA_ORIGEM_ONEDRIVE) or not os.path.exists(ARQUIVO_SEXO):
        print("ALERTA: O ambiente de teste não foi encontrado. Verifique PASTA_ORIGEM_ONEDRIVE e ARQUIVO_SEXO.")
    else:
        run_full_etl()