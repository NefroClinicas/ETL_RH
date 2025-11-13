import pandas as pd
import os
import time
from datetime import date
import re
from gender_guesser_br import Genero
import unicodedata

PASTA_ORIGEM_ONEDRIVE = r"C:\Users\LuisGuilhermeMoraesd\Nefroclinicas Serviço de Nefrologia e Dialise Ltda\Nefroclinicas - 07 - DADOS (1)"
PASTA_TRABALHO = r"C:\Users\LuisGuilhermeMoraesd\OneDrive - Nefroclinicas Serviço de Nefrologia e Dialise Ltda\Área de Trabalho\Limpar_Dados"

# Nomes dos arquivos de saída e auxiliares
ARQUIVO_SEXO = os.path.join(PASTA_TRABALHO, "Nomes_e_Sexo_Inferido.xlsx")
ARQUIVO_INTERMEDIARIO_CSV = os.path.join(PASTA_TRABALHO, "Base_BI_Consolidada2.csv")
ARQUIVO_FINAL_MODELADO = os.path.join(PASTA_TRABALHO, "Base_MODELADA_PowerBI_V4.xlsx")

# Arquivos Auxiliares (Faltas e Absenteísmo)
ARQUIVO_FALTAS = os.path.join(PASTA_TRABALHO, "faltas_cpf.csv")
ARQUIVO_ABS = os.path.join(PASTA_TRABALHO, "abs_atualizado.csv")

# Nome/Prefixo do arquivo único que contém todas as unidades (procura por nomes que comecem com isso)
ARQUIVO_PREFIXO_UNICO = "relatório turnover - att"

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

def extrair_primeiro_nome(nome_completo):
    if pd.isna(nome_completo):
        return None
    nome = str(nome_completo).strip()
    nome_limpo_completo = re.sub(r'[^a-zA-Z\s]', ' ', nome) 
    palavras = [p for p in nome_limpo_completo.split() if p]
    if not palavras:
        return None
    return palavras[0].capitalize()

def inferir_sexo_br(primeiro_nome):
    if not primeiro_nome:
        return None 
    try:
        resultado = Genero(primeiro_nome)() 
        if resultado == 'masculino':
            return 'Masculino'
        elif resultado == 'feminino':
            return 'Feminino'
        else:
            return None
    except Exception:
        return None 

def formatar_cpf(cpf):
    cpf = str(cpf)
    cpf = re.sub(r'[^0-9]', '', cpf)
    if len(cpf) == 11:
        return f'{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}'
    else:
        return None

def padronizar_nome_pessoa(nome):
    if pd.isna(nome):
        return None
    nome = str(nome).strip()
    nfkd = unicodedata.normalize('NFKD', nome)
    nome_sem_acento = nfkd.encode('ASCII', 'ignore').decode('utf-8')
    nome_limpo = re.sub(r'[^A-Z\s]', '', nome_sem_acento.upper())
    nome_limpo = re.sub(r'\s+', ' ', nome_limpo)  # remove espaços duplos
    return nome_limpo.strip()

def carregar_dados_sexo(caminho_arquivo):
    print(f"\n-> Carregando dados de Sexo de: {os.path.basename(caminho_arquivo)}")
    try:
        df_sexo = pd.read_excel(caminho_arquivo)
        df_sexo.columns = df_sexo.columns.str.strip()
        if 'Nome' not in df_sexo.columns or 'Sexo' not in df_sexo.columns:
            print("  ERRO: O arquivo de sexo não contém as colunas 'Nome' e/ou 'Sexo'.")
            return None
        df_sexo = df_sexo[['Nome', 'Sexo']].copy()
        df_sexo['Nome'] = df_sexo['Nome'].astype(str).str.strip().str.upper()
        df_sexo.drop_duplicates(subset=['Nome'], keep='first', inplace=True)
        valores_a_anular = ['DESCONHECIDO', 'AMBOS', '']
        df_sexo['Sexo'] = df_sexo['Sexo'].astype(str).str.upper()
        df_sexo.loc[df_sexo['Sexo'].isin(valores_a_anular), 'Sexo'] = pd.NA
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
    nome_arquivo = os.path.basename(caminho_arquivo_local)
    if nome_arquivo.lower().endswith(('.xls', '.xlsx')):
        print(f"  -> Processando com Pandas: {nome_arquivo} (Lendo como Excel)")
        try:
            df = pd.read_excel(caminho_arquivo_local)
        except Exception as e:
            print(f"  ERRO CRÍTICO ao ler {nome_arquivo} como Excel: {e}")
            return None
    else:
        print(f"  -> Processando com Pandas: {nome_arquivo} (Lendo como CSV, tentando delimitadores)")
        try:
            df = pd.read_csv(caminho_arquivo_local, sep=';', encoding='latin1') 
        except Exception:
            try:
                df = pd.read_csv(caminho_arquivo_local, sep=',', encoding='latin1')
            except Exception as e:
                print(f"  ERRO CRÍTICO ao ler {nome_arquivo} como CSV. Falha com ';' e ','.")
                print(f"  Detalhes: {e}")
                return None

    df.columns = df.columns.str.strip()
    colunas_reais = df.columns.tolist()
    if 'Valor Salário' not in colunas_reais and 'Salário Simulado' in colunas_reais:
        df.rename(columns={'Salário Simulado': 'Valor Salário'}, inplace=True)
        print(f"  AVISO: 'Salário Simulado' renomeado para 'Valor Salário' em {nome_arquivo}.")
    colunas_presentes = [col.strip() for col in colunas_desejadas if col.strip() in df.columns]
    colunas_ausentes = [col.strip() for col in colunas_desejadas if col.strip() not in df.columns]
    if colunas_ausentes:
        print(f"  AVISO: Colunas não encontradas e IGNORADAS em {nome_arquivo}: {colunas_ausentes}")
    try:
        df_selecionado = df[colunas_presentes].copy()
        if 'Nome' in df_selecionado.columns:
            df_selecionado['Nome'] = df_selecionado['Nome'].astype(str).str.strip().str.upper()
        return df_selecionado
    except KeyError as e:
        print(f"  ERRO: Seleção final falhou após o mapeamento. Detalhes: {e}")
        return None

def limpar_chave(df, coluna_id):
    if coluna_id in df.columns:
        if df[coluna_id].dtype in ['int64', 'float64']:
            df[coluna_id] = df[coluna_id].astype(str).str.replace(r'\.0$', '', regex=True)
        df[coluna_id] = df[coluna_id].astype(str).str.strip().str.upper().str.replace('NAN', '')
    return df

def calcular_idade_faixa_etaria(df):
    if 'Nascimento' in df.columns:
        print("Calculando 'Idade Atual' e 'Faixa Etária'...")
        hoje = pd.to_datetime(date.today())
        df['Nascimento'] = pd.to_datetime(df['Nascimento'], errors='coerce', dayfirst=True)
        df['Idade Atual'] = (hoje - df['Nascimento']).dt.days / 365.25
        df['Idade Atual'] = df['Idade Atual'].apply(lambda x: int(x) if pd.notna(x) else pd.NA).astype('Int64')
        bins = [0, 20, 25, 30, 35, 40, 50, 60, 100]
        labels = [
            "1. Abaixo de 20", "2. 20 - 24 Anos", "3. 25 - 29 Anos", "4. 30 - 34 Anos",
            "5. 35 - 39 Anos", "6. 40 - 49 Anos", "7. 50 - 59 Anos", "8. 60 ou Mais"
        ]
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

def etl_processa_csv_auxiliar(caminho_arquivo, nome_tabela):
    print(f"\n-> Processando arquivo auxiliar: {os.path.basename(caminho_arquivo)} para a tabela {nome_tabela}")
    try:
        df = pd.read_csv(caminho_arquivo, sep=';', encoding='utf-8')
    except FileNotFoundError:
        print(f"  AVISO: Arquivo '{os.path.basename(caminho_arquivo)}' não encontrado. Pulando.")
        return None
    except Exception as e:
        try:
            df = pd.read_csv(caminho_arquivo, sep=',', encoding='latin1')
        except Exception as e:
            print(f"  ERRO CRÍTICO ao ler {os.path.basename(caminho_arquivo)} como CSV. Detalhes: {e}")
            return None
    df.columns = df.columns.str.strip().str.replace(r'[^a-zA-Z0-9\s\(\)%]', '', regex=True)
    if 'Nome' in df.columns:
        df['Nome'] = df['Nome'].astype(str).str.strip().str.upper()
        print("  Coluna 'Nome' padronizada (Upper, Strip).")
    colunas_para_limpar = ['Previsto', 'Ausencia', 'Presenca']
    for col in colunas_para_limpar:
        if col in df.columns:
            df[col] = (
                df[col]
                .astype(str)
                .str.strip()
                .str.replace(r'(?<=\d{2}):\d{2}$', '', regex=True)
                .apply(lambda x: re.sub(r'^(\d{1,3}):(\d{1})$', r'\1:0\2', x))
                .str.strip()
            )
            print(f"  ✅ Coluna '{col}' padronizada para HH:MM (sem segundos).")
        else:
            print(f"  ⚠️ Coluna '{col}' não encontrada para limpeza.")
    print(f"  Tabela '{nome_tabela}' processada com {len(df)} linhas.")
    return df

# =======================================================================
# 3. LÓGICA PRINCIPAL: CONSOLIDAÇÃO (ETL) - atualizado para arquivo único
# =======================================================================

def etl_consolida_e_salva_csv(colunas_desejadas):
    lista_dataframes = []
    print("\n" + "="*70)
    print(f"ETAPA 1/2: CONSOLIDAÇÃO E INFERÊNCIA DE SEXO | Início: {time.ctime()}")
    print("="*70)

    # PASSO 0: Carregar Dados Auxiliares (Sexo)
    df_sexo = carregar_dados_sexo(ARQUIVO_SEXO)

    # 1. BUSCAR O ARQUIVO ÚNICO PELO PREFIXO
    try:
        arquivos_na_pasta = [
            f for f in os.listdir(PASTA_ORIGEM_ONEDRIVE)
            if os.path.isfile(os.path.join(PASTA_ORIGEM_ONEDRIVE, f)) and not f.startswith('.')
        ]
    except FileNotFoundError:
        print(f"\nERRO: Pasta de origem não encontrada: {PASTA_ORIGEM_ONEDRIVE}.")
        return None

    arquivos_match = [
        os.path.join(PASTA_ORIGEM_ONEDRIVE, f)
        for f in arquivos_na_pasta
        if f.lower().startswith(ARQUIVO_PREFIXO_UNICO) and f.lower().endswith(('.xls', '.xlsx'))
    ]

    if not arquivos_match:
        print("\n" + "="*70)
        print(f"ERRO: Nenhum arquivo começando com '{ARQUIVO_PREFIXO_UNICO}' (.xls/.xlsx) foi encontrado na pasta especificada.")
        print(f"Caminho verificado: {PASTA_ORIGEM_ONEDRIVE}")
        print("="*70)
        return None

    # pega o primeiro encontrado (se houver mais de um, usa o primeiro)
    arquivo_unico = arquivos_match[0]
    print(f"-> Arquivo único selecionado para processamento: {os.path.basename(arquivo_unico)}")

    # 2. Processamento do arquivo único (mantendo a lógica original)
    df_processado = transformar_e_selecionar(arquivo_unico, colunas_desejadas)

    if df_processado is not None and not df_processado.empty:
        df_processado['Origem_Arquivo'] = os.path.basename(arquivo_unico)
        lista_dataframes.append(df_processado)
        print(f" Arquivo único processado com sucesso. Linhas: {len(df_processado)}")
    else:
        print(" ERRO: Não foi possível processar o arquivo único.")
        return None

    # 3. Combinação Final (apenas um arquivo agora)
    if lista_dataframes:
        CONTEUDO_FINAL = pd.concat(lista_dataframes, ignore_index=True, sort=False)

        # --- MERGE E INFERÊNCIA CONDICIONAL DE SEXO ---
        CONTEUDO_FINAL.drop(columns=['Sexo'], inplace=True, errors='ignore')
        if 'Sexo' not in CONTEUDO_FINAL.columns:
            CONTEUDO_FINAL['Sexo'] = pd.NA

        if df_sexo is not None and 'Nome' in CONTEUDO_FINAL.columns:
            print("\n-> Realizando Merge com a informação de Sexo do Excel (Prioridade Manual)...")
            df_merged = pd.merge(
                CONTEUDO_FINAL.drop(columns=['Sexo'], errors='ignore'),
                df_sexo,
                on='Nome',
                how='left'
            )
            CONTEUDO_FINAL = df_merged
            print(f" Merge concluído. {CONTEUDO_FINAL['Sexo'].count()} valores de Sexo manual/existente carregados.")
        else:
            print(" AVISO: Arquivo de Sexo manual não carregado. Pulando merge.")

        condicao_inferir = CONTEUDO_FINAL['Sexo'].isna()
        df_a_inferir = CONTEUDO_FINAL[condicao_inferir].copy()

        if not df_a_inferir.empty:
            total_a_inferir = len(df_a_inferir)
            print(f"\n -> Aplicando Inferência do IBGE (Lenta) a {total_a_inferir} registros NA/Nulos...")
            df_a_inferir['Primeiro_Nome'] = df_a_inferir['Nome'].apply(extrair_primeiro_nome)
            novos_sexos = df_a_inferir['Primeiro_Nome'].apply(inferir_sexo_br)
            sucesso_br = novos_sexos.dropna()
            CONTEUDO_FINAL.loc[sucesso_br.index, 'Sexo'] = sucesso_br
            print(f" {sucesso_br.count()} valores de Sexo preenchidos por Inferência.")
        else:
            print(" Nenhum valor nulo (NA) de Sexo restante para inferência. Passo ignorado.")

        # Garante que todas as colunas desejadas estejam presentes
        for col in [str(coluna).strip() for coluna in colunas_desejadas if coluna is not None]:
            if col not in CONTEUDO_FINAL.columns:
                CONTEUDO_FINAL[col] = pd.NA

        # Salvar o Resultado Intermediário (CSV)
        colunas_finais_ordenadas = [c.strip() for c in colunas_desejadas] + ['Origem_Arquivo']
        colunas_para_salvar = [col for col in colunas_finais_ordenadas if col in CONTEUDO_FINAL.columns]
        CONTEUDO_FINAL[colunas_para_salvar].to_csv(ARQUIVO_INTERMEDIARIO_CSV, index=False, encoding='utf-8')

        print("\n" + "="*70)
        print(f"ETAPA 1/2 COMPLETA! CSV Intermediário salvo em: {ARQUIVO_INTERMEDIARIO_CSV}")
        print(f"Total de linhas consolidadas: {len(CONTEUDO_FINAL)}")
        print("="*70)
        return CONTEUDO_FINAL
    else:
        print("\nNenhum arquivo pôde ser processado com sucesso.")
        return None

# =======================================================================
# 4. LÓGICA PRINCIPAL: MODELAGEM (ELT) - MANTIDA
# =======================================================================

def etl_modela_e_salva_excel(df_input, colunas_desejadas):
    if df_input is None or df_input.empty:
        print("ERRO: O DataFrame consolidado está vazio ou não foi gerado.")
        return

    print("\n" + "="*70)
    print("ETAPA 2/2: MODELAGEM E CRIAÇÃO DE MODELO ESTRELA/SNOWFLAKE")
    print("="*70)

    df = df_input.copy()

    # Processamento das tabelas auxiliares
    df_faltas = etl_processa_csv_auxiliar(ARQUIVO_FALTAS, 'Fato_Faltas')
    df_abs = etl_processa_csv_auxiliar(ARQUIVO_ABS, 'Fato_Absenteismo')

    # 1. SELEÇÃO DE COLUNAS
    colunas_para_selecao = [col for col in colunas_desejadas if col in df.columns]
    df = df[colunas_para_selecao + ['Origem_Arquivo']].copy()
    print(f"DataFrame filtrado para {len(df.columns)} colunas desejadas e existentes.")

    # 2. LIMPEZA E TRATAMENTO DE DADOS
    print("\nIniciando limpeza e padronização das chaves de relacionamento...")
    colunas_para_limpar = ['Nome', 'CPF', 'Empresa', 'Cadastro', 'Cargo', 'C.Custo', 'Filial']
    for col in colunas_para_limpar:
        df = limpar_chave(df, col)
    print("Limpeza de chaves de relacionamento (strip e upper) concluída.")

    df.replace('', pd.NA, inplace=True)
    df.dropna(subset=['CPF', 'Nome'], inplace=True)
    df = df[df['CPF'] != 'NAN']
    df = df[df['Nome'] != 'NAN']
    print(f"Linhas com CPF e/ou Nome vazios removidas. Linhas restantes: {len(df)}")

    if 'Descrição (T. Adm)' in df.columns:
        df['Descrição (T. Adm)'] = df['Descrição (T. Adm)'].fillna('NÃO INFORMADO')
        print("Valores nulos em 'Descrição (T. Adm)' preenchidos.")
    if 'Sexo' in df.columns:
        df['Sexo'] = df['Sexo'].fillna('Não Definido/Inferido')
        print("Valores nulos em 'Sexo' preenchidos com 'Não Definido/Inferido'.")

    # 3. CONVERSÃO DE DATAS
    COLUNAS_DE_DATA = [col for col in colunas_desejadas if 'Data' in col or 'Admissão' in col or 'Nascimento' in col or 'Última Simulação' in col]
    print("\nIniciando conversão de colunas de data...")
    for col in COLUNAS_DE_DATA:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce', dayfirst=True)
    print("Conversão de datas concluída.")

    # ===================================================================
    # 4. GERAÇÃO DAS TABELAS DIMENSÃO
    # ===================================================================
    print("\nGerando Tabelas Dimensão com ID Sequencial (Chaves Suplementares)...")

    # Dim Empresa
    dim_empresa = df[['Empresa', 'Nome (Empresa)']].copy().dropna(subset=['Empresa'])
    dim_empresa.drop_duplicates(subset=['Empresa'], inplace=True)
    dim_empresa.insert(0, 'Empresa_ID', range(1, len(dim_empresa) + 1))
    print(f"Dimensão Empresa (Pai) criada com {len(dim_empresa)} registros únicos.")

    # Dim Filial
    colunas_filial_origem = ['Filial', 'Apelido (Filial)', 'Empresa']
    dim_filial = df[[col for col in colunas_filial_origem if col in df.columns]].copy().dropna(subset=['Filial'])
    dim_filial.drop_duplicates(subset=['Filial', 'Empresa'], inplace=True)
    if 'Apelido (Filial)' in dim_filial.columns:
        dim_filial['Apelido (Filial)'] = dim_filial['Apelido (Filial)'].fillna('NÃO INFORMADO')
    dim_filial.insert(0, 'Filial_ID', range(1, len(dim_filial) + 1))
    dim_filial = pd.merge(dim_filial, dim_empresa[['Empresa', 'Empresa_ID']], on='Empresa', how='left')
    dim_filial.drop(columns=['Empresa'], inplace=True, errors='ignore')
    print(f"Dimensão Filial (Filho) criada com {len(dim_filial)} registros únicos e ligada à Empresa (FK Empresa_ID).")
    colunas_pessoa = [
        'Nome', 'CPF', 'Cadastro', 'Nascimento', 'Sexo',
        'Estado Civil', 'Descrição (Estado Civil)', 'Instrução', 'Descrição (Instrução)',
        'Raça/Etnia', 'Descrição (Raça/Etnia)', 'Dependentes IR', 'Dependentes Saf',
        'Dep. Saldo FGTS', 'Cadastramento PIS', 'Nome (Cadastro O. Contrato)'
    ]
    cols_existentes_pessoa = [col for col in colunas_pessoa if col in df.columns]
    dim_pessoa = df[cols_existentes_pessoa].copy()

    # Padroniza nome (remove acentos, deixa maiúsculo e sem espaços duplos)
    dim_pessoa['Nome_Padrao'] = dim_pessoa['Nome'].apply(padronizar_nome_pessoa)

    # Remove pontos e traços do CPF, mas não usa mais como critério de duplicidade
    dim_pessoa['CPF'] = (
        dim_pessoa['CPF']
        .astype(str)
        .str.replace(r'[^0-9]', '', regex=True)
        .replace({'': pd.NA, 'nan': pd.NA, 'NaN': pd.NA})
    )

    # Converte nascimento para data (se ainda não estiver)
    dim_pessoa['Nascimento'] = pd.to_datetime(dim_pessoa['Nascimento'], errors='coerce', dayfirst=True)

    # Cria chave composta Nome + Data Nascimento
    dim_pessoa['Chave_Nome_Nasc'] = (
        dim_pessoa['Nome_Padrao'].fillna('') + "_" + dim_pessoa['Nascimento'].astype(str).fillna('')
    )

    # Remove duplicatas
    antes = len(dim_pessoa)
    dim_pessoa.drop_duplicates(subset=['Chave_Nome_Nasc'], inplace=True, keep='first')
    depois = len(dim_pessoa)
    print(f"✅ Removidas {antes - depois} duplicatas baseadas em Nome + Data de Nascimento.")

    # Gera o ID sequencial
    dim_pessoa.insert(0, 'Pessoa_ID', range(1, len(dim_pessoa) + 1))

    # Calcula idade e faixa etária
    dim_pessoa = calcular_idade_faixa_etaria(dim_pessoa)

    # Padroniza novamente o Nome final
    if 'Nome' in dim_pessoa.columns:
        dim_pessoa['Nome'] = dim_pessoa['Nome'].apply(padronizar_nome_pessoa)
    if 'Nome (Cadastro O. Contrato)' in dim_pessoa.columns:
        dim_pessoa['Nome (Cadastro O. Contrato)'] = dim_pessoa['Nome (Cadastro O. Contrato)'].apply(padronizar_nome_pessoa)

    print(f"Dimensão Pessoa criada com {len(dim_pessoa)} registros únicos (base Nome + Nascimento).")


    # Dim Cargo
    dim_cargo = df[['Cargo', 'Título Reduzido (Cargo)']].copy().dropna(subset=['Cargo'])
    dim_cargo.drop_duplicates(subset=['Cargo'], inplace=True)
    dim_cargo.insert(0, 'Cargo_ID', range(1, len(dim_cargo) + 1))
    print(f"Dimensão Cargo criada com {len(dim_cargo)} registros únicos.")

    # Dim CCusto
    dim_ccusto = df[['C.Custo', 'Descrição (C.Custo)']].copy().dropna(subset=['C.Custo'])
    dim_ccusto.drop_duplicates(subset=['C.Custo'], inplace=True)
    dim_ccusto.insert(0, 'CCusto_ID', range(1, len(dim_ccusto) + 1))
    print(f"Dimensão C.Custo criada com {len(dim_ccusto)} registros únicos.")

    # ===================================================================
    # 5. CRIAÇÃO DA TABELA FATO (Fato_Contratos)
    # ===================================================================
    print("\nPreparando a Tabela Fato com todos os novos IDs...")
    df_fato = df.copy()

    # Merge Pessoa por CPF
    if 'CPF' in df_fato.columns and 'CPF' in dim_pessoa.columns:
        df_fato = pd.merge(df_fato, dim_pessoa[['Pessoa_ID', 'CPF']], on='CPF', how='left')

    faltando_pessoa = df_fato['Pessoa_ID'].isna().sum()
    if faltando_pessoa > 0 and 'Nome' in df_fato.columns:
        print(f"⚠️ {faltando_pessoa} registros sem Pessoa_ID via CPF. Tentando merge por Nome padronizado...")
        df_fato['Nome_Padrao'] = df_fato['Nome'].apply(padronizar_nome_pessoa)
        df_fato = pd.merge(
            df_fato.drop(columns=['Pessoa_ID'], errors='ignore'),
            dim_pessoa[['Pessoa_ID', 'Nome']].rename(columns={'Nome': 'Nome_Padrao'}),
            on='Nome_Padrao', how='left'
        )
    df_fato.drop(columns=['Nome_Padrao'], inplace=True, errors='ignore')
    print(f"✅ Merge final de Pessoa_ID concluído. Total de registros com ID: {df_fato['Pessoa_ID'].notna().sum()}")

    # Merge Cargo, CCusto, Empresa e Filial
    df_fato = pd.merge(df_fato, dim_cargo[['Cargo', 'Cargo_ID']], on='Cargo', how='left')
    df_fato = pd.merge(df_fato, dim_ccusto[['C.Custo', 'CCusto_ID']], on='C.Custo', how='left')
    df_fato = pd.merge(df_fato, dim_empresa[['Empresa', 'Empresa_ID']], on='Empresa', how='left')
    dim_filial_chaves = dim_filial[['Filial_ID', 'Filial', 'Empresa_ID']]
    df_fato = pd.merge(df_fato, dim_filial_chaves, on=['Filial', 'Empresa_ID'], how='left', suffixes=('_drop', ''))
    print("Todos os novos IDs (Surrogate Keys) foram adicionados à Tabela Fato.")

    # Remoção de colunas que viraram IDs
    colunas_para_descartar = ['Nome', 'CPF', 'Cargo', 'C.Custo', 'Empresa', 'Filial', 'Sexo']
    df_fato.drop(columns=[col for col in colunas_para_descartar if col in df_fato.columns], inplace=True, errors='ignore')

    # Seleção final das colunas da tabela fato
    colunas_fato_desejadas = [
        'Pessoa_ID', 'Cargo_ID', 'CCusto_ID', 'Filial_ID', 'Empresa_ID',
        'Cadastro', 'Admissão', 'Data Afastamento', 'Data Salário', 'Data Cargo',
        'Data C.Custo', 'Data de Reintegração', 'Data Vínculo', 'Última Simulação',
        'Data Inclusão', '% Desempenho', '% Insalubridade', '% Periculosidade',
        '% Reajuste', '% FGTS', '% ISS', 'Dependentes IR', 'Dependentes Saf',
        'Situação', 'Descrição (Situação)', 'Causa', 'Descrição (Causa)',
        'Escala', 'Descrição (Escala)', 'Opção FGTS', 'Período Pagto',
        'Descrição (Período Pagto)', 'Descrição (T. Adm)', 'Descrição (T. Contrato)',
        'Descrição (Cat. eSocial)', 'Descrição (Motivo Alt. Salário)',
        'Recebe 13° Salário', 'Código Fornecedor', 'Origem_Arquivo'
    ]
    df_fato_final = df_fato[[col for col in colunas_fato_desejadas if col in df_fato.columns]].copy()
    print("Tabela Fato final criada com as colunas de IDs e medidas.")

    # ===================================================================
    # 6. SALVAMENTO EM MÚLTIPLAS ABAS (EXCEL)
    # ===================================================================
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
            dim_pessoa.to_excel(writer, sheet_name='Dim_Pessoa', index=False)
            dim_cargo.drop(columns=['Cargo']).to_excel(writer, sheet_name='Dim_Cargo', index=False)
            dim_ccusto.drop(columns=['C.Custo']).to_excel(writer, sheet_name='Dim_CCusto', index=False)
            dim_filial.to_excel(writer, sheet_name='Dim_Filial', index=False)
            dim_empresa.drop(columns=['Empresa']).to_excel(writer, sheet_name='Dim_Empresa', index=False)

            if df_faltas is not None:
                df_faltas.to_excel(writer, sheet_name='Fato_Faltas', index=False)
                print("Tabela auxiliar 'Fato_Faltas' salva.")
            if df_abs is not None:
                df_abs.to_excel(writer, sheet_name='Fato_Absenteismo', index=False)
                print("Tabela auxiliar 'Fato_Absenteismo' salva.")

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
    df_consolidado = etl_consolida_e_salva_csv(COLUNAS_DESEJADAS)
    if df_consolidado is not None:
        etl_modela_e_salva_excel(df_consolidado, COLUNAS_DESEJADAS)

if __name__ == "__main__":
    # Verificação da existência do arquivo único pelo prefixo
    try:
        arquivos_na_pasta = [
            f for f in os.listdir(PASTA_ORIGEM_ONEDRIVE)
            if os.path.isfile(os.path.join(PASTA_ORIGEM_ONEDRIVE, f)) and not f.startswith('.')
        ]
    except FileNotFoundError:
        print(f"ALERTA: Pasta de origem não encontrada: {PASTA_ORIGEM_ONEDRIVE}")
        arquivos_na_pasta = []

    encontrados = [f for f in arquivos_na_pasta if f.lower().startswith(ARQUIVO_PREFIXO_UNICO) and f.lower().endswith(('.xls', '.xlsx'))]
    if not encontrados:
        print(f"ALERTA: Nenhum arquivo começando com '{ARQUIVO_PREFIXO_UNICO}' (.xls/.xlsx) encontrado em: {PASTA_ORIGEM_ONEDRIVE}")
        print("Coloque o arquivo na pasta ou ajuste ARQUIVO_PREFIXO_UNICO se o nome mudou.")
    elif not os.path.exists(ARQUIVO_SEXO):
        print("ALERTA: Arquivo de sexo não encontrado. Verifique ARQUIVO_SEXO.")
    else:
        run_full_etl()
