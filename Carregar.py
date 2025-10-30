import pandas as pd
import os
import time



PASTA_ORIGEM_ONEDRIVE = r"C:\Users\LuisGuilhermeMoraesd\Nefroclinicas Serviço de Nefrologia e Dialise Ltda\Nefroclinicas - 07 - DADOS (1)"

# Diretório de trabalho onde o arquivo final e o arquivo de sexo estão localizados
PASTA_TRABALHO = r"C:\Users\LuisGuilhermeMoraesd\OneDrive - Nefroclinicas Serviço de Nefrologia e Dialise Ltda\Área de Trabalho\Limpar_Dados"

ARQUIVO_FINAL_CSV = os.path.join(PASTA_TRABALHO, "Base_BI_Consolidada2.csv")

# NOVO: Caminho para o arquivo que contém o Nome e o Sexo (ajuste o nome se for diferente)
# IMPORTANTE: O arquivo é .xlsx, por isso a função de carregamento foi corrigida.
ARQUIVO_SEXO = os.path.join(PASTA_TRABALHO, "Nomes_e_Sexo_Inferido.xlsx")


# Lista EXATA das colunas a serem extraídas (incluindo 'Sexo')
COLUNAS_DESEJADAS = [
    'Nome', 
    'Sexo', # <-- NOVA COLUNA ADICIONADA AQUI
    'Empresa',
    'Cadastro',
    'Admissão',
    'Cargo',
    'C.Custo',
    'Descrição (C.Custo)',
    'Data Afastamento',
    'Título Reduzido (Cargo)',
    'Descrição (Raça/Etnia)',
    'Descrição (Cat. eSocial)',
    'Causa',
    'Descrição (Causa)',
    'Escala',
    'Descrição (Escala)',
    'Filial',
    'Apelido (Filial)',
    'Código Fornecedor',
    'Descrição (Motivo Alt. Salário)',
    'Data Adicionais',
    'Data Aposentadoria',
    'Data Cargo',
    'Data C.Custo',
    'Data Ult. Alt. Cat.',
    'Data de Chegada',
    'Data Escala',
    'Data Estabilidade',
    'Data Escala VTR',
    'Data Filial',
    'Data Histórico de Contrato',
    'Data Inclusão',
    'Data Local', 
    'Nascimento', 
    'Opção FGTS',
    'Data Posto',
    'Data Ass. PPR',
    'Data de Reintegração',
    'Data Salário',
    'Data Cat. SEFIP',
    'Última Simulação',
    'Data Sindicato',
    'Data FGTS',
    'Data Vínculo',
    'Cadastramento PIS',
    'Dependentes IR',
    'Dependentes Saf',
    'Dep. Saldo FGTS',
    'Estado Civil',
    'Descrição (Estado Civil)',
    'Instrução',
    'Descrição (Instrução)',
    'Nome (Empresa)', 
    'Nome (Cadastro O. Contrato)', 
    'Nome (Empresa O. Contrato)',
    'Descrição (Tipo O. Contrato)',
    '% Desempenho',
    '% Insalubridade',
    '% Base IR Transportista',
    '% ISS',
    '% FGTS',
    'Período Pagto',
    'Descrição (Período Pagto)',
    '% Periculosidade',
    '% Reajuste',
    '% Base INSS Transportista',
    'Raça/Etnia',
    'Recebe 13° Salário',
    'Situação',
    'Descrição (Situação)',
    'Descrição (T. Adm)',
    'Descrição (T. Contrato)'
]


# =======================================================================
# 2. FUNÇÕES DE TRANSFORMAÇÃO E SELEÇÃO
# =======================================================================

def carregar_dados_sexo(caminho_arquivo):
    """
    Carrega o arquivo de sexo (agora como EXCEL) e prepara o DataFrame para merge.
    CORREÇÃO: Uso de pd.read_excel ao invés de pd.read_csv.
    """
    print(f"\n-> Carregando dados de Sexo de: {os.path.basename(caminho_arquivo)}")
    try:
        # Tenta carregar o arquivo Excel (.xlsx)
        df_sexo = pd.read_excel(caminho_arquivo)
        
        # Garante que as colunas 'Nome' e 'Sexo' existam e limpa nomes de colunas
        df_sexo.columns = df_sexo.columns.str.strip()

        # Verifica se as colunas necessárias para o merge existem
        if 'Nome' not in df_sexo.columns or 'Sexo' not in df_sexo.columns:
            print("  ERRO: O arquivo de sexo não contém as colunas 'Nome' e/ou 'Sexo' (verifique a grafia no Excel).")
            return None

        # Seleciona apenas as colunas de interesse
        df_sexo = df_sexo[['Nome', 'Sexo']].copy()
        
        # Limpa espaços em branco nos Nomes do arquivo de Sexo para garantir o merge
        df_sexo['Nome'] = df_sexo['Nome'].astype(str).str.strip()
        
        # Remove duplicatas, mantendo a primeira ocorrência (se houver nomes duplicados)
        df_sexo.drop_duplicates(subset=['Nome'], keep='first', inplace=True)
        
        print(f"  Sucesso: {len(df_sexo)} nomes únicos carregados com dados de Sexo.")
        return df_sexo
        
    except FileNotFoundError:
        print(f"  AVISO: Arquivo de sexo não encontrado no caminho: {caminho_arquivo}")
        return None
    except Exception as e:
        print(f"  ERRO ao processar o arquivo de sexo (EXCEL): {e}")
        return None


def transformar_e_selecionar(caminho_arquivo_local):
    
    nome_arquivo = os.path.basename(caminho_arquivo_local)
    print(f"  -> Processando com Pandas: {nome_arquivo} (Lendo como CSV)")

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
    if 'Valor Salário' not in colunas_reais and 'Salário Simulado' in colunas_reais:
        df.rename(columns={'Salário Simulado': 'Valor Salário'}, inplace=True)
        print(f"  AVISO: 'Salário Simulado' renomeado para 'Valor Salário' em {nome_arquivo}.")
    
    # 3. Lista de Colunas Disponíveis
    # Filtra COLUNAS_DESEJADAS (exceto 'Sexo', que será adicionada depois)
    colunas_base_desejadas = [col for col in COLUNAS_DESEJADAS if col != 'Sexo']
    
    colunas_presentes = [col.strip() for col in colunas_base_desejadas if col.strip() in df.columns]

    colunas_ausentes = [col.strip() for col in colunas_base_desejadas if col.strip() not in df.columns]
    
    if colunas_ausentes:
        print(f"  AVISO: Colunas não encontradas e IGNORADAS em {nome_arquivo}: {colunas_ausentes}")


    # 4. Seleção de Colunas
    try:
        df_selecionado = df[colunas_presentes].copy() # Adiciona .copy() para evitar SettingWithCopyWarning
        
        # Limpa espaços em branco nos Nomes do DataFrame principal também
        if 'Nome' in df_selecionado.columns:
            df_selecionado['Nome'] = df_selecionado['Nome'].astype(str).str.strip()

        return df_selecionado
    except KeyError as e:
        print(f"  ERRO: Seleção final falhou após o mapeamento. Detalhes: {e}")
        return None

# =======================================================================
# 3. LÓGICA PRINCIPAL: BUSCA, PROCESSAMENTO, CONSOLIDAÇÃO E MERGE
# =======================================================================

def etl_local_drive():
    
    lista_dataframes = []
    
    print(f"Iniciando ETL Local (Leitura forçada como CSV): {time.ctime()}")

    # ===================================================================
    # PASSO 0: Carregar Dados Auxiliares (Sexo)
    # ===================================================================
    # Chamada atualizada para ler o arquivo Excel
    df_sexo = carregar_dados_sexo(ARQUIVO_SEXO) 
    
    # 1. BUSCAR TODOS OS ARQUIVOS DA PASTA
    todos_arquivos = [
        os.path.join(PASTA_ORIGEM_ONEDRIVE, f)
        for f in os.listdir(PASTA_ORIGEM_ONEDRIVE)
        if os.path.isfile(os.path.join(PASTA_ORIGEM_ONEDRIVE, f)) and not f.startswith('.') 
    ]
    
    # 2. FILTRA OS ARQUIVOS QUE PROVAVELMENTE PODEM SER CSVs
    arquivos_csv_compativeis = [
        f for f in todos_arquivos 
        if f.lower().endswith('.xls') or f.lower().endswith('.csv')
    ]

    if not arquivos_csv_compativeis:
        print("\n" + "="*50)
        print("NENHUM ARQUIVO COMPATÍVEL (.xls ou .csv) ENCONTRADO na pasta especificada.")
        print(f"Caminho verificado: {PASTA_ORIGEM_ONEDRIVE}")
        print("="*50)
        return

    print(f"-> {len(arquivos_csv_compativeis)} arquivos compatíveis encontrados. Iniciando consolidação...")
    
    # 3. Processamento e Consolidação
    for arquivo in arquivos_csv_compativeis:
        df_processado = transformar_e_selecionar(arquivo)
        
        if df_processado is not None:
            df_processado['Origem_Arquivo'] = os.path.basename(arquivo)
            lista_dataframes.append(df_processado)

    # 4. Combinação Final (Concatenar)
    if lista_dataframes:
        CONTEUDO_FINAL = pd.concat(lista_dataframes, ignore_index=True, sort=True)
        
        # ===============================================================
        # PASSO 5: MERGE COM DADOS DE SEXO (USANDO 'Nome')
        # ===============================================================
        
        # Trata o caso onde a coluna 'Sexo' já existe na base consolidada (se vier de algum arquivo de origem)
        # e a remove para garantir que vamos usar o 'Sexo' inferido do Excel.
        if 'Sexo' in CONTEUDO_FINAL.columns:
             print("-> Removendo coluna 'Sexo' existente antes do Merge para usar o Sexo Inferido.")
             CONTEUDO_FINAL.drop(columns=['Sexo'], inplace=True, errors='ignore')
             
        if df_sexo is not None and 'Nome' in CONTEUDO_FINAL.columns:
            print("\n-> Realizando Merge dos dados consolidados com a informação de Sexo Inferido (Excel)...")
            
            # Limpa o 'Nome' na base consolidada (feito acima, mas re-confirma)
            CONTEUDO_FINAL['Nome'] = CONTEUDO_FINAL['Nome'].astype(str).str.strip()
            
            # Merge (Left Join) - Adiciona a coluna 'Sexo' do Excel na base consolidada.
            CONTEUDO_FINAL = pd.merge(
                CONTEUDO_FINAL,
                df_sexo,
                on='Nome',
                how='left'
            )
            print(f"  Merge concluído. {CONTEUDO_FINAL['Sexo'].count()} valores de Sexo adicionados.")
        else:
            # Garante que a coluna 'Sexo' exista mesmo que o merge não tenha ocorrido
            if 'Sexo' not in CONTEUDO_FINAL.columns:
                CONTEUDO_FINAL['Sexo'] = pd.NA
            print("  AVISO: Merge de Sexo ignorado (Arquivo de sexo não carregado ou Nome ausente).")

        
        # 6. Adiciona as colunas ausentes na consolidação final (se alguma faltou em todos)
        for col in [c.strip() for c in COLUNAS_DESEJADAS]:
            if col not in CONTEUDO_FINAL.columns:
                CONTEUDO_FINAL[col] = pd.NA
        
        # 7. Salvar o Resultado Final 
        # ATENÇÃO: Salva o CSV com as colunas na ORDEM ORIGINAL desejada
        colunas_finais_ordenadas = [c.strip() for c in COLUNAS_DESEJADAS] + ['Origem_Arquivo']
        
        # Filtra as colunas que realmente existem após a concatenação
        colunas_para_salvar = [col for col in colunas_finais_ordenadas if col in CONTEUDO_FINAL.columns]

        CONTEUDO_FINAL[colunas_para_salvar].to_csv(ARQUIVO_FINAL_CSV, index=False, encoding='utf-8')
        
        print("\n" + "="*50)
        print(f"ETL COMPLETA! Dados prontos para o Power BI.")
        print(f"Total de linhas consolidadas: {len(CONTEUDO_FINAL)}")
        print(f"Arquivo CSV salvo em: {ARQUIVO_FINAL_CSV}")
        print("="*50)
        
    else:
        print("\nNenhum arquivo pôde ser processado com sucesso.")

# Execução do programa
if __name__ == "__main__":
    etl_local_drive()
