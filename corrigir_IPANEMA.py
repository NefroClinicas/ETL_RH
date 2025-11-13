import pandas as pd
import os
import shutil

# Caminhos dos arquivos
PASTA = r"C:\Users\LuisGuilhermeMoraesd\Nefroclinicas Serviço de Nefrologia e Dialise Ltda\Nefroclinicas - 07 - DADOS (1)"
ARQUIVO_COMPLETO = os.path.join(PASTA, "Relatório Turnover - ATT.xlsx")
ARQUIVO_IPANEMA_CORRETO = os.path.join(PASTA, "ncipa - turnover - Copiar.xlsx")

# Nome da unidade que deve ser substituída
FILTRO_IPANEMA = "Nefroclinicas Ipanema Servico de Nefrolo"

print("🔄 Lendo bases...")

# Faz backup automático antes de sobrescrever
backup_path = ARQUIVO_COMPLETO.replace(".xlsx", "_BACKUP.xlsx")
if not os.path.exists(backup_path):
    shutil.copy2(ARQUIVO_COMPLETO, backup_path)
    print(f"💾 Backup criado: {backup_path}")
else:
    print("⚠️ Backup já existente, mantendo o arquivo anterior de segurança.")

# Lê as duas planilhas
df_completo = pd.read_excel(ARQUIVO_COMPLETO)
df_ipanema = pd.read_excel(ARQUIVO_IPANEMA_CORRETO)

print(f"✅ Base completa: {len(df_completo)} linhas")
print(f"✅ Base Ipanema corrigida: {len(df_ipanema)} linhas")

# Normaliza nomes de colunas para evitar erros de comparação
df_completo.columns = df_completo.columns.str.strip()
df_ipanema.columns = df_ipanema.columns.str.strip()

# Garante que as colunas sejam compatíveis
colunas_comuns = [c for c in df_ipanema.columns if c in df_completo.columns]
df_ipanema = df_ipanema[colunas_comuns]
df_completo = df_completo[colunas_comuns]

# Filtra apenas as linhas com o nome da empresa de Ipanema
mask_ipanema = df_completo["Nome (Empresa)"].astype(str).str.contains(FILTRO_IPANEMA, case=False, na=False)
df_sem_ipanema = df_completo[~mask_ipanema].copy()

print(f"🗑️ Removendo {mask_ipanema.sum()} linhas antigas de Ipanema...")
print(f"➕ Inserindo {len(df_ipanema)} linhas corretas de Ipanema...")

# Junta tudo novamente
df_final = pd.concat([df_sem_ipanema, df_ipanema], ignore_index=True)

# Salva o resultado no próprio arquivo original (sobrescreve)
df_final.to_excel(ARQUIVO_COMPLETO, index=False)

print("\n✅ SUBSTITUIÇÃO DEFINITIVA CONCLUÍDA!")
print(f"📂 Arquivo sobrescrito: {ARQUIVO_COMPLETO}")
print(f"Total de linhas na nova base: {len(df_final)}")
print(f"🛡️ Backup mantido em: {backup_path}")
