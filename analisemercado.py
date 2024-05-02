import pandas as pd
import matplotlib.pyplot as plt


dados_geral = pd.read_excel(r'C:\Users\ligui\Downloads\desafio_economapas.xlsx', sheet_name=None)

df_base = dados_geral['BASE']
df_local = dados_geral['LOCAL']
df_mercado = dados_geral['MERCADO']
df_objetivo = dados_geral['OBJETIVO']
df_origem = dados_geral['ORIGEM']
df_porte = dados_geral['PORTE']

all_data = pd.DataFrame({
    'LEAD_ID': pd.concat([df_base['LEAD_ID'], df_local['LEAD_ID'], df_mercado['LEAD_ID'], df_objetivo['LEAD_ID'], df_origem['LEAD_ID'], df_porte['LEAD_ID']], axis=0, ignore_index=True),
    'DATA CADASTRO': df_base['DATA CADASTRO'],
    'VENDIDO': df_base['VENDIDO'],
    'MERCADO': pd.concat([df_base['MERCADO'], df_mercado['MERCADO']], ignore_index=True),
    'ORIGEM': df_origem['ORIGEM'],
    'SUB-ORIGEM': df_origem['SUB-ORIGEM'],
    'MERCADO2': pd.concat([df_base['MERCADO'], df_mercado['MERCADO']], ignore_index=True),  
    'LOCAL': df_local['LOCAL'],
    'UF': df_local['LOCAL'].str[-2:], # UF
    'PORTE': df_porte['PORTE'],
    'OBJETIVO': df_objetivo['OBJETIVO']
})


nome_do_arquivo_excel = 'dados_processados.xlsx'
all_data.to_excel(nome_do_arquivo_excel, index=False)

print("Exportação para Excel concluída com sucesso!")

all_data_filtrado = all_data[all_data['VENDIDO'] == 'SIM'] 
dados_processados = all_data_filtrado


contagem_por_uf = dados_processados['UF'].value_counts()
contagem_por_uf = dados_processados['UF'].value_counts().sort_values(ascending=False)

plt.figure(figsize=(10, 6))
contagem_por_uf.plot(kind='bar', color='skyblue')
plt.title('Contagem de IDs por UF')
plt.xlabel('UF')
plt.ylabel('Contagem de IDs')
plt.xticks(rotation=45) 
plt.grid(axis='y', linestyle='--', alpha=0.7) 
plt.tight_layout() 
plt.show()

contagem_por_origem = dados_processados['ORIGEM'].value_counts()
contagem_por_origem = contagem_por_origem.sort_values(ascending=False)

plt.figure(figsize=(10, 6))
contagem_por_origem.plot(kind='bar', color='skyblue')
plt.title('Contagem de IDs por Origem')
plt.xlabel('Origem')
plt.ylabel('Contagem de IDs')
plt.xticks(rotation=45)  
plt.grid(axis='y', linestyle='--', alpha=0.7)  
plt.tight_layout()  
plt.show()

contagem_por_mercado = dados_processados['MERCADO'].value_counts()
contagem_por_mercado = contagem_por_mercado.sort_values(ascending=False)

plt.figure(figsize=(10, 6))
contagem_por_mercado.plot(kind='bar', color='skyblue')
plt.title('Contagem de IDs por Mercado')
plt.xlabel('Mercado')
plt.ylabel('Contagem de IDs')
plt.xticks(rotation=45)  
plt.grid(axis='y', linestyle='--', alpha=0.7)  
plt.tight_layout() 
plt.show()

contagem_por_porte = dados_processados['PORTE'].value_counts()

plt.figure(figsize=(10, 6))
contagem_por_porte.plot(kind='bar', color='skyblue')
plt.title('Contagem de LEAD_ID por Porte da Empresa')
plt.xlabel('Porte da Empresa')
plt.ylabel('Contagem de LEAD_ID')
plt.xticks(rotation=45) 
plt.grid(axis='y', linestyle='--', alpha=0.7)
plt.tight_layout()
plt.show()



