import re
import pandas as pd

# Ler os arquivos Excel
site_list = pd.read_excel('SiteList.xlsx')
results = pd.read_excel('Results.xlsx')

# Filtrar os dados para o ano de 2023
site_list_2023 = site_list[site_list['Year'] == 2023]
results_2023 = results[results['Year'] == 2023]

# Função para extrair o Estado do Sitelist
def extrair_estado(texto):
    padrao_estado = r'([A-Z]{2})$'
    resultado = re.search(padrao_estado, texto)
    if resultado:
        return resultado.group(1)
    else:
        return None

site_list_2023 = site_list_2023.copy()
site_list_2023.loc[:, 'State'] = site_list_2023['Site Name'].apply(extrair_estado)

# Selecionar as colunas desejadas
site_data = site_list_2023[['Site Name', 'State']]
results_data = results_2023[['Site ID', 'Equipment', 'Signal (%)', 'Quality (0-10)', 'Mbps']]

# Salvar relatorio
results_data.loc[:, 'Site ID'] = results_data['Site ID'].astype(str)

# Adicionar as colunas 'Site' e 'State' ao DataFrame results_data
results_data = pd.merge(results_data, site_data, how='left', left_on='Site ID', right_on='Site Name')


results_data.drop(columns=['Site Name'], inplace=True)

site_data = site_data.sort_values(by=['State', 'Site Name'], ascending=[True, True])
results_data = results_data.sort_values(by=['State', 'Site ID'], ascending=[True, True])


sites_not_in_results = site_data[~site_data['Site Name'].isin(results_data['Site ID'])]
missing_results = sites_not_in_results[sites_not_in_results['State'].notna()]

with pd.ExcelWriter('relatorio.xlsx', engine='xlsxwriter') as writer:
    results_data.to_excel(writer, sheet_name='Relatorio', index=False)
    missing_results[['Site Name', 'State']].to_excel(writer, sheet_name='Sites não presentes nos Results', index=False)

# Printar as informações no console
print("Sites com 0 de Qualidade:")
print(results_data[results_data['Quality (0-10)'] == 0][['Site ID', 'Equipment']])

print("\nSites com mais de 80 Mbps:")
print(results_data[results_data['Mbps'] > 80][['Site ID', 'Equipment']])

print("\nSites que não estão presentes no Results:")
print(missing_results[['Site Name', 'State']])
