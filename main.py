import pandas as pd
import openpyxl

siteList = pd.read_excel('SiteList.xlsx')
results = pd.read_excel('Results.xlsx')


# Encontrando resultados dos sites presentes na planilha SiteList

#sites_id_list = siteList['Site ID']
#sites_id_results = results['Site ID']

sites_results = siteList.merge(results, how='inner', on=['Site ID', 'Site Name', 'Year', 'Equipment'])

sites_orded = sites_results.sort_values('State')

print(sites_orded)

sites_orded.to_excel('quality-report-2023.xlsx', sheet_name='Resultados sites')

sites_Alert_num = sites_orded[sites_orded['Alerts'].str.contains('Yes')].shape[0]

print(f"Numero de sites com Alerta ativado: {sites_Alert_num}")  # Numero de sites com alertas ativos

media_qualidade = sites_orded['Quality (0-10)'].mean()

print(f"Media da qualidade dos sites: {media_qualidade}")

sites_Mbps = sites_orded[sites_results['Mbps'] < 10].shape[0]

print(f"Numero de sites com Mbps menor que 10: {sites_Mbps}")


