import pandas as pd

# Planilha RFB
plan_enriched = pd.ExcelFile(r'.\input.xlsx')
plan_enriched_data = pd.read_excel(plan_enriched, 'RFB')
plan_enriched_data = pd.DataFrame(plan_enriched_data, columns= ['CNPJ', 'Sócios', 'Faturamento'])

# Lista com informaçoes dos cnpjs
cnpjs = []

# Lista de resultados
matches_res = []

# Armazenando informações da RFB
for rfb in plan_enriched_data.itertuples(index=False):
        socios = str(rfb[1]).split(',')
        cnpj_total = {
            "cnpj": rfb[0],
            "socios": socios if socios != ['nan'] else [],
            "faturamento": rfb[2]
        }
        cnpjs.append(cnpj_total)

# Socios em comum (lista já deve estar ordenada)
for cnpj in cnpjs:
    res = {
        "CNPJ": str(cnpj['cnpj']),
        "Dono": None,
        "Faturamento Dono": None,
        "Empresas com Sócios em Comum": [],
        "Sócios": []
    }
    for other_cnpj in cnpjs:
        for socio in cnpj['socios']:
            if socio in other_cnpj['socios']:
                res['Sócios'].append(socio)
                if other_cnpj['cnpj'] not in res['Empresas com Sócios em Comum']:
                    res['Empresas com Sócios em Comum'].append(other_cnpj['cnpj'])
                if res['Dono'] == None:
                    res['Dono'] = other_cnpj['cnpj']
                if res['Faturamento Dono'] == None:
                    res['Faturamento Dono'] = other_cnpj['faturamento']
    res['Sócios'] = list(set(res['Sócios']))
    res['Sócios'] = sorted(res['Sócios'])
    res['Sócios'] = ', '.join(str(x) for x in res['Sócios'])
    res['Empresas com Sócios em Comum'] = ', '.join(str(x) for x in res['Empresas com Sócios em Comum'])
    matches_res.append(res)

# Criar planilha final
plan_matches = pd.DataFrame(matches_res)
with pd.ExcelWriter("output.xlsx") as w:
    plan_matches.to_excel(w, index=False)