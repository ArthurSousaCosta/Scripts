import pandas as pd

# Planilha RFB
plan_enriched = pd.ExcelFile(r'.\input.xlsx')
plan_enriched_data = pd.read_excel(plan_enriched, 'RFB')
plan_enriched_data = pd.DataFrame(plan_enriched_data, columns= ['CNPJ', 'Raiz CNPJ','Nome Fantasia', 'Sócios'])

# Lista com informaçoes dos cnpjs
cnpjs = []

# Lista de resultados
matches_res = []

# Armazenando informações da RFB
for rfb in plan_enriched_data.itertuples(index=False):
        socios = str(rfb[3]).split(',')
        cnpj_total = {
            "cnpj": rfb[0],
            "raiz_cnpj": rfb[1],
            "nome_fantasia": rfb[2] if rfb[2] != 'nan' else '',
            "socios": socios if socios != ['nan'] else []
        }
        cnpjs.append(cnpj_total)

# Raiz, sócio, nome fantasia em comum
for cnpj in cnpjs:
    res = {
        "Grupo": [],
        "Raiz CNPJ": [],
        "Nome Fantasia": [],
        "Sócios": []      
    }
    for other_cnpj in cnpjs:
        if cnpj['raiz_cnpj'] == other_cnpj['raiz_cnpj']:
            res['Grupo'].append(other_cnpj['cnpj'])
            res['Raiz CNPJ'].append(other_cnpj['raiz_cnpj'])
            continue
        if cnpj['nome_fantasia'] == other_cnpj['nome_fantasia'] and cnpj['nome_fantasia'] != '':
            res['Grupo'].append(other_cnpj['cnpj'])
            res['Nome Fantasia'].append(other_cnpj['nome_fantasia'])
            continue
        for socio in cnpj['socios']:
            if socio in other_cnpj['socios']:
                res['Sócios'].append(socio)
                if other_cnpj['cnpj'] not in res['Grupo']:
                    res['Grupo'].append(other_cnpj['cnpj'])
    res['Grupo'] = ', '.join(str(x) for x in res['Grupo'])
    res['Sócios'] = list(set(res['Sócios']))
    res['Sócios'] = sorted(res['Sócios'])
    res['Sócios'] = ', '.join(str(x) for x in res['Sócios'])
    res['Nome Fantasia'] = list(set(res['Nome Fantasia']))
    res['Nome Fantasia'] = sorted(res['Nome Fantasia'])
    res['Nome Fantasia'] = ', '.join(str(x) for x in res['Nome Fantasia'])
    res['Raiz CNPJ'] = list(set(res['Raiz CNPJ']))
    res['Raiz CNPJ'] = sorted(res['Raiz CNPJ'])
    res['Raiz CNPJ'] = ', '.join(str(x) for x in res['Raiz CNPJ'])
    matches_res.append(res)

# Criar planilha final
plan_matches = pd.DataFrame(matches_res)
with pd.ExcelWriter("output.xlsx") as w:
    plan_matches.to_excel(w, index=False)