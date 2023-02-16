import pandas as pd
import networkx as nx
from itertools import chain, pairwise

# Planilha RFB
plan_enriched = pd.ExcelFile(r'.\RWS - Parte 1.xlsx')
plan_enriched_data = pd.read_excel(plan_enriched, 'Sheet1')
plan_enriched_data = pd.DataFrame(plan_enriched_data, columns= ['Empresas com Sócios em Comum'])

# Lista com informaçoes dos cnpjs
grupos = []

# Lista de resultados
matches_res = []

# Armazenando informações da RFB
for rfb in plan_enriched_data.itertuples(index=False):
    grupos.append(rfb[0])

# Lista ordenada e com 14 zeros
for i in range(len(grupos)):
    grupos[i] = grupos[i].split(", ")
    for j in range(len(grupos[i])):
        grupos[i][j] = grupos[i][j].zfill(14)
    p = "p" + str(grupos[i][0])
    grupos[i].append(p)

G = nx.Graph()
G = nx.from_edgelist(chain.from_iterable(pairwise(e) for e in grupos))
G.add_nodes_from(set.union(*map(set, grupos)))
grupos = list(nx.connected_components(G))

for i in range(len(grupos)):
    grupos[i] = list(grupos[i])

for g in grupos:
    res = {
        "Possíveis Donos": [],
        "Grupo": ''
    }
    for cnpj in g:
        if str(cnpj[0]) == 'p':
            res['Possíveis Donos'].append(cnpj[1:])
    g_clean = (x for x in g if x[0] != 'p')
    res['Grupo'] = list(set(g_clean))
    res['Grupo'] = ', '.join(str(x) for x in res['Grupo'])
    res['Possíveis Donos'] = ', '.join(str(x) for x in res['Possíveis Donos'])
    matches_res.append(res)

# Criar planilha final
plan_matches = pd.DataFrame(matches_res)
with pd.ExcelWriter("RWS - Parte 2.xlsx") as w:
    plan_matches.to_excel(w, index=False)