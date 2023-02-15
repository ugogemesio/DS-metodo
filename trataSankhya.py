import pandas as pd
import regex as re
df= pd.read_excel('SUB-OS.xls')
df.columns = df.iloc[1,:]
df = df.iloc[2:]
import numpy as np
import math 
df2 = df['Data Fechamento']

df = df[['Situação', 'Problema (Ordem de Serviço)','Solução',"Nome (Endereço)","ISDN","Núm. OS"]]

df = df.rename(columns={"Problema (Ordem de Serviço)": "Problema","Nome (Endereço)": "UNIDADE", "ISDN": "RAMAL"})


ramal = pd.Series(df['Problema']).str.extract('rama\w*\w* *:* *([\d\d\d\d,* *e*\-]+)', flags=re.IGNORECASE)

campus = pd.Series(df['Problema']).str.extract('camp\w+ *:* *([\w \-]+)', flags=re.IGNORECASE)

andar_sala = pd.Series(df['Problema']).str.extract('Andar *e* *s*a*l*a* *:* *([\w ,\-]+)', flags=re.IGNORECASE)

unidade = pd.Series(df['Problema']).str.extract('unidade *:* *([\w ,\-]+)', flags=re.IGNORECASE)

setor = pd.Series(df['Problema']).str.extract('setor *:* *([\w ,\-]+)', flags=re.IGNORECASE)
predio_bloco = pd.Series(df['Problema']).str.extract('pr\wdio */* *bloco *:* *([\w ,\-]+)', flags=re.IGNORECASE)

problema =df['Problema']
solucao = df['Solução']

columns = pd.Index(["RAMAL", "CAMPUS", "PRÉDIO/BLOCO","ANDAR E SALA", 'TIPO DE SERVIÇO SOLICITADO'])
index = pd.Index(list(range(312)))
tabela = pd.DataFrame(index=index)
tabela['RAMAL'] = ramal
tabela['CAMPUS'] = campus
tabela['ANDAR E SALA'] = andar_sala
tabela["PRÉDIO/BLOCO"] = predio_bloco
tabela['SETOR'] = setor
#tipo_servico = pd.Series(list(295))
i = 0

tabela = tabela.iloc[2:,]
tabela['RAMAL'][2] = '2386'
tabela['CAMPUS'][2] = 'Valonguinho'
tabela['ANDAR E SALA'][2] = ' 2º andar, sala 214'
tabela['PRÉDIO/BLOCO'][2] = '30'
solucao2=solucao.dropna()
problema2=problema.dropna()

i = 0
c=[]
for index, value in problema2.items():
    if  ' instal' in value:
        c.append('INSTALAÇÃO')
    else:
        c.append('MANUTENÇÃO, AGENDAMENTO OU OUTRA TRATATIVA')

x = pd.Series(c)
tabela['TIPO DE SERVIÇO SOLICITADO'] = c
df2.to_excel("datafecha.xlsx")
