import pandas as pd
import json
import re

# Ler dados principais da nova planilha
df = pd.read_excel('/home/ubuntu/upload/RETENÇÃODEMAISCURSOSEEAD26.1(1).xlsx', sheet_name='DADOS', header=3)
df.columns = ['NOME', 'MATRICULA', 'CONTATO', 'CAMPUS', 'CURSO', 'DATA_SOLICITACAO', 'DATA_CONTATO', 'RESULTADO', 'MOTIVO', 'INSIGHT', 'SITUACAO']

# Remover linhas sem nome ou sem resultado válido
df = df.dropna(subset=['NOME'])
df = df[df['RESULTADO'].notna()]

# Preencher NaN em campos categóricos
df['MOTIVO'] = df['MOTIVO'].fillna('NÃO INFORMADO')
df['SITUACAO'] = df['SITUACAO'].fillna('NÃO INFORMADO')
df['CAMPUS'] = df['CAMPUS'].fillna('NÃO INFORMADO').str.strip()

# Normalizar datas
df['DATA_SOLICITACAO'] = pd.to_datetime(df['DATA_SOLICITACAO'], errors='coerce')
df['DATA_CONTATO'] = pd.to_datetime(df['DATA_CONTATO'], errors='coerce')

# Normalizar nomes de cursos
def normalizar_curso(curso):
    if pd.isna(curso):
        return 'NÃO INFORMADO'
    curso = str(curso)
    curso = re.sub(r'\s*\(\d+\)', '', curso)
    curso = re.sub(r'\s*EAD\d+', '', curso)
    curso = curso.strip().upper()
    mapa = {
        'ADMINISTRAÇÃO ()': 'ADMINISTRAÇÃO',
        'ADMINISTRACAO': 'ADMINISTRAÇÃO',
        'MEDICINA VETERINÁRIA': 'MEDICINA VETERINÁRIA',
        'MEDICINA VETERINARIA': 'MEDICINA VETERINÁRIA',
        'EDUCAÇÃO FISÍCA': 'EDUCAÇÃO FÍSICA',
        'EDUCAÇÃO FÍSICA': 'EDUCAÇÃO FÍSICA',
        'CIÊNCIAS CONTABÉIS': 'CIÊNCIAS CONTÁBEIS',
        'INTELIGÊNCIA ARTIFICIAL ()': 'INTELIGÊNCIA ARTIFICIAL',
        'ANÁLISE E DESENVOLVIMENTO SISTEMAS EAD': 'ADS EAD',
        'MARKETING ()': 'MARKETING EAD',
    }
    return mapa.get(curso, curso)

df['CURSO_NORM'] = df['CURSO'].apply(normalizar_curso)

# Modalidade
df['MODALIDADE'] = df['CURSO'].apply(lambda x: 'EAD' if 'EAD' in str(x).upper() else 'Presencial')

# Mês de solicitação
df['MES_ANO'] = df['DATA_SOLICITACAO'].dt.strftime('%Y-%m')
df['MES_LABEL'] = df['DATA_SOLICITACAO'].dt.strftime('%b/%Y')

# ---- KPIs gerais ----
total = len(df)
retidos = int((df['RESULTADO'] == 'RETIDO').sum())
desistencias = int((df['RESULTADO'] == 'DESISTÊNCIA').sum())
taxa_retencao = round(retidos / total * 100, 1)
taxa_desistencia = round(desistencias / total * 100, 1)
em_processo = int((df['SITUACAO'].str.strip() == 'EM PROCESSO').sum())
finalizados = int((df['SITUACAO'].str.strip() == 'FINALIZADO').sum())

kpis = {
    'total': total,
    'retidos': retidos,
    'desistencias': desistencias,
    'taxa_retencao': taxa_retencao,
    'taxa_desistencia': taxa_desistencia,
    'em_processo': em_processo,
    'finalizados': finalizados
}

# ---- Resultado por campus ----
resultado_campus = pd.crosstab(df['CAMPUS'], df['RESULTADO']).reset_index()
resultado_campus.columns = [str(c) for c in resultado_campus.columns]
resultado_campus_list = resultado_campus.to_dict('records')

# ---- Motivos ----
motivos = df['MOTIVO'].value_counts().reset_index()
motivos.columns = ['motivo', 'total']
motivos_list = motivos.to_dict('records')

# ---- Resultado por motivo ----
motivo_resultado = pd.crosstab(df['MOTIVO'], df['RESULTADO']).reset_index()
motivo_resultado.columns = [str(c) for c in motivo_resultado.columns]
motivo_resultado_list = motivo_resultado.to_dict('records')

# ---- Cursos com mais desistências ----
cursos_desist = df[df['RESULTADO'] == 'DESISTÊNCIA']['CURSO_NORM'].value_counts().head(12).reset_index()
cursos_desist.columns = ['curso', 'desistencias']
cursos_desist_list = cursos_desist.to_dict('records')

# ---- Solicitações por mês ----
mes_data = df.groupby(['MES_ANO', 'MES_LABEL', 'RESULTADO']).size().reset_index(name='count')
mes_pivot = mes_data.pivot_table(index=['MES_ANO', 'MES_LABEL'], columns='RESULTADO', values='count', fill_value=0).reset_index()
mes_pivot.columns = [str(c) for c in mes_pivot.columns]
mes_list = mes_pivot.sort_values('MES_ANO').to_dict('records')

# ---- Modalidade ----
modalidade = df.groupby(['MODALIDADE', 'RESULTADO']).size().reset_index(name='count')
modalidade_list = modalidade.to_dict('records')

# ---- Resultado por modalidade ----
mod_resultado = pd.crosstab(df['MODALIDADE'], df['RESULTADO']).reset_index()
mod_resultado.columns = [str(c) for c in mod_resultado.columns]
mod_resultado_list = mod_resultado.to_dict('records')

# ---- Situação das solicitações ----
situacao = df['SITUACAO'].str.strip().value_counts().reset_index()
situacao.columns = ['situacao', 'total']
situacao_list = situacao.to_dict('records')

# ---- Campus distribuição ----
campus_dist = df['CAMPUS'].value_counts().reset_index()
campus_dist.columns = ['campus', 'total']
campus_dist_list = campus_dist.to_dict('records')

# ---- Tabela de dados completa (sem contato) ----
df_table = df[['NOME', 'CAMPUS', 'CURSO_NORM', 'RESULTADO', 'MOTIVO', 'SITUACAO', 'DATA_SOLICITACAO', 'MODALIDADE']].copy()
df_table['DATA_SOLICITACAO'] = df_table['DATA_SOLICITACAO'].dt.strftime('%d/%m/%Y')
df_table['SITUACAO'] = df_table['SITUACAO'].str.strip()
df_table.columns = ['Nome', 'Campus', 'Curso', 'Resultado', 'Motivo', 'Situação', 'Data Solicitação', 'Modalidade']
tabela_list = df_table.to_dict('records')

# Compilar tudo
data_json = {
    'kpis': kpis,
    'resultado_campus': resultado_campus_list,
    'motivos': motivos_list,
    'motivo_resultado': motivo_resultado_list,
    'cursos_desistencias': cursos_desist_list,
    'solicitacoes_mes': mes_list,
    'modalidade': modalidade_list,
    'modalidade_resultado': mod_resultado_list,
    'situacao': situacao_list,
    'campus_distribuicao': campus_dist_list,
    'tabela': tabela_list
}

with open('/home/ubuntu/dashboard/data.json', 'w', encoding='utf-8') as f:
    json.dump(data_json, f, ensure_ascii=False, indent=2)

print('Dados atualizados com sucesso!')
print(f'Total registros: {total}')
print(f'Taxa de retenção: {taxa_retencao}%')
print(f'Taxa de desistência: {taxa_desistencia}%')
print(f'Em processo: {em_processo}')
print(f'Finalizados: {finalizados}')
