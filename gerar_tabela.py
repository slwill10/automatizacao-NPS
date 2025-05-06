import pandas as pd
from nps_gestor import calcular_nps_gestor
from nps_aluno import calcular_nps_aluno
from nps_contratante import calcular_nps_contratante

df_gestor = pd.read_excel('planilha_de_gestor.xlsx')
df_contratante_and_aluno = pd.read_excel('tabela_alunos_e_contratante.xlsx')

def gerar_tabela_resumo_nps(writer):
    nps_dados = {}

    # GESTOR
    if not df_gestor.empty:
        nps_dados["gestor"] = {
            **calcular_nps_gestor(df_gestor),
            "total": len(df_gestor)
        }
    else:
        nps_dados["gestor"] = {
            "nps": 0.0,
            "promotores": 0,
            "detratores": 0,
            "neutros": 0,
            "total": 0
        }

    # ALUNO
    if not df_contratante_and_aluno.empty:
        nps_dados["aluno"] = {
            **calcular_nps_aluno(df_contratante_and_aluno),
            "total": len(df_contratante_and_aluno)
        }
    else:
        nps_dados["aluno"] = {
            "nps": 0.0,
            "promotores": 0,
            "detratores": 0,
            "neutros": 0,
            "total": 0
        }

    # CONTRATANTE
    if not df_contratante_and_aluno.empty:
        nps_dados["contratante"] = {
            **calcular_nps_contratante(df_contratante_and_aluno),
            "total": len(df_contratante_and_aluno)
        }
    else:
        nps_dados["contratante"] = {
            "nps": 0.0,
            "promotores": 0,
            "detratores": 0,
            "neutros": 0,
            "total": 0
        }

    resumo_df = pd.DataFrame([
        {
            'Categoria': 'NPS gestor',
            'Total de respondentes': nps_dados['gestor']['total'],
            '%': f"{nps_dados['gestor']['nps']:.1f}%" if nps_dados['gestor']['total'] > 0 else ""
        },
        {
            'Categoria': 'NPS aluno',
            'Total de respondentes': nps_dados['aluno']['total'],
            '%': f"{nps_dados['aluno']['nps']:.1f}%" if nps_dados['aluno']['total'] > 0 else ""
        },
        {
            'Categoria': 'NPS contratante_pós vendas',
            'Total de respondentes': nps_dados['contratante']['total'],
            '%': f"{nps_dados['contratante']['nps']:.1f}%" if nps_dados['contratante']['total'] > 0 else ""
        },
        {
            'Categoria': 'NPS ESR',
            'Total de respondentes': sum([nps_dados[c]['total'] for c in nps_dados]),
            '%': f"{sum([nps_dados[c]['nps'] * nps_dados[c]['total'] for c in nps_dados]) / max(1, sum([nps_dados[c]['total'] for c in nps_dados])):.1f}%"
            if sum([nps_dados[c]['total'] for c in nps_dados]) > 0 else ""
        }
    ])

    resumo_df.to_excel(writer, index=False, sheet_name="Resumo NPS")
    worksheet = writer.sheets["Resumo NPS"]
    workbook = writer.book

    header_format = workbook.add_format({'bold': True, 'bg_color': '#FBE4D5', 'border': 1})
    red_text = workbook.add_format({'font_color': 'red', 'border': 1})
    border_format = workbook.add_format({'border': 1})

    for col_num, value in enumerate(resumo_df.columns.values):
        worksheet.write(0, col_num, value, header_format)

    for row_idx, (_, row_data) in enumerate(resumo_df.iterrows()):
        total = row_data['Total de respondentes']
        for col_idx, value in enumerate(row_data):
            fmt = border_format
            if row_idx == 3 and col_idx == 2 and total == 0:
                fmt = red_text
            worksheet.write(row_idx + 1, col_idx, value, fmt)

    worksheet.set_column('A:A', 5)
    worksheet.set_column('B:B', 35)
    worksheet.set_column('C:D', 20)

    linha_inicial = len(resumo_df) + 4 
    dados_detalhados = [
        {
            'Perfil de público': 'NPS gestor',
            'Promotores': nps_dados['gestor']['promotores'],
            'Detratores': nps_dados['gestor']['detratores'],
            'Neutros': nps_dados['gestor']['neutros'],
        },
        {
            'Perfil de público': 'NPS aluno',
            'Promotores': nps_dados['aluno']['promotores'],
            'Detratores': nps_dados['aluno']['detratores'],
            'Neutros': nps_dados['aluno']['neutros'],
        },
        {
            'Perfil de público': 'NPS contratante',
            'Promotores': nps_dados['contratante']['promotores'],
            'Detratores': nps_dados['contratante']['detratores'],
            'Neutros': nps_dados['contratante']['neutros'],
        }
    ]

    total_promotores = sum(d['Promotores'] for d in dados_detalhados)
    total_detratores = sum(d['Detratores'] for d in dados_detalhados)
    total_neutros = sum(d['Neutros'] for d in dados_detalhados)
    total_geral = total_promotores + total_detratores + total_neutros

    dados_detalhados.append({
        'Perfil de público': 'Total',
        'Promotores': total_promotores,
        'Detratores': total_detratores,
        'Neutros': total_neutros
    })

    for i, d in enumerate(dados_detalhados[:-1]):
        d['% Promotores'] = f"{d['Promotores'] / total_geral:.0%}" if total_geral else "0%"
        d['% Detratores'] = f"{d['Detratores'] / total_geral:.0%}" if total_geral else "0%"
        d['% Neutros'] = f"{d['Neutros'] / total_geral:.0%}" if total_geral else "0%"
    dados_detalhados[-1]['% Promotores'] = ''
    dados_detalhados[-1]['% Detratores'] = ''
    dados_detalhados[-1]['% Neutros'] = ''

    tabela_detalhada_df = pd.DataFrame(dados_detalhados)

    for col_num, value in enumerate(tabela_detalhada_df.columns):
        worksheet.write(linha_inicial, col_num, value, header_format)

    for row_idx, (_, row) in enumerate(tabela_detalhada_df.iterrows()):
        for col_idx, value in enumerate(row):
            val = value
            fmt = border_format

            if isinstance(value, str) and "%" in value:
                if 'Promotores' in tabela_detalhada_df.columns[col_idx]:
                    fmt = workbook.add_format({'bg_color': '#92D050', 'border': 1})
                elif 'Detratores' in tabela_detalhada_df.columns[col_idx]:
                    fmt = workbook.add_format({'bg_color': '#FF0000', 'border': 1})
                elif 'Neutros' in tabela_detalhada_df.columns[col_idx]:
                    fmt = workbook.add_format({'bg_color': '#FFC000', 'border': 1})
            elif tabela_detalhada_df.columns[col_idx] in ['Promotores', 'Detratores', 'Neutros']:
                cor = {
                    'Promotores': '#92D050',
                    'Detratores': '#FF0000',
                    'Neutros': '#FFC000'
                }.get(tabela_detalhada_df.columns[col_idx])
                fmt = workbook.add_format({'bg_color': cor, 'border': 1})

            worksheet.write(linha_inicial + row_idx + 1, col_idx, val, fmt)

    worksheet.set_column('A:H', 15)
