import pandas as pd
import os

def ler_arquivo(arquivo_path):
    if isinstance(arquivo_path, str):
        extensao = arquivo_path.split('.')[-1].lower()
    else:
        extensao = arquivo_path.name.split('.')[-1].lower()

    if extensao == 'csv':
        return pd.read_csv(arquivo_path)
    elif extensao == 'xlsx':
        return pd.read_excel(arquivo_path, engine='openpyxl')
    elif extensao == 'xls':
        return pd.read_excel(arquivo_path)  
    else:
        raise ValueError(f"Formato de arquivo não suportado: {extensao}")    

def gerar_aba_contratante(writer, arquivo):
    df = ler_arquivo(arquivo)
    df = df[df['Qual a sua relação com a ESR? Selecione as opções aplicáveis.'] == 'Contratante'].copy()

    nota = 'Qual é a probabilidade de você recomendar a ESR a um(a) amigo(a) ou colega?'

    df = df[pd.to_numeric(df[nota], errors='coerce').notnull()].copy()
    df[nota] = df[nota].astype(int)

    df['nota_do_aluno'] = df[nota]
    total = len(df)

    contagem_notas = {i: 0 for i in range(11)}

    neutros = 0
    promotores = 0
    detratores = 0 

    for n in df['nota_do_aluno']:
        contagem_notas[int(n)] += 1
        if n >= 9:
            promotores += 1
        elif 7 <= n <= 8:
            neutros += 1
        elif n <= 6:
            detratores += 1

    total_respostas = len(df['nota_do_aluno'])

    porc_promotores = round((promotores / total) * 100, 2)
    porc_neutros = round((neutros / total) * 100, 2)
    porc_detratores = round((detratores / total) * 100, 2)
    nps = round(((promotores - detratores) / total_respostas) * 100, 2)

    df.to_excel(writer, index=False, sheet_name="NPS Contratante")
    workbook = writer.book
    worksheet = writer.sheets["NPS Contratante"]

    linha_base = df.shape[0] + 10

    bold = workbook.add_format({'bold': True, 'border': 1})

    formatos_cor = {
        'verde': workbook.add_format({'bg_color': '#63BE7B', 'border': 1, 'bold': True}),
        'amarelo': workbook.add_format({'bg_color': '#FBE983', 'border': 1, 'bold': True}),
        'vermelho': workbook.add_format({'bg_color': '#F8696B', 'border': 1, 'bold': True, 'font_color': 'white'}),
        'cinza': workbook.add_format({'bg_color': '#D9E1F2', 'border': 1, 'bold': True}),
        'preto': workbook.add_format({'border': 1}),
    }

    worksheet.write(linha_base, 1, "Digitar as quantidade de notas recebidas", bold)

    for i, nota in enumerate(range(10, -1, -1)):
        cor = (
            'verde' if nota >= 9 else
            'amarelo' if nota >= 7 else
            'vermelho'
        )
        linha = linha_base + 1 + i
        worksheet.write(linha, 0, f"Quantidade de NOTAS {nota}:", formatos_cor[cor])
        worksheet.write(linha, 1, contagem_notas[nota], formatos_cor['preto'])

    linha_resumo = linha_base + 13
    grupos = [
        ("Clientes Promotores:", promotores, f"{porc_promotores:.2f}%", 'verde'),
        ("Clientes Neutros:", neutros, f"{porc_neutros:.2f}%", 'amarelo'),
        ("Clientes Detratores:", detratores, f"{porc_detratores:.2f}%", 'vermelho'),
    ]

    for i, (titulo, qtd, perc, cor) in enumerate(grupos):
        linha = linha_resumo + i
        worksheet.write(linha, 0, titulo, formatos_cor[cor])
        worksheet.write(linha, 1, qtd, formatos_cor['preto'])
        worksheet.write(linha, 2, perc, formatos_cor['preto'])

    linha_nps = linha_resumo + 4
    worksheet.write(linha_nps, 0, "Seu NPS:", formatos_cor['cinza'])
    worksheet.write(linha_nps, 1, promotores - detratores, formatos_cor['cinza'])
    worksheet.write(linha_nps, 2, f"{nps:.1f}%", formatos_cor['cinza'])

    worksheet.set_column('A:A', 35)
    worksheet.set_column('B:C', 15)

def calcular_nps_contratante(df):
    df = df[df['Qual a sua relação com a ESR? Selecione as opções aplicáveis.'] == 'Contratante'].copy()
    nota = 'Qual é a probabilidade de você recomendar a ESR a um(a) amigo(a) ou colega?'
    df = df[pd.to_numeric(df[nota], errors='coerce').notnull()]
    df[nota] = df[nota].astype(int)
    
    promotores = df[df[nota] >= 9].shape[0]
    neutros = df[(df[nota] >= 7) & (df[nota] <= 8)].shape[0]
    detratores = df[df[nota] <= 6].shape[0]

    total = len(df)

    porc_promotores = round((promotores / total) * 100, 2) if total else 0
    porc_neutros = round((neutros / total) * 100, 2) if total else 0
    porc_detratores = round((detratores / total) * 100, 2) if total else 0
    nps = round(((promotores - detratores) / total) * 100, 2) if total else 0

    return {
        "promotores": promotores,
        "neutros": neutros,
        "detratores": detratores,
        "porc_promotores": porc_promotores,
        "porc_neutros": porc_neutros,
        "porc_detratores": porc_detratores,
        "nps": nps
    }
