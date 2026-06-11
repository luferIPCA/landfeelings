#!/usr/bin/env python3
"""
excel_para_json.py — LandFeelings
Importa o Excel de metadados e actualiza os index.json correspondentes.

Uso:
  python excel_para_json.py

Lê:  catalogo_metadados.xlsx
Actualiza: images/Catalogos/ANO/index.json (usando ID como chave)
"""

import os
import json

try:
    import openpyxl
except ImportError:
    import subprocess
    subprocess.run(["pip", "install", "openpyxl", "--break-system-packages", "-q"])
    import openpyxl

BASE   = os.path.join(os.path.dirname(__file__), 'images', 'Catalogos')
EXCEL  = os.path.join(os.path.dirname(__file__), 'catalogo_metadados.xlsx')

# Mapeamento colunas Excel → campos JSON (mesma ordem do gerar_excel.py)
CAMPOS = ['id', 'ficheiro', 'ano', 'titulo', 'descricao', 'autor',
          'preco', 'altura', 'largura', 'peso', 'hashtags']

def ler_excel():
    """Lê o Excel e devolve dicionário {id: dados}."""
    if not os.path.exists(EXCEL):
        print(f'Ficheiro não encontrado: {EXCEL}')
        print('Corre primeiro: python gerar_excel.py')
        return {}

    wb   = openpyxl.load_workbook(EXCEL)
    ws   = wb.active
    dados = {}

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:  # ID vazio → ignorar linha
            continue

        peca = {}
        for i, campo in enumerate(CAMPOS):
            val = row[i] if i < len(row) else ''
            peca[campo] = str(val).strip() if val is not None else ''

        # Converter hashtags de string para lista
        hashtags_str = peca.get('hashtags', '')
        if hashtags_str:
            peca['hashtags'] = [h.strip() for h in hashtags_str.split(',') if h.strip()]
        else:
            peca['hashtags'] = []

        dados[peca['id']] = peca

    print(f'Excel lido: {len(dados)} peças')
    return dados

def actualizar_json(dados_excel):
    """Actualiza os index.json com os metadados do Excel."""
    if not os.path.exists(BASE):
        print(f'Pasta não encontrada: {BASE}')
        return

    anos = sorted([d for d in os.listdir(BASE) if os.path.isdir(os.path.join(BASE, d))])
    total_actualizadas = 0

    for ano in anos:
        index_path = os.path.join(BASE, ano, 'index.json')
        if not os.path.exists(index_path):
            print(f'  {ano}: sem index.json — ignorado')
            continue

        try:
            with open(index_path, 'r', encoding='utf-8') as f:
                pecas = json.load(f)
        except Exception as e:
            print(f'  {ano}: erro ao ler JSON — {e}')
            continue

        actualizadas = 0
        for peca in pecas:
            pid = peca.get('id', '')
            if pid and pid in dados_excel:
                dados = dados_excel[pid]
                # Actualizar campos (manter ficheiro e id originais)
                for campo in ['titulo', 'descricao', 'autor', 'data', 'preco',
                               'altura', 'largura', 'peso', 'hashtags']:
                    peca[campo] = dados.get(campo, peca.get(campo, ''))
                actualizadas += 1

        with open(index_path, 'w', encoding='utf-8') as f:
            json.dump(pecas, f, ensure_ascii=False, indent=2)

        total_actualizadas += actualizadas
        print(f'  {ano}: {actualizadas}/{len(pecas)} peças actualizadas → index.json guardado')

    print(f'\nConcluído! {total_actualizadas} peças actualizadas no total.')

def main():
    print('LandFeelings — Importador Excel → JSON')
    print(f'Excel: {EXCEL}')
    print(f'Base:  {BASE}\n')

    dados = ler_excel()
    if not dados:
        return

    actualizar_json(dados)
    print('\nRecarrega o browser para ver as alterações.')

if __name__ == '__main__':
    main()
