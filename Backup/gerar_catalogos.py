#!/usr/bin/env python3
"""
gerar_catalogos.py — LandFeelings
Varre as pastas em images/Catalogos/ e gera um index.json em cada uma.
Atribui IDs únicos no formato LF-ANO-NNN a cada peça nova.

Uso:
  python gerar_catalogos.py

Executar sempre que adicionar novas imagens a uma pasta de catálogo.
Preserva metadados existentes e adiciona apenas entradas novas.
"""

import os
import json

EXTENSOES = {'.jpg', '.jpeg', '.png', '.webp', '.gif'}
BASE = os.path.join(os.path.dirname(__file__), 'images', 'Catalogos')

def gerar_id(ano, n):
    return f'LF-{ano}-{n:03d}'

def proximo_contador(existentes, ano):
    nums = []
    for item in existentes.values():
        pid = item.get('id', '')
        prefixo = f'LF-{ano}-'
        if pid.startswith(prefixo):
            try:
                nums.append(int(pid[len(prefixo):]))
            except ValueError:
                pass
    return max(nums, default=0) + 1

def processar_pasta(pasta_path, ano):
    index_path = os.path.join(pasta_path, 'index.json')
    existentes = {}
    if os.path.exists(index_path):
        try:
            with open(index_path, 'r', encoding='utf-8') as f:
                dados = json.load(f)
            existentes = {item['ficheiro']: item for item in dados}
        except Exception as e:
            print(f'  Aviso: erro ao ler {index_path}: {e}')

    imagens = sorted([
        f for f in os.listdir(pasta_path)
        if os.path.splitext(f)[1].lower() in EXTENSOES
    ])

    if not imagens:
        print(f'  {ano}: sem imagens')
        return

    contador = proximo_contador(existentes, ano)
    novas = 0
    resultado = []

    for img in imagens:
        if img in existentes:
            entrada = existentes[img]
            if not entrada.get('id'):
                entrada['id'] = gerar_id(ano, contador)
                contador += 1
            for campo in ['preco', 'altura', 'largura', 'peso']:
                if campo not in entrada:
                    entrada[campo] = ''
            if 'hashtags' not in entrada:
                entrada['hashtags'] = []
            resultado.append(entrada)
        else:
            resultado.append({
                'id':        gerar_id(ano, contador),
                'ficheiro':  img,
                'titulo':    '',
                'descricao': '',
                'autor':     '',
                'data':      '',
                'preco':     '',
                'altura':    '',
                'largura':   '',
                'peso':      '',
                'hashtags':  []
            })
            contador += 1
            novas += 1

    with open(index_path, 'w', encoding='utf-8') as f:
        json.dump(resultado, f, ensure_ascii=False, indent=2)

    print(f'  {ano}: {len(imagens)} imagens ({novas} novas) -> index.json atualizado')

def main():
    if not os.path.exists(BASE):
        print(f'Pasta nao encontrada: {BASE}')
        return

    print('LandFeelings -- Gerador de catalogos')
    print(f'Base: {BASE}\n')

    anos = sorted([d for d in os.listdir(BASE) if os.path.isdir(os.path.join(BASE, d))])
    if not anos:
        print('Nenhuma pasta de ano encontrada.')
        return

    for ano in anos:
        processar_pasta(os.path.join(BASE, ano), ano)

    print(f'\nConcluido! {len(anos)} pasta(s) processada(s).')
    print('Proximos passos:')
    print('  1. python gerar_excel.py   -> gera Excel para preenchimento')
    print('  2. Preencher o Excel com metadados')
    print('  3. python excel_para_json.py -> importa para os index.json')

if __name__ == '__main__':
    main()
