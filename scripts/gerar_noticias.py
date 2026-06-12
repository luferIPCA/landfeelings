#!/usr/bin/env python3
"""
gerar_noticias.py — LandFeelings
Lê os ficheiros Markdown de _conteudos/noticias/ e gera data/noticias.json
para o site exibir na secção de Notícias.

Uso:
  python scripts/gerar_noticias.py

Ficheiro gerado:
  data/noticias.json
"""

import os, json, re
from datetime import datetime

# Raiz do projeto (dois níveis acima de scripts/)
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PASTA = os.path.join(BASE, '_conteudos', 'noticias')
DATA_DIR = os.path.join(BASE, 'data')
OUTPUT = os.path.join(DATA_DIR, 'noticias.json')

def parsear_frontmatter(texto):
    """Extrai o frontmatter YAML e o conteúdo do ficheiro Markdown."""
    padrao = re.match(r'^---\s*\n(.*?)\n---\s*\n(.*)', texto, re.DOTALL)
    if not padrao:
        return {}, texto.strip()
    
    fm_texto = padrao.group(1)
    conteudo = padrao.group(2).strip()
    
    fm = {}
    for linha in fm_texto.split('\n'):
        if ':' in linha:
            chave, _, valor = linha.partition(':')
            fm[chave.strip()] = valor.strip()
    
    return fm, conteudo

def formatar_data(data_str):
    """Converte a data para formato legível em português."""
    meses = ['janeiro','fevereiro','março','abril','maio','junho',
             'julho','agosto','setembro','outubro','novembro','dezembro']
    try:
        data_str = re.sub(r'([+-]\d{2}:\d{2}|Z)$', '', data_str.strip())
        dt = datetime.fromisoformat(data_str)
        return f"{dt.day} de {meses[dt.month-1]} de {dt.year}"
    except:
        return data_str

def main():
    if not os.path.exists(PASTA):
        print(f'Pasta não encontrada: {PASTA}')
        print('Cria primeiro uma notícia no painel Decap CMS.')
        return

    ficheiros = sorted(
        [f for f in os.listdir(PASTA) if f.endswith('.md')],
        reverse=True
    )

    if not ficheiros:
        print('Nenhum ficheiro Markdown encontrado em _conteudos/noticias/')
        return

    # Criar pasta data/ se não existir
    os.makedirs(DATA_DIR, exist_ok=True)

    noticias = []
    for nome in ficheiros:
        caminho = os.path.join(PASTA, nome)
        with open(caminho, 'r', encoding='utf-8') as f:
            texto = f.read()
        
        fm, conteudo = parsear_frontmatter(texto)
        
        noticia = {
            'id': nome.replace('.md', ''),
            'titulo': fm.get('title', 'Sem título'),
            'data': fm.get('date', ''),
            'data_formatada': formatar_data(fm.get('date', '')),
            'resumo': fm.get('resumo', ''),
            'imagem': fm.get('imagem', ''),
            'conteudo': conteudo,
        }
        noticias.append(noticia)
        print(f'  ✓ {noticia["titulo"]} ({noticia["data_formatada"]})')

    with open(OUTPUT, 'w', encoding='utf-8') as f:
        json.dump(noticias, f, ensure_ascii=False, indent=2)

    print(f'\n✓ {len(noticias)} notícia(s) exportada(s) para data/noticias.json')

if __name__ == '__main__':
    main()
