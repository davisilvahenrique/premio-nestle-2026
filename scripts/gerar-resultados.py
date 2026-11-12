"""Lê a planilha de votos e reescreve o bloco RESULTADOS dentro de resultados.html.

Uso: python scripts/gerar-resultados.py "C:\\caminho\\Prêmio Nestlé 2026.xlsx"

Só as contagens por opção vão pro HTML (sem nome de quem votou).
"""
import json
import re
import sys
import unicodedata
from collections import Counter
from pathlib import Path

import openpyxl

HTML = Path(__file__).resolve().parent.parent / 'resultados.html'

# título da coluna na planilha (aba "Votos") -> chave da categoria em resultados.html
COLUNAS = {
    'Premio Saka': 'premio_saka',
    'Prêmio CEO': 'premio_ceo',
    'Prêmio Piscineira': 'premio_piscineira',
    'Casal do Ano': 'casal_ano',
    'Golden Boy': 'golden_boy',
    'Fioti do Ano': 'fioti_ano',
    'Resenha do Ano': 'resenha_ano',
    'AcacIA': 'versao_acacio',
    'Story do Samuca': 'story_samuca',
    'AmigA Nestlé do Ano': 'amiga_ano',
    'Pior Amigo do Ano': 'pior_amigo',
    'Amigo Nestlé do Ano': 'amigo_ano',
    'Prêmio Neve': 'premio_neve',
}
# valor na planilha -> nome exibido (opção renomeada no site depois que o voto foi salvo)
RENOMEAR = {'Niggacio': 'Aggacio'}

nfc = lambda s: unicodedata.normalize('NFC', str(s)).strip()


def main(xlsx):
    ws = openpyxl.load_workbook(xlsx, data_only=True)['Votos']
    cab = [nfc(c.value) if c.value is not None else '' for c in ws[1]]
    colunas = {nfc(k): v for k, v in COLUNAS.items()}
    faltando = [k for k in colunas if k not in cab]
    if faltando:
        sys.exit('Colunas não encontradas na planilha: ' + ', '.join(faltando))
    linhas = [r for r in ws.iter_rows(min_row=2, values_only=True) if r[0]]
    categorias = {}
    for titulo, chave in colunas.items():
        idx = cab.index(titulo)
        c = Counter(RENOMEAR.get(nfc(r[idx]), nfc(r[idx])) for r in linhas if r[idx])
        categorias[chave] = dict(c.most_common())
    dados = {'votantes': len(linhas), 'categorias': categorias}
    bloco = 'const RESULTADOS = ' + json.dumps(dados, ensure_ascii=False, indent=2).replace('\n', '\n    ') + ';'
    html = HTML.read_text(encoding='utf-8')
    novo, n = re.subn(r'(// <RESULTADOS>\n    ).*?(\n    // </RESULTADOS>)', lambda m: m.group(1) + bloco + m.group(2), html, flags=re.S)
    if n != 1:
        sys.exit('Marcadores // <RESULTADOS> não encontrados em resultados.html')
    HTML.write_text(novo, encoding='utf-8', newline='')
    print(f'{len(linhas)} votantes, {len(categorias)} categorias gravadas em {HTML.name}')


if __name__ == '__main__':
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    main(sys.argv[1])
