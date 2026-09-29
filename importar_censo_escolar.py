#!/usr/bin/env python3
"""Extrai da Sinopse Estatística da Educação Básica 2025 (INEP) os dados de educação especial por município.

Tabelas usadas (linhas de município, identificadas pelo código IBGE de 7 dígitos na 4ª coluna):
  1.56  matrículas da educação especial por dependência administrativa
  1.58  matrículas da educação especial por tipo de deficiência / TEA / altas habilidades (todas as redes)
  3.45  estabelecimentos com educação especial por dependência administrativa

Saída: dados/censo_escolar_2025.json
  {cod: [tot, federal, estadual, municipal, privada, tea, intelectual, fisica, auditiva, visual, multipla, ahsd, esc_estadual, esc_municipal, esc_total]}
"""
import json, os, sys
import openpyxl

BASE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(BASE, 'fontes_brutas/sinopse_2025.xlsx')
OUT = os.path.join(BASE, 'dados/censo_escolar_2025.json')


def linhas_municipio(ws):
    for r in ws.iter_rows(values_only=True):
        cod = r[3] if len(r) > 3 else None
        if cod is not None and str(cod).strip().isdigit() and len(str(cod).strip()) == 7:
            yield str(cod).strip(), r


def n(v):
    try: return int(v)
    except (TypeError, ValueError): return 0


def main():
    wb = openpyxl.load_workbook(SRC, read_only=True)
    d = {}
    for cod, r in linhas_municipio(wb['1.56']):          # Total | Pública | Federal | Estadual | Municipal | Privada
        d[cod] = [n(r[4]), n(r[6]), n(r[7]), n(r[8]), n(r[9])] + [0]*10
    for cod, r in linhas_municipio(wb['1.58']):          # Total | Cegueira | Baixa Visão | Surdez | Def. Auditiva | Surdocegueira | Física | Intelectual | Múltipla | TEA | AH/SD | Visão Monocular
        x = d.setdefault(cod, [0]*15)
        x[5] = n(r[13]); x[6] = n(r[11]); x[7] = n(r[10]); x[8] = n(r[7]) + n(r[8]) + n(r[9])
        x[9] = n(r[5]) + n(r[6]) + n(r[15]); x[10] = n(r[12]); x[11] = n(r[14])
    for cod, r in linhas_municipio(wb['3.45']):          # Total | Pública | Federal | Estadual | Municipal | Privada
        x = d.setdefault(cod, [0]*15)
        x[12] = n(r[7]); x[13] = n(r[8]); x[14] = n(r[4])
    json.dump(d, open(OUT, 'w'), separators=(',', ':'))
    tot = [sum(v[i] for v in d.values()) for i in range(15)]
    print(f'{len(d)} municípios · educação especial {tot[0]:,} (estadual {tot[2]:,}, municipal {tot[3]:,}, privada {tot[4]:,}) · TEA {tot[5]:,} · escolas c/ ed. especial {tot[14]:,} (municipais {tot[13]:,})')


if __name__ == '__main__':
    sys.exit(main())
