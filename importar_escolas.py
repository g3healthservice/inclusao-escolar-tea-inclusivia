#!/usr/bin/env python3
"""Mapa de escolas — lê os microdados do Censo Escolar (INEP) e gera, para as redes pública estadual e municipal:

  dados/escolas_resumo.json   {cod_mun: [mEsc, mComEsp, mSala, mComEspSemSala, eEsc, eComEsp, eSala, eComEspSemSala]}  (embutido no painel)
  escolas/<UF>.json           {cod_mun: {"m": [[cod_inep, nome, matriculas, mat_ed_especial, sala_recursos], ...], "e": [...]}}  (carregado sob demanda)

Uso: python3 importar_escolas.py fontes_brutas/microdados_2024.zip
Considera só escolas em atividade (TP_SITUACAO_FUNCIONAMENTO = 1). Rede: 2 = estadual (inclui a rede do DF), 3 = municipal.
"""
import csv, io, json, os, re, sys, zipfile

BASE = os.path.dirname(os.path.abspath(__file__))
PART = {'de', 'da', 'do', 'das', 'dos', 'e', 'em', 'a', 'o', 'no', 'na'}
SIGLAS = {'EM', 'EMEF', 'EMEI', 'EE', 'EMPGES', 'CEI', 'CEU', 'CMEI', 'EMEB', 'EEEFM', 'EEEF', 'EEEM', 'CIEP', 'CE', 'CEF', 'CED', 'CEM', 'EC', 'JI', 'CAIC', 'CEEBJA', 'APAE', 'II', 'III', 'IV', 'VI', 'VII'}
UF_COD = {'11':'RO','12':'AC','13':'AM','14':'RR','15':'PA','16':'AP','17':'TO','21':'MA','22':'PI','23':'CE','24':'RN','25':'PB','26':'PE','27':'AL',
          '28':'SE','29':'BA','31':'MG','32':'ES','33':'RJ','35':'SP','41':'PR','42':'SC','43':'RS','50':'MS','51':'MT','52':'GO','53':'DF'}


def nome_bonito(s):
    out = []
    for i, w in enumerate(re.split(r'(\s+)', s.strip())):
        if not w.strip(): out.append(w); continue
        u = w.upper()
        if u in SIGLAS: out.append(u)
        elif i > 0 and w.lower() in PART: out.append(w.lower())
        else: out.append(w[:1].upper() + w[1:].lower())
    return ''.join(out)


def n(v):
    try: return int(v)
    except (TypeError, ValueError): return 0


def main():
    src = sys.argv[1] if len(sys.argv) > 1 else os.path.join(BASE, 'fontes_brutas/microdados_2024.zip')
    z = zipfile.ZipFile(src)
    f = [x for x in z.namelist() if re.search(r'microdados_ed_basica_\d{4}\.csv$', x)][0]
    ano = re.search(r'(\d{4})\.csv$', f).group(1)
    por_uf, resumo = {}, {}
    with z.open(f) as fh:
        for r in csv.DictReader(io.TextIOWrapper(fh, encoding='latin-1'), delimiter=';'):
            if r['TP_SITUACAO_FUNCIONAMENTO'] != '1' or r['TP_DEPENDENCIA'] not in ('2', '3'): continue
            mun, rede = r['CO_MUNICIPIO'], ('m' if r['TP_DEPENDENCIA'] == '3' else 'e')
            bas, esp, sala = n(r.get('QT_MAT_BAS')), n(r['QT_MAT_ESP']), 1 if r['IN_SALA_ATENDIMENTO_ESPECIAL'] == '1' else 0
            if bas == 0 and esp == 0: continue
            uf = UF_COD[mun[:2]]
            por_uf.setdefault(uf, {}).setdefault(mun, {'m': [], 'e': []})[rede].append([r['CO_ENTIDADE'], nome_bonito(r['NO_ENTIDADE']), bas, esp, sala])
            s = resumo.setdefault(mun, [0] * 8); o = 0 if rede == 'm' else 4
            s[o] += 1; s[o + 1] += esp > 0; s[o + 2] += sala; s[o + 3] += (esp > 0 and not sala)
    os.makedirs(os.path.join(BASE, 'escolas'), exist_ok=True)
    tot = 0
    for uf, d in por_uf.items():
        for m in d.values():
            for k in ('m', 'e'): m[k].sort(key=lambda x: (-x[3], -x[2]))
        json.dump({'ano': int(ano), 'mun': d}, open(os.path.join(BASE, f'escolas/{uf}.json'), 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))
        tot += sum(len(m['m']) + len(m['e']) for m in d.values())
    json.dump({'ano': int(ano), 'mun': resumo}, open(os.path.join(BASE, 'dados/escolas_resumo.json'), 'w'), separators=(',', ':'))
    tam = sum(os.path.getsize(os.path.join(BASE, 'escolas', x)) for x in os.listdir(os.path.join(BASE, 'escolas')))
    print(f'Censo Escolar {ano}: {tot:,} escolas públicas (estaduais e municipais) em {len(resumo):,} municípios · {len(por_uf)} arquivos por UF, {tam/1e6:.1f} MB')
    g = resumo.get('3518701'); print('Guarujá:', g)


if __name__ == '__main__':
    sys.exit(main())
