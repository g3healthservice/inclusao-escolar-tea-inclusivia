#!/usr/bin/env python3
"""Coleta, via API do SICONFI (Tesouro), o gasto com Educação Especial (subfunção 12.367) e com
Educação (função 12) de cada município e estado — DCA Anexo I-E (despesas por função/subfunção).

Saída: dados/siconfi_educacao_especial.json  {cod_ibge: [ano, edesp_empenhado, edesp_liquidado, educacao_empenhado]}
Usa o exercício mais recente entregue (2025; se ausente, 2024). Retomável: pula entes já coletados.
"""
import json, os, sys, time, urllib.request, urllib.parse
from concurrent.futures import ThreadPoolExecutor, as_completed

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, 'dados/siconfi_educacao_especial.json')
API = 'https://apidatalake.tesouro.gov.br/ords/siconfi/tt/dca'
ANOS = [2025, 2024]
UFS = {'11':'RO','12':'AC','13':'AM','14':'RR','15':'PA','16':'AP','17':'TO','21':'MA','22':'PI','23':'CE','24':'RN','25':'PB','26':'PE','27':'AL',
       '28':'SE','29':'BA','31':'MG','32':'ES','33':'RJ','35':'SP','41':'PR','42':'SC','43':'RS','50':'MS','51':'MT','52':'GO','53':'DF'}


def consulta(ente, ano):
    q = urllib.parse.urlencode({'an_exercicio': ano, 'no_anexo': 'DCA-Anexo I-E', 'id_ente': ente})
    for tent in range(4):
        try:
            with urllib.request.urlopen(API + '?' + q, timeout=60) as r:
                return json.load(r).get('items', [])
        except Exception:
            time.sleep(2 + 3 * tent)
    return None


def extrai(items):
    v = {}
    for r in items:
        c, col = str(r.get('conta', '')), r.get('coluna', '')
        if c.startswith('12.367') and col == 'Despesas Empenhadas': v['emp'] = r.get('valor')
        if c.startswith('12.367') and col == 'Despesas Liquidadas': v['liq'] = r.get('valor')
        if c.startswith('12 - Educação') and col == 'Despesas Empenhadas': v['edu'] = r.get('valor')
    return v


def coleta(ente):
    for ano in ANOS:
        it = consulta(ente, ano)
        if it is None: return ente, None            # falha de rede: tentar de novo depois
        if it:                                        # entregou este exercício
            v = extrai(it)
            return ente, [ano, round(v.get('emp') or 0, 2), round(v.get('liq') or 0, 2), round(v.get('edu') or 0, 2)]
    return ente, [0, 0, 0, 0]                         # não entregou 2025 nem 2024


def main():
    M = json.load(open(os.path.join(BASE, 'dados/tea_municipios_censo2022.json'), encoding='utf-8'))
    entes = [m[0] for m in M] + list(UFS.keys())
    res = json.load(open(OUT)) if os.path.exists(OUT) else {}
    falta = [e for e in entes if e not in res]
    print(f'{len(res)} já coletados; faltam {len(falta)}', flush=True)
    t0 = time.time(); feitos = 0
    with ThreadPoolExecutor(max_workers=6) as ex:
        futs = [ex.submit(coleta, e) for e in falta]
        for f in as_completed(futs):
            e, v = f.result(); feitos += 1
            if v is not None: res[e] = v
            if feitos % 200 == 0:
                json.dump(res, open(OUT, 'w'), separators=(',', ':'))
                print(f'{feitos}/{len(falta)} · {time.time()-t0:.0f}s', flush=True)
    json.dump(res, open(OUT, 'w'), separators=(',', ':'))
    anos = {}
    for v in res.values(): anos[v[0]] = anos.get(v[0], 0) + 1
    print('concluído:', len(res), 'entes · por exercício:', anos, flush=True)


if __name__ == '__main__':
    sys.exit(main())
