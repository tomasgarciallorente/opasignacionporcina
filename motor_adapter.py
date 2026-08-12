# -*- coding: utf-8 -*-
"""Puente entre los datos que vienen de Supabase (snapshot de stock subido por Excel +
tipificación cargada por los tipificadores) y el motor ya validado en
`Análisis de entregas/Asignación de mercadería/` (asignacion_engine.py,
run_asignacion_stock_real.py, stock_real.py) — no reescribe la lógica de negocio, solo arma
las estructuras de datos que el motor espera y expone una función única `generar_reparto()`
para la app."""
import sys
import os
import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import run_asignacion_stock_real as ras  # noqa: E402
from asignacion_engine import HistoricalData, dia_de_reparto  # noqa: E402
from stock_real import peso_bin, TIPIF_OTHER  # noqa: E402

DIAS_PY_A_ES = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo']

# Mismo ranking que se usó a mano el 2026-08-11: bloques ordenados de más a menos
# favorecidos por capones magros, según el % histórico de tipificación baja (-1/0/1) —
# se recalcula en tiempo real desde el histórico cargado, no queda hardcodeado.
MAGRO_CATS = ('-1', '0', '1')


def _parse_fecha(v):
    if isinstance(v, datetime.date):
        return v
    return datetime.date.fromisoformat(str(v)[:10])


def construir_stock_rows(snapshot_rows, animales_tipificados):
    """snapshot_rows: filas de porcino_stock_snapshot (Excel subido, 'resto' del stock).
    animales_tipificados: filas de porcino_tipificacion_animales+lote (carga de hoy de los
    tipificadores, ver db.fetch_animales_recientes). Devuelve la lista de dicts que espera
    asignar_capon/asignar_chancha (mismo shape que stock_real.read_stock_bd()) — evita
    duplicar un correlativo si aparece en ambas fuentes (la carga del tipificador manda,
    por si el snapshot está desactualizado para esa pieza puntual)."""
    correlativos_tipificados = {(a['mercaderia'], a['correlativo']) for a in animales_tipificados}
    rows = []
    for a in animales_tipificados:
        kg = float(a['kg'])
        rows.append({
            'correlativo': a['correlativo'],
            'proveedor': a['proveedor'],
            'fecha_faena': _parse_fecha(a['fecha_faena']),
            'tipif_raw': a['nivel_grasa'],
            'tipif': a['nivel_grasa'] or TIPIF_OTHER,
            'kg': kg,
            'peso': peso_bin(kg),
            'merc': a['mercaderia'],
            'es_real': True,
            'observacion': a.get('observacion'),
        })
    for s in snapshot_rows:
        if s.get('comprometido'):
            continue
        if (s['mercaderia'], s['correlativo']) in correlativos_tipificados:
            continue
        kg = float(s['kg']) if s.get('kg') else None
        rows.append({
            'correlativo': s['correlativo'],
            'proveedor': s.get('proveedor'),
            'fecha_faena': _parse_fecha(s['fecha_faena']) if s.get('fecha_faena') else datetime.date.today(),
            'tipif_raw': s.get('nivel_grasa'),
            'tipif': s.get('nivel_grasa') or TIPIF_OTHER,
            'kg': kg,
            'peso': peso_bin(kg) if kg else None,
            'merc': s['mercaderia'],
            'es_real': True,
            'observacion': None,
        })
    return rows


def ranking_magro(hist, merc='Capón'):
    """Bloques ordenados de más a menos favorecidos por capones magros, según
    tipif_pct_capon (mismo criterio usado a mano el 2026-08-11)."""
    pct = hist.tipif_pct_capon if merc == 'Capón' else hist.tipif_pct_chancha
    shares = hist.bloque_shares(merc)
    filas = []
    for code in shares:
        p = pct.get(code, {})
        magro = sum(v for k, v in p.items() if k in MAGRO_CATS)
        filas.append((magro, code))
    filas.sort(reverse=True)
    return [code for _magro, code in filas]


def reasignar_golpes_cortes(asignado, shares, ranking):
    """Mismo criterio aplicado a mano hoy: cada correlativo con observación (golpe/corte)
    se mueve, vía swap 1 a 1 con un animal del mismo bloque de peso, al bloque más
    favorecido por magros que tenga un candidato disponible — sin alterar el total ni el
    mix de peso de ningún bloque. Devuelve la lista de movimientos (para mostrar y para
    marcar 'reasignado_por_golpe_corte' al guardar el resultado)."""
    con_observacion = [
        (code, f) for code, filas in asignado.items() for f in filas if f.get('observacion')
    ]
    movimientos = []
    ranking_local = list(ranking)
    for origen_code, fila in con_observacion:
        destino_code, swap = None, None
        for cand in ranking_local:
            if cand == origen_code:
                continue
            candidatos = [g for g in asignado.get(cand, []) if g['peso'] == fila['peso']]
            if candidatos:
                destino_code, swap = cand, candidatos[0]
                break
        if destino_code is None:
            movimientos.append({
                'correlativo': fila['correlativo'], 'motivo': fila['observacion'],
                'origen': origen_code, 'destino': None,
                'detalle': 'Ningún bloque del ranking tenía un animal del mismo peso para swapear.',
            })
            continue
        asignado[origen_code].remove(fila)
        asignado[destino_code].remove(swap)
        asignado[origen_code].append(swap)
        asignado[destino_code].append(fila)
        ranking_local.remove(destino_code)
        ranking_local.append(destino_code)
        movimientos.append({
            'correlativo': fila['correlativo'], 'motivo': fila['observacion'],
            'origen': origen_code, 'destino': destino_code,
            'detalle': f'Swap con correlativo {swap["correlativo"]} ({shares[destino_code][0]}).',
        })
    return movimientos


def generar_reparto(snapshot_rows, animales_tipificados, bloque_rows, calidad_rows, dia=None):
    """Corre el motor completo (Capón + Chancha + reasignación de golpes/cortes). dia: 'Martes'
    etc. — default, mañana hábil desde hoy (mismo criterio que dia_de_reparto()).

    modo='prorrateo' (Tomás, 2026-08-12: "quiero que el stock se reparta completo a los
    bloques en proporción de las cantidades que trabajan y en relación a la calidad que
    trabajan... no me puedo quedar con stock sin asignar") — reparte el 100% del stock
    disponible ahora mismo entre los bloques, proporcional a su participación histórica
    (capa1) y su perfil de peso/tipificación histórico (capa2/3). Ya no usa cupo/cuota
    manual — el reparto siempre vacía todo el stock disponible, sin sobrante."""
    if dia is None:
        dia = DIAS_PY_A_ES[dia_de_reparto(datetime.date.today()).weekday()]

    hist = HistoricalData.from_rows(bloque_rows, calidad_rows)
    hist.redirigir('V04', 'V11')  # cartera de V04 (dejó de trabajar) sigue yendo a V11

    stock_rows = construir_stock_rows(snapshot_rows, animales_tipificados)

    shares_c, target_c, matrix_c, asignado_c, disp_c, asig_c, sobrante_c, cupo_c = ras.asignar_capon(
        hist, stock_rows, dia=dia, modo='prorrateo')
    shares_h, target_h, matrix_h, asignado_h, disp_h, asig_h, sobrante_h, cupo_h = ras.asignar_chancha(
        hist, stock_rows, dia=dia, modo='prorrateo')

    ranking_c = ranking_magro(hist, 'Capón')
    movimientos_c = reasignar_golpes_cortes(asignado_c, shares_c, ranking_c)
    ranking_h = ranking_magro(hist, 'Chancha')
    movimientos_h = reasignar_golpes_cortes(asignado_h, shares_h, ranking_h)

    return {
        'dia': dia,
        'capon': {'shares': shares_c, 'target': target_c, 'matrix': matrix_c, 'asignado': asignado_c,
                  'disponible': disp_c, 'asignado_total': asig_c, 'sobrante': sobrante_c,
                  'movimientos_golpes_cortes': movimientos_c},
        'chancha': {'shares': shares_h, 'target': target_h, 'matrix': matrix_h, 'asignado': asignado_h,
                    'disponible': disp_h, 'asignado_total': asig_h, 'sobrante': sobrante_h,
                    'movimientos_golpes_cortes': movimientos_h},
    }
