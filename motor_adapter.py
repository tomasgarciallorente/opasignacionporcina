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

from asignacion_engine import HistoricalData, dia_de_reparto  # noqa: E402
from stock_real import peso_bin, TIPIF_OTHER  # noqa: E402
from asignacion_semana import asignar_semana, fechas_reparto_semana  # noqa: E402

DIAS_PY_A_ES = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo']

# Mismo ranking que se usó a mano el 2026-08-11: bloques ordenados de más a menos
# favorecidos por capones magros, según el % histórico de tipificación baja (-1/0/1) —
# se recalcula en tiempo real desde el histórico cargado, no queda hardcodeado.
MAGRO_CATS = ('-1', '0', '1')


def _parse_fecha(v):
    if isinstance(v, datetime.date):
        return v
    return datetime.date.fromisoformat(str(v)[:10])


def construir_stock_rows(snapshot_rows, animales_tipificados, ya_repartidos=None):
    """snapshot_rows: filas de porcino_stock_snapshot (Excel subido, 'resto' del stock).
    animales_tipificados: filas de porcino_tipificacion_animales+lote (carga de hoy de los
    tipificadores, ver db.fetch_animales_recientes). ya_repartidos: {(mercaderia,
    correlativo)} ya asignado y guardado en una corrida anterior (ver
    db.fetch_correlativos_ya_repartidos) — se excluye del pool para que no se vuelva a
    ofrecer (Tomás, 2026-08-12: lo no asignado hoy se arrastra a mañana, pero lo YA asignado
    no vuelve a aparecer). Devuelve la lista de dicts que espera asignar_capon/asignar_chancha
    (mismo shape que stock_real.read_stock_bd()) — evita duplicar un correlativo si aparece
    en snapshot Y en tipificación (la carga del tipificador manda, por si el snapshot está
    desactualizado para esa pieza puntual)."""
    ya_repartidos = ya_repartidos or set()
    correlativos_tipificados = {(a['mercaderia'], a['correlativo']) for a in animales_tipificados}
    rows = []
    for a in animales_tipificados:
        if (a['mercaderia'], a['correlativo']) in ya_repartidos:
            continue
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
        if (s['mercaderia'], s['correlativo']) in ya_repartidos:
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
        if fila not in asignado.get(origen_code, []):
            # Ya se movió como pareja de swap de otro animal con observación de este
            # mismo bloque (procesado antes en este loop) — no hay nada más que hacer.
            continue
        destino_code, swap = None, None
        for cand in ranking_local:
            if cand == origen_code:
                continue
            candidatos = [g for g in asignado.get(cand, []) if g['peso'] == fila['peso'] and not g.get('observacion')]
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


def generar_reparto(snapshot_rows, animales_tipificados, bloque_rows, calidad_rows, dia=None,
                     ya_repartidos=None):
    """Corre el motor completo (Capón + Chancha + reasignación de golpes/cortes). dia: 'Martes'
    etc. — default, mañana hábil desde hoy (mismo criterio que dia_de_reparto()).

    Corrección de Tomás (2026-08-12): "el reparto del día NO debe asignar el 100% del stock,
    sino lo que le corresponde al cupo del día — lo que SÍ debe pasar es que el 100% del stock
    termine repartido entre los bloques, según cuánto trabaja cada uno." O sea: modo='cupo'
    (respeta el cupo histórico de CADA día, no vacía todo de una), pero lo que sobra hoy no se
    pierde — se excluye de futuras corridas todo lo que ya se guardó (ver ya_repartidos /
    db.fetch_correlativos_ya_repartidos), así el sobrante de hoy queda disponible para que
    mañana lo tome otra corrida, hasta agotar el 100% en sucesivos días.

    Fix de Tomás (2026-08-13): "no me está repartiendo con criterio FIFO a rajatabla, eran más
    de 160 capones de ayer y antes de ayer y me reparte un total de 118 del 11 y del 12" — este
    reparto de UN día llamaba directo a asignar_capon/asignar_chancha con el pool entero, sin el
    filtro "vieja vs nueva" que sí tiene el semanal (ver asignar_semana): dentro de cada
    categoría de peso/tipificación el orden SÍ era FIFO, pero nada impedía tomar faena de HOY
    para una categoría mientras quedaba faena de días anteriores sin tocar en OTRA categoría.
    Se corrige reusando asignar_semana() con una lista de un solo día — es el mismo motor
    validado del semanal (1° pasada solo contra 'vieja', recién si no alcanza el cupo se abre
    'nueva'), así el reparto de un día y el de la semana quedan con la MISMA disciplina en vez
    de dos caminos que podían divergir."""
    if dia is None:
        dia = DIAS_PY_A_ES[dia_de_reparto(datetime.date.today()).weekday()]

    hist = HistoricalData.from_rows(bloque_rows, calidad_rows)
    hist.redirigir('V04', 'V11')  # cartera de V04 (dejó de trabajar) sigue yendo a V11

    stock_rows = construir_stock_rows(snapshot_rows, animales_tipificados, ya_repartidos=ya_repartidos)

    dias_fechas_semana = fechas_reparto_semana()
    fecha_dia = next((f for d, f in dias_fechas_semana if d == dia), None)
    if fecha_dia is None:
        # dia pedido ya no está en lo que queda de esta semana (p.ej. se corrió a mano un día
        # que ya pasó) — sin una fecha de referencia no hay 'futuro' que filtrar, se usa hoy.
        fecha_dia = datetime.date.today()

    out = {'dia': dia}
    for merc_key, merc in [('capon', 'Capón'), ('chancha', 'Chancha')]:
        pool = [r for r in stock_rows if r['merc'] == merc]
        resultados, _pool_pendiente = asignar_semana(hist, pool, merc, [(dia, fecha_dia)])
        r = resultados[0]
        ranking = ranking_magro(hist, merc)
        movimientos = reasignar_golpes_cortes(r['correlativos'], r['shares'], ranking)
        out[merc_key] = {
            'shares': r['shares'], 'target': r['asignado'], 'matrix': r['matrix'],
            'asignado': r['correlativos'], 'disponible': r['disponible_dia'],
            'asignado_total': r['total_asignado_dia'], 'sobrante': r['sobrante'],
            'cupo': r['presupuesto'], 'movimientos_golpes_cortes': movimientos,
        }
    return out


def generar_reparto_semanal(snapshot_rows, animales_tipificados, bloque_rows, calidad_rows, ya_repartidos=None):
    """Misma cascada que generar_reparto() pero para TODOS los días que quedan de la semana
    actual (desde el día de reparto de hoy hasta el viernes) de una — Tomás, 2026-08-12:
    "quiero que hagas tanto una asignación para el reparto del día siguiente... como también
    la asignación semanal". Reusa asignacion_semana.asignar_semana tal cual está validada en
    el motor de escritorio: el sobrante de un día pasa como pool al siguiente, y el déficit
    (cuando el stock no alcanza el cupo de un día) se arrastra y se intenta cubrir más
    adelante en la semana — nunca se pierde silenciosamente."""
    hist = HistoricalData.from_rows(bloque_rows, calidad_rows)
    hist.redirigir('V04', 'V11')

    stock_rows = construir_stock_rows(snapshot_rows, animales_tipificados, ya_repartidos=ya_repartidos)
    dias_fechas = fechas_reparto_semana()
    dias = [d for d, _f in dias_fechas]

    out = {'dias': dias, 'semana_iso': dias_fechas[0][1].isocalendar()[1] if dias_fechas else None}
    for merc_key, merc in [('capon', 'Capón'), ('chancha', 'Chancha')]:
        pool = [r for r in stock_rows if r['merc'] == merc]
        resultados, pool_pendiente = asignar_semana(hist, pool, merc, dias_fechas)

        ranking = ranking_magro(hist, merc)
        movimientos_por_dia = {}
        for r in resultados:
            movimientos_por_dia[r['dia']] = reasignar_golpes_cortes(r['correlativos'], r['shares'], ranking)

        out[merc_key] = {
            'resultados': resultados,  # [{'dia','shares','presupuesto','asignado','correlativos','sobrante',...}]
            'sobrante_fin_semana': pool_pendiente,
            'movimientos_golpes_cortes_por_dia': movimientos_por_dia,
        }
    return out
