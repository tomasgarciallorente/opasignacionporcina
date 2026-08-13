# -*- coding: utf-8 -*-
"""Parsers para la página 'Actualizar datos' — reciben un archivo subido (file-like, de
st.file_uploader) y devuelven filas listas para guardar en Supabase (ver db.py). Reusan la
lógica de lectura del motor (stock_real.py, asignacion_engine.py) pero SIN depender de rutas
locales (OP_CARGAS de stock_real.comprometido() vive en la máquina de Tomás, no en el
servidor de Streamlit) — para el stock comprometido, en esta app se usa solo la señal de
Observaciones ('SALE ...'), no el cruce contra la lista de proveedores reales. Simplificación
deliberada para la v1: cubre el caso real encontrado hoy (Rubiolo), a costa de ser un poco
menos estricta que el motor de escritorio."""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import openpyxl  # noqa: E402
from stock_real import peso_bin, clean_tipif  # noqa: E402
from asignacion_engine import DIAS_SEMANA  # noqa: E402


def parse_stock_bd(file) -> list[dict]:
    """Mismo formato de hoja que STOCK (BD) de 2B-OP Stock y Entregas Porcino 2026.xlsx —
    header en la fila donde la columna D dice 'CORRELATIVO'. Devuelve filas listas para
    db.replace_stock_snapshot().

    read_only=True + iter_rows() en vez de load_workbook() normal + ws.cell() por celda
    (Tomás, 2026-08-13: "es muy lenta la carga") — el Excel de Dropbox pesa ~6MB con estilos
    de todo el libro; cargarlo entero y acceder celda a celda tardaba MÁS DE 100 SEGUNDOS
    (medido). read_only con iter_rows() salta el parseo de estilos/formato y termina en
    fracciones de segundo — mismo resultado, sin reescribir la lógica de filtrado."""
    wb = openpyxl.load_workbook(file, data_only=True, read_only=True)
    ws = wb['STOCK (BD)'] if 'STOCK (BD)' in wb.sheetnames else wb.active
    all_rows = list(ws.iter_rows(values_only=True))
    wb.close()

    def _get(row, idx):
        return row[idx] if idx < len(row) else None

    header_idx = None
    for i, row in enumerate(all_rows[:20]):
        if str(_get(row, 3) or '').strip().upper() == 'CORRELATIVO':
            header_idx = i
            break
    if header_idx is None:
        raise ValueError("No encontré la fila de header (columna D = 'CORRELATIVO'). ¿Es la hoja STOCK (BD)?")

    rows = []
    for row in all_rows[header_idx + 1:]:
        correlativo = _get(row, 3)
        kg = _get(row, 4)
        if not correlativo or not isinstance(kg, (int, float)) or kg <= 0:
            continue
        merc_raw = str(_get(row, 6) or '').strip().upper()
        if merc_raw in ('CA', 'MEI'):
            merc = 'Capón'
        elif merc_raw == 'CH':
            merc = 'Chancha'
        else:
            continue
        observaciones = _get(row, 13)
        comprometido = bool(observaciones and 'sale' in str(observaciones).lower())
        fecha = _get(row, 1)
        tipif_raw = _get(row, 2)
        rows.append({
            'correlativo': int(correlativo),
            'proveedor': str(_get(row, 0) or '').strip() or None,
            'fecha_faena': fecha.date().isoformat() if hasattr(fecha, 'date') else None,
            'kg': float(kg),
            'mercaderia': merc,
            'nivel_grasa': clean_tipif(tipif_raw) if tipif_raw is not None else None,
            'comprometido': comprometido,
        })
    return rows


def parse_historico(file) -> tuple[list[dict], list[dict]]:
    """Mismo formato que 'Promedio diario Capón y Chancha por vendedor.xlsx' — reusa
    HistoricalData (asignacion_engine.py) para no reimplementar el parseo, y aplana el
    resultado a filas para db.replace_historico()."""
    from asignacion_engine import HistoricalData
    hist = HistoricalData(xlsx_path=file)

    bloque_rows = []
    for merc, plan in [('Capón', hist.plan_capon), ('Chancha', hist.plan_chancha)]:
        for code, name, _row, dias in plan:
            for dia in DIAS_SEMANA + ['Total']:
                bloque_rows.append({
                    'bloque_codigo': code, 'bloque_nombre': name, 'mercaderia': merc,
                    'dia': dia, 'cupo': float(dias.get(dia, 0.0)),
                })

    calidad_rows = []
    for merc, eje, pct_dict in [
        ('Capón', 'peso', hist.peso_pct),
        ('Capón', 'tipif', hist.tipif_pct_capon),
        ('Chancha', 'tipif', hist.tipif_pct_chancha),
    ]:
        for code, categorias in pct_dict.items():
            for categoria, pct in categorias.items():
                calidad_rows.append({
                    'bloque_codigo': code, 'mercaderia': merc, 'eje': eje,
                    'categoria': categoria, 'pct': float(pct),
                })
    return bloque_rows, calidad_rows
