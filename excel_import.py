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
    db.replace_stock_snapshot()."""
    wb = openpyxl.load_workbook(file, data_only=True)
    ws = wb['STOCK (BD)'] if 'STOCK (BD)' in wb.sheetnames else wb.active
    header_row = None
    for r in range(1, 20):
        if str(ws.cell(r, 4).value or '').strip().upper() == 'CORRELATIVO':
            header_row = r
            break
    if header_row is None:
        raise ValueError("No encontré la fila de header (columna D = 'CORRELATIVO'). ¿Es la hoja STOCK (BD)?")

    rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        correlativo = ws.cell(r, 4).value
        kg = ws.cell(r, 5).value
        if not correlativo or not isinstance(kg, (int, float)) or kg <= 0:
            continue
        merc_raw = str(ws.cell(r, 7).value or '').strip().upper()
        if merc_raw in ('CA', 'MEI'):
            merc = 'Capón'
        elif merc_raw == 'CH':
            merc = 'Chancha'
        else:
            continue
        observaciones = ws.cell(r, 14).value
        comprometido = bool(observaciones and 'sale' in str(observaciones).lower())
        fecha = ws.cell(r, 2).value
        tipif_raw = ws.cell(r, 3).value
        rows.append({
            'correlativo': int(correlativo),
            'proveedor': str(ws.cell(r, 1).value or '').strip() or None,
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
