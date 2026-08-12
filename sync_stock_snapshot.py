# -*- coding: utf-8 -*-
"""Sincroniza el stock porcino actual (STOCK (BD) del Excel de Dropbox) a Supabase
(porcino_stock_snapshot) para que la app puntual de asignación siempre tenga el stock
fresco, sin que nadie tenga que subirlo a mano desde 'Actualizar datos'. Corre en esta
máquina (tiene el Dropbox sincronizado) — pedido de Tomás, 2026-08-11: "necesito que el
proceso sea rápido y simple" para Ignacio/Gonzalo. Programado cada 2 horas en horario
laboral vía Task Scheduler (ver README.md)."""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import excel_import
from supabase import create_client

URL = "https://ubryvnogujmhfjkdvbor.supabase.co"
KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InVicnl2bm9ndWptaGZqa2R2Ym9yIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc3OTgxODY1MSwiZXhwIjoyMDk1Mzk0NjUxfQ.meznQ7dVBNUsjwsCOKgr7TH7vx_xiRpPIijbT6RTHew"
STOCK_XLSX = r'C:\Users\Gtecomercial\Dropbox\01-Abasto\03-Entregas de carnes\2B-OP Stock y Entregas Porcino 2026.xlsx'


def main():
    sb = create_client(URL, KEY)
    rows = excel_import.parse_stock_bd(STOCK_XLSX)
    sb.table("porcino_stock_snapshot").delete().neq("id", 0).execute()
    for i in range(0, len(rows), 500):
        sb.table("porcino_stock_snapshot").insert(rows[i:i + 500]).execute()
    print(f'{len(rows)} filas sincronizadas ({sum(1 for r in rows if r["comprometido"])} comprometidas, excluidas del pool).')


if __name__ == '__main__':
    main()
