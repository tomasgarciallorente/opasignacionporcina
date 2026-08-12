# -*- coding: utf-8 -*-
"""Capa de acceso a Supabase para la app puntual de asignación porcina — tablas propias
(porcino_*), no toca stock_piezas ni nada de crm-extendido. Usa el service_role key
(bypassa RLS) porque el login de esta app es una contraseña compartida, no Supabase Auth
por usuario — ver plan (moonlit-prancing-willow.md)."""
import datetime
import uuid

import streamlit as st
from supabase import create_client, Client


@st.cache_resource
def get_client() -> Client:
    return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_SERVICE_ROLE_KEY"])


# --- Tipificación (carga del tipificador) ---

def insert_lote(proveedor, fecha_faena, mercaderia, correlativo_inicial, cantidad, foto_url, creado_por):
    sb = get_client()
    res = sb.table("porcino_tipificacion_lotes").insert({
        "proveedor": proveedor,
        "fecha_faena": fecha_faena.isoformat(),
        "mercaderia": mercaderia,
        "correlativo_inicial": correlativo_inicial,
        "cantidad": cantidad,
        "foto_url": foto_url,
        "creado_por": creado_por,
    }).execute()
    return res.data[0]["id"]


def insert_animales(lote_id, animales):
    """animales: [{correlativo, kg, nivel_grasa, observacion}]"""
    sb = get_client()
    rows = [{"lote_id": lote_id, **a} for a in animales]
    sb.table("porcino_tipificacion_animales").insert(rows).execute()


def fetch_animales_recientes(dias=7):
    """Animales tipificados en los últimos N días, con datos del lote (proveedor,
    fecha_faena, mercaderia) — pool para 'Generar reparto'."""
    sb = get_client()
    desde = (datetime.date.today() - datetime.timedelta(days=dias)).isoformat()
    lotes = sb.table("porcino_tipificacion_lotes").select("*").gte("fecha_faena", desde).execute().data
    if not lotes:
        return []
    lotes_por_id = {l["id"]: l for l in lotes}
    animales = sb.table("porcino_tipificacion_animales").select("*").in_(
        "lote_id", list(lotes_por_id.keys())
    ).execute().data
    out = []
    for a in animales:
        lote = lotes_por_id[a["lote_id"]]
        out.append({
            "correlativo": a["correlativo"],
            "kg": float(a["kg"]),
            "nivel_grasa": a["nivel_grasa"],
            "observacion": a["observacion"],
            "proveedor": lote["proveedor"],
            "fecha_faena": lote["fecha_faena"],
            "mercaderia": lote["mercaderia"],
        })
    return out


def upload_foto(file_bytes, filename_hint, content_type):
    sb = get_client()
    ext = filename_hint.rsplit(".", 1)[-1] if "." in filename_hint else "jpg"
    path = f"{datetime.date.today().isoformat()}/{uuid.uuid4().hex}.{ext}"
    sb.storage.from_("romaneos-fotos").upload(path, file_bytes, {"content-type": content_type})
    return sb.storage.from_("romaneos-fotos").get_public_url(path)


def insert_fotos(lote_id, urls):
    """Varias fotos por tropa (Tomás, 2026-08-11) — ver porcino_tipificacion_fotos."""
    if not urls:
        return
    sb = get_client()
    sb.table("porcino_tipificacion_fotos").insert(
        [{"lote_id": lote_id, "foto_url": u} for u in urls]
    ).execute()


def fetch_fotos(lote_id):
    sb = get_client()
    return [r["foto_url"] for r in sb.table("porcino_tipificacion_fotos").select("foto_url").eq("lote_id", lote_id).execute().data]


# --- Snapshots subidos por Excel ("Actualizar datos") ---

def replace_stock_snapshot(rows):
    """rows: [{correlativo, proveedor, fecha_faena, kg, mercaderia, nivel_grasa, comprometido}]
    Reemplaza el snapshot entero (delete + insert) — siempre es 'el último Excel subido'."""
    sb = get_client()
    sb.table("porcino_stock_snapshot").delete().neq("id", 0).execute()
    if rows:
        for i in range(0, len(rows), 500):
            sb.table("porcino_stock_snapshot").insert(rows[i:i + 500]).execute()


def fetch_stock_snapshot():
    sb = get_client()
    return sb.table("porcino_stock_snapshot").select("*").execute().data


def fetch_stock_snapshot_info():
    """Última vez que se sincronizó el snapshot y cuántas filas trajo — para que se vea en
    'Generar reparto' si el stock está fresco antes de correr el motor (Tomás, 2026-08-11:
    '¿cómo sé si se actualizó bien?')."""
    sb = get_client()
    res = sb.table("porcino_stock_snapshot").select("subido_at").order("subido_at", desc=True).limit(1).execute().data
    total = sb.table("porcino_stock_snapshot").select("id", count="exact").execute()
    return {
        "ultimo_sync": res[0]["subido_at"] if res else None,
        "total_filas": total.count or 0,
    }


def fetch_tipificacion_hoy_info():
    """Lotes de tipificación cargados hoy (fecha_faena = hoy) — cantidad de animales y de
    qué proveedor/mercadería, para el mismo panel de 'Generar reparto'."""
    sb = get_client()
    hoy = datetime.date.today().isoformat()
    lotes = sb.table("porcino_tipificacion_lotes").select("*").eq("fecha_faena", hoy).execute().data
    return lotes


def replace_historico(bloque_rows, calidad_rows):
    """bloque_rows: [{bloque_codigo, bloque_nombre, mercaderia, dia, cupo}]
    calidad_rows: [{bloque_codigo, mercaderia, eje, categoria, pct}]"""
    sb = get_client()
    sb.table("porcino_historico_bloque").delete().neq("id", 0).execute()
    sb.table("porcino_historico_calidad").delete().neq("id", 0).execute()
    if bloque_rows:
        for i in range(0, len(bloque_rows), 500):
            sb.table("porcino_historico_bloque").insert(bloque_rows[i:i + 500]).execute()
    if calidad_rows:
        for i in range(0, len(calidad_rows), 500):
            sb.table("porcino_historico_calidad").insert(calidad_rows[i:i + 500]).execute()


def fetch_historico():
    sb = get_client()
    bloque = sb.table("porcino_historico_bloque").select("*").execute().data
    calidad = sb.table("porcino_historico_calidad").select("*").execute().data
    return bloque, calidad


# --- Resultado de una corrida de reparto ---

def insert_reparto_resultados(dia_reparto, filas):
    """filas: [{mercaderia, bloque_codigo, correlativo, kg, nivel_grasa, observacion,
    reasignado_por_golpe_corte}]"""
    sb = get_client()
    corrida_at = datetime.datetime.now(datetime.timezone.utc).isoformat()
    rows = [{"dia_reparto": dia_reparto, "corrida_at": corrida_at, **f} for f in filas]
    for i in range(0, len(rows), 500):
        sb.table("porcino_reparto_resultados").insert(rows[i:i + 500]).execute()
