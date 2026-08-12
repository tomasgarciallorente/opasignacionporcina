# -*- coding: utf-8 -*-
"""Herramienta puntual de asignación de mercadería porcina — Origen Pampa.
Login simple (contraseña compartida) -> Cargar tipificación -> Generar reparto ->
Actualizar datos. Ver plan: C:\\Users\\Gtecomercial\\.claude\\plans\\moonlit-prancing-willow.md
"""
import datetime
import io

import streamlit as st

import db
import excel_import
from motor_adapter import generar_reparto, construir_stock_rows
from run_asignacion_stock_real import write_resumen_dia, write_correlativos, write_sobrante
import openpyxl

st.set_page_config(page_title='Asignación porcina — Origen Pampa', page_icon='🐖', layout='wide')

OBSERVACIONES = [
    'Ninguna', 'Golpe en cadera', 'Golpe interno', 'Golpe interno feo',
    'Corte en paleta (cuero)', 'Otro',
]
GRADOS = ['-1', '0', '1', '1+', '1++', '2', '2+', '3', '3+', '4']
DIAS_HABILES = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes']


# --- Login ---
def check_password():
    if st.session_state.get('autenticado'):
        return True
    st.title('🐖 Asignación porcina — Origen Pampa')
    pw = st.text_input('Contraseña', type='password')
    if st.button('Entrar'):
        if pw == st.secrets.get('APP_PASSWORD'):
            st.session_state['autenticado'] = True
            st.rerun()
        else:
            st.error('Contraseña incorrecta.')
    return False


if not check_password():
    st.stop()


# --- Navegación ---
pagina = st.sidebar.radio('Página', ['Cargar tipificación', 'Generar reparto', 'Actualizar datos'])
st.sidebar.button('Cerrar sesión', on_click=lambda: st.session_state.pop('autenticado', None))


# ============================== CARGAR TIPIFICACIÓN ==============================
if pagina == 'Cargar tipificación':
    st.title('Cargar tipificación de una tropa')

    if 'tropa_generada' not in st.session_state:
        st.session_state['tropa_generada'] = None

    with st.form('form_tropa'):
        col1, col2 = st.columns(2)
        proveedor = col1.text_input('Proveedor', placeholder='Ej. ISOWEAN S.A.')
        fecha_faena = col2.date_input('Fecha de faena', value=datetime.date.today())
        col3, col4, col5 = st.columns(3)
        mercaderia = col3.radio('Mercadería', ['Capón', 'Chancha'], horizontal=True)
        correlativo_inicial = col4.number_input('Correlativo inicial', min_value=1, step=1)
        cantidad = col5.number_input('Cantidad de cabezas', min_value=1, max_value=500, step=1)
        generar = st.form_submit_button('Generar planilla de carga')

    if generar:
        st.session_state['tropa_generada'] = {
            'proveedor': proveedor, 'fecha_faena': fecha_faena, 'mercaderia': mercaderia,
            'correlativo_inicial': int(correlativo_inicial), 'cantidad': int(cantidad),
        }

    tropa = st.session_state['tropa_generada']
    if tropa:
        st.divider()
        st.subheader(f"{tropa['mercaderia']} — {tropa['cantidad']} cabezas, correlativo {tropa['correlativo_inicial']}..{tropa['correlativo_inicial'] + tropa['cantidad'] - 1}")
        st.caption('Un renglón por animal — Kg + grado (tocá el botón) + observación si tiene golpe/corte.')

        filas = []
        for i in range(tropa['cantidad']):
            correlativo = tropa['correlativo_inicial'] + i
            c1, c2, c3 = st.columns([1, 2, 2])
            c1.markdown(f"**{correlativo}**")
            kg = c1.number_input('Kg', min_value=1.0, max_value=300.0, step=1.0, key=f'kg_{correlativo}', label_visibility='collapsed')
            grado = c2.radio('Grado', GRADOS, horizontal=True, key=f'grado_{correlativo}', label_visibility='collapsed')
            obs = c3.selectbox('Observación', OBSERVACIONES, key=f'obs_{correlativo}', label_visibility='collapsed')
            filas.append({'correlativo': correlativo, 'kg': kg, 'nivel_grasa': grado,
                          'observacion': None if obs == 'Ninguna' else obs})

        st.divider()
        fotos = st.file_uploader('Fotos del romaneo (opcional, respaldo de toda la tropa — podés subir varias)',
                                  type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)
        creado_por = st.text_input('Tu nombre', placeholder='Ej. Ignacio Beas')

        if st.button('Confirmar y guardar', type='primary'):
            if not proveedor:
                st.error('Falta el proveedor.')
            elif not creado_por:
                st.error('Falta tu nombre.')
            else:
                lote_id = db.insert_lote(tropa['proveedor'], tropa['fecha_faena'], tropa['mercaderia'],
                                          tropa['correlativo_inicial'], tropa['cantidad'], None, creado_por)
                db.insert_animales(lote_id, filas)
                if fotos:
                    urls = [db.upload_foto(f.getvalue(), f.name, f.type) for f in fotos]
                    db.insert_fotos(lote_id, urls)
                st.success(f"Guardado — {tropa['cantidad']} animales cargados"
                           + (f", {len(fotos)} foto(s)." if fotos else "."))
                st.session_state['tropa_generada'] = None
                st.rerun()


# ============================== GENERAR REPARTO ==============================
elif pagina == 'Generar reparto':
    st.title('Generar reparto por bloques')

    bloque_rows, calidad_rows = db.fetch_historico()
    if not bloque_rows:
        st.warning('Todavía no se subió el Excel histórico (Promedio diario Capón y Chancha por '
                   'vendedor.xlsx) — andá a "Actualizar datos" antes de generar un reparto.')
        st.stop()

    # --- Estado del stock, deduplicado de verdad (Tomás, 2026-08-12: "está poniendo un stock
    # que no existe" — el total anterior sumaba snapshot + tipificación sin descontar los
    # correlativos que aparecen en ambos lados) ---
    snap_info = db.fetch_stock_snapshot_info()
    lotes_hoy = db.fetch_tipificacion_hoy_info()
    snapshot_rows = db.fetch_stock_snapshot()
    animales_tipificados = db.fetch_animales_recientes()
    stock_rows = construir_stock_rows(snapshot_rows, animales_tipificados)
    animales_hoy = sum(l['cantidad'] for l in lotes_hoy)
    n_capon = sum(1 for r in stock_rows if r['merc'] == 'Capón')
    n_chancha = sum(1 for r in stock_rows if r['merc'] == 'Chancha')

    with st.container(border=True):
        st.markdown('**Stock disponible ahora mismo (ya deduplicado)**')
        c1, c2, c3 = st.columns(3)
        if snap_info['ultimo_sync']:
            ts = datetime.datetime.fromisoformat(snap_info['ultimo_sync'].replace('Z', '+00:00'))
            ts_local = ts.astimezone().strftime('%d/%m %H:%M')
            c1.metric('Última sync del Excel', ts_local, f"{snap_info['total_filas']} piezas en el snapshot")
        else:
            c1.metric('Última sync del Excel', 'nunca')
        c2.metric('Tipificado hoy', f'{animales_hoy} animales', f'{len(lotes_hoy)} tropa(s)')
        c3.metric('Total disponible real', f'{len(stock_rows)}', f'{n_capon} Capón / {n_chancha} Chancha')
        if lotes_hoy:
            st.caption('Tropas de hoy: ' + ', '.join(
                f"{l['proveedor']} ({l['mercaderia']}, {l['cantidad']})" for l in lotes_hoy))

        with st.expander(f'Ver detalle — {len(stock_rows)} piezas'):
            st.dataframe(
                [{'Correlativo': r['correlativo'], 'Mercadería': r['merc'], 'Kg': r['kg'],
                  'Grado': r['tipif'], 'Proveedor': r['proveedor'],
                  'Fecha faena': r['fecha_faena'], 'Observación': r.get('observacion') or ''}
                 for r in sorted(stock_rows, key=lambda r: (r['merc'], r['correlativo']))],
                use_container_width=True, hide_index=True, height=300,
            )

    dia = st.selectbox('Día de reparto', DIAS_HABILES, index=1)
    st.caption('El reparto siempre distribuye el 100% del stock disponible entre los bloques, '
               'proporcional a cuánto trabaja cada uno históricamente (cantidad y calidad) — '
               'sin dejar sobrante.')

    if st.button('Generar propuesta', type='primary'):
        if not stock_rows:
            st.error('No hay stock cargado — ni tipificación de hoy ni snapshot de Excel subido.')
            st.stop()

        resultado = generar_reparto(snapshot_rows, animales_tipificados, bloque_rows, calidad_rows, dia=dia)
        st.session_state['resultado'] = resultado

    resultado = st.session_state.get('resultado')
    if resultado:
        st.divider()
        st.caption(f"Reparto para el {resultado['dia']}")
        for merc_key, merc_label in [('capon', 'Capón'), ('chancha', 'Chancha')]:
            d = resultado[merc_key]
            st.subheader(merc_label)
            c1, c2, c3 = st.columns(3)
            c1.metric('Disponible', d['disponible'])
            c2.metric('Asignado', d['asignado_total'])
            c3.metric('Sobrante', len(d['sobrante']))

            filas_resumen = []
            for code, (nombre, _share) in d['shares'].items():
                filas_resumen.append({'Bloque': nombre, 'Asignado': d['target'].get(code, 0)})
            st.dataframe(filas_resumen, use_container_width=True, hide_index=True)

            if d['movimientos_golpes_cortes']:
                st.markdown('**Golpes / cortes — reasignados por afinidad a magros**')
                st.dataframe(d['movimientos_golpes_cortes'], use_container_width=True, hide_index=True)

            with st.expander(f'Detalle de correlativos por bloque — {merc_label}'):
                for code, (nombre, _share) in d['shares'].items():
                    animales = d['asignado'].get(code, [])
                    if not animales:
                        continue
                    st.markdown(f'**{nombre}** ({len(animales)})')
                    st.dataframe(
                        [{'Correlativo': a['correlativo'], 'Kg': a['kg'], 'Grado': a['tipif'],
                          'Observación': a.get('observacion') or ''} for a in animales],
                        use_container_width=True, hide_index=True,
                    )

        # --- Excel de descarga, mismo formato que el motor de escritorio ---
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        ws0 = wb.create_sheet('Resumen del día')
        write_resumen_dia(ws0, resultado['dia'], {
            'Capón': (resultado['capon']['shares'], {}, resultado['capon']['target'],
                      resultado['capon']['disponible'], resultado['capon']['asignado_total']),
            'Chancha': (resultado['chancha']['shares'], {}, resultado['chancha']['target'],
                        resultado['chancha']['disponible'], resultado['chancha']['asignado_total']),
        })
        for merc_key, merc_label, con_peso in [('capon', 'Capón', True), ('chancha', 'Chancha', False)]:
            d = resultado[merc_key]
            ws_det = wb.create_sheet(f'{merc_label} - Correlativos')
            write_correlativos(ws_det, merc_label, d['asignado'], d['shares'])
            ws_sob = wb.create_sheet(f'{merc_label} - Sobrante')
            write_sobrante(ws_sob, merc_label, d['sobrante'])

        buf = io.BytesIO()
        wb.save(buf)
        st.download_button('Descargar Excel', buf.getvalue(),
                            file_name=f"Propuesta reparto - {resultado['dia']} - {datetime.date.today().isoformat()}.xlsx")

        if st.button('Guardar este resultado'):
            filas_guardar = []
            for merc_key, merc_label in [('capon', 'Capón'), ('chancha', 'Chancha')]:
                d = resultado[merc_key]
                movidos = {m['correlativo'] for m in d['movimientos_golpes_cortes'] if m['destino']}
                for code, (_nombre, _s) in d['shares'].items():
                    for a in d['asignado'].get(code, []):
                        filas_guardar.append({
                            'mercaderia': merc_label, 'bloque_codigo': code, 'correlativo': a['correlativo'],
                            'kg': a['kg'], 'nivel_grasa': a['tipif'], 'observacion': a.get('observacion'),
                            'reasignado_por_golpe_corte': a['correlativo'] in movidos,
                        })
            db.insert_reparto_resultados(resultado['dia'], filas_guardar)
            st.success('Guardado.')


# ============================== ACTUALIZAR DATOS ==============================
elif pagina == 'Actualizar datos':
    st.title('Actualizar datos desde Excel')
    st.caption('Subí los mismos archivos que ya usás — no hace falta tocar nada de ellos.')

    st.subheader('Stock (STOCK (BD) del Excel de Dropbox)')
    st.caption("Exportá/copiá el archivo '2B-OP Stock y Entregas Porcino 2026.xlsx' y subilo acá. "
               "Reemplaza el snapshot anterior entero.")
    f_stock = st.file_uploader('Excel de stock', type=['xlsx'], key='up_stock')
    if f_stock is not None and st.button('Cargar stock'):
        try:
            rows = excel_import.parse_stock_bd(f_stock)
            db.replace_stock_snapshot(rows)
            st.success(f'Cargado — {len(rows)} piezas de stock.')
        except Exception as e:
            st.error(f'Error al leer el Excel: {e}')

    st.divider()
    st.subheader('Histórico de bloques (Promedio diario Capón y Chancha por vendedor.xlsx)')
    st.caption('Cupo por día de semana y % de peso/tipificación por bloque — reemplaza el histórico anterior entero.')
    f_hist = st.file_uploader('Excel histórico', type=['xlsx'], key='up_hist')
    if f_hist is not None and st.button('Cargar histórico'):
        try:
            bloque_rows, calidad_rows = excel_import.parse_historico(f_hist)
            db.replace_historico(bloque_rows, calidad_rows)
            st.success(f'Cargado — {len(bloque_rows)} filas de cupo, {len(calidad_rows)} filas de calidad.')
        except Exception as e:
            st.error(f'Error al leer el Excel: {e}')
