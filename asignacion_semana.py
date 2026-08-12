# -*- coding: utf-8 -*-
"""
Motor unificado semana completa — porcino (Tomás, 2026-08-04):
"[el stock real] lo debés considerar todo hasta acabarlo... los que te sobraron hoy se van a
asignar mañana para reparto del día siguiente. Y así ponerlos en planilla hasta acabar stock.
Y en los capones proyectados, aclarar que no existen realmente en el stock, sino que son
presupuesto."

Reemplaza el diseño anterior (Lunes = motor "Stock real" separado, Martes-Viernes = motor
"Proyectado" separado, cada uno con su propia lógica) por UN solo pool por mercadería: stock
real (correlativo/kg/tipificación ciertos) + filas sintéticas "PROY-" por cada faena todavía
no ejecutada del plan de Compras (es_real=False, sin correlativo real). El pool se consume
día por día (Lunes..Viernes) con la MISMA cascada (_asignar_generico) — el FIFO por fecha de
faena ya deja el stock real (fecha más vieja) primero y lo proyectado (fecha futura) después,
así que el "hasta acabar stock real, después proyectado" sale solo del orden, sin lógica aparte.

"Cupo" (columna Presupuesto) = el objetivo ANTES de tocar el stock disponible (histórico o
cuota real, según HistoricalData.cupo_o_cuota). "Asignado" = lo que efectivamente sale ese día
una vez cruzado contra el pool disponible (real + proyectado) — pueden diferir si el pool no
alcanza. Tomás, 2026-08-04: "el cupo es lo que presupuestamos asignarle previamente, y luego lo
asignado es lo que se va a repartir al día siguiente."
"""
import datetime

from asignacion_engine import dia_de_reparto
from stock_real import filas_manuales
from run_asignacion_stock_real import asignar_capon, asignar_chancha

DIAS_SEMANA = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes']


def dias_desde_hoy(hoy=None):
    """Días de reparto Lunes..Viernes desde el día de reparto de hoy en adelante (si hoy cae
    fuera de esa ventana — ya pasó la semana — devuelve la lista completa como respaldo)."""
    hoy = hoy or datetime.date.today()
    dia_inicio = DIAS_SEMANA[dia_de_reparto(hoy).weekday()] if dia_de_reparto(hoy).weekday() < 5 else 'Lunes'
    if dia_inicio not in DIAS_SEMANA:
        return list(DIAS_SEMANA)
    return DIAS_SEMANA[DIAS_SEMANA.index(dia_inicio):]


def fechas_reparto_semana(hoy=None):
    """[(dia_nombre, fecha_reparto), ...] — la fecha calendario real de cada día en
    dias_desde_hoy(), dentro de la semana actual (Lunes a Viernes). Necesaria para el filtro de
    "faena todavía no llegó" de asignar_semana (ver más abajo) — con solo el nombre del día no
    alcanza para comparar contra fecha_faena real."""
    hoy = hoy or datetime.date.today()
    inicio = dia_de_reparto(hoy)
    lunes_semana = inicio - datetime.timedelta(days=inicio.weekday())
    out = []
    for i, nombre in enumerate(DIAS_SEMANA):
        fecha = lunes_semana + datetime.timedelta(days=i)
        if fecha >= inicio:
            out.append((nombre, fecha))
    return out


def _normalizar_fecha(f):
    if hasattr(f, 'date') and callable(f.date):
        return f.date()
    return f


def construir_pool_semana(stock_rows, plan_rows, merc):
    """stock_rows: de read_stock_bd (+ filas_manuales 'PEND' si hace falta) — reales, es_real=True.
    plan_rows: de run_asignacion.read_plan_input (fecha/proveedor/merc/cant) — se convierten en
    filas sintéticas 'PROY' (es_real=False), fecha_faena = la del plan (futura respecto al
    stock real, así el FIFO las deja para después)."""
    real = [r for r in stock_rows if r['merc'] == merc]
    proy_entries = [{'proveedor': p['proveedor'], 'fecha_faena': p['fecha'], 'merc': merc, 'cantidad': p['cant']}
                     for p in plan_rows if p['merc'] == merc]
    proy = filas_manuales(proy_entries, prefijo='PROY', es_real=False)
    return real + proy


def asignar_semana(hist, pool_inicial, merc, dias_fechas, cuota_semana=None, asignar_fn=None):
    """Corre la cascada día por día sobre el MISMO pool, pasando el sobrante de un día al
    siguiente. dias_fechas: [(dia_nombre, fecha_reparto), ...] de fechas_reparto_semana().

    FIFO real, en dos capas (Tomás, 2026-08-05/06, después de ver que faena vieja quedaba sin
    salir dos veces seguidas):
    (1) Ningún día ve faena de un día FUTURO — dia_de_reparto(fecha_faena) tiene que ser <= la
        fecha de reparto de hoy, si no, queda en 'futuros' hasta que le toque (esto ya evita que
        el proyectado del jueves tape el cupo del martes).
    (2) DENTRO de lo elegible ese día, no alcanza con que el FIFO ordene los correlativos por
        fecha dentro de cada categoría — hay que agotar TODA la faena que ya estaba disponible
        de un día anterior ("vieja") antes de tocar la que recién se vuelve elegible HOY
        ("nueva"), aunque las categorías no calcen perfecto con lo que pide cada bloque ese día.
        Tomás: "el miércoles no quiero repartir faena de martes porque hay faena de lunes
        todavía." Se resuelve en dos pasadas: 1° el cupo completo del día contra 'vieja' sola;
        2° SOLO si esa pasada no alcanzó a cubrir el cupo total, el faltante (por bloque) se
        cubre con 'nueva' (cupo_override, ver run_asignacion_stock_real._asignar_generico).

    asignar_fn: función (hist, filas, dia=, cuota_semana=, dias_referencia=, cupo_override=) ->
    (shares, target_i, matrix, asignado, disp, asig, sobrante, cupo) — inyectable para poder
    reusar este mismo motor con otra especie (2026-08-06: bovino usa asignar_categoria de
    run_asignacion_stock_real_bovino.py). Si no se pasa, cae al selector viejo de porcino
    (capón/chancha por nombre de mercadería) para no romper las llamadas existentes.

    Arrastre de faltante entre días (Tomás, 2026-08-07): "si tengo un presupuesto de 950, no
    puede ser que el asignado sea menor y me quede un saldo remanente sin repartir... esto puede
    pasar el lunes que el stock es bajo, pero después en la semana se debe compensar." Cada día
    el cupo objetivo es el cupo base de ese día (histórico o cuota, MISMA forma de siempre) MÁS
    lo que quedó sin cubrir de días anteriores (deficit_acumulado) — así, si el stock alcanza
    más adelante en la semana, el faltante de un día de poco stock se termina cubriendo, y el
    total asignado de la semana converge al presupuesto (siempre que el stock total de la
    semana alcance) en vez de perderse."""
    if asignar_fn is None:
        asignar_fn = asignar_capon if merc == 'Capón' else asignar_chancha
    dias = [d for d, _f in dias_fechas]
    resultados = []
    pool_pendiente = list(pool_inicial)
    deficit_acumulado = {}
    for dia, fecha_reparto_dia in dias_fechas:
        vieja, nueva, futuros = [], [], []
        for f in pool_pendiente:
            reparto_f = dia_de_reparto(_normalizar_fecha(f['fecha_faena']))
            if reparto_f < fecha_reparto_dia:
                vieja.append(f)
            elif reparto_f == fecha_reparto_dia:
                nueva.append(f)
            else:
                futuros.append(f)

        cupo_base = {code: v for code, (_n, v) in hist.cupo_o_cuota(
            merc, dia=dia, cuota_semana=cuota_semana, dias_referencia=dias).items()}
        cupo_dia = {code: v + deficit_acumulado.get(code, 0.0) for code, v in cupo_base.items()}

        shares, target_i, matrix, asignado, disp1, asig1, sobrante_vieja, cupo = asignar_fn(
            hist, vieja, dia=dia, cupo_override=cupo_dia)

        total_cupo_dia = sum(cupo.values())
        total_asig_pase1 = sum(target_i.values())
        target_i2, asignado2, sobrante_nueva, disp2, asig2 = {}, {}, nueva, 0, 0
        if nueva and total_asig_pase1 < total_cupo_dia - 1e-6:
            cupo_restante = {code: max(round(cupo.get(code, 0)) - target_i.get(code, 0), 0) for code in cupo}
            _sh2, target_i2, _mx2, asignado2, disp2, asig2, sobrante_nueva, _cu2 = asignar_fn(
                hist, nueva, dia=dia, cupo_override=cupo_restante)

        target_total = {code: target_i.get(code, 0) + target_i2.get(code, 0) for code in cupo}
        asignado_total = {code: asignado.get(code, []) + asignado2.get(code, []) for code in shares}
        sobrante_total = sobrante_vieja + sobrante_nueva

        deficit_acumulado = {code: max(cupo_dia.get(code, 0.0) - target_total.get(code, 0), 0.0)
                              for code in cupo_dia}

        resultados.append({
            'dia': dia, 'shares': shares, 'presupuesto': cupo, 'asignado': target_total,
            'matrix': matrix, 'correlativos': asignado_total, 'disponible_dia': disp1 + disp2,
            'total_asignado_dia': asig1 + asig2, 'sobrante': sobrante_total,
        })
        pool_pendiente = sobrante_total + futuros
    return resultados, pool_pendiente
