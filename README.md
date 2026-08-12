# Asignación porcina — app puntual

Herramienta chica y separada de `crm-extendido` para que los tipificadores (Ignacio Beas,
Gonzalo Maldonado) carguen la tipificación de cada tropa desde el celular y se pueda generar
el reparto por bloques al toque — mismo motor validado en `Análisis de entregas/Asignación de
mercadería/` (asignacion_engine.py, run_asignacion_stock_real.py), sin reescribirlo.

Ver el plan completo en `C:\Users\Gtecomercial\.claude\plans\moonlit-prancing-willow.md`.

## Probar ahora mismo, en esta máquina

Ya está corriendo local. Abrí **http://localhost:8765** en el navegador de esta computadora
(contraseña en `.streamlit/secrets.toml`, campo `APP_PASSWORD` — cambiala antes de compartir
el link con nadie).

Si se cerró, volver a levantarla:
```
cd "Análisis de entregas/Asignación de mercadería/app_web"
python -m streamlit run app.py
```

## Para que Ignacio y Gonzalo entren desde el celular (deploy real)

Necesita 2 cosas que solo vos podés hacer (son tu cuenta, no la mía):

1. **Un repo de GitHub** con el contenido de esta carpeta (`app_web/` — el `.gitignore` ya
   excluye `secrets.toml`, no se sube nunca). Si querés que yo lo arme, decime y lo hago con
   `git`/`gh` una vez que tengas `gh auth login` corrido en esta máquina — hoy no está
   instalado/logueado.
2. **Conectar ese repo a Streamlit Community Cloud** (share.streamlit.io, gratis, entrás con
   tu cuenta de GitHub) — apunta a `app_web/app.py` como archivo principal. En "Settings →
   Secrets" del panel de Streamlit Cloud, pegás el mismo contenido de
   `.streamlit/secrets.toml` (SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY, APP_PASSWORD).

Una vez desplegada, te da una URL pública (`https://<algo>.streamlit.app`) — instalable como
acceso directo en el celular (no hace falta App Store).

## Qué hace cada página

- **Cargar tipificación**: el tipificador arma una tropa (proveedor, fecha, mercadería,
  correlativo inicial, cantidad) y carga Kg + grado (botonera `-1` a `4`) + observación
  (golpe/corte) por animal. Foto de respaldo opcional.
- **Generar reparto**: cupo semanal por bloque (editable) + botón que corre el motor completo
  (cascada volumen→peso→tipificación, reasignación de golpes/cortes por afinidad a magros) —
  muestra el resultado en pantalla y permite descargar el mismo Excel de siempre.
- **Actualizar datos**: subir el Excel de stock (STOCK (BD)) y el histórico de cupo/calidad
  por bloque — reemplazan el snapshot anterior entero. Cualquiera con esos archivos lo puede
  actualizar, no hace falta que yo corra nada.

## Simplificaciones deliberadas de esta v1 (ver plan para el detalle)
- Sin RBAC real — una sola contraseña compartida.
- El filtro de "stock comprometido" en el Excel subido usa solo la señal de Observaciones
  ('SALE ...'), no el cruce contra la lista de proveedores reales (sí lo hace la versión de
  escritorio, `stock_real.py`).
- Bovino: no está, mismo patrón para más adelante.
