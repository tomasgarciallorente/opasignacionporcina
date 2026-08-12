-- Esquema de la herramienta puntual de asignación porcina (Streamlit), separada de
-- crm-extendido pero en el MISMO proyecto Supabase (ubryvnogujmhfjkdvbor) para facilitar
-- la fusión más adelante. No toca stock_piezas ni ninguna tabla existente del CRM.
--
-- Aplicar una sola vez con scripts/apply_schema.py (no forma parte de las migraciones
-- versionadas de crm-extendido — es un esquema propio de esta app chica).

CREATE TABLE IF NOT EXISTS public.porcino_tipificacion_lotes (
  id                  BIGINT GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
  proveedor           TEXT NOT NULL,
  fecha_faena         DATE NOT NULL,
  mercaderia          TEXT NOT NULL CHECK (mercaderia IN ('Capón', 'Chancha')),
  correlativo_inicial BIGINT NOT NULL,
  cantidad            INTEGER NOT NULL CHECK (cantidad > 0),
  foto_url            TEXT,
  creado_por          TEXT,
  creado_at           TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS public.porcino_tipificacion_animales (
  id            BIGINT GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
  lote_id       BIGINT NOT NULL REFERENCES public.porcino_tipificacion_lotes(id) ON DELETE CASCADE,
  correlativo   BIGINT NOT NULL,
  kg            NUMERIC NOT NULL CHECK (kg > 0),
  nivel_grasa   TEXT NOT NULL CHECK (nivel_grasa IN ('-1', '0', '1', '1+', '1++', '2', '2+', '3', '3+', '4')),
  observacion   TEXT,  -- golpe en cadera / golpe interno / golpe interno feo / corte en paleta (cuero) / otro / NULL
  creado_at     TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  UNIQUE (lote_id, correlativo)
);

CREATE INDEX IF NOT EXISTS idx_tipif_animales_lote ON public.porcino_tipificacion_animales (lote_id);
CREATE INDEX IF NOT EXISTS idx_tipif_lotes_fecha ON public.porcino_tipificacion_lotes (fecha_faena);

-- Varias fotos por tropa (Tomás, 2026-08-11: "necesito poder cargar varias fotos de
-- romaneos") -- reemplaza el uso de la columna foto_url singular de porcino_tipificacion_lotes
-- (esa columna queda sin usar de acá en más, no se borra para no romper filas viejas).
CREATE TABLE IF NOT EXISTS public.porcino_tipificacion_fotos (
  id        BIGINT GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
  lote_id   BIGINT NOT NULL REFERENCES public.porcino_tipificacion_lotes(id) ON DELETE CASCADE,
  foto_url  TEXT NOT NULL,
  creado_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS idx_tipif_fotos_lote ON public.porcino_tipificacion_fotos (lote_id);
ALTER TABLE public.porcino_tipificacion_fotos ENABLE ROW LEVEL SECURITY;

-- Resultado de una corrida de "Generar reparto" — para poder repasar el historial de
-- propuestas sin volver a correr el motor, y como base para exportar el Excel.
CREATE TABLE IF NOT EXISTS public.porcino_reparto_resultados (
  id            BIGINT GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
  corrida_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  dia_reparto   TEXT NOT NULL,
  mercaderia    TEXT NOT NULL CHECK (mercaderia IN ('Capón', 'Chancha')),
  bloque_codigo TEXT NOT NULL,
  correlativo   BIGINT NOT NULL,
  kg            NUMERIC,
  nivel_grasa   TEXT,
  observacion   TEXT,
  reasignado_por_golpe_corte BOOLEAN NOT NULL DEFAULT false
);

CREATE INDEX IF NOT EXISTS idx_reparto_corrida ON public.porcino_reparto_resultados (corrida_at);

-- --- Snapshots subidos por Excel (página "Actualizar datos") ---
-- Se REEMPLAZAN enteros en cada subida (delete + insert, ver db.py) — no hay versionado,
-- siempre es "el último Excel que alguien subió". subido_at identifica el lote de la subida.
CREATE TABLE IF NOT EXISTS public.porcino_stock_snapshot (
  id            BIGINT GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
  correlativo   BIGINT NOT NULL,
  proveedor     TEXT,
  fecha_faena   DATE,
  kg            NUMERIC,
  mercaderia    TEXT NOT NULL CHECK (mercaderia IN ('Capón', 'Chancha')),
  nivel_grasa   TEXT,   -- puede venir vacío si el Excel todavía no lo tiene tipificado
  comprometido  BOOLEAN NOT NULL DEFAULT false,  -- ver comprometido() en stock_real.py
  subido_at     TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS idx_stock_snapshot_correlativo ON public.porcino_stock_snapshot (correlativo);

-- Histórico de cupo/calidad por bloque (de "Promedio diario Capón y Chancha por vendedor.xlsx")
CREATE TABLE IF NOT EXISTS public.porcino_historico_bloque (
  id            BIGINT GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
  bloque_codigo TEXT NOT NULL,
  bloque_nombre TEXT NOT NULL,
  mercaderia    TEXT NOT NULL CHECK (mercaderia IN ('Capón', 'Chancha')),
  dia           TEXT NOT NULL,   -- Lunes|Martes|Miércoles|Jueves|Viernes|Total
  cupo          NUMERIC NOT NULL,
  subido_at     TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS public.porcino_historico_calidad (
  id            BIGINT GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
  bloque_codigo TEXT NOT NULL,
  mercaderia    TEXT NOT NULL CHECK (mercaderia IN ('Capón', 'Chancha')),
  eje           TEXT NOT NULL CHECK (eje IN ('peso', 'tipif')),
  categoria     TEXT NOT NULL,
  pct           NUMERIC NOT NULL,
  subido_at     TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- RLS simple: esta app no usa Supabase Auth por usuario (login es una contraseña
-- compartida a nivel de la app Streamlit, ver plan) — el service role key de Streamlit
-- bypassea RLS igual, pero lo dejamos activado y sin policies de acceso público por las
-- dudas de que alguna vez se exponga la anon key.
ALTER TABLE public.porcino_tipificacion_lotes ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.porcino_tipificacion_animales ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.porcino_reparto_resultados ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.porcino_stock_snapshot ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.porcino_historico_bloque ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.porcino_historico_calidad ENABLE ROW LEVEL SECURITY;
