DROP VIEW IF EXISTS staging.discrepancias_servicios CASCADE;
DROP TABLE IF EXISTS staging.merge_asphia_portafolio;

CREATE SCHEMA IF NOT EXISTS staging;

WITH mapeo_semantico AS (
    SELECT * FROM (VALUES
        ('Activo',      'ACTIVO'),
        ('Activo',      'ACTIVO_CON_CAMBIO'),
        ('Suspendido',  'SUSPENDIDO'),
        ('Cancelado',   'CANCELADO'),
        ('En Mantenimiento', 'EN_MANTENIMIENTO'),
        ('Rollback',    'ROLL_BACK')
    ) AS m(estado_portaf, estado_asph_eq)
),

asphia_valida AS (
    SELECT *
    FROM clean.asphia_clean
    WHERE idservicio NOT LIKE 'SIN_ID_%'
),

portafolio_enriquecido AS (
    SELECT 
        p.id_servicio,
        p.cuenta_cliente,
        p.estado AS estado_portafolio,
        a.estado_ciclo AS estado_asphia,
        a.cambio,
        a.n_acceso,
        a.anillo,
        a.cliente AS cliente_asphia,
        a.historial_actividades,
        a.fecha AS fecha_ultimo_evento,
        CASE 
            WHEN a.idservicio IS NULL THEN 'SOLO_EN_PORTAFOLIO'
            WHEN p.estado = a.estado_ciclo THEN 'EXACTAMENTE_IGUAL'
            WHEN EXISTS (
                SELECT 1 FROM mapeo_semantico ms
                WHERE ms.estado_portaf = p.estado AND ms.estado_asph_eq = a.estado_ciclo
            ) THEN 'EQUIVALENTE_SEGUN_MAPEO'
            ELSE 'DIFERENTE_SIN_EQUIVALENCIA'
        END AS tipo_coincidencia
    FROM 
        raw.portafolio_origin p
    LEFT JOIN asphia_valida a ON p.id_servicio::TEXT = a.idservicio::TEXT
)

SELECT 
    id_servicio,
    cuenta_cliente,
    estado_portafolio AS estado,
    tipo_coincidencia,
    cambio,
    n_acceso,
    anillo,
    cliente_asphia AS cliente,
    historial_actividades,
    estado_asphia AS estado_ciclo,
    'NINGUNO' AS historial_descartados
INTO staging.merge_asphia_portafolio
FROM portafolio_enriquecido;

CREATE INDEX IF NOT EXISTS idx_merge_id ON staging.merge_asphia_portafolio(id_servicio);
CREATE INDEX IF NOT EXISTS idx_merge_tipo ON staging.merge_asphia_portafolio(tipo_coincidencia);
CREATE INDEX IF NOT EXISTS idx_merge_estado_portafolio ON staging.merge_asphia_portafolio(estado);
CREATE INDEX IF NOT EXISTS idx_merge_estado_asphia ON staging.merge_asphia_portafolio(estado_ciclo);

CREATE OR REPLACE VIEW staging.discrepancias_servicios AS
SELECT *
FROM staging.merge_asphia_portafolio
WHERE tipo_coincidencia NOT IN ('EXACTAMENTE_IGUAL', 'EQUIVALENTE_SEGUN_MAPEO');

CREATE OR REPLACE VIEW staging.asphia_sin_correspondencia AS
SELECT *
FROM clean.asphia_clean
WHERE idservicio LIKE 'SIN_ID_%';