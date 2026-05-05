DROP VIEW IF EXISTS staging.discrepancias_servicios CASCADE;
DROP TABLE IF EXISTS staging.servicios_conciliados;

CREATE SCHEMA IF NOT EXISTS staging;

WITH mapeo_semantico AS (
    SELECT * FROM (VALUES
        ('Activo',      'ACTIVO'),
        ('Activo',      'ACTIVO +'),
        ('Suspendido',  'CANCELADO'),
        ('Cancelado',   'CANCELADO')
    ) AS m(estado_portaf, estado_asph_eq)
),

asphia_ordenado AS (
    SELECT 
        idservicio,
        cambio,
        n_acceso,
        anillo,
        cliente,
        historial_actividades,
        estado_ciclo,
        fecha,
        CASE 
            WHEN estado_ciclo IN ('ACTIVO', 'ACTIVO +') THEN 1
            ELSE 2
        END AS prioridad,
        ctid AS row_id
    FROM clean.asphia_clean
),

mejor_registro AS (
    SELECT DISTINCT ON (idservicio)
        idservicio,
        cambio,
        n_acceso,
        anillo,
        cliente,
        historial_actividades,
        estado_ciclo,
        fecha,
        prioridad,
        row_id
    FROM asphia_ordenado
    ORDER BY idservicio, prioridad ASC, fecha DESC
),

historial_descartados AS (
    SELECT 
        a.idservicio,
        STRING_AGG(
            'fecha=' || a.fecha::TEXT || 
            ', estado=' || COALESCE(a.estado_ciclo, 'NULL') ||
            ', n_acceso=' || COALESCE(a.n_acceso, 'NULL') ||
            ', anillo=' || COALESCE(a.anillo, 'NULL'),
            ' | ' ORDER BY a.fecha DESC
        ) AS otros_registros
    FROM asphia_ordenado a
    LEFT JOIN mejor_registro m ON a.idservicio = m.idservicio AND a.row_id = m.row_id
    WHERE m.idservicio IS NULL
    GROUP BY a.idservicio
),

portafolio_enriquecido AS (
    SELECT 
        p.id_servicio,
        p.cuenta_cliente,
        p.estado AS estado_portafolio,
        m.estado_ciclo AS estado_asphia,
        m.cambio,
        m.n_acceso,
        m.anillo,
        m.cliente AS cliente_asphia,
        m.historial_actividades,
        m.fecha AS fecha_ultimo_evento,
        h.otros_registros,
        CASE 
            WHEN m.idservicio IS NULL THEN 'SOLO_EN_PORTAFOLIO'
            WHEN p.estado = m.estado_ciclo THEN 'EXACTAMENTE_IGUAL'
            WHEN EXISTS (
                SELECT 1 FROM mapeo_semantico ms
                WHERE ms.estado_portaf = p.estado AND ms.estado_asph_eq = m.estado_ciclo
            ) THEN 'EQUIVALENTE_SEGUN_MAPEO'
            ELSE 'DIFERENTE_SIN_EQUIVALENCIA'
        END AS tipo_coincidencia
    FROM 
        raw.portafolio_superior p
    LEFT JOIN mejor_registro m ON p.id_servicio::TEXT = m.idservicio::TEXT
    LEFT JOIN historial_descartados h ON p.id_servicio::TEXT = h.idservicio
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
    COALESCE(otros_registros, 'NINGUNO') AS historial_descartados
INTO staging.servicios_conciliados
FROM portafolio_enriquecido;

CREATE INDEX IF NOT EXISTS idx_sc_id ON staging.servicios_conciliados(id_servicio);
CREATE INDEX IF NOT EXISTS idx_sc_tipo ON staging.servicios_conciliados(tipo_coincidencia);
CREATE INDEX IF NOT EXISTS idx_sc_estado_portafolio ON staging.servicios_conciliados(estado);
CREATE INDEX IF NOT EXISTS idx_sc_estado_asphia ON staging.servicios_conciliados(estado_ciclo);

CREATE OR REPLACE VIEW staging.discrepancias_servicios AS
SELECT *
FROM staging.servicios_conciliados
WHERE tipo_coincidencia NOT IN ('EXACTAMENTE_IGUAL', 'EQUIVALENTE_SEGUN_MAPEO');