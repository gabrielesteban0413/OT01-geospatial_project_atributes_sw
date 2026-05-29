DROP VIEW IF EXISTS staging.discrepancias_servicios CASCADE;
DROP TABLE IF EXISTS staging.merge_asphia_portafolio CASCADE;

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

DROP TABLE IF EXISTS staging.merge_corporativo_sync CASCADE;

CREATE TABLE staging.merge_corporativo_sync AS
WITH 
corporativo_clean AS (
    SELECT 
        cable_acceso,
        idservicio AS idservicio_corp,
        anillo AS anillo_corp,
        responsable_red_externa,
        incorporacion_cable,
        TRIM(idservicio) AS idservicio_trim,
        SPLIT_PART(TRIM(cable_acceso), '/', 1) AS cable_principal,
        TRIM(anillo) AS anillo_trim
    FROM raw.corporativo_origin
),

merge_ref AS (
    SELECT 
        id_servicio,
        cuenta_cliente,
        estado,
        tipo_coincidencia,
        cambio,
        n_acceso,
        anillo,
        cliente,
        historial_actividades,
        estado_ciclo
    FROM staging.merge_asphia_portafolio
),

unificados AS (
    SELECT 
        COALESCE(m.id_servicio, c.idservicio_trim) AS id_servicio,
        m.cuenta_cliente,
        m.estado,
        m.tipo_coincidencia,
        m.cambio,
        m.n_acceso AS cable_merge,
        m.anillo AS anillo_merge,
        m.cliente,
        m.historial_actividades,
        m.estado_ciclo,
        c.cable_acceso AS cable_corp,
        c.anillo_corp AS anillo_corp,
        c.responsable_red_externa,
        c.incorporacion_cable,
        CASE 
            WHEN m.id_servicio IS NOT NULL AND c.idservicio_trim IS NOT NULL THEN 'AMBOS'
            WHEN m.id_servicio IS NOT NULL THEN 'SOLO_MERGE'
            ELSE 'SOLO_CORPORATIVO'
        END AS origen
    FROM merge_ref m
    FULL OUTER JOIN corporativo_clean c ON m.id_servicio = c.idservicio_trim
)

SELECT 
    CASE 
        WHEN cable_merge IS NOT NULL AND cable_corp IS NOT NULL AND cable_merge != cable_corp 
            THEN cable_merge || '|' || cable_corp
        WHEN cable_merge IS NOT NULL THEN cable_merge
        ELSE cable_corp
    END AS cable,
    
    CASE 
        WHEN anillo_merge IS NOT NULL AND anillo_corp IS NOT NULL AND anillo_merge != anillo_corp 
            THEN anillo_merge || '|' || anillo_corp
        WHEN anillo_merge IS NOT NULL THEN anillo_merge
        ELSE anillo_corp
    END AS anillo,
    
    id_servicio,
    cuenta_cliente,
    estado,
    tipo_coincidencia,
    cambio,
    cliente,
    historial_actividades,
    estado_ciclo,
    responsable_red_externa,
    incorporacion_cable,
    origen,
    CASE 
        WHEN estado_ciclo IN ('ACTIVO', 'ACTIVO_CON_CAMBIO') THEN 'ACTIVO'
        WHEN estado_ciclo IN ('CANCELADO', 'SUSPENDIDO', 'ROLL_BACK', 'EN_MANTENIMIENTO') THEN 'INACTIVO'
        ELSE 'SIN_INFO'
    END AS estado_final
FROM unificados
ORDER BY id_servicio;

CREATE INDEX IF NOT EXISTS idx_corp_sync_cable ON staging.merge_corporativo_sync(cable);
CREATE INDEX IF NOT EXISTS idx_corp_sync_idservicio ON staging.merge_corporativo_sync(id_servicio);
CREATE INDEX IF NOT EXISTS idx_corp_sync_estado ON staging.merge_corporativo_sync(estado_final);