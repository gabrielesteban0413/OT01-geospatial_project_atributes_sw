DROP TABLE IF EXISTS clean.asphia_clean CASCADE;
CREATE SCHEMA IF NOT EXISTS clean;

CREATE TABLE clean.asphia_clean AS
WITH 
grupos AS (
    SELECT 
        n_acceso AS fibra,
        anillo,
        MIN(fecha) AS fecha_inicio,
        MAX(fecha) AS fecha_fin,
        MIN(cliente) AS cliente,
        MIN(direccion) AS direccion,
        MIN(servicio) AS servicio,
        MIN(equipo) AS equipo,
        MIN(cuadrilla) AS cuadrilla,
        COUNT(*) AS total_eventos,
        COALESCE(
            MIN(CASE WHEN idservicio IS NOT NULL 
                      AND idservicio != 'None' 
                      AND idservicio != '' 
                 THEN idservicio END),
            'SIN_ID_' || REPLACE(n_acceso, ' ', '') || '_' || REPLACE(anillo, ' ', '')
        ) AS idservicio_clean,
        (ARRAY_AGG(actividad ORDER BY fecha DESC, idcliente DESC))[1] AS ultima_actividad,
        (ARRAY_AGG(cambio ORDER BY fecha DESC, idcliente DESC))[1] AS ultimo_cambio,
        (ARRAY_AGG(observacion ORDER BY fecha DESC, idcliente DESC))[1] AS ultima_observacion,
        MAX(fecha) AS ultima_fecha
    FROM raw.asphia_origin
    WHERE n_acceso IS NOT NULL AND anillo IS NOT NULL
    GROUP BY n_acceso, anillo
),

historial AS (
    SELECT 
        n_acceso AS fibra,
        anillo,
        STRING_AGG(
            COALESCE(actividad, '?') || 
            CASE WHEN cambio IS NOT NULL AND cambio != '' THEN ' (' || cambio || ')' ELSE '' END,
            ' → ' ORDER BY fecha, idcliente
        ) AS historial_actividades,
        STRING_AGG(COALESCE(observacion, ''), ' | ' ORDER BY fecha, idcliente) AS historial_observaciones
    FROM raw.asphia_origin
    WHERE n_acceso IS NOT NULL AND anillo IS NOT NULL
    GROUP BY n_acceso, anillo
),

estado_ciclos AS (
    SELECT 
        g.fibra,
        g.anillo,
        CASE 
            WHEN g.ultima_actividad IN ('RETIRO', 'DESPROGRAMACION') THEN 'CANCELADO'
            WHEN g.ultima_actividad = 'ROLL BACK' THEN 'ROLL_BACK'
            WHEN g.ultima_actividad = 'SUSPENSION' THEN 'SUSPENDIDO'
            WHEN g.ultima_actividad = 'REACTIVACION' THEN 'ACTIVO'
            WHEN g.ultima_actividad LIKE 'CAMBIO%' THEN 'ACTIVO_CON_CAMBIO'
            WHEN g.ultima_actividad = 'MANTENIMIENTO' THEN 'EN_MANTENIMIENTO'
            WHEN g.ultima_actividad = 'INSTALACION' AND g.total_eventos = 1 THEN 'ACTIVO'
            WHEN g.ultima_actividad = 'INSTALACION' AND g.total_eventos > 1 THEN 'ACTIVO_CON_CAMBIO'
            ELSE 'ACTIVO'
        END AS estado_ciclo
    FROM grupos g
),

instalacion_event AS (
    SELECT 
        n_acceso AS fibra,
        anillo,
        MIN(fecha) AS instalacion_fecha,
        (ARRAY_AGG(cambio ORDER BY fecha))[1] AS instalacion_cambio,
        (ARRAY_AGG(cuadrilla ORDER BY fecha))[1] AS instalacion_cuadrilla,
        (ARRAY_AGG(bw ORDER BY fecha))[1] AS instalacion_bw,
        (ARRAY_AGG(und_bw ORDER BY fecha))[1] AS instalacion_und_bw,
        (ARRAY_AGG(n_hilo_troncal ORDER BY fecha))[1] AS instalacion_n_hilo_troncal
    FROM raw.asphia_origin
    WHERE actividad = 'INSTALACION'
      AND n_acceso IS NOT NULL AND anillo IS NOT NULL
    GROUP BY n_acceso, anillo
)

SELECT 
    g.idservicio_clean AS idservicio,
    COALESCE(ie.instalacion_cambio, g.ultimo_cambio) AS cambio,
    g.ultima_actividad AS actividad,
    g.fibra AS n_acceso,
    g.anillo,
    g.cliente,
    g.servicio,
    g.equipo,
    g.direccion,
    COALESCE(ie.instalacion_fecha, g.ultima_fecha) AS fecha,
    g.ultima_observacion AS observacion,
    COALESCE(ie.instalacion_cuadrilla, g.cuadrilla) AS cuadrilla,
    ie.instalacion_bw AS bw,
    ie.instalacion_und_bw AS und_bw,
    NULL AS n_troncal,
    ie.instalacion_n_hilo_troncal AS n_hilo_troncal,
    h.historial_observaciones,
    h.historial_actividades,
    ec.estado_ciclo,
    1 AS ciclo,
    CASE 
        WHEN ie.instalacion_fecha IS NULL THEN 'SIN_EVENTO_INSTALACION'
        WHEN g.idservicio_clean LIKE 'SIN_ID_%' THEN 'IDSERVICIO_INFERIDO'
        ELSE 'OK'
    END AS observacion_caso_atipico
FROM grupos g
LEFT JOIN historial h ON g.fibra = h.fibra AND g.anillo = h.anillo
LEFT JOIN estado_ciclos ec ON g.fibra = ec.fibra AND g.anillo = ec.anillo
LEFT JOIN instalacion_event ie ON g.fibra = ie.fibra AND g.anillo = ie.anillo
ORDER BY g.fecha_inicio;

CREATE INDEX IF NOT EXISTS idx_asphia_clean_idservicio ON clean.asphia_clean(idservicio);
CREATE INDEX IF NOT EXISTS idx_asphia_clean_n_acceso_anillo ON clean.asphia_clean(n_acceso, anillo);
CREATE INDEX IF NOT EXISTS idx_asphia_clean_estado ON clean.asphia_clean(estado_ciclo);
CREATE INDEX IF NOT EXISTS idx_asphia_clean_cliente ON clean.asphia_clean(cliente);