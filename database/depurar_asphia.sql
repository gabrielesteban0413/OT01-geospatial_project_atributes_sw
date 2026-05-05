-- Opcional: fijar el orden de búsqueda (raw para origen, clean para destino)
SET search_path TO raw, clean;

-- Eliminar tabla si existe (en el esquema clean)
DROP TABLE IF EXISTS clean.asphia_clean CASCADE;

-- Crear tabla limpia en el esquema clean
CREATE TABLE clean.asphia_clean AS
WITH 
ciclos_base AS (
    SELECT 
        n_acceso AS fibra,
        anillo,
        MIN(fecha) AS fecha_inicio,
        MAX(fecha) AS fecha_fin,
        MIN(idcliente) AS idcliente,
        MIN(diseno) AS diseno,
        MIN(version) AS version,
        MIN(n_hilo) AS n_hilo,
        MIN(cliente) AS cliente,
        MIN(direccion) AS direccion,
        MIN(servicio) AS servicio,
        MIN(l_cable_acceso) AS l_cable_acceso,
        MIN(consecutivo_cable_acceso) AS consecutivo_cable_acceso,
        MIN(disenador) AS disenador,
        MIN(cuadrilla) AS cuadrilla,
        MIN(ciudad) AS ciudad,
        MIN(bw) AS bw,
        MIN(und_bw) AS und_bw,
        MIN(equipo) AS equipo,
        MIN(n_hilo_troncal) AS n_hilo_troncal,
        MIN(n_troncal) AS n_troncal,
        MIN(l_cable_troncal) AS l_cable_troncal,
        MIN(odf) AS odf,
        COUNT(*) AS total_eventos,
        (ARRAY_AGG(cambio ORDER BY fecha DESC, idcliente DESC))[1] AS cambio,
        (ARRAY_AGG(observacion ORDER BY fecha DESC, idcliente DESC))[1] AS observacion,
        (ARRAY_AGG(actividad ORDER BY fecha DESC, idcliente DESC))[1] AS ultima_actividad,
        (ARRAY_AGG(idservicio ORDER BY fecha DESC, idcliente DESC))[1] AS idservicio
    FROM raw.asphia_origin
    WHERE idservicio IS NOT NULL 
      AND idservicio != 'None'
    GROUP BY n_acceso, anillo
),
 
ciclos_numerados AS (
    SELECT 
        *,
        ROW_NUMBER() OVER (ORDER BY fecha_inicio) AS ciclo_num
    FROM ciclos_base
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
      AND idservicio IS NOT NULL 
      AND idservicio != 'None'
    GROUP BY n_acceso, anillo
),

historial AS (
    SELECT 
        n_acceso AS fibra,
        anillo,
        STRING_AGG(
            actividad || ' (' || COALESCE(cambio, '') || ', ' || COALESCE(idservicio, '') || ')', 
            ' → ' 
            ORDER BY fecha, idcliente
        ) AS historial_actividades,
        STRING_AGG(COALESCE(observacion, ''), ' | ' ORDER BY fecha, idcliente) AS historial_observaciones
    FROM raw.asphia_origin
    WHERE idservicio IS NOT NULL 
      AND idservicio != 'None'
    GROUP BY n_acceso, anillo
),

estado_ciclos AS (
    SELECT 
        cn.fibra,
        cn.anillo,
        cn.ciclo_num,
        CASE 
            WHEN cn.total_eventos = 1 AND cn.ultima_actividad = 'INSTALACION' 
                THEN 'ACTIVO'
            WHEN cn.ultima_actividad IN ('RETIRO', 'DESPROGRAMACION') 
                THEN 'CANCELADO'
            WHEN cn.ultima_actividad = 'ROLL BACK' 
                THEN 'ROLL_BACK'
            ELSE 'ACTIVO +'
        END AS estado_ciclo
    FROM ciclos_numerados cn
)

SELECT 
    cn.idservicio,
    COALESCE(ie.instalacion_cambio, cn.cambio) AS cambio,
    cn.ultima_actividad AS actividad,
    cn.fibra AS n_acceso,
    cn.anillo,
    cn.cliente,
    cn.servicio,
    cn.equipo,
    cn.direccion,
    COALESCE(ie.instalacion_fecha, cn.fecha_inicio) AS fecha,
    cn.observacion,
    COALESCE(ie.instalacion_cuadrilla, cn.cuadrilla) AS cuadrilla,
    COALESCE(ie.instalacion_bw, cn.bw) AS bw,
    COALESCE(ie.instalacion_und_bw, cn.und_bw) AS und_bw,
    cn.n_troncal,
    COALESCE(ie.instalacion_n_hilo_troncal, cn.n_hilo_troncal) AS n_hilo_troncal,
    h.historial_observaciones,
    h.historial_actividades,
    ec.estado_ciclo,
    cn.ciclo_num AS ciclo,
    'SIN_CASOS_ATIPICOS' AS observacion_caso_atipico
FROM ciclos_numerados cn
INNER JOIN historial h 
    ON cn.fibra = h.fibra AND cn.anillo = h.anillo
INNER JOIN estado_ciclos ec 
    ON cn.fibra = ec.fibra AND cn.anillo = ec.anillo
LEFT JOIN instalacion_event ie
    ON cn.fibra = ie.fibra AND cn.anillo = ie.anillo
ORDER BY cn.fecha_inicio;

-- Índices sobre la tabla en esquema clean
CREATE INDEX IF NOT EXISTS idx_spc_fibra_anillo ON clean.asphia_clean(n_acceso, anillo);
CREATE INDEX IF NOT EXISTS idx_spc_idservicio ON clean.asphia_clean(idservicio);
CREATE INDEX IF NOT EXISTS idx_spc_cliente ON clean.asphia_clean(cliente);
CREATE INDEX IF NOT EXISTS idx_spc_estado ON clean.asphia_clean(estado_ciclo);
CREATE INDEX IF NOT EXISTS idx_spc_fecha ON clean.asphia_clean(fecha);

COMMENT ON TABLE clean.asphia_clean IS 
    'Tabla depurada de servicios. Cada registro representa un ciclo independiente
     identificado por fibra (n_acceso) + anillo. Las columnas cambio, fecha,
     cuadrilla, bw, und_bw y n_hilo_troncal corresponden al evento de INSTALACION de ese ciclo.';

COMMENT ON COLUMN clean.asphia_clean.estado_ciclo IS 
    'Estado final del ciclo: ACTIVO_SIN_CAMBIOS (solo instalación), 
     ACTIVO_CON_CAMBIOS (con adiciones/ampliaciones), 
     FINALIZADO (retirado o desprogramado),
     ROLL_BACK_AL_FINAL (terminó en rollback - requiere revisión)';

COMMENT ON COLUMN clean.asphia_clean.observacion_caso_atipico IS 
    'Campo reservado para detección de casos atípicos.';