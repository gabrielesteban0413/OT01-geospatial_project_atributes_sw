DROP TABLE IF EXISTS staging.mg_centrales CASCADE;
DROP FUNCTION IF EXISTS staging.refrescar_mg_centrales() CASCADE;

CREATE TABLE staging.mg_centrales (
    "equipo central" TEXT,
    puerto           TEXT,
    odf_patcheo      TEXT,
    cassetera        TEXT,
    port_pacheo      TEXT,
    fibra            TEXT,
    odf_troncal      TEXT,
    hilo             TEXT,
    mg_anillo        TEXT,
    status           TEXT,
    etiqueta_nce     TEXT,
    acceso           TEXT,
    corporativo_estatus     TEXT,
    corporativo_obs         TEXT,
    corporativo_idservicio  TEXT,
    observacion      TEXT,
    comments         TEXT
);

CREATE OR REPLACE FUNCTION staging.refrescar_mg_centrales()
RETURNS void LANGUAGE plpgsql AS $$
DECLARE
    sql_final TEXT;
BEGIN
    sql_final := $q$
        WITH base AS (
            -- NCE
            SELECT
                "equipo central"::text AS "equipo central",
                'NCE'::text            AS origen,
                "port_ec"::text        AS puerto,
                "etiqueta"::text       AS anillo,
                NULL::text             AS pacheo,
                COALESCE("administrative_status", '') ||
                    CASE
                        WHEN "operational_status" IS NOT NULL
                             AND "administrative_status" IS NOT NULL
                        THEN ' - '
                        ELSE ''
                    END ||
                    COALESCE("operational_status", '') AS status,
                (regexp_match("etiqueta", '_([0-9]+)_'))[1]::text AS fibra,
                NULL::text             AS odf_troncal,
                (regexp_match("etiqueta", '[Hh]([0-9]+)_'))[1]::text AS hilo,
                NULL::text             AS observacion,
                "etiqueta"::text       AS etiqueta_nce,
                NULL::text             AS puerto_baseodf
            FROM raw.r_nce
            WHERE "equipo central" IS NOT NULL

            UNION ALL

            -- VIAS
            SELECT
                "equipo central"::text AS "equipo central",
                'VIAS'::text           AS origen,
                "port_ec"::text        AS puerto,
                "anillo"::text         AS anillo,
                "pacheo"::text         AS pacheo,
                NULL::text             AS status,
                "troncal_cable"::text  AS fibra,
                "odf_calle"::text      AS odf_troncal,
                "hilo"::text           AS hilo,
                "observacion"::text    AS observacion,
                NULL::text             AS etiqueta_nce,
                NULL::text             AS puerto_baseodf
            FROM raw.r_vias
            WHERE "equipo central" IS NOT NULL

            UNION ALL

            -- BASEODF
            SELECT
                NULL::text             AS "equipo central",
                'BASEODF'::text        AS origen,
                NULL::text             AS puerto,
                NULL::text             AS anillo,
                NULL::text             AS pacheo,
                NULL::text             AS status,
                "num_troncal_actual"::text AS fibra,
                "odf"::text            AS odf_troncal,
                "idhilo"::text         AS hilo,
                "obshil_tro"::text     AS observacion,
                NULL::text             AS etiqueta_nce,
                "puerto"::text         AS puerto_baseodf
            FROM raw.r_baseodf
        ),
        por_puerto AS (
            SELECT
                "equipo central",
                puerto,
                array_agg(DISTINCT anillo) FILTER (WHERE anillo IS NOT NULL) AS anillos,
                array_agg(DISTINCT pacheo) FILTER (WHERE pacheo IS NOT NULL) AS pacheos,
                array_to_string(array_agg(DISTINCT status)       FILTER (WHERE status IS NOT NULL AND status <> ''), ' | ') AS status,
                array_to_string(array_agg(DISTINCT fibra)        FILTER (WHERE fibra IS NOT NULL AND fibra <> ''),   ' | ') AS fibra,
                array_to_string(array_agg(DISTINCT odf_troncal)  FILTER (WHERE odf_troncal IS NOT NULL AND origen <> 'BASEODF'), ' | ') AS odf_nce_vias,
                array_to_string(array_agg(DISTINCT hilo)         FILTER (WHERE hilo IS NOT NULL AND hilo <> ''),    ' | ') AS hilo,
                array_to_string(array_agg(DISTINCT observacion)  FILTER (WHERE observacion IS NOT NULL AND observacion <> '' AND origen <> 'BASEODF'), ' | ') AS observacion_otras,
                array_to_string(array_agg(DISTINCT etiqueta_nce) FILTER (WHERE etiqueta_nce IS NOT NULL), ' | ') AS etiqueta_nce,
                array_to_string(array_agg(DISTINCT origen || ' °° ' || anillo) FILTER (WHERE anillo IS NOT NULL), ' | ') AS anillos_str,
                COUNT(DISTINCT origen) AS n_origenes
            FROM base
            WHERE puerto IS NOT NULL
            GROUP BY "equipo central", puerto
        ),
        por_fibra_hilo AS (
            SELECT
                fibra,
                hilo,
                array_to_string(array_agg(DISTINCT odf_troncal)   FILTER (WHERE odf_troncal IS NOT NULL), ' | ') AS odf_troncal,
                array_to_string(array_agg(DISTINCT puerto_baseodf) FILTER (WHERE puerto_baseodf IS NOT NULL), ' | ') AS puerto_baseodf
            FROM base
            WHERE origen = 'BASEODF'
              AND fibra IS NOT NULL
              AND hilo IS NOT NULL
            GROUP BY fibra, hilo
        ),
        asphia_raw AS (
            SELECT
                "n_acceso"::text       AS acceso,
                "n_hilo_troncal"::text AS hilo,
                split_part("n_acceso"::text, '-', 1) AS fibra
            FROM raw.r_asphia
            WHERE "n_acceso" IS NOT NULL
              AND "n_hilo_troncal" IS NOT NULL
              AND "n_hilo_troncal"::text <> '0'
        ),
        -- CORPORATIVO: columnas reales (troncal, hilo_troncal, estatus, observaciones, idservicio)
        corporativo_raw AS (
            SELECT
                "hilo_troncal"::text AS hilo,
                split_part("troncal"::text, '-', 1) AS fibra,
                "estatus"::text       AS estatus,
                "observaciones"::text AS observaciones,
                "idservicio"::text    AS idservicio
            FROM raw.r_clientes_corp
            WHERE "hilo_troncal" IS NOT NULL
              AND "troncal" IS NOT NULL
        ),
        coincidencia AS (
            SELECT
                a."equipo central",
                a.puerto,
                (
                    SELECT CASE WHEN length(x) <= length(y) THEN x ELSE y END
                    FROM unnest(a.anillos) x, unnest(a.anillos) y
                    WHERE x <> y
                      AND (
                          x ILIKE '%' || y || '%'
                          OR y ILIKE '%' || x || '%'
                          OR y ILIKE '%' ||
                             (regexp_split_to_array(x, '_'))
                                [array_length(regexp_split_to_array(x,'_'),1)] || '%'
                          OR x ILIKE '%' ||
                             (regexp_split_to_array(y, '_'))
                                [array_length(regexp_split_to_array(y,'_'),1)] || '%'
                      )
                    LIMIT 1
                ) AS anillo_corto
            FROM por_puerto a
        ),
        pacheo_desglosado AS (
            SELECT
                a."equipo central",
                a.puerto,
                NULLIF((regexp_match(a.pacheos[1], '^\s*([0-9]+)\s*/'))[1], '') AS odf_patcheo,
                NULLIF((regexp_match(a.pacheos[1], '([A-Za-z])\s*$'))[1], '')    AS cassetera,
                CASE
                    WHEN a.pacheos[1] ~* 'POS\s+[0-9]+\s*\.\s*[0-9]+'
                    THEN (regexp_match(a.pacheos[1], 'POS\s+([0-9]+)\s*\.\s*([0-9]+)'))[1]
                         || '+' ||
                         (regexp_match(a.pacheos[1], 'POS\s+([0-9]+)\s*\.\s*([0-9]+)'))[2]
                    WHEN a.pacheos[1] ~ '\.[0-9]+\s*/\s*[0-9]+'
                    THEN (regexp_match(a.pacheos[1], '\.([0-9]+)\s*/\s*([0-9]+)'))[1]
                         || '+' ||
                         (regexp_match(a.pacheos[1], '\.([0-9]+)\s*/\s*([0-9]+)'))[2]
                    WHEN a.pacheos[1] ~ '\.[0-9]+'
                    THEN (regexp_match(a.pacheos[1], '\.([0-9]+)'))[1]
                    WHEN a.pacheos[1] ~* 'POS\s+[0-9]+'
                    THEN (regexp_match(a.pacheos[1], 'POS\s+([0-9]+)'))[1]
                    ELSE NULL
                END AS port_pacheo
            FROM por_puerto a
            WHERE a.pacheos IS NOT NULL
        ),
        asphia_match AS (
            SELECT
                b."equipo central",
                b.puerto,
                COUNT(DISTINCT a.acceso) AS n_accesos,
                MIN(a.acceso)            AS unico_acceso
            FROM por_puerto b
            LEFT JOIN asphia_raw a
                   ON a.hilo  = b.hilo
                  AND a.fibra = b.fibra
            GROUP BY b."equipo central", b.puerto
        ),
        corporativo_match AS (
            SELECT
                c.fibra,
                c.hilo,
                COUNT(DISTINCT COALESCE(c.idservicio,'') || '|' || COALESCE(c.estatus,'') || '|' || COALESCE(c.observaciones,'')) AS n_registros,
                array_to_string(array_agg(DISTINCT c.estatus)       FILTER (WHERE c.estatus IS NOT NULL AND c.estatus <> ''), ' | ') AS estatus,
                array_to_string(array_agg(DISTINCT c.observaciones) FILTER (WHERE c.observaciones IS NOT NULL AND c.observaciones <> ''), ' | ') AS observaciones,
                array_to_string(array_agg(DISTINCT c.idservicio)    FILTER (WHERE c.idservicio IS NOT NULL AND c.idservicio <> ''), ' | ') AS idservicio
            FROM corporativo_raw c
            GROUP BY c.fibra, c.hilo
        )
        SELECT
            b."equipo central",
            b.puerto,
            p.odf_patcheo,
            p.cassetera,
            p.port_pacheo,
            b.fibra,
            COALESCE(b.odf_nce_vias, '') ||
                CASE
                    WHEN f.odf_troncal IS NOT NULL
                         AND COALESCE(b.odf_nce_vias, '') <> ''
                    THEN ' |ODFS(' || f.odf_troncal || ')'
                    WHEN f.odf_troncal IS NOT NULL
                    THEN '|ODFS(' || f.odf_troncal || ')'
                    ELSE ''
                END AS odf_troncal,
            b.hilo,
            CASE
                WHEN c.anillo_corto IS NOT NULL
                    THEN c.anillo_corto
                WHEN b.anillos_str IS NULL
                    THEN '- NCE'
                ELSE
                    '- ' || b.anillos_str
            END AS "mg_anillo",
            b.status,
            b.etiqueta_nce,
            CASE
                WHEN am.n_accesos = 1 THEN am.unico_acceso
                WHEN am.n_accesos > 1 THEN 'MULTIPLES(' || am.n_accesos || ')'
                ELSE NULL
            END AS acceso,
            cm.estatus       AS corporativo_estatus,
            cm.observaciones AS corporativo_obs,
            cm.idservicio    AS corporativo_idservicio,
            COALESCE(b.observacion_otras, '') ||
                CASE
                    WHEN f.puerto_baseodf IS NOT NULL
                         AND COALESCE(b.observacion_otras, '') <> ''
                    THEN ' (' || f.puerto_baseodf || ')'
                    WHEN f.puerto_baseodf IS NOT NULL
                    THEN '(' || f.puerto_baseodf || ')'
                    ELSE ''
                END AS observacion,
            '...' AS comments
        FROM por_puerto b
        LEFT JOIN coincidencia c USING ("equipo central", puerto)
        LEFT JOIN pacheo_desglosado p USING ("equipo central", puerto)
        LEFT JOIN por_fibra_hilo f
               ON f.fibra = b.fibra
              AND f.hilo  = b.hilo
        LEFT JOIN asphia_match am USING ("equipo central", puerto)
        LEFT JOIN corporativo_match cm
               ON cm.fibra = b.fibra
              AND cm.hilo  = b.hilo
    $q$;

    DROP TABLE IF EXISTS staging.mg_centrales_data CASCADE;
    EXECUTE 'CREATE TABLE staging.mg_centrales_data AS ' || sql_final;

    TRUNCATE staging.mg_centrales;
    INSERT INTO staging.mg_centrales
    SELECT * FROM staging.mg_centrales_data;

    DROP TABLE staging.mg_centrales_data;

    CREATE INDEX IF NOT EXISTS idx_mgc_equipo
        ON staging.mg_centrales ("equipo central");
    CREATE INDEX IF NOT EXISTS idx_mgc_puerto
        ON staging.mg_centrales (puerto);
    BEGIN
        EXECUTE 'CREATE INDEX IF NOT EXISTS idx_mgc_equipo_trgm
                 ON staging.mg_centrales
                 USING gin ("equipo central" gin_trgm_ops)';
    EXCEPTION WHEN undefined_object THEN
        RAISE NOTICE 'pg_trgm no disponible, se omite índice trigram';
    END;

    ANALYZE staging.mg_centrales;
END;
$$;

SELECT staging.refrescar_mg_centrales();