DROP TABLE IF EXISTS clean.baseodf_clean CASCADE;
CREATE TABLE clean.baseodf_clean (
    fibra TEXT,
    equipo TEXT,
    port_ec TEXT,
    odf_troncal TEXT,
    idhilo TEXT,
    cap_troncal TEXT,
    anillo TEXT,
    central TEXT,
    informacion_adicional TEXT,
    fecha_actualizacion TIMESTAMP DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_clean_baseodf_fibra ON clean.baseodf_clean(fibra);
CREATE INDEX IF NOT EXISTS idx_clean_baseodf_equipo ON clean.baseodf_clean(equipo);
CREATE INDEX IF NOT EXISTS idx_clean_baseodf_port_ec ON clean.baseodf_clean(port_ec);
CREATE INDEX IF NOT EXISTS idx_clean_baseodf_odf_troncal ON clean.baseodf_clean(odf_troncal);
CREATE INDEX IF NOT EXISTS idx_clean_baseodf_idhilo ON clean.baseodf_clean(idhilo);
CREATE INDEX IF NOT EXISTS idx_clean_baseodf_anillo ON clean.baseodf_clean(anillo);
CREATE INDEX IF NOT EXISTS idx_clean_baseodf_central ON clean.baseodf_clean(central);

DROP TRIGGER IF EXISTS trg_clean_update_fecha_actualizacion ON clean.baseodf_clean;

CREATE OR REPLACE FUNCTION clean.update_fecha_actualizacion()
RETURNS TRIGGER AS $$
BEGIN
    NEW.fecha_actualizacion = NOW();
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

CREATE TRIGGER trg_clean_update_fecha_actualizacion
    BEFORE UPDATE ON clean.baseodf_clean
    FOR EACH ROW
    EXECUTE FUNCTION clean.update_fecha_actualizacion();

TRUNCATE TABLE clean.baseodf_clean;

INSERT INTO clean.baseodf_clean (
    fibra,
    equipo,
    port_ec,
    odf_troncal,
    idhilo,
    cap_troncal,
    anillo,
    central,
    informacion_adicional,
    fecha_actualizacion
)
SELECT 
    NULLIF(TRIM(num_troncal_actual), '') AS fibra,
    COALESCE(
        clean.extraer_equipo_central(puerto),
        clean.extraer_equipo_central(obshil_tro),
        clean.extraer_equipo_central(nem_anillo)
    ) AS equipo,
    COALESCE(
        clean.extraer_port_ec(puerto),
        clean.extraer_port_ec(obshil_tro),
        clean.extraer_port_ec(nem_anillo)
    ) AS port_ec,
    NULLIF(TRIM(odf), '') AS odf_troncal,
    NULLIF(NULLIF(TRIM(idhilo), '0'), '') AS idhilo,
    NULLIF(NULLIF(TRIM(cap_troncal), '0'), '') AS cap_troncal,
    NULLIF(TRIM(nem_anillo), '') AS anillo,
    NULLIF(TRIM(central), '') AS central,
    CONCAT_WS(' | ',
        CASE WHEN puerto IS NOT NULL AND puerto != '' AND puerto != '0' THEN puerto END,
        CASE WHEN tipo_anillo IS NOT NULL AND tipo_anillo != '' AND tipo_anillo != '0' THEN tipo_anillo END,
        CASE WHEN obshil_tro IS NOT NULL AND obshil_tro != '' AND obshil_tro != '-' AND obshil_tro != '0' THEN obshil_tro END
    ) AS informacion_adicional,
    NOW()
FROM raw.baseodf_origin
WHERE idhil_tro IS NOT NULL 
  AND idhil_tro != '' 
  AND idhil_tro != 'nan';