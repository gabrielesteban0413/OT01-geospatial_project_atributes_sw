-- 1. Ver los primeros 10 servicios
SELECT * FROM servicios_depurados LIMIT 10;

-- 2. Ver el caso específico CCCH00013942
SELECT * FROM servicios_depurados WHERE idservicio = 'CCCH00013942';

-- 3. Ver servicios con casos atípicos
SELECT idservicio, actividad_final, total_eventos, observacion_caso_atipico 
FROM servicios_depurados 
WHERE observacion_caso_atipico IS NOT NULL 
  AND observacion_caso_atipico != ''
  AND observacion_caso_atipico != ' '
LIMIT 20;

-- 4. Estadísticas de casos atípicos
SELECT 
    CASE 
        WHEN observacion_caso_atipico LIKE '%ROLL_BACK%' THEN 'ROLL_BACK'
        WHEN observacion_caso_atipico LIKE '%MULTIPLES_INSTALACIONES%' THEN 'MULTIPLES_INSTALACIONES'
        WHEN observacion_caso_atipico LIKE '%FIBRA_CAMBIADA%' THEN 'FIBRA_CAMBIADA'
        WHEN observacion_caso_atipico LIKE '%INSTALACION_DESPUES%' THEN 'REACTIVACION'
        WHEN observacion_caso_atipico LIKE '%RETIRO_SIN_INSTALACION%' THEN 'RETIRO_SIN_INSTALACION'
        WHEN observacion_caso_atipico LIKE '%ADICION_SIN_INSTALACION%' THEN 'ADICION_SIN_INSTALACION'
        ELSE 'SIN_CASOS'
    END as tipo_caso,
    COUNT(*) as cantidad
FROM servicios_depurados
GROUP BY tipo_caso
ORDER BY cantidad DESC;