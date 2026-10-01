

--cruces central--
	--lsd--
	
	SELECT 
	    Columna1 AS equipo,
	    REPLACE(Columna8, 'GigabitEthernet', '') AS port_ec,
	    Columna15 AS etiqueta,
	    Columna14 AS status
	FROM raw.lsd_origin 
	WHERE Columna15 LIKE '%MEADALBOSB75%';
	--WHERE Columna15 LIKE '%70987-9%' ;
	--WHERE Columna1 LIKE '%BOBAHU931201' AND Columna8 '%2/1/4%';  AND Columna15 LIKE '%95%' ;
			
	
	



	--vias--
	SELECT 
	    
		EQUIPO AS equipo,
		INTERFAZ_PUERTO_GE AS Port_ec,
	    fibra AS fibra,
	    HILO AS hilo,
		pacheo AS pacheo
		
	FROM raw.vias_origin
	--WHERE "fibra" LIKE '%71222%';
	WHERE "equipo" LIKE '%BOCEHU930603%' AND hilo LIKE '30';




	SELECT * FROM staging.merge_corporativo_sync
WHERE cardiseño_origi IN (
    SELECT regexp_split_to_table(
        '
		

70235-34
71239-68


',
        '\n'
    )
);

	

	

		