--cruces central--
	--lsd--
		SELECT 
		    Columna1 AS equipo,
		    Columna8 AS port_ec,
		    Columna15 AS etiqueta,
		    Columna14 AS status
		FROM raw.lsd_origin 
		WHERE Columna15 LIKE '%MEADALBOCO134%' 
		--WHERE Columna15 LIKE '%70987-9%' ;
		--WHERE Columna1 LIKE '%BOBAHU931201' AND Columna8 '%2/1/4%';  AND Columna15 LIKE '%95%' ;
		WHERE 
	--vias--
	SELECT 
	    
		EQUIPO AS equipo,
		INTERFAZ_PUERTO_GE AS Port_ec,
	    fibra AS fibra,
	    HILO AS hilo,
		pacheo AS pacheo,
		odf_calle AS odf,
		observacion AS observacion
		
		
	FROM raw.vias_origin
	WHERE "fibra" LIKE '%76017%';
	WHERE "equipo" LIKE '%BOCOHU930602%' AND hilo LIKE '79';
	
	

	

		