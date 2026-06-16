_block
    ruta_fuente << "C:\\A_GS1_PROYECTOS\\00_MODULOS\\Z_CALIDAD\\Z_DELETE\\FIBRA CORRUCTA - C.71212.shp\\CNT_04_SHEATH_WITH_LOC\\FIND\\Z_F_B_00_RESULTADO1.txt"
    ruta_salida << "C:\\A_GS1_PROYECTOS\\00_MODULOS\\Z_CALIDAD\\Z_DELETE\\FIBRA CORRUCTA - C.71212.shp\\CNT_04_SHEATH_WITH_LOC\\FIND\\1Z_F_B_01_RESULTADO1.txt"

    archivo_fuente << external_text_input_stream.new(ruta_fuente)
    archivo_salida << external_text_output_stream.new(ruta_salida)

    vista << gis_program_manager.cached_dataset(:gis)
    vista.checkpoint("FindBuilding")
    buildings << vista.collection(:building)

    resultados << rope.new()

    _for una_linea _over 1.upto(100000)
    _loop
        linea << archivo_fuente.get_line()
        _if linea _is _unset
        _then
            _leave
        _endif

        nombre_a_buscar << linea.write_string

        id_a_buscar << nombre_a_buscar.as_number()

        posibles << buildings.select(predicate.eq(:id, id_a_buscar))
        un_objeto << posibles.an_element()

        _if un_objeto _isnt _unset
        _then
            name_texto << un_objeto.name.write_string
            resultados.add(nombre_a_buscar + "," + name_texto)
        _else
            resultados.add(nombre_a_buscar + "&N/A")
        _endif
    _endloop

    archivo_fuente.close()

    _for r _over resultados.elements()
    _loop
        archivo_salida.write(r)
        archivo_salida.newline()
    _endloop
    archivo_salida.close()

    show("--------GS-FIND---------")
_endblock
