_block
    ruta_fuente << "C:\\A_GS1_PROYECTOS\\0_Documents_gs\\src\\Private_Collections\\ConnSheathWithLoc\\find\\TX_SWL_Find_name.txt"
    ruta_salida << "C:\\A_GS1_PROYECTOS\\0_Documents_gs\\src\\Private_Collections\\ConnSheathWithLoc\\find\\TX00_SWL_Find_output.txt"

    archivo_fuente << external_text_input_stream.new(ruta_fuente)
    archivo_salida << external_text_output_stream.new(ruta_salida)

    vista << gis_program_manager.cached_dataset(:gis)
    vista.checkpoint("FindSheathWithLocByName")
    sheath_with_locs << vista.collection(:sheath_with_loc)

    resultados << rope.new()

    _for una_linea _over 1.upto(100000)
    _loop
        linea << archivo_fuente.get_line()
        _if linea _is _unset
        _then
            _leave
        _endif

        nombre_a_buscar << linea.write_string

        posibles << sheath_with_locs.select(predicate.eq(:name, nombre_a_buscar))
        un_objeto << posibles.an_element()

        _if un_objeto _isnt _unset
        _then
            id_texto << un_objeto.id.write_string
            resultados.add( id_texto + "|" + nombre_a_buscar)
        _else
            resultados.add("N/A " + nombre_a_buscar )
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