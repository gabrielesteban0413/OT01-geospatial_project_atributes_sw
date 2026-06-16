_block
    ruta_fuente << "C:\\A_GS1_PROYECTOS\\00_MODULOS\\Z_CALIDAD\\Z_AUTOMATIZACION\\INFR_01_BUILDING\\TXT\\B_name.txt"

    vista << gis_program_manager.cached_dataset(:gis)
    vista.checkpoint("underground_route")
    underground_routes << vista.collection(:underground_route)

    archivo_fuente << external_text_input_stream.new(ruta_fuente)
    lista_resultados << rope.new()

    contador_exitosos << 0
    contador_fallidos << 0

    _for una_linea _over 1.upto(10000)
    _loop
        linea << archivo_fuente.get_line()
        _if linea _is _unset
        _then
            _leave
        _endif

        _try
            id_underground_route << linea.write_string.as_number()
            posibles_underground_routes << underground_routes.select(predicate.eq(:id, id_underground_route))
            un_underground_route << posibles_underground_routes.an_element()

            _if un_underground_route _isnt _unset
            _then
                un_underground_route.delete()
                lista_resultados.add(linea.write_string + "&&&EXITOSO")
                contador_exitosos << contador_exitosos + 1
            _else
                lista_resultados.add(linea.write_string + "&&&FALLIDO")
                contador_fallidos << contador_fallidos + 1
            _endif
        _when error
            lista_resultados.add(linea.write_string + "&&&FALLIDO")
            contador_fallidos << contador_fallidos + 1
        _endtry
    _endloop

    archivo_fuente.close()
    vista.commit()

    archivo_fuente << external_text_output_stream.new(ruta_fuente)
    _for resultado _over lista_resultados.elements()
    _loop
        archivo_fuente.write(resultado)
        archivo_fuente.newline()
    _endloop
    archivo_fuente.close()

    show("--------------DELETE------------.")
    show("Total eliminados: ", contador_exitosos)
    show("Total fallidos: ", contador_fallidos)
_endblock
