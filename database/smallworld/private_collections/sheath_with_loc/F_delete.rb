_block
    ruta_fuente << "C:\\A_GS1_PROYECTOS\\0_Documents_gs\\scripts\\monitoring\\private_collections\\00_out.txt"

    vista << gis_program_manager.cached_dataset(:gis)
    vista.checkpoint("sheath_with_loc")
    sheath_with_locs << vista.collection(:sheath_with_loc)

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
            id_sheath_with_loc << linea.write_string.as_number()
            posibles_sheath_with_locs << sheath_with_locs.select(predicate.eq(:id, id_sheath_with_loc))
            un_sheath_with_loc << posibles_sheath_with_locs.an_element()

            _if un_sheath_with_loc _isnt _unset
            _then
                un_sheath_with_loc.delete()
                lista_resultados.add(linea.write_string + "|EXITOSO")
                contador_exitosos << contador_exitosos + 1
            _else
                lista_resultados.add(linea.write_string + "|FALLIDO")
                contador_fallidos << contador_fallidos + 1
            _endif
        _when error
            lista_resultados.add(linea.write_string + "|FALLIDO")
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

