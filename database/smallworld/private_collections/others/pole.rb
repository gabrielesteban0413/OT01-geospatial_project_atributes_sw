_block
    ruta_fuente << "C:\\A_GS1_PROYECTOS\\00_MODULOS\\Z_CALIDAD\\Z_AUTOMATIZACION\\SRC\\postess.txt"
    ruta_exitosos << "C:\\A_GS1_PROYECTOS\\exitosos_log.txt"

    vista << gis_program_manager.cached_dataset(:gis)
    vista.checkpoint("pole")
    poles << vista.collection(:pole)

    address_exitosos << rope.new()
    archivo_fuente << external_text_input_stream.new(ruta_fuente)

    _for una_linea _over 1.upto(100)
    _loop
        linea << archivo_fuente.get_line()
        _if linea _is _unset
        _then
            archivo_fuente.close()
            _leave
        _endif

        texto << linea.write_string.split_by("&&&", _true)
        _if texto.size <> 2
        _then
            _continue
        _endif

        id_pole << texto[1].as_number()
        usage_value << texto[2].write_string

        posibles_poles << poles.select(predicate.eq(:id, id_pole))
        un_pole << posibles_poles.an_element()

        _if un_pole _isnt _unset
        _then
            _try
                un_pole.usage << usage_value
                address_exitosos.add(linea)
            _when error
            _endtry
        _endif
    _endloop

    archivo_fuente.close()
    vista.commit()

    show("Registros de poles actualizados: ", address_exitosos.size)

    _if address_exitosos.size <> 0
    _then
        archivo_exitosos << external_text_output_stream.new(ruta_exitosos)
        _for exito _over address_exitosos.elements()
        _loop
            archivo_exitosos.write(exito)
            archivo_exitosos.newline()
        _endloop
        archivo_exitosos.close()
    _endif
_endblock