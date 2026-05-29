_block
    ruta_fuente << "C:\\A_GS1_PROYECTOS\\0_Documents_gs\\scripts\\monitoring\\private_collections\\00_out1.txt"

    vista << gis_program_manager.cached_dataset(:gis)
    vista.checkpoint("mit_rme_port")
    mit_rme_ports << vista.collection(:mit_rme_port)

    archivo_fuente << external_text_input_stream.new(ruta_fuente)
    lineas_procesadas << rope.new()

    contador_exitosos << 0
    contador_fallidos << 0

    _for una_linea _over 1.upto(100000)
    _loop
        linea << archivo_fuente.get_line()
        _if linea _is _unset
        _then
            _leave
        _endif

        texto << linea.write_string.split_by("|", _true)
        _if texto.size <> 2
        _then
            # Mantener la línea original y agregar estado FALLIDO
            lineas_procesadas.add(linea + "|FALLIDO - Formato incorrecto")
            contador_fallidos << contador_fallidos + 1
            _continue
        _endif

        id_mit_rme_port << texto[1].as_number()
        nuevo_valor << texto[2].as_number()

        posibles_mit_rme_ports << mit_rme_ports.select(predicate.eq(:id, id_mit_rme_port))
        un_mit_rme_port << posibles_mit_rme_ports.an_element()

        _if un_mit_rme_port _isnt _unset
        _then
            _try
                un_mit_rme_port.direccion_cliente_id << nuevo_valor
                # Agregar estado EXITOSO a la línea original
                lineas_procesadas.add(linea + "|EXITOSO")
                contador_exitosos << contador_exitosos + 1
            _when error
                # Agregar estado FALLIDO a la línea original
                lineas_procesadas.add(linea + "|FALLIDO - Error en actualización")
                contador_fallidos << contador_fallidos + 1
            _endtry
        _else
            # Agregar estado FALLIDO a la línea original
            lineas_procesadas.add(linea + "|FALLIDO - Puerto no encontrado")
            contador_fallidos << contador_fallidos + 1
        _endif
    _endloop

    archivo_fuente.close()
    vista.commit()

    # Sobrescribir el archivo original con las líneas procesadas (incluyendo el estado)
    archivo_salida << external_text_output_stream.new(ruta_fuente)
    _for nueva_linea _over lineas_procesadas.elements()
    _loop
        archivo_salida.write(nueva_linea)
        archivo_salida.newline()
    _endloop
    archivo_salida.close()

    show("--------------RESULTADOS--------.")
    show("exitosos: ", contador_exitosos)
    show("fallidos: ", contador_fallidos)
_endblock