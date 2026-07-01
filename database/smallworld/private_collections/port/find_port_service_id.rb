_block
    ruta_fuente << "C:\A_GS1_PROYECTOS\0_Documents_gs\database\smallworld\private_collections\00_out.txt"
    ruta_salida << "C:\A_GS1_PROYECTOS\0_Documents_gs\database\smallworld\private_collections\00_find.txt"

    archivo_fuente << external_text_input_stream.new(ruta_fuente)
    archivo_salida << external_text_output_stream.new(ruta_salida)

    vista << gis_program_manager.cached_dataset(:gis)
    vista.checkpoint("FIND_PT_CB")
    mit_rme_ports << vista.collection(:mit_rme_port)

    resultados << rope.new()

    _for una_linea _over 1.upto(100000)
    _loop
        linea << archivo_fuente.get_line()
        _if linea _is _unset
        _then
            _leave
        _endif

        nombre_a_buscar << linea.write_string

        posibles << mit_rme_ports.select(predicate.eq(:id_servicio, nombre_a_buscar))
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