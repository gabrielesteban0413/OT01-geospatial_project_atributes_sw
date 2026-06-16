_block
    ruta_fuente << "C:\\Users\\Jeinvel2\\Desktop\\B01_delete.txt"
    #"C:\A_GS1_PROYECTOS\0_Documents_gs\src\private_collections\ConnBay\B01_delete.txt"

    vista << gis_program_manager.cached_dataset(:gis)
    vista.checkpoint("mit_bay")
    mit_bays << vista.collection(:mit_bay)

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
            id_mit_bay << linea.write_string.as_number()
            posibles_mit_bays << mit_bays.select(predicate.eq(:id, id_mit_bay))
            un_mit_bay << posibles_mit_bays.an_element()

            _if un_mit_bay _isnt _unset
            _then
                un_mit_bay.delete()
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

