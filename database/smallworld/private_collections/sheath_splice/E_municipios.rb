_block
    ruta_fuente << "C:\\Users\\Jeinvel2\\Desktop\\F01_Municipio-EMPALMES1.txt"
    #"C:\\A_GS1_PROYECTOS\\0_Documents_gs\\src\\private_collections\\ConnSheathWithLoc\\edit\\F01_Municipio.txt"
    #"C:\\Users\\gelvsieg\\Desktop\\nodo.txt"

    vista << gis_program_manager.cached_dataset(:gis)
    vista.checkpoint("FI_MUN_A")
    sheath_splices << vista.collection(:sheath_splice)

    archivo_fuente << external_text_input_stream.new(ruta_fuente)
    lista_modificada << rope.new()

    contador_exitosos << 0
    contador_fallidos << 0

    _for una_linea _over 1.upto(100000)
    _loop
        linea << archivo_fuente.get_line()
        _if linea _is _unset
        _then
            _leave
        _endif

        linea_str << linea.write_string
        texto << linea_str.split_by("|", _true)

        _if texto.size <> 2
        _then
            lista_modificada.add(linea_str + "|FALLIDO|FORMATO_INCORRECTO")
            contador_fallidos << contador_fallidos + 1
            _continue
        _endif

        id_sheath_splice << texto[1].as_number()
        municipio_id_value << texto[2].as_number()

        posibles_sheath_splices << sheath_splices.select(predicate.eq(:id, id_sheath_splice))
        un_sheath_splice << posibles_sheath_splices.an_element()

        _if un_sheath_splice _isnt _unset
        _then
            _try
                un_sheath_splice.municipio_id << municipio_id_value
                valor_nuevo << un_sheath_splice.municipio_id
                _if valor_nuevo = municipio_id_value
                _then
                    lista_modificada.add(linea_str + "|EXITOSO")
                    contador_exitosos << contador_exitosos + 1
                _else
                    lista_modificada.add(linea_str + "|FALLIDO|NO_CAMBIO")
                    contador_fallidos << contador_fallidos + 1
                _endif
            _when error  
                lista_modificada.add(linea_str + "|FALLIDO|ERROR_EN_EJECUCION")
                contador_fallidos << contador_fallidos + 1
            _endtry
        _else
            lista_modificada.add(linea_str + "|FALLIDO|NO_ENCONTRADO")
            contador_fallidos << contador_fallidos + 1
        _endif
    _endloop

    archivo_fuente.close()

    vista.commit()

    archivo_salida << external_text_output_stream.new(ruta_fuente)
    _for nueva_linea _over lista_modificada.elements()
    _loop
        archivo_salida.write(nueva_linea)
        archivo_salida.newline()
    _endloop
    archivo_salida.close()

    show("--------------GS--EDIT-FI-NODO---------")
    show("exitosos: ", contador_exitosos)
    show("fallidos: ", contador_fallidos)
_endblock