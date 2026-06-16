_block
    input_file_path << "C:\\A_GS1_PROYECTOS\\0_Documents_gs\\src\\Private_Collections\\ConnSheathWithLoc\\find\\TX_SWL_Find_id.txt"
    output_file_path << "C:\\A_GS1_PROYECTOS\\0_Documents_gs\\src\\Private_Collections\\ConnSheathWithLoc\\find\\TX00_SWL_Find_output.txt"

    input_file << external_text_input_stream.new(input_file_path)
    output_file << external_text_output_stream.new(output_file_path)

    vista << gis_program_manager.cached_dataset(:gis)

    sheath_with_locs << vista.collection(:sheath_with_loc)
    results << rope.new()

    _for una_linea _over 1.upto(100000)
    _loop
        linea << input_file.get_line()
        _if linea _is _unset
        _then
            _leave
        _endif

        id << linea.write_string
        id_a_buscar << id.as_number()

        posibles << sheath_with_locs.select(predicate.eq(:id, id_a_buscar))
        un_objeto << posibles.an_element()

        _if un_objeto _isnt _unset
        _then
            name_old << un_objeto.nombre_antiguo.write_string
            name_text << un_objeto.name.write_string
            results.add( '"' + id + '"|' + name_text ) # + "|" + name_text
        _else
            results.add(id + "&!")
        _endif
    _endloop

    input_file.close()

    _for r _over results.elements()
    _loop
        output_file.write(r)
        output_file.newline()
    _endloop
    output_file.close()

    show("--------GS-FIND---------")
_endblock
