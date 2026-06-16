_block
    vista << gis_program_manager.cached_dataset(:gis)
    underground_routes << vista.collection(:underground_route)
    
    id_underground_route << 612802304610685781
    
    show("=== INTENTANDO ACCESO DIRECTO ===")
    
    posibles << underground_routes.select(predicate.eq(:id, id_underground_route))
    un_objeto << posibles.an_element()
    show("Objeto: ", un_objeto)
    
    # Ver todos los campos sin tocar geometría
    _try
        show("Campos del objeto:")
        _for campo _over un_objeto.record_exemplar.all_fields.elements()
        _loop
            _try
                valor << un_objeto.perform(campo.name)
                show("  ", campo.name, " = ", valor)
            _when error
                show("  ", campo.name, " = ERROR AL LEER")
            _endtry
        _endloop
    _when error
        show("Error leyendo campos")
    _endtry
    
    # Intentar borrado directo sin tocar campos
    show("=== INTENTANDO DELETE DIRECTO ===")
    _try
        vista.checkpoint("test_delete")
        mut << un_objeto.as_mutable()
        show("as_mutable OK: ", mut)
        mut.delete()
        show("DELETE OK")
        vista.commit()
    _when error
        show("Error en delete, intentando rollback")
        _try
            vista.rollback()
        _when error
        _endtry
    _endtry
_endblock