
_block
    vista << gis_program_manager.cached_dataset(:gis)
    underground_routes << vista.collection(:underground_route)
    id_underground_route << 612802304610685781

    posibles << underground_routes.select(predicate.eq(:id, id_underground_route))
    un_objeto << posibles.an_element()

    show("=== TODOS LOS METODOS ===")
    _for m _over un_objeto.record_exemplar.method_table.elements()
    _loop
        show(m.name)
    _endloop
    show("=== FIN ===")
_endblock
$