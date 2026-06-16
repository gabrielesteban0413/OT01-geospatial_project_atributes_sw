_block
    vista << gis_program_manager.cached_dataset(:gis)
    design_ds << gis_program_manager.cached_dataset(:design_admin)
    alt_map << design_ds.collection(:swg_dsn_scheme_alternative_map)
    esquemas << design_ds.collection(:swg_dsn_scheme)
    proyectos << design_ds.collection(:swg_dsn_project)
    buildings << vista.collection(:building)
    un_objeto << buildings.select(predicate.eq(:id, 3492000048800395440)).an_element()
    target_version << un_objeto.ds!version
    show("ds!version del objeto: ", target_version)
    mejor_alt << _unset
    mejor_version << -1
    _for alt _over alt_map.select(predicate.eq(:dataset_name, "gis")).elements()
    _loop
        v << alt.ds!version
        _if v <= target_version _andif v > mejor_version
        _then
            mejor_version << v
            mejor_alt << alt
        _endif
    _endloop
    _if mejor_alt _isnt _unset
    _then
        scheme << esquemas.select(predicate.eq(:id, mejor_alt.scheme_id)).an_element()
        _if scheme _isnt _unset
        _then
            proyecto << proyectos.select(predicate.eq(:id, scheme.swg_dsn_project_id)).an_element()
            show("=== RESPONSABLE DEL OBJETO ===")
            show("  Alternativa:       ", mejor_alt.alternative_name)
            show("  Nombre diseño:     ", scheme.name)
            show("  Owner:             ", scheme.owner)
            show("  Modificado por:    ", scheme.geaendert_von)
            show("  Fecha modif:       ", scheme.geaendert_am)
            show("  Status:            ", scheme.status)
            _if proyecto _isnt _unset
            _then
                show("  Proyecto nombre:   ", proyecto.name)
                show("  Proyecto job:      ", proyecto.job_title)
                show("  Grupo dueño:       ", proyecto.owning_group)
                show("  Proyecto mod por:  ", proyecto.geaendert_von)
                show("  Fecha mod proy:    ", proyecto.geaendert_am)
            _endif
        _else
            show("Scheme no encontrado para id: ", mejor_alt.scheme_id)
        _endif
    _else
        show("No se encontro alternativa para ds!version: ", target_version)
    _endif
_endblock
$