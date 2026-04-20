_block
    vista << gis_program_manager.cached_dataset(:gis)
    
    target_checkpoint << 612802304610685789
    
    show("=== BUSCANDO CHECKPOINT: ", target_checkpoint, " ===")
    
    _for key _over vista.alternative_history.keys.elements()
    _loop
        estado << vista.alternative_history[key]
        
        _if estado.checkpoint = target_checkpoint
        _then
            show("--- ENCONTRADO ---")
            show("  key:        ", key)
            show("  checkpoint: ", estado.checkpoint)
            show("  mode:       ", estado.mode)
        _endif
    _endloop
_endblock
$







#saber usuarios

_block
    vista << gis_program_manager.cached_dataset(:gis)
    design_ds << gis_program_manager.cached_dataset(:design_admin)
    alt_map << design_ds.collection(:swg_dsn_scheme_alternative_map)
    esquemas << design_ds.collection(:swg_dsn_scheme)

    # Listar todas las alternativas del dataset gis
    show("=== ALTERNATIVAS EN DATASET GIS ===")
    _for alt _over alt_map.select(predicate.eq(:dataset_name, "gis")).elements()
    _loop
        scheme << esquemas.select(predicate.eq(:id, alt.scheme_id)).an_element()
        _if scheme _isnt _unset
        _then
            show("---")
            show("  alternativa:    ", alt.alternative_name)
            show("  scheme_id:      ", alt.scheme_id)
            show("  nombre:         ", scheme.name)
            show("  creado por:     ", scheme.erfasst_von)
            show("  fecha creacion: ", scheme.erfasst_am)
        _endif
    _endloop
    show("=== FIN ===")
_endblock
$





# saber checkpoin
_block
    vista << gis_program_manager.cached_dataset(:gis)
    design_ds << gis_program_manager.cached_dataset(:design_admin)
    alt_map << design_ds.collection(:swg_dsn_scheme_alternative_map)
    esquemas << design_ds.collection(:swg_dsn_scheme)

    # Paso 1: Obtener el objeto y su checkpoint
    underground_routes << vista.collection(:underground_route)
    un_objeto << underground_routes.select(predicate.eq(:id, 612802304610685781)).an_element()
    checkpoint << un_objeto.perform(:ds!version)
    nombre_plan << "PLAN_" + checkpoint.write_string
    show("Checkpoint del objeto: ", checkpoint)
    show("Buscando plan:         ", nombre_plan)

    # Paso 2: Buscar el plan en alternative_map
    alt << alt_map.select(predicate.eq(:alternative_name, nombre_plan)).an_element()

    _if alt _isnt _unset
    _then
        # Paso 3: Obtener el scheme asociado
        scheme << esquemas.select(predicate.eq(:id, alt.scheme_id)).an_element()

        _if scheme _isnt _unset
        _then
            show("=== HISTORIAL DEL OBJETO ===")
            show("  Plan:            ", nombre_plan)
            show("  Nombre diseño:   ", scheme.name)
            show("  Creado por:      ", scheme.erfasst_von)
            show("  Fecha creacion:  ", scheme.erfasst_am)
            show("  Modificado por:  ", scheme.geaendert_von)
            show("  Fecha modif:     ", scheme.geaendert_am)
            show("  Status:          ", scheme.status)
        _else
            show("Scheme no encontrado para id: ", alt.scheme_id)
        _endif
    _else
        show("Plan no encontrado: ", nombre_plan)
        show("Verificar si el checkpoint corresponde directamente al nombre del plan")
    _endif
_endblock
$





MagikSF>
_block
    vista << gis_program_manager.cached_dataset(:gis)

    # El ds!version es el numero de version del dataset cuando se guardo el objeto
    # Buscar que alternativa tenia esa version
    show("=== ALTERNATIVAS ACTIVAS EN EL DATASET GIS ===")
    _for key _over vista.alternative_history.keys.elements()
    _loop
        estado << vista.alternative_history[key]
        show("  key: ", key, " | checkpoint: ", estado.checkpoint, " | mode: ", estado.mode)
    _endloop
_endblock
$





MagikSF>
_block
    vista << gis_program_manager.cached_dataset(:gis)

    # El ds!version es el numero de version del dataset cuando se guardo el objeto
    # Buscar que alternativa tenia esa version
    show("=== ALTERNATIVAS ACTIVAS EN EL DATASET GIS ===")
    _for key _over vista.alternative_history.keys.elements()
    _loop
        estado << vista.alternative_history[key]
        show("  key: ", key, " | checkpoint: ", estado.checkpoint, " | mode: ", estado.mode)
    _endloop
_endblock
$
"=== ALTERNATIVAS ACTIVAS EN EL DATASET GIS ==="
"  key: " :|***top***| " | checkpoint: " unset " | mode: " :readonly
"  key: " :\||Engineering Design| " | checkpoint: " unset " | mode: " :readonly
"  key: " :\||Engineering Design|\||PLAN_34195| " | checkpoint: " unset " | mode: " :write
MagikSF>











MagikSF>
_block
    vista << gis_program_manager.cached_dataset(:gis)
    version_buscada << 84938486

    # product_history existe en la lista de metodos de la vista
    show("=== PRODUCT HISTORY ===")
    _protect
        hist << vista.product_history()
        show("product_history: ", hist)
        _for entry _over hist.elements()
        _loop
            show("  entry: ", entry)
        _endloop
    _protection
        show("product_history fallo")
    _endprotect

    # alternative_history existe en la lista de metodos
    show("=== ALTERNATIVE HISTORY ===")
    _protect
        ah << vista.alternative_history
        show("alternative_history: ", ah)
        _for entry _over ah.elements()
        _loop
            show("  ", entry)
        _endloop
    _protection
        show("alternative_history fallo")
    _endprotect

    # username del creador via ds!version
    show("=== USERNAME ACTUAL ===")
    show("usuario actual: ", vista.current_username)
    show("username: ", vista.username)
_endblock
$
        ah << vista.alternative_history
           ^
Global ah does not exist: create it? (Y) y
Defining global user:ah
"=== PRODUCT HISTORY ==="
"product_history: " a sw:sorted_collection(36)
"  entry: " sw_gis!datamodel_history25:(sw_kernel,ds_src,datamodel_history,2,Install)
"  entry: " sw_gis!datamodel_history25:(sw_core,datastore_geometry_vector,base,1,Install)
"  entry: " sw_gis!datamodel_history25:(sw_core,datastore_geometry_vector,dimension,1,Install)
"  entry: " sw_gis!datamodel_history25:(sw_core,datastore_geometry_vector,raster,1,Install)
"  entry: " sw_gis!datamodel_history25:(sw_core,datastore_geometry_vector,simple,1,Install)
"  entry: " sw_gis!datamodel_history25:(sw_core,datastore_geometry_vector,tin,1,Install)
"  entry: " sw_gis!datamodel_history25:(sw_core,datastore_geometry_vector,topology,1,Install)
"  entry: " sw_gis!datamodel_history25:(sw_kernel,ds_src,dd,2,Install)
"  entry: " sw_gis!datamodel_history25:(sw_core,datastore_base,drawing,1,Install)
"  entry: " sw_gis!datamodel_history25:(sw_core,drafting,drafting,1,Install)
"  entry: " sw_gis!datamodel_history25:(pni,cable.model,cable.model,31000,Install)
"  entry: " sw_gis!datamodel_history25:(pni,conduit.model,conduit.model,31016,Install)
"  entry: " sw_gis!datamodel_history25:(pni,connectivity.model,connectivity.model,31000,Install)
"  entry: " sw_gis!datamodel_history25:(pni,copper.model,copper.model,31003,Install)
"  entry: " sw_gis!datamodel_history25:(pni,fiber.model,fiber.model,31011,Install)
"  entry: " sw_gis!datamodel_history25:(pni,ftth.model,ftth.model,31000,Install)
"  entry: " sw_gis!datamodel_history25:(pni,landbase.model,landbase.model,31002,Install)
"  entry: " sw_gis!datamodel_history25:(pni,mdu.model,mdu.model,31000,Install)
"  entry: " sw_gis!datamodel_history25:(pni,microwave.model,microwave.model,31000,Install)
"  entry: " sw_gis!datamodel_history25:(pni,misc.model,misc.model,31015,Install)
"  entry: " sw_gis!datamodel_history25:(pni,network.model,network.model,31000,Install)
"  entry: " sw_gis!datamodel_history25:(pni,powering.model,powering.model,31001,Install)
"  entry: " sw_gis!datamodel_history25:(pni,rf.model,rf.model,31010,Install)
"  entry: " sw_gis!datamodel_history25:(pni,rme.model,rme.model,31006,Install)
"  entry: " sw_gis!datamodel_history25:(pni,schematics.model,schematics.model,31001,Install)
"  entry: " sw_gis!datamodel_history25:(pni,specs.model,specs.model,31020,Install)
"  entry: " sw_gis!datamodel_history25:(pni,structure.model,structure.model,31010,Install)
"  entry: " sw_gis!datamodel_history25:(pni,ftth.model,ftth.model,31001,Install)
"  entry: " sw_gis!datamodel_history25:(pni,ftth.model,ftth.model,31002,Install)
"  entry: " sw_gis!datamodel_history25:(pni,ftth.model,ftth.model,31003,Install)
"  entry: " sw_gis!datamodel_history25:(pni,ftth.model,ftth.model,31004,Install)
"  entry: " sw_gis!datamodel_history25:(pni,ftth.model,ftth.model,31005,Install)
"  entry: " sw_gis!datamodel_history25:(pni,structure.model,structure.model,31011,Install)
"  entry: " sw_gis!datamodel_history25:(pni,structure.model,structure.model,31012,Install)
"  entry: " sw_gis!datamodel_history25:(pni,structure.model,structure.model,31013,Install)
"  entry: " sw_gis!datamodel_history25:(pni,structure.model,structure.model,31014,Install)
"product_history fallo"
"=== ALTERNATIVE HISTORY ==="
"alternative_history: " sw:hash_table(3)
"  " a ds_alternative_state(checkpoint: unset,  mode: readonly)
"  " a ds_alternative_state(checkpoint: unset,  mode: readonly)
"  " a ds_alternative_state(checkpoint: unset,  mode: write)
"alternative_history fallo"
"=== USERNAME ACTUAL ==="
"usuario actual: " "jeinvel2"
"username: " "jeinvel2"
MagikSF>





_block
    vista << gis_program_manager.cached_dataset(:gis)
    underground_routes << vista.collection(:underground_route)
    posibles << underground_routes.select(predicate.eq(:id, 612802304610685781))
    un_objeto << posibles.an_element()

    # Ver todas las propiedades del instance_metadata
    show("=== INSTANCE METADATA COMPLETO ===")
    meta << un_objeto.instance_metadata
    _for key _over meta.keys.elements()
    _loop
        show("  ", key, " = ", meta[key])
    _endloop
    show("=== FIN ===")
_endblock
$
    meta << un_objeto.instance_metadata
         ^
Global meta does not exist: create it? (Y) y
Defining global user:meta
"=== INSTANCE METADATA COMPLETO ==="
"  " :class_name " = " :underground_route34389
"  " :instance_format " = " :slotted
"  " :instance_data_type " = " :pointer
"  " :behaviour_sinks " = " sw:weak_set(0)
"  " :behaviour_sources " = " property_list(1)
"  " :method_slot_names " = " property_list(2)
"  " :instance_slot_data " = " sw:simple_vector:[1-2]
"  " :exemplar " = " a underground_route34389
"=== FIN ==="
MagikSF>
_block






_block
    vista << gis_program_manager.cached_dataset(:gis)
    underground_routes << vista.collection(:underground_route)
    id_underground_route << 612802304610685781

    posibles << underground_routes.select(predicate.eq(:id, id_underground_route))
    un_objeto << posibles.an_element()

    show("=== INSTANCE METADATA ===")
    show(un_objeto.instance_metadata)
_endblock
$
"=== INSTANCE METADATA ==="
property_list(8)
MagikSF>
show(gis_program_manager.cached_dataset(:gis).history())
$
a ds_alternative_state(checkpoint: unset,  mode: write)
MagikSF>



















MagikSF>
_block
    vista << gis_program_manager.cached_dataset(:gis)
    version_buscada << 84938486

    # Buscar en el historial de transacciones del dataset
    show("=== HISTORIAL DE TRANSACCIONES ===")
    _protect
        _for trans _over vista.transaction_log.elements()
        _loop
            _if trans.version = version_buscada
            _then
                show("Version: ", trans.version)
                show("Usuario: ", trans.user)
                show("Fecha:   ", trans.date)
                show("Hora:    ", trans.time)
            _endif
        _endloop
    _protection
        show("transaction_log no disponible")
    _endprotect

    # Alternativa: changelog del dataset
    _protect
        _for entry _over vista.change_log.elements()
        _loop
            _if entry.version = version_buscada
            _then
                show("Entry: ", entry)
            _endif
        _endloop
    _protection
        show("change_log no disponible")
    _endprotect

    # Alternativa: directamente del ds_version_manager
    _protect
        vm << vista.version_manager
        info << vm.version_info(version_buscada)
        show("version_manager info: ", info)
    _protection
        show("version_manager no disponible")
    _endprotect

    show("=== FIN ===")
_endblock
$
    version_buscada << 84938486
                    ^
Global version_buscada does not exist: create it? (Y) y
Defining global user:version_buscada
        vm << vista.version_manager
           ^
Global vm does not exist: create it? (Y) y
Defining global user:vm
        info << vm.version_info(version_buscada)
             ^
Global info does not exist: create it? (Y) y
Defining global user:info
"=== HISTORIAL DE TRANSACCIONES ==="
**** Error: Object mit_gis_ds_view(Gis) does not understand message transaction_log
     does_not_understand(object=mit_gis_ds_view(Gis), selector=:transaction_log, arguments=sw:simple_vector:[1-0], iterator?=False, private?=False)

---- traceback: cli (heavy_thread 1294962) ----
time=12/03/2026 13:44:51
sw!version=4.3.0  (swaf)
os_text_encoding=cp1252
!snapshot_traceback?!=False

condition(information).raise(:does_not_understand, {:object, mit_gis_ds_view(Gis), :selector, :transaction_log, ... <size=10>})
mit_gis_ds_view(Gis).does_not_understand(a sw:message, False)
mit_gis_ds_view(Gis).sys!send_error(:transaction_log, method_table for sw:mit_gis_ds_view, False, 1, sw:simple_vector:[1-0])
*** top level ***()
a sw:magik_rep.process(sw:simple_vector:[1-5])
a sw:magik_rep.cli(a sw:terminal, "MagikSF> ")
cli()
light_thread_launcher_proc_990928()
"transaction_log no disponible"
MagikSF>