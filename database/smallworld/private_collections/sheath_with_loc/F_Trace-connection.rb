_block

_local ruta_fuente << "C:\\A_GS1_PROYECTOS\\0_Documents_gs\\database\\smallworld\\private_collections\\00_find_Trace-connection.txt"
_local ruta_salida << "C:\\A_GS1_PROYECTOS\\0_Documents_gs\\database\\smallworld\\private_collections\\00_out_Trace-connection.txt"
_local input_file << external_text_input_stream.new(ruta_fuente)
_local output_file << external_text_output_stream.new(ruta_salida)
_local vista << gis_program_manager.cached_dataset(:gis)

# ---- FUNCION AUXILIAR: obtiene la descripcion del/los shelf(es) de un mit_bay ----
# Sin recursion. Recorre:
#   1) el shelf directo del rme_container del bay (poco comun)
#   2) los shelves de cada sub-rme_container (lo habitual)
_local shelf_desc_for_bay << _proc(bay)
    _local cont << _unset
    _try
        cont << bay.rme_container
    _when error
        cont << _unset
    _endtry
    _if cont _is _unset
    _then _return "SIN_SHELF" _endif

    _local partes << rope.new()

    # 1) shelf directo del rme_container (por si acaso)
    _local sh_dir << _unset
    _try
        sh_dir << cont.mit_shelf
    _when error
        sh_dir << _unset
    _endtry
    _if sh_dir _isnt _unset
    _then
        _local d << _unset
        _try
            d << sh_dir.description
        _when error
            d << _unset
        _endtry
        _if d _is _unset _orif d.write_string.trim_spaces().size = 0
        _then d << "SIN_DESCRIPCION" _endif
        partes.add(d.write_string)
    _endif

    # 2) shelves de los sub-containers (un solo nivel)
    _local subs << _unset
    _try
        subs << cont.rme_containers
    _when error
        subs << _unset
    _endtry

    _if subs _isnt _unset
    _then
        _for sub _over subs.fast_elements()
        _loop
            _local sh << _unset
            _try
                sh << sub.mit_shelf
            _when error
                sh << _unset
            _endtry
            _if sh _isnt _unset
            _then
                _local d << _unset
                _try
                    d << sh.description
                _when error
                    d << _unset
                _endtry
                _if d _is _unset _orif d.write_string.trim_spaces().size = 0
                _then d << "SIN_DESCRIPCION" _endif
                partes.add(d.write_string)
            _endif
        _endloop
    _endif

    _if partes.size = 0
    _then _return "SIN_SHELF" _endif

    _local salida << partes[1]
    _for k _over 2.upto(partes.size)
    _loop
        salida << salida + "; " + partes[k]
    _endloop
    _return salida
_endproc

output_file.write("SHEATH_ID|SHEATH_NAME|SHEATH_OLDNAME|HILO|ID_HILO|'NOMBRE'|ESTADO|TOPOLOGIA|TECNOLOGIA|STATUS|USO|RANGO|SPLICE_ENTRADA_ID|SPLICE_ENTRADA_NOMBRE|SHELF_ENTRADA_DESC|SPLICE_SALIDA_ID|SPLICE_SALIDA_NOMBRE|SHELF_SALIDA_DESC")
output_file.newline()

_for una_linea _over 1.upto(1000000)
_loop
    _local linea << input_file.get_line()
    _if linea _is _unset _then _leave _endif
    _local id_texto << linea.write_string.trim_spaces()
    _if id_texto.size = 0 _then _continue _endif
    _try
        _local id_sheath << id_texto.as_number()
        _local sheath << vista.collection(:sheath_with_loc).select(predicate.eq(:id, id_sheath)).an_element()
        _if sheath _is _unset
        _then
            output_file.write(id_texto + "|ERROR_NO_ENCONTRADA|-|-|-|-|-|-|-|-|-|-|-|-|-|-|-|-")
            output_file.newline()
            _continue
        _endif
        _local total_hilos << sheath.fiber_count
        _local nombre_funda << sheath.name.write_string
        _local line_counts << sheath.copper_line_of_counts
        _local nombre_antiguo_funda << sheath.nombre_antiguo
        _if nombre_antiguo_funda _is _unset
        _then nombre_antiguo_funda << "-"
        _else nombre_antiguo_funda << nombre_antiguo_funda.write_string
        _endif
        _local fibra_a_objeto << rope.new()
        _for i _over 1.upto(total_hilos) _loop fibra_a_objeto.add(_unset) _endloop
        _if line_counts _isnt _unset
        _then
            _for elem _over line_counts.fast_elements()
            _loop
                _local low << elem.actual_low_range
                _local high << elem.actual_high_range
                _if low _isnt _unset _andif high _isnt _unset
                _then
                    _for fibra _over low.upto(high)
                    _loop
                        _if fibra >= 1 _andif fibra <= total_hilos
                        _then
                            _if fibra_a_objeto[fibra] _is _unset
                            _then fibra_a_objeto[fibra] << elem
                            _endif
                        _endif
                    _endloop
                _endif
            _endloop
        _endif

        # ---- RECOLECTAR TODAS LAS CONEXIONES DE LA FUNDA (splice O bastidor, por pin) ----
        _local conexiones_de_funda << rope.new()
        _local vistos_funda << set.new()

        _for cable _over sheath.mit_cables.fast_elements()
        _loop
            _local pins << _unset
            _try
                pins << cable.mit_sheath_with_loc_pins
            _when error
                pins << _unset
            _endtry
            _if pins _isnt _unset _andif pins.size > 0
            _then
                _for p _over pins.fast_elements()
                _loop
                    _local sp << _unset
                    _try
                        sp << p.sheath_splice
                    _when error
                        sp << _unset
                    _endtry

                    # ---- si hay splice, se registra como splice y NO se busca bastidor ----
                    _if sp _isnt _unset
                    _then
                        _local clave << "SPLICE:" + sp.id.write_string
                        _if _not vistos_funda.includes?(clave)
                        _then
                            vistos_funda.add(clave)
                            conexiones_de_funda.add(simple_vector.new_with(:splice, sp, _unset))
                        _endif
                    _else
                        # ---- no hay splice: buscar si el pin esta conectado a un bastidor ----
                        _try
                            _local strw_conn << p.strw_connect_point
                            _if strw_conn _isnt _unset
                            _then
                                _for bay _over vista.collection(:mit_bay).fast_elements()
                                _loop
                                    _try
                                        _local strw_bay << bay.strw_connect_point
                                        _if strw_bay _isnt _unset _andif strw_bay.id = strw_conn.id
                                        _then
                                            _local clave2 << "BASTIDOR:" + bay.id.write_string
                                            _if _not vistos_funda.includes?(clave2)
                                            _then
                                                vistos_funda.add(clave2)
                                                conexiones_de_funda.add(simple_vector.new_with(:bastidor, _unset, bay))
                                            _endif
                                            _leave
                                        _endif
                                    _when error
                                        # se ignora este bastidor puntual
                                    _endtry
                                _endloop
                            _endif
                        _when error
                            # se ignora este pin puntual
                        _endtry
                    _endif
                _endloop
            _endif
        _endloop

        # ---- ASIGNAR ENTRADA/SALIDA ----
        _local con_entrada << _unset
        _local con_salida << _unset

        _if conexiones_de_funda.size > 0
        _then
            con_entrada << conexiones_de_funda[1]
            _if conexiones_de_funda.size >= 2
            _then
                con_salida << conexiones_de_funda[conexiones_de_funda.size]
            _endif
        _endif

        _for i _over 1.upto(total_hilos)
        _loop
            _local elem << fibra_a_objeto[i]
            _local nombre << ""
            _local estado << ""
            _local Id << "-"
            _local topologia << "-"
            _local tecnologia << "-"
            _local status << "-"
            _local uso << "-"
            _local rango << "-"
            _if elem _isnt _unset
            _then
                Id << elem.id.write_string
                _if Id _is _unset _then Id << "-" _endif
                nombre << elem.designation
                _if nombre _is _unset _then nombre << "" _endif
                _if nombre = "UNDESIGNATED" _then nombre << "" _endif
                _if nombre = ""
                _then estado << "LIBRE"
                _else estado << "FUNCIONANDO"
                _endif
                topologia << elem.tipo_topologia
                _if topologia _is _unset _then topologia << "-" _endif
                tecnologia << elem.tipo_tecnologia
                _if tecnologia _is _unset _then tecnologia << "-" _endif
                status << elem.physical_status
                _if status _is _unset _then status << "-" _endif
                uso << elem.tipo_uso
                _if uso _is _unset _then uso << "-" _endif
                _local low2 << elem.actual_low_range
                _local high2 << elem.actual_high_range
                _if low2 _isnt _unset _andif high2 _isnt _unset
                _then rango << low2.write_string + "-" + high2.write_string
                _endif
            _else
                estado << "LIBRE"
            _endif

            _local sp_ent_id << "-"
            _local sp_ent_nombre << "-"
            _local shelf_ent_desc << "-"
            _local sp_sal_id << "-"
            _local sp_sal_nombre << "-"
            _local shelf_sal_desc << "-"

            _if con_entrada _isnt _unset
            _then
                _try
                    _if con_entrada[1] = :splice
                    _then
                        sp_ent_id << con_entrada[2].id.write_string
                        sp_ent_nombre << con_entrada[2].name.write_string
                    _elif con_entrada[1] = :bastidor
                    _then
                        sp_ent_id << con_entrada[3].id.write_string
                        sp_ent_nombre << "BASTIDOR " + con_entrada[3].number.write_string +
                                         " (SALON:" + con_entrada[3].no_salon.write_string +
                                         " FILA:" + con_entrada[3].no_fila.write_string + ")"
                        shelf_ent_desc << shelf_desc_for_bay(con_entrada[3])
                    _endif
                _when error
                    sp_ent_id << "ERROR"
                    sp_ent_nombre << "ERROR"
                    shelf_ent_desc << "ERROR"
                _endtry
            _endif

            _if con_salida _isnt _unset
            _then
                _try
                    _if con_salida[1] = :splice
                    _then
                        sp_sal_id << con_salida[2].id.write_string
                        sp_sal_nombre << con_salida[2].name.write_string
                    _elif con_salida[1] = :bastidor
                    _then
                        sp_sal_id << con_salida[3].id.write_string
                        sp_sal_nombre << "BASTIDOR " + con_salida[3].number.write_string +
                                         " (SALON:" + con_salida[3].no_salon.write_string +
                                         " FILA:" + con_salida[3].no_fila.write_string + ")"
                        shelf_sal_desc << shelf_desc_for_bay(con_salida[3])
                    _endif
                _when error
                    sp_sal_id << "ERROR"
                    sp_sal_nombre << "ERROR"
                    shelf_sal_desc << "ERROR"
                _endtry
            _endif

            # ---- AQUI SE AGREGAN LAS COMILLAS AL NOMBRE ----
            _local nombre_con_comillas << "'" + nombre + "'"

            _local fila << id_texto + "|" + nombre_funda + "|" + nombre_antiguo_funda + "|" +
                          i.write_string + "|" + Id + "|" + nombre_con_comillas + "|" + estado + "|" +
                          topologia + "|" + tecnologia + "|" + status + "|" +
                          uso + "|" + rango + "|" +
                          sp_ent_id + "|" + sp_ent_nombre + "|" + shelf_ent_desc + "|" +
                          sp_sal_id + "|" + sp_sal_nombre + "|" + shelf_sal_desc
            output_file.write(fila)
            output_file.newline()
        _endloop
    _when error
        output_file.write(id_texto + "|ERROR_PROCESANDO|-|-|-|-|-|-|-|-|-|-|-|-|-|-|-|-")
        output_file.newline()
    _endtry
_endloop
input_file.close()
output_file.close()
show("--------TRACE_TABLE_MASIVO---------")
_endblock