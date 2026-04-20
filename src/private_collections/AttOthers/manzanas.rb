_block
	
	
	archivo << external_text_input_stream.new("\\appsignet\Usuarios_GIS\SHP_Cargues_IDECA\Manzana\Txt_Borrar_ID_Manzana.txt")
	no_existe_manzana << rope.new()
	contador << 0
	manzanas << gis_program_manager.cached_Dataset(:ds_cartografia).collection(:manzana)
	_for una_linea _over 1.upto(100000000)
	_loop
		contador +<<1
		linea << archivo.get_line()
		_if contador _mod 100 = 0
		_then
			w(contador)
		_endif 

		_if linea _is _unset _then
			w("contador lineas", contador)
			_leave
		_endif

		_try 
			a_id << linea.write_string
			mio << manzanas.at(a_id.as_number())
			_if mio _is _unset _then 
				no_existe_manzana.add(a_id)
				_continue 
			_endif
			
			mio.delete()
		_when  error
			archivo.close()
		_endtry 
	_endloop
	archivo.close()
_endblock 
$	
