clear all
set more off

////////////////////////////////////////////////////////////////////////////////
//                                                                            //
//          INSTITUTO DE INVESTIGACIONES SOCIO-ECONÓMICAS IISEC               //
//                                                                            //
////////////////////////////////////////////////////////////////////////////////

* ============================================================
* 04_PREPARAR_EXCELS_GRAFICOS_DEMOGRAFIA_SSR_COMPLETO.do
* ============================================================
*
* Objetivo:
*     Preparar archivos Excel individuales para alimentar el
*     script de gráficos de indicadores de Demografía, Salud
*     Sexual y Reproductiva.
*
* Entrada:
*     Excel maestro generado por el script de cálculo:
*         indicadores_demografia_ssr_2005_20XX.xlsx
*
* Salida:
*     Archivos individuales por código de indicador:
*         06_001_01.xlsx
*         06_001_02.xlsx
*         ...
*         06_003_22.xlsx
*
* Carpeta de salida:
*     G:\Unidades compartidas\1_INDICADORES_ODSB_2022\01COPIA_IISEC_ODSB\!EH_ARMONIZADA\_salud
*
* Estructuras de salida:
*
*     Indicador simple:
*         _in  : Año | Bolivia
*         _in1 : Año | Departamento | Valor
*                Orden departamental fijo: 01 Chuquisaca, 02 La Paz, ..., 09 Pando
*
*     Indicador con dos tipos:
*         _in  : Año | Tipo | Valor
*         _in1 : Año | Departamento | Tipo | Valor
*
*     Indicador doble en columnas:
*         _in  : Año | Bolivia1 | Bolivia2
*         _in1 : Año | Departamento | Valor1 | Valor2
*
*     Indicador por sexo:
*         _in  : Año | Bolivia1 | Bolivia2
*         _in1 : Año | Departamento | Valor1 | Valor2
*         _in2 : Año | Bolivia1 | Bolivia2      // copia para compatibilidad
*         _in3 : Año | Departamento | Valor1 | Valor2 // copia para compatibilidad
*
* Nota:
*     Si algún indicador falla porque Stata no encuentra la hoja,
*     revisar la sección "MAPEO DE INDICADORES" al final y ajustar
*     el nombre de la hoja del Excel maestro.
* ============================================================

* ============================================================
* 0. DIRECTORIOS DEL PROYECTO
* ============================================================

	global path "G:/Unidades compartidas/1_INDICADORES_ODSB_2022/01COPIA_IISEC_ODSB"

	* Verificar que Stata pueda leer la ruta raíz.
	capture dir "$path/*"

	if _rc != 0 {
		di as error "ERROR: Stata no puede leer la ruta raíz:"
		di as error "$path"
		di as error "Verifica manualmente con:"
		di as error `"dir "G:/Unidades compartidas/1_INDICADORES_ODSB_2022/*""'
		error 601
	}

	* Directorios principales.
	global work "$path/INDICADORES/6_SALUD"

	global do   "$work/_do"
	global gph  "$work/_gph"
	global xlsx "$work/_xlsx"
	global out  "$work/_out"

	* Entradas.
	global in          "$path/!EH_ARMONIZADA/_out"
	global in_graficos "$path/!EH_ARMONIZADA/_salud"

	* Salidas.
	global out_dta      "$out/Bases DTA"
	global out_xlsx     "$out/Indicadores Excel"
	global out_aud      "$out/Auditorías"
	global out_graficos "$path/!EH_ARMONIZADA/_salud"

	* Excel maestro de indicadores.
	global master "$out_xlsx/indicadores_demografia_ssr_2005_2024.xlsx"

	* Crear carpetas si no existen.
	capture mkdir "$out"
	capture mkdir "$out_dta"
	capture mkdir "$out_xlsx"
	capture mkdir "$out_aud"
	capture mkdir "$out_graficos"

	* Controles básicos.
	capture dir "$work/*"

	if _rc != 0 {
		di as error "ERROR: Stata no puede leer la carpeta de trabajo:"
		di as error "$work"
		error 601
	}

	capture dir "$out_graficos/*"

	if _rc != 0 {
		di as error "ERROR: Stata no puede leer la carpeta de salida para gráficos:"
		di as error "$out_graficos"
		error 601
	}

	confirm file "$master"

	di as result "============================================================"
	di as result "DIRECTORIOS CONFIGURADOS"
	di as result "path:         $path"
	di as result "work:         $work"
	di as result "out_xlsx:     $out_xlsx"
	di as result "out_graficos: $out_graficos"
	di as result "master:       $master"
	di as result "============================================================"

* ------------------------------------------------------------
* 1. Listas de territorios
* ------------------------------------------------------------

	* Orden oficial por código geográfico:
	* 00 Bolivia, 01 Chuquisaca, 02 La Paz, 03 Cochabamba,
	* 04 Oruro, 05 Potosí, 06 Tarija, 07 Santa Cruz, 08 Beni, 09 Pando.
	local geo_vars  BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO
	local geo_codes 00      01         02     03         04    05     06     07          08   09
	local geo_labs  Bolivia Chuquisaca LaPaz  Cochabamba Oruro Potosí Tarija SantaCruz  Beni Pando

	local dep_vars  CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO
	local dep_codes 01         02     03         04    05     06     07          08   09
	local dep_labs  Chuquisaca LaPaz  Cochabamba Oruro Potosí Tarija SantaCruz  Beni Pando

	global GEO_VARS  `geo_vars'
	global GEO_CODES `geo_codes'
	global GEO_LABS  `geo_labs'

	global DEP_VARS  `dep_vars'
	global DEP_CODES `dep_codes'
	global DEP_LABS  `dep_labs'

* ============================================================
* 2. PROGRAMAS AUXILIARES
* ============================================================

* ------------------------------------------------------------
* 2.1 Importar hoja con candidatos
* ------------------------------------------------------------

	capture program drop _importa_hoja_maestra
	program define _importa_hoja_maestra
		syntax, Sheets(string)

		local found 0

		foreach sh of local sheets {
			capture noisily import excel using "$master", sheet("`sh'") firstrow clear

			if _rc == 0 {
				local found 1
				di as result "Hoja importada desde Excel maestro: `sh'"
				continue, break
			}
		}

		if `found' == 0 {
			di as error "ERROR: No se encontró ninguna de estas hojas en el Excel maestro: `sheets'"
			error 601
		}
	end

* ------------------------------------------------------------
* 2.2 Estandarizar columna de año
* ------------------------------------------------------------

	capture program drop _estandariza_anio
	program define _estandariza_anio

		capture confirm variable Año
		if _rc == 0 {
			exit
		}

		capture confirm variable anio
		if _rc == 0 {
			rename anio Año
			exit
		}

		capture confirm variable ANIO
		if _rc == 0 {
			rename ANIO Año
			exit
		}

		capture confirm variable year
		if _rc == 0 {
			rename year Año
			exit
		}

		di as error "ERROR: No se encontró variable de año. Se esperaba Año, anio, ANIO o year."
		error 459
	end

* ------------------------------------------------------------
* 2.3 Confirmar columnas territoriales simples
* ------------------------------------------------------------

	capture program drop _confirma_geo_simple
	program define _confirma_geo_simple
		foreach v of global GEO_VARS {
			capture confirm variable `v'
			if _rc != 0 {
				di as error "ERROR: No se encontró la columna territorial `v'."
				error 459
			}
		}
	end

* ------------------------------------------------------------
* 2.4 Exportar indicador simple
* ------------------------------------------------------------

	capture program drop exporta_simple
	program define exporta_simple
		syntax, Archivo(string) Sheets(string)

		di as result "============================================================"
		di as result "Preparando indicador simple: `archivo'"
		di as result "============================================================"

		local outfile "$out_graficos/`archivo'.xlsx"
		capture erase `"`outfile'"'

		_importa_hoja_maestra, sheets("`sheets'")
		_estandariza_anio
		_confirma_geo_simple

		* Nacional: _in
		preserve
			keep Año BOLIVIA
			rename BOLIVIA Bolivia
			sort Año
			export excel using "$out_graficos/`archivo'.xlsx", sheet("_in") firstrow(variables) replace
		restore

		* Departamental: _in1
		preserve
			keep Año $DEP_VARS

			rename CHUQUISACA valor01
			rename LA_PAZ valor02
			rename COCHABAMBA valor03
			rename ORURO valor04
			rename POTOSI valor05
			rename TARIJA valor06
			rename SANTA_CRUZ valor07
			rename BENI valor08
			rename PANDO valor09

			reshape long valor, i(Año) j(codigo_geo) string

			gen str20 Departamento = ""
			replace Departamento = "Chuquisaca" if codigo_geo == "01"
			replace Departamento = "LaPaz"      if codigo_geo == "02"
			replace Departamento = "Cochabamba" if codigo_geo == "03"
			replace Departamento = "Oruro"      if codigo_geo == "04"
			replace Departamento = "Potosí"     if codigo_geo == "05"
			replace Departamento = "Tarija"     if codigo_geo == "06"
			replace Departamento = "SantaCruz"  if codigo_geo == "07"
			replace Departamento = "Beni"       if codigo_geo == "08"
			replace Departamento = "Pando"      if codigo_geo == "09"

			rename valor Valor
			gen byte orden_geo = real(codigo_geo)
			keep Año orden_geo Departamento Valor
			sort Año orden_geo
			drop orden_geo
			order Año Departamento Valor

			export excel using "$out_graficos/`archivo'.xlsx", sheet("_in1") firstrow(variables) sheetreplace
		restore

		di as result "Archivo exportado: $out_graficos/`archivo'.xlsx"
	end

* ------------------------------------------------------------
* 2.5 Exportar indicador con dos tipos en formato largo
*     _in  : Año | Tipo | Valor
*     _in1 : Año | Departamento | Tipo | Valor
* ------------------------------------------------------------

	capture program drop exporta_tipo2
	program define exporta_tipo2
		syntax, Archivo(string) Sheets(string) Suf1(string) Suf2(string) Lab1(string) Lab2(string)

		di as result "============================================================"
		di as result "Preparando indicador de dos tipos: `archivo'"
		di as result "============================================================"

		local outfile "$out_graficos/`archivo'.xlsx"
		capture erase `"`outfile'"'

		_importa_hoja_maestra, sheets("`sheets'")
		_estandariza_anio

		* Confirmar columnas nacionales.
		capture confirm variable BOLIVIA_`suf1'
		if _rc != 0 {
			di as error "ERROR: Falta BOLIVIA_`suf1' en la hoja maestra."
			error 459
		}

		capture confirm variable BOLIVIA_`suf2'
		if _rc != 0 {
			di as error "ERROR: Falta BOLIVIA_`suf2' en la hoja maestra."
			error 459
		}

		* Nacional: _in
		preserve
			keep Año BOLIVIA_`suf1' BOLIVIA_`suf2'
			rename BOLIVIA_`suf1' valor1
			rename BOLIVIA_`suf2' valor2
			reshape long valor, i(Año) j(tipo_id)
			gen str40 Tipo = ""
			replace Tipo = "`lab1'" if tipo_id == 1
			replace Tipo = "`lab2'" if tipo_id == 2
			rename valor Valor
			gen byte orden_tipo = tipo_id
			keep Año orden_tipo Tipo Valor
			sort Año orden_tipo
			drop orden_tipo
			order Año Tipo Valor
			export excel using "$out_graficos/`archivo'.xlsx", sheet("_in") firstrow(variables) replace
		restore

		* Departamental: _in1
		tempfile base_tipo2
		preserve
			clear
			set obs 0
			gen double Año = .
			gen byte orden_geo = .
			gen byte orden_tipo = .
			gen str20 Departamento = ""
			gen str40 Tipo = ""
			gen double Valor = .
			save `base_tipo2', replace
		restore

		local depvars $DEP_VARS
		local deplabs $DEP_LABS
		local n : word count `depvars'

		forvalues i = 1/`n' {
			local g : word `i' of `depvars'
			local d : word `i' of `deplabs'

			foreach s in `suf1' `suf2' {
				capture confirm variable `g'_`s'
				if _rc != 0 {
					di as error "ERROR: Falta la columna `g'_`s'."
					error 459
				}
			}

			preserve
				keep Año `g'_`suf1' `g'_`suf2'
				rename `g'_`suf1' valor1
				rename `g'_`suf2' valor2
				reshape long valor, i(Año) j(tipo_id)
				gen str20 Departamento = "`d'"
				gen str40 Tipo = ""
				replace Tipo = "`lab1'" if tipo_id == 1
				replace Tipo = "`lab2'" if tipo_id == 2
				rename valor Valor
				gen byte orden_geo = `i'
				gen byte orden_tipo = tipo_id
				keep Año orden_geo orden_tipo Departamento Tipo Valor
				append using `base_tipo2'
				save `base_tipo2', replace
			restore
		}

		use `base_tipo2', clear
		sort Año orden_geo orden_tipo
		drop orden_geo orden_tipo
		order Año Departamento Tipo Valor
		export excel using "$out_graficos/`archivo'.xlsx", sheet("_in1") firstrow(variables) sheetreplace

		di as result "Archivo exportado: $out_graficos/`archivo'.xlsx"
	end

* ------------------------------------------------------------
* 2.6 Exportar indicador con cuatro tipos en formato largo
*     Para distribución de grupos etarios.
* ------------------------------------------------------------

	capture program drop exporta_tipo4
	program define exporta_tipo4
		syntax, Archivo(string) Sheets(string) Suf1(string) Suf2(string) Suf3(string) Suf4(string) Lab1(string) Lab2(string) Lab3(string) Lab4(string)

		di as result "============================================================"
		di as result "Preparando indicador de cuatro tipos: `archivo'"
		di as result "============================================================"

		local outfile "$out_graficos/`archivo'.xlsx"
		capture erase `"`outfile'"'

		_importa_hoja_maestra, sheets("`sheets'")
		_estandariza_anio

		* Nacional: _in
		preserve
			keep Año BOLIVIA_`suf1' BOLIVIA_`suf2' BOLIVIA_`suf3' BOLIVIA_`suf4'
			rename BOLIVIA_`suf1' valor1
			rename BOLIVIA_`suf2' valor2
			rename BOLIVIA_`suf3' valor3
			rename BOLIVIA_`suf4' valor4
			reshape long valor, i(Año) j(tipo_id)
			gen str40 Tipo = ""
			replace Tipo = "`lab1'" if tipo_id == 1
			replace Tipo = "`lab2'" if tipo_id == 2
			replace Tipo = "`lab3'" if tipo_id == 3
			replace Tipo = "`lab4'" if tipo_id == 4
			rename valor Valor
			gen byte orden_tipo = tipo_id
			keep Año orden_tipo Tipo Valor
			sort Año orden_tipo
			drop orden_tipo
			order Año Tipo Valor
			export excel using "$out_graficos/`archivo'.xlsx", sheet("_in") firstrow(variables) replace
		restore

		* Departamental: _in1
		tempfile base_tipo4
		preserve
			clear
			set obs 0
			gen double Año = .
			gen byte orden_geo = .
			gen byte orden_tipo = .
			gen str20 Departamento = ""
			gen str40 Tipo = ""
			gen double Valor = .
			save `base_tipo4', replace
		restore

		local depvars $DEP_VARS
		local deplabs $DEP_LABS
		local n : word count `depvars'

		forvalues i = 1/`n' {
			local g : word `i' of `depvars'
			local d : word `i' of `deplabs'

			preserve
				keep Año `g'_`suf1' `g'_`suf2' `g'_`suf3' `g'_`suf4'
				rename `g'_`suf1' valor1
				rename `g'_`suf2' valor2
				rename `g'_`suf3' valor3
				rename `g'_`suf4' valor4
				reshape long valor, i(Año) j(tipo_id)
				gen str20 Departamento = "`d'"
				gen str40 Tipo = ""
				replace Tipo = "`lab1'" if tipo_id == 1
				replace Tipo = "`lab2'" if tipo_id == 2
				replace Tipo = "`lab3'" if tipo_id == 3
				replace Tipo = "`lab4'" if tipo_id == 4
				rename valor Valor
				gen byte orden_geo = `i'
				gen byte orden_tipo = tipo_id
				keep Año orden_geo orden_tipo Departamento Tipo Valor
				append using `base_tipo4'
				save `base_tipo4', replace
			restore
		}

		use `base_tipo4', clear
		sort Año orden_geo orden_tipo
		drop orden_geo orden_tipo
		order Año Departamento Tipo Valor
		export excel using "$out_graficos/`archivo'.xlsx", sheet("_in1") firstrow(variables) sheetreplace

		di as result "Archivo exportado: $out_graficos/`archivo'.xlsx"
	end

* ------------------------------------------------------------
* 2.7 Exportar indicador doble en columnas
*     _in  : Año | Bolivia1 | Bolivia2
*     _in1 : Año | Departamento | Valor1 | Valor2
* ------------------------------------------------------------

	capture program drop exporta_doble_columnas
	program define exporta_doble_columnas
		syntax, Archivo(string) Sheets(string) Suf1(string) Suf2(string)

		di as result "============================================================"
		di as result "Preparando indicador doble en columnas: `archivo'"
		di as result "============================================================"

		local outfile "$out_graficos/`archivo'.xlsx"
		capture erase `"`outfile'"'

		_importa_hoja_maestra, sheets("`sheets'")
		_estandariza_anio

		* Nacional: _in
		preserve
			keep Año BOLIVIA_`suf1' BOLIVIA_`suf2'
			rename BOLIVIA_`suf1' Bolivia1
			rename BOLIVIA_`suf2' Bolivia2
			sort Año
			export excel using "$out_graficos/`archivo'.xlsx", sheet("_in") firstrow(variables) replace
		restore

		* Departamental: _in1
		tempfile base_doble
		preserve
			clear
			set obs 0
			gen double Año = .
			gen byte orden_geo = .
			gen str20 Departamento = ""
			gen double Valor1 = .
			gen double Valor2 = .
			save `base_doble', replace
		restore

		local depvars $DEP_VARS
		local deplabs $DEP_LABS
		local n : word count `depvars'

		forvalues i = 1/`n' {
			local g : word `i' of `depvars'
			local d : word `i' of `deplabs'

			preserve
				keep Año `g'_`suf1' `g'_`suf2'
				rename `g'_`suf1' Valor1
				rename `g'_`suf2' Valor2
				gen byte orden_geo = `i'
				gen str20 Departamento = "`d'"
				keep Año orden_geo Departamento Valor1 Valor2
				append using `base_doble'
				save `base_doble', replace
			restore
		}

		use `base_doble', clear
		sort Año orden_geo
		drop orden_geo
		order Año Departamento Valor1 Valor2
		export excel using "$out_graficos/`archivo'.xlsx", sheet("_in1") firstrow(variables) sheetreplace

		di as result "Archivo exportado: $out_graficos/`archivo'.xlsx"
	end

* ------------------------------------------------------------
* 2.8 Exportar indicador por sexo
*     Archivo independiente:
*         _in  / _in1
*     Compatibilidad con formato antiguo:
*         _in2 / _in3
* ------------------------------------------------------------

	capture program drop exporta_sexo
	program define exporta_sexo
		syntax, Archivo(string) Sheets(string)

		di as result "============================================================"
		di as result "Preparando indicador por sexo: `archivo'"
		di as result "============================================================"

		local outfile "$out_graficos/`archivo'.xlsx"
		capture erase `"`outfile'"'

		_importa_hoja_maestra, sheets("`sheets'")
		_estandariza_anio

		* Nacional: _in e _in2
		preserve
			keep Año BOLIVIA_HOMBRE BOLIVIA_MUJER
			rename BOLIVIA_HOMBRE Bolivia1
			rename BOLIVIA_MUJER Bolivia2
			sort Año
			export excel using "$out_graficos/`archivo'.xlsx", sheet("_in") firstrow(variables) replace
			export excel using "$out_graficos/`archivo'.xlsx", sheet("_in2") firstrow(variables) sheetreplace
		restore

		* Departamental: _in1 e _in3
		tempfile base_sexo
		preserve
			clear
			set obs 0
			gen double Año = .
			gen byte orden_geo = .
			gen str20 Departamento = ""
			gen double Valor1 = .
			gen double Valor2 = .
			save `base_sexo', replace
		restore

		local depvars $DEP_VARS
		local deplabs $DEP_LABS
		local n : word count `depvars'

		forvalues i = 1/`n' {
			local g : word `i' of `depvars'
			local d : word `i' of `deplabs'

			preserve
				keep Año `g'_HOMBRE `g'_MUJER
				rename `g'_HOMBRE Valor1
				rename `g'_MUJER Valor2
				gen byte orden_geo = `i'
				gen str20 Departamento = "`d'"
				keep Año orden_geo Departamento Valor1 Valor2
				append using `base_sexo'
				save `base_sexo', replace
			restore
		}

		use `base_sexo', clear
		sort Año orden_geo
		drop orden_geo
		order Año Departamento Valor1 Valor2
		export excel using "$out_graficos/`archivo'.xlsx", sheet("_in1") firstrow(variables) sheetreplace
		export excel using "$out_graficos/`archivo'.xlsx", sheet("_in3") firstrow(variables) sheetreplace

		di as result "Archivo exportado: $out_graficos/`archivo'.xlsx"
	end

* ------------------------------------------------------------
* 2.9 Agregar hojas por sexo al mismo archivo del indicador total
*     Uso principal:
*         06_003_11.xlsx -> _in, _in1, _in2, _in3
*         06_003_12.xlsx -> _in, _in1, _in2, _in3
* ------------------------------------------------------------

	capture program drop agrega_sexo_mismo_archivo
	program define agrega_sexo_mismo_archivo
		syntax, Archivo(string) Sheets(string)

		di as result "============================================================"
		di as result "Agregando hojas por sexo al archivo: `archivo'"
		di as result "============================================================"

		local outfile "$out_graficos/`archivo'.xlsx"
		capture confirm file `"`outfile'"'
		if _rc != 0 {
			di as error "ERROR: No existe el archivo base `outfile'. Primero debe exportarse el indicador total."
			error 601
		}

		_importa_hoja_maestra, sheets("`sheets'")
		_estandariza_anio

		* Nacional por sexo: _in2
		preserve
			keep Año BOLIVIA_HOMBRE BOLIVIA_MUJER
			rename BOLIVIA_HOMBRE Bolivia1
			rename BOLIVIA_MUJER Bolivia2
			sort Año
			export excel using "$out_graficos/`archivo'.xlsx", sheet("_in2") firstrow(variables) sheetreplace
		restore

		* Departamental por sexo: _in3
		tempfile base_sexo_anexo
		preserve
			clear
			set obs 0
			gen double Año = .
			gen byte orden_geo = .
			gen str20 Departamento = ""
			gen double Valor1 = .
			gen double Valor2 = .
			save `base_sexo_anexo', replace
		restore

		local depvars $DEP_VARS
		local deplabs $DEP_LABS
		local n : word count `depvars'

		forvalues i = 1/`n' {
			local g : word `i' of `depvars'
			local d : word `i' of `deplabs'

			preserve
				keep Año `g'_HOMBRE `g'_MUJER
				rename `g'_HOMBRE Valor1
				rename `g'_MUJER Valor2
				gen byte orden_geo = `i'
				gen str20 Departamento = "`d'"
				keep Año orden_geo Departamento Valor1 Valor2
				append using `base_sexo_anexo'
				save `base_sexo_anexo', replace
			restore
		}

		use `base_sexo_anexo', clear
		sort Año orden_geo
		drop orden_geo
		order Año Departamento Valor1 Valor2
		export excel using "$out_graficos/`archivo'.xlsx", sheet("_in3") firstrow(variables) sheetreplace

		di as result "Hojas _in2 y _in3 agregadas a: $out_graficos/`archivo'.xlsx"
	end

* ============================================================
* 3. MAPEO DE INDICADORES Y EXPORTACIÓN
* ============================================================
*
* Nota:
*     Cada llamada incluye varios nombres candidatos de hoja.
*     Si el nombre real en el Excel maestro es distinto, agregarlo
*     a la lista sheets().
* ============================================================

* ------------------------------------------------------------
* 06_001 | Población grupos especiales
* ------------------------------------------------------------

	exporta_simple, archivo("06_001_01") sheets("num_mujeres_edad_fertil")
	exporta_simple, archivo("06_001_02") sheets("pct_mujeres_edad_fertil")
	exporta_simple, archivo("06_001_03") sheets("num_embarazos")
	exporta_simple, archivo("06_001_04") sheets("num_partos")
	exporta_simple, archivo("06_001_05") sheets("num_abortos")

	exporta_tipo2, archivo("06_001_06") sheets("pct_partos_abortos") suf1("PARTOS") suf2("ABORTOS") lab1("Partos") lab2("Abortos")

	exporta_simple, archivo("06_001_07") sheets("num_nacimientos")
	exporta_simple, archivo("06_001_08") sheets("num_hijos_nacidos_vivos")
	exporta_simple, archivo("06_001_09") sheets("num_hijos_nacidos_muertos")

	exporta_tipo2, archivo("06_001_10") sheets("dist_hnv_hnm") suf1("VIVOS") suf2("MUERTOS") lab1("HNV") lab2("HNM")

* ------------------------------------------------------------
* 06_002 | Salud sexual y reproductiva
* ------------------------------------------------------------

	exporta_simple, archivo("06_002_01") sheets("tasa_aborto_general")
	exporta_simple, archivo("06_002_02") sheets("tasa_fecundidad_general")
	exporta_simple, archivo("06_002_03") sheets("tasa_mortalidad_fetal")
	exporta_simple, archivo("06_002_04") sheets("tasa_embarazo_estimada")

	* ------------------------------------------------------------
* 06_003 | Demografía
* ------------------------------------------------------------

	exporta_simple, archivo("06_003_01") sheets("tasa_bruta_natalidad")
	exporta_simple, archivo("06_003_02") sheets("indice_masculinidad")
	exporta_simple, archivo("06_003_03") sheets("tasa_crecimiento_pob")

	exporta_tipo4, archivo("06_003_04") sheets("dist_grupos_etarios") suf1("0_4") suf2("5_14") suf3("15_64") suf4("65MAS") lab1("0-4años") lab2("5-14años") lab3("15-64años") lab4("65añosy+")

	exporta_simple, archivo("06_003_05") sheets("indice_dependencia_pot")
	exporta_simple, archivo("06_003_06") sheets("indice_envejecimiento")
	exporta_simple, archivo("06_003_07") sheets("indice_dependencia_juvenil")
	exporta_simple, archivo("06_003_08") sheets("indice_dependencia_senil")
	exporta_simple, archivo("06_003_09") sheets("edad_mediana_poblacion")
	exporta_simple, archivo("06_003_10") sheets("edad_media_poblacion")

* 06_003_11:
*     El script de gráficos actual espera el total y el sexo
*     en el mismo archivo:
*         _in  / _in1  = total
*         _in2 / _in3  = por sexo
	exporta_simple, archivo("06_003_11") sheets("indice_reemplazamiento_pea")
	agrega_sexo_mismo_archivo, archivo("06_003_11") sheets("irpea_sexo")

* Archivo separado opcional para la matriz/código 06_003_11a.
	exporta_sexo, archivo("06_003_11a") sheets("irpea_sexo")

* 06_003_12:
*     El script de gráficos actual espera el total y el sexo
*     en el mismo archivo:
*         _in  / _in1  = total
*         _in2 / _in3  = por sexo
	exporta_simple, archivo("06_003_12") sheets("indice_estructura_pea")
	agrega_sexo_mismo_archivo, archivo("06_003_12") sheets("iepea_sexo")

* Archivo separado opcional para la matriz/código 06_003_12a.
	exporta_sexo, archivo("06_003_12a") sheets("iepea_sexo")

	exporta_simple, archivo("06_003_13") sheets("indice_maternidad")
	exporta_simple, archivo("06_003_14") sheets("indice_friz")

	exporta_doble_columnas, archivo("06_003_15") sheets("indice_sundbarg") suf1("IS_JOVENES") suf2("IS_MAYORES")

	exporta_simple, archivo("06_003_16") sheets("indice_burgdofer")
	exporta_simple, archivo("06_003_17") sheets("indice_generacional_ancianos")
	exporta_simple, archivo("06_003_18") sheets("indice_vejez")
	exporta_simple, archivo("06_003_19") sheets("indice_sobreenvejecimiento")
	exporta_simple, archivo("06_003_20") sheets("indice_juventud")
	exporta_simple, archivo("06_003_21") sheets("indice_infancia")
	exporta_simple, archivo("06_003_22") sheets("indice_pot_reprod_fem")

* ============================================================
* 4. CIERRE
* ============================================================

	di as result "============================================================"
	di as result "EXPORTACIÓN DE EXCELS PARA GRÁFICOS FINALIZADA"
	di as result "Carpeta de salida:"
	di as result "$out_graficos"
	di as result "============================================================"

	di as result "Verificación rápida de archivos creados/actualizados:"
	local archivos 06_001_01 06_001_02 06_001_03 06_001_04 06_001_05 06_001_06 06_001_07 06_001_08 06_001_09 06_001_10 06_002_01 06_002_02 06_002_03 06_002_04 06_003_01 06_003_02 06_003_03 06_003_04 06_003_05 06_003_06 06_003_07 06_003_08 06_003_09 06_003_10 06_003_11 06_003_11a 06_003_12 06_003_12a 06_003_13 06_003_14 06_003_15 06_003_16 06_003_17 06_003_18 06_003_19 06_003_20 06_003_21 06_003_22
	foreach a of local archivos {
		capture confirm file "$out_graficos/`a'.xlsx"
		if _rc == 0 {
			di as result "OK: `a'.xlsx"
		}
		else {
			di as error "FALTA: `a'.xlsx"
		}
	}

