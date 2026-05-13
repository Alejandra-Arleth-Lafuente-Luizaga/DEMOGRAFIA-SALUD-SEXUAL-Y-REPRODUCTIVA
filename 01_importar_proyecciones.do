clear all
set more off

////////////////////////////////////////////////////////////////////////////////
//                                                                            //
//          INSTITUTO DE INVESTIGACIONES SOCIO-ECONÓMICAS IISEC               //
//                                                                            //
////////////////////////////////////////////////////////////////////////////////

**  PROYECTO: 
**			INDICADORES DE DEMOGRAFÍA, SALUD SEXUAL Y REPRODUCTIVA
**            Automatización de indicadores de Demografía, Salud Sexual y Reproductiva

**  ARCHIVO:  01_importar_proyecciones.do

**  OBJETIVO GENERAL:
**     Importar desde Excel las bases de proyecciones poblacionales preparadas y
**     convertirlas en bases .dta reproducibles.

**  OBJETIVOS ESPECÍFICOS:
**     1. Importar proyecciones de población de grupos especiales.
**     2. Importar proyecciones de población por grupos de edad en salud.
**    3. Importar proyecciones de población por edades simples y quinquenales.
**     4. Guardar las bases resultantes en:
**            _out/Bases DTA
**    5. Exportar auditorías en:
**            _out/Auditorías
**        solo cuando existan duplicados u otros problemas relevantes.

** ENTRADAS:
**     1. PROYECCIONES POBLACION GRUPOS ESPECIALES.xlsx
**     2. PROYECCIONES POBLACION GENERAL GRUPOS DE EDAD EN SALUD.xlsx
**     3. PROYECCIONES POBLACION GENERAL POR GRUPOS DE EDAD SIMPLES Y QUINQUENALES.xlsx

** SALIDAS:
**     En _out/Bases DTA:
**         1. pob_grupos_especiales_2005_20XX.dta
**         2. pob_edad_salud_2005_20XX.dta
**         3. pob_edades_simples_raw_2012_20XX.dta

**     En _out/Auditorías:
**         1. auditoria_duplicados_grupos_especiales.xlsx, si corresponde.
**         2. auditoria_duplicados_edad_salud.xlsx, si corresponde.
**         3. auditoria_missing_geo_raw_edades.xlsx, si corresponde.

** AUTORA:
**     Alejandra Arleth Lafuente-Luizaga

** FECHA DE CREACIÓN:
**     06-may-2026

** ÚLTIMA ACTUALIZACIÓN:
**     09-may-2026

** SUPUESTOS DE ESTRUCTURA:
**     - Cada archivo Excel tiene una hoja por año.
**     - Las hojas se llaman con el año: 2005, 2006, ..., 20XX.
**     - Las tres primeras filas contienen título, subtítulo y encabezados.
**    - Los datos empiezan desde la fila 4.
**     - La primera columna contiene el identificador geográfico.
**     - Los archivos fuente están en:
**         _xlsx/Demografía_Salud_Sexual_y_Reproductiva/BD PREPARADA
**     - Los nombres de archivos no usan tilde en POBLACION.

** INSTRUCCIONES PARA ACTUALIZAR A 20XX:
**     - Agregar hojas 20XX a los Excel fuente.
**     - Verificar que mantengan la misma estructura.
**     - Cambiar:
**         local yfin_esp     2024
**         local yfin_salud   2024
**        local yfin_edades  2024
**      por:
**         local yfin_esp     2026
**         local yfin_salud   2026
**         local yfin_edades  2026

** NOTAS:
**    - Este script solo importa bases fuente.
**    - No calcula indicadores.
**    - No exporta indicadores finales.
**    - Los indicadores se calculan en 03_calcular_indicadores.do.
**   - Este script usa "/" en las rutas para evitar errores con "\" en Stata.

***************************************************************************************
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

	di as result "============================================================"
	di as result "DIRECTORIOS CONFIGURADOS"
	di as result "path:         $path"
	di as result "work:         $work"
	di as result "in:           $in"
	di as result "in_graficos:  $in_graficos"
	di as result "out_dta:      $out_dta"
	di as result "out_xlsx:     $out_xlsx"
	di as result "out_graficos: $out_graficos"
	di as result "master:       $master"
	di as result "============================================================"

* ============================================================
* 1. ARCHIVOS FUENTE Y PARÁMETROS
* ============================================================

	global xlsx_prep "$xlsx/Demografía_Salud_Sexual_y_Reproductiva/BD PREPARADA"

	global file_esp    "$xlsx_prep/PROYECCIONES POBLACION GRUPOS ESPECIALES.xlsx"
	global file_salud  "$xlsx_prep/PROYECCIONES POBLACION GENERAL GRUPOS DE EDAD EN SALUD.xlsx"
	global file_edades "$xlsx_prep/PROYECCIONES POBLACION GENERAL POR GRUPOS DE EDAD SIMPLES Y QUINQUENALES.xlsx"

	local yini_esp     2005
	local yfin_esp     2024

	local yini_salud   2005
	local yfin_salud   2024

	local yini_edades  2012
	local yfin_edades  2024

* ============================================================
* 2. IMPORTAR GRUPOS ESPECIALES
* ============================================================
*
* Archivo fuente:
*     PROYECCIONES POBLACION GRUPOS ESPECIALES.xlsx
*
* Variables esperadas:
*     A: identificador geográfico
*     B: mujeres en edad fértil
*     C: embarazos
*     D: abortos
*     E: partos
*     F: nacimientos
*     G: hijos nacidos vivos
*     H: hijos nacidos muertos
*
* Salida:
*     _out/Bases DTA/pob_grupos_especiales_2005_20XX.dta
* ============================================================

	tempfile base_esp
	local first = 1

	forvalues y = `yini_esp'/`yfin_esp' {

		di as text "Importando grupos especiales: `y'"

		import excel using "$file_esp", sheet("`y'") clear allstring

		* Eliminar título, subtítulo y encabezado.
		drop in 1/3

		* Conservar columnas necesarias.
		keep A B C D E F G H

		rename A geo_raw
		rename B mef
		rename C embarazos
		rename D abortos
		rename E partos
		rename F nacimientos
		rename G hnv
		rename H hnm

		gen anio = `y'

		* Limpiar identificador geográfico.
		replace geo_raw = strtrim(geo_raw)
		drop if missing(geo_raw)

		* Convertir variables numéricas.
		foreach v in mef embarazos abortos partos nacimientos hnv hnm {
			replace `v' = subinstr(`v', ",", "", .)
			replace `v' = subinstr(`v', " ", "", .)
			destring `v', replace force
		}

		* Extraer código y nombre geográfico.
		gen codigo_geo = regexs(1) if regexm(geo_raw, "^([0-9]+)")
		gen nombre_geo = strtrim(regexr(geo_raw, "^[0-9]+\.?", ""))

		* Clasificar nivel geográfico.
		gen nivel_geo = .
		replace nivel_geo = 0 if codigo_geo == "00"
		replace nivel_geo = 1 if strlen(codigo_geo) == 2 & codigo_geo != "00"
		replace nivel_geo = 2 if strlen(codigo_geo) == 4
		replace nivel_geo = 3 if strlen(codigo_geo) == 6

		label define nivel_geo 0 "Nacional" 1 "Departamental" 2 "Provincial" 3 "Municipal", replace
		label values nivel_geo nivel_geo

		order anio codigo_geo nombre_geo nivel_geo geo_raw

		if `first' == 1 {
			save `base_esp', replace
			local first = 0
		}
		else {
			append using `base_esp'
			save `base_esp', replace
		}
	}

	use `base_esp', clear
	compress

	save "$out_dta/pob_grupos_especiales_`yini_esp'_`yfin_esp'.dta", replace


* ============================================================
* 2A. CONTROL DE DUPLICADOS: GRUPOS ESPECIALES
* ============================================================
* Llave esperada:
*     anio + codigo_geo
*
* Si existen duplicados, se exporta una auditoría.
* No se eliminan automáticamente.
* ============================================================

	use "$out_dta/pob_grupos_especiales_`yini_esp'_`yfin_esp'.dta", clear

	duplicates tag anio codigo_geo, gen(dup_key)

	count if dup_key > 0

	if r(N) > 0 {
		di as error "ADVERTENCIA: Hay duplicados en pob_grupos_especiales."

		preserve
			keep if dup_key > 0
			sort anio codigo_geo
			export excel using "$out_aud/auditoria_duplicados_grupos_especiales.xlsx", firstrow(variables) replace
		restore
	}

	drop dup_key

* ============================================================
* 3. IMPORTAR GRUPOS DE EDAD EN SALUD
* ============================================================
* Archivo fuente:
*     PROYECCIONES POBLACION GENERAL GRUPOS DE EDAD EN SALUD.xlsx
*
* Nota:
*     La estructura cambia entre años antiguos y recientes.
*     Por eso las columnas adicionales se renombran como col_h, col_i, etc.
*
* Salida:
*     _out/Bases DTA/pob_edad_salud_2005_20XX.dta
* ============================================================

	tempfile base_salud
	local first = 1

	forvalues y = `yini_salud'/`yfin_salud' {

		di as text "Importando grupos de edad en salud: `y'"

		import excel using "$file_salud", sheet("`y'") clear allstring

		* Eliminar título, subtítulo y encabezado.
		drop in 1/3

		rename A geo_raw

		gen anio = `y'

		* Renombrar columnas comunes.
		capture rename B edad_0
		capture rename C edad_1
		capture rename D edad_2
		capture rename E edad_3
		capture rename F edad_4
		capture rename G grupo_0_4

		* Columnas adicionales según estructura anual.
		capture rename H col_h
		capture rename I col_i
		capture rename J col_j
		capture rename K col_k
		capture rename L col_l
		capture rename M col_m
		capture rename N col_n
		capture rename O col_o

		* Limpiar identificador geográfico.
		replace geo_raw = strtrim(geo_raw)
		drop if missing(geo_raw)

		* Extraer código y nombre geográfico.
		gen codigo_geo = regexs(1) if regexm(geo_raw, "^([0-9]+)")
		gen nombre_geo = strtrim(regexr(geo_raw, "^[0-9]+\.?", ""))

		* Normalizar códigos 0, 1, ..., 9 a 00, 01, ..., 09.
		gen codigo_num = real(codigo_geo)
		replace codigo_geo = string(codigo_num, "%02.0f") if !missing(codigo_num) & codigo_num < 10
		drop codigo_num

		* Clasificar nivel geográfico.
		gen nivel_geo = .
		replace nivel_geo = 0 if codigo_geo == "00"
		replace nivel_geo = 1 if strlen(codigo_geo) == 2 & codigo_geo != "00"
		replace nivel_geo = 2 if strlen(codigo_geo) == 4
		replace nivel_geo = 3 if strlen(codigo_geo) == 6

		label define nivel_geo 0 "Nacional" 1 "Departamental" 2 "Provincial" 3 "Municipal", replace
		label values nivel_geo nivel_geo

		* Convertir a numéricas todas las columnas de población.
		ds anio codigo_geo nombre_geo nivel_geo geo_raw, not
		local numvars `r(varlist)'

		foreach v of local numvars {
			replace `v' = subinstr(`v', ",", "", .)
			replace `v' = subinstr(`v', " ", "", .)
			destring `v', replace force
		}

		order anio codigo_geo nombre_geo nivel_geo geo_raw

		if `first' == 1 {
			save `base_salud', replace
			local first = 0
		}
		else {
			append using `base_salud'
			save `base_salud', replace
		}
	}

	use `base_salud', clear

	compress

	save "$out_dta/pob_edad_salud_`yini_salud'_`yfin_salud'.dta", replace

* ============================================================
* 3A. CONTROL DE DUPLICADOS: GRUPOS DE EDAD EN SALUD
* ============================================================

	use "$out_dta/pob_edad_salud_`yini_salud'_`yfin_salud'.dta", clear

	duplicates tag anio codigo_geo, gen(dup_key)

	count if dup_key > 0

	if r(N) > 0 {

		di as error "ADVERTENCIA: Hay duplicados en pob_edad_salud."

		preserve
			keep if dup_key > 0
			sort anio codigo_geo
			export excel using "$out_aud/auditoria_duplicados_edad_salud.xlsx", firstrow(variables) replace
		restore
	}

	drop dup_key

* ============================================================
* 3B. CREAR BASE AUXILIAR DE POBLACIÓN TOTAL
* ============================================================
*
* Archivo fuente:
*     PROYECCIONES POBLACION GENERAL GRUPOS DE EDAD EN SALUD.xlsx
*
* Objetivo:
*     Extraer la población total desde la última columna con dato
*     de cada fila geográfica.
*
* Justificación:
*     En el Excel, la columna "Total general" cambia de posición
*     entre años. Para 2005–2011 aparece en L y para 2012–20XX
*     aparece en O. Por eso no se usa una letra fija.
*
* Resultado:
*     _out/Bases DTA/pob_total_salud_2005_20XX.dta
*
* Variables principales:
*     anio
*     codigo_geo
*     nombre_geo
*     pob_total
*
* Universo:
*     Se conserva toda la geografía disponible.
*     Los indicadores luego filtran Bolivia y departamentos.
* ============================================================

	tempfile pob_total_salud
	local first_pob = 1

	forvalues y = `yini_salud'/`yfin_salud' {

		di as result "Extrayendo población total desde Excel de salud: `y'"

		import excel using "$file_salud", sheet("`y'") clear allstring

		gen fila_excel = _n

		rename A geo_raw

		replace geo_raw = strtrim(geo_raw)
		replace geo_raw = subinstr(geo_raw, char(160), " ", .)
		replace geo_raw = strtrim(geo_raw)

* ------------------------------------------------------------
* Identificar último dato no vacío hacia la derecha
* ------------------------------------------------------------

		ds geo_raw fila_excel, not
		local columnas `r(varlist)'

		gen strL valor_ultima_col = ""
		gen str32 var_ultima_col = ""
		gen strL encabezado_ultima_col = ""

		foreach v of varlist `columnas' {

			capture confirm string variable `v'

			if _rc == 0 {
				replace valor_ultima_col = `v' if strtrim(`v') != ""
				replace var_ultima_col = "`v'" if strtrim(`v') != ""
				replace encabezado_ultima_col = strtrim(`v'[3]) if strtrim(`v') != ""
			}
		}

* ------------------------------------------------------------
* Mantener solo filas geográficas
* ------------------------------------------------------------

		keep if regexm(geo_raw, "^[ ]*[0-9]+")

		gen codigo_geo = regexs(1) if regexm(geo_raw, "^[ ]*([0-9]+)")
		gen nombre_geo = strtrim(regexr(geo_raw, "^[ ]*[0-9]+\.?", ""))

		* Normalizar códigos 0, 1, ..., 9 a 00, 01, ..., 09.
		gen codigo_num = real(codigo_geo)
		replace codigo_geo = string(codigo_num, "%02.0f") if !missing(codigo_num) & codigo_num < 10
		drop codigo_num

		gen anio = `y'

* ------------------------------------------------------------
* Convertir población total a número
* ------------------------------------------------------------

		gen strL pob_total_txt = valor_ultima_col

		replace pob_total_txt = strtrim(pob_total_txt)
		replace pob_total_txt = subinstr(pob_total_txt, char(160), "", .)
		replace pob_total_txt = subinstr(pob_total_txt, " ", "", .)

		gen n_puntos = length(pob_total_txt) - length(subinstr(pob_total_txt, ".", "", .))
		gen n_comas  = length(pob_total_txt) - length(subinstr(pob_total_txt, ",", "", .))

		gen pos_punto = strpos(pob_total_txt, ".")
		gen pos_coma  = strpos(pob_total_txt, ",")

		gen despues_punto = length(pob_total_txt) - pos_punto
		gen despues_coma  = length(pob_total_txt) - pos_coma

		gen strL pob_total_clean = pob_total_txt

		* Miles con puntos: 12.169.501
		replace pob_total_clean = subinstr(pob_total_clean, ".", "", .) if n_puntos >= 2 & n_comas == 0

		* Un solo punto como separador de miles: 123.456
		replace pob_total_clean = subinstr(pob_total_clean, ".", "", .) if n_puntos == 1 & n_comas == 0 & despues_punto == 3

		* Miles con comas: 12,169,501
		replace pob_total_clean = subinstr(pob_total_clean, ",", "", .) if n_comas >= 2 & n_puntos == 0

		* Una sola coma como separador de miles: 123,456
		replace pob_total_clean = subinstr(pob_total_clean, ",", "", .) if n_comas == 1 & n_puntos == 0 & despues_coma == 3

		* Coma decimal: 9294063,16
		replace pob_total_clean = subinstr(pob_total_clean, ",", ".", .) if n_comas == 1 & n_puntos == 0 & despues_coma != 3

		* Formato europeo: 9.294.063,16
		replace pob_total_clean = subinstr(pob_total_clean, ".", "", .) if n_puntos > 0 & n_comas > 0 & pos_punto < pos_coma
		replace pob_total_clean = subinstr(pob_total_clean, ",", ".", .) if n_puntos > 0 & n_comas > 0 & pos_punto < pos_coma

		* Formato inglés: 9,294,063.16
		replace pob_total_clean = subinstr(pob_total_clean, ",", "", .) if n_puntos > 0 & n_comas > 0 & pos_coma < pos_punto

		destring pob_total_clean, gen(pob_total) force

		drop pob_total_txt pob_total_clean n_puntos n_comas pos_punto pos_coma despues_punto despues_coma

		label variable pob_total "Población total desde última columna con dato"

		* Clasificar nivel geográfico.
		gen nivel_geo = .
		replace nivel_geo = 0 if codigo_geo == "00"
		replace nivel_geo = 1 if strlen(codigo_geo) == 2 & codigo_geo != "00"
		replace nivel_geo = 2 if strlen(codigo_geo) == 4
		replace nivel_geo = 3 if strlen(codigo_geo) == 6

		label define nivel_geo 0 "Nacional" 1 "Departamental" 2 "Provincial" 3 "Municipal", replace
		label values nivel_geo nivel_geo

		keep anio codigo_geo nombre_geo nivel_geo geo_raw var_ultima_col encabezado_ultima_col valor_ultima_col pob_total

		order anio codigo_geo nombre_geo nivel_geo geo_raw var_ultima_col encabezado_ultima_col valor_ultima_col pob_total

		if `first_pob' == 1 {
			save `pob_total_salud', replace
			local first_pob = 0
		}
		else {
			append using `pob_total_salud'
			save `pob_total_salud', replace
		}
	}

	use `pob_total_salud', clear

	compress

* ------------------------------------------------------------
* 3B.1 CONTROL Y CORRECCIÓN DE ESCALA: POB_TOTAL
* ------------------------------------------------------------
* Problema detectado:
*     Algunas filas departamentales pueden quedar multiplicadas
*     por 1.000 por problemas de conversión decimal en el Excel fuente.
*
* Criterio:
*     Ningún territorio subnacional debe tener población total mayor
*     que Bolivia nacional en el mismo año.
*
* Corrección:
*     Si pob_total > pob_total_bolivia y pob_total / 1000 queda en
*     una escala plausible, se corrige:
*
*         pob_total = pob_total / 1000
*
* Ejemplo detectado:
*     2,630,380,725  ->  2,630,380.725
* ------------------------------------------------------------

	gen double aux_pob_bolivia = pob_total if codigo_geo == "00"

	bysort anio: egen double pob_total_bolivia = max(aux_pob_bolivia)

	drop aux_pob_bolivia

	gen byte error_escala_pob_total = 0

	replace error_escala_pob_total = 1 if nivel_geo != 0 & pob_total > pob_total_bolivia & (pob_total / 1000) < pob_total_bolivia & !missing(pob_total, pob_total_bolivia)

	count if error_escala_pob_total == 1

	di as result "Casos corregidos por posible error de escala en pob_total: " r(N)

	if r(N) > 0 {

		preserve

			keep if error_escala_pob_total == 1

			gen double pob_total_original = pob_total
			gen double pob_total_corregido = pob_total / 1000

			format pob_total_bolivia pob_total_original pob_total_corregido %15.2fc

			sort anio codigo_geo

			export excel using "$out_aud/auditoria_pob_total_error_escala.xlsx", firstrow(variables) replace

		restore

		replace pob_total = pob_total / 1000 if error_escala_pob_total == 1
	}

	gen byte error_pob_total_post = 0

	replace error_pob_total_post = 1 if nivel_geo != 0 & pob_total > pob_total_bolivia & !missing(pob_total, pob_total_bolivia)

	count if error_pob_total_post == 1

	if r(N) > 0 {

		di as error "ERROR: Persisten territorios con población total mayor que Bolivia."

		preserve
			keep if error_pob_total_post == 1
			sort anio codigo_geo
			export excel using "$out_aud/error_pob_total_departamento_mayor_bolivia.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	drop pob_total_bolivia error_escala_pob_total error_pob_total_post

* ------------------------------------------------------------
* 3B.2 Control de población total faltante
* ------------------------------------------------------------

	count if missing(pob_total) & inlist(nivel_geo, 0, 1)

	if r(N) > 0 {

		di as error "ADVERTENCIA: Hay pob_total faltante para Bolivia o departamentos."

		preserve
			keep if missing(pob_total) & inlist(nivel_geo, 0, 1)
			sort anio codigo_geo
			export excel using "$out_aud/auditoria_missing_pob_total_salud.xlsx", firstrow(variables) replace
		restore
	}

	save "$out_dta/pob_total_salud_`yini_salud'_`yfin_salud'.dta", replace

	export excel using "$out_aud/auditoria_pob_total_salud_`yini_salud'_`yfin_salud'.xlsx", firstrow(variables) replace

* ============================================================
* 4. IMPORTAR EDADES SIMPLES Y QUINQUENALES
* ============================================================
*
* Archivo fuente:
*     PROYECCIONES POBLACION GENERAL POR GRUPOS DE EDAD SIMPLES Y QUINQUENALES.xlsx
*
* Nota:
*     Esta base queda en formato raw porque contiene una estructura jerárquica:
*         unidad geográfica
*         1.Mujer
*         2.Hombre
*
*     La limpieza definitiva se realiza en:
*         02_limpiar_edades_hombre_mujer.do
*
* Salida:
*     _out/Bases DTA/pob_edades_simples_raw_2012_20XX.dta
* ============================================================

	tempfile base_edades
	local first = 1

	forvalues y = `yini_edades'/`yfin_edades' {

		di as text "Importando edades simples y quinquenales: `y'"

		import excel using "$file_edades", sheet("`y'") clear allstring

		* Eliminar título, subtítulo y encabezado.
		drop in 1/3

		rename A geo_raw

		gen anio = `y'

		* Limpiar identificador.
		replace geo_raw = strtrim(geo_raw)
		drop if missing(geo_raw)

		* Convertir a numéricas todas las columnas excepto identificadores.
		ds anio geo_raw, not
		local numvars `r(varlist)'

		foreach v of local numvars {
			replace `v' = subinstr(`v', ",", "", .)
			replace `v' = subinstr(`v', " ", "", .)
			destring `v', replace force
		}

		order anio geo_raw

		if `first' == 1 {
			save `base_edades', replace
			local first = 0
		}
		else {
			append using `base_edades'
			save `base_edades', replace
		}
	}

	use `base_edades', clear
	compress

	save "$out_dta/pob_edades_simples_raw_`yini_edades'_`yfin_edades'.dta", replace

* ============================================================
* 4A. CONTROL BÁSICO: EDADES SIMPLES RAW
* ============================================================
* Esta base puede tener más de una fila por anio + geo_raw porque contiene:
*     unidad geográfica
*     1.Mujer
*     2.Hombre
*
* Por tanto, no se exige unicidad anio + geo_raw en esta etapa.
* ============================================================

	use "$out_dta/pob_edades_simples_raw_`yini_edades'_`yfin_edades'.dta", clear

	count if missing(geo_raw)

	if r(N) > 0 {
		di as error "ADVERTENCIA: Hay filas sin geo_raw en edades simples raw."

		preserve
			keep if missing(geo_raw)
			export excel using "$out_aud/auditoria_missing_geo_raw_edades.xlsx", firstrow(variables) replace
		restore
	}

* ============================================================
* 5. CONTROLES MÍNIMOS DE IMPORTACIÓN
* ============================================================

	di as result "============================================================"
	di as result "CONTROL DE IMPORTACIÓN"
	di as result "============================================================"

	use "$out_dta/pob_grupos_especiales_`yini_esp'_`yfin_esp'.dta", clear

	di as text "Base: grupos especiales"
	describe
	tab anio
	tab nivel_geo, missing

	use "$out_dta/pob_edad_salud_`yini_salud'_`yfin_salud'.dta", clear

	di as text "Base: edad en salud"
	describe
	tab anio
	tab nivel_geo, missing

	use "$out_dta/pob_edades_simples_raw_`yini_edades'_`yfin_edades'.dta", clear

	di as text "Base: edades simples y quinquenales raw"
	describe
	tab anio

	di as result "============================================================"
	di as result "BASES IMPORTADAS Y GUARDADAS EN:"
	di as result "$out_dta"
	di as result "FIN DEL SCRIPT 01_importar_proyecciones.do"
	di as result "============================================================"