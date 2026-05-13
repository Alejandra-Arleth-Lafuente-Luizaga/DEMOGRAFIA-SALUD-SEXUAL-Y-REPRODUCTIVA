capture restore
clear all
set more off

////////////////////////////////////////////////////////////////////////////////
//                                                                            //
//          INSTITUTO DE INVESTIGACIONES SOCIO-ECONÓMICAS IISEC               //
//                                                                            //
////////////////////////////////////////////////////////////////////////////////

** PROYECTO:
**     Automatización de indicadores de Demografía, Salud Sexual y Reproductiva

** ARCHIVO:
**     03_calcular_indicadores.do

** OBJETIVO GENERAL:
**     Calcular los indicadores oficiales de Demografía, Salud Sexual y Reproductiva
**     a partir de las bases limpias generadas por:

**         01_importar_proyecciones.do
**         02_limpiar_edades_hombre_mujer.do

** OBJETIVOS ESPECÍFICOS:
**     1. Leer bases limpias desde:_out/Bases DTA
**     2. Calcular indicadores en el orden del Excel/diccionario.
**     3. Exportar todos los indicadores a un único archivo Excel:
**            _out/Indicadores Excel/indicadores_demografia_ssr_2005_20XX.xlsx
**     4. Guardar cada indicador en una pestaña distinta del mismo libro Excel.
**     5. No guardar bases .dta de indicadores.

** ESTRUCTURA DEL EXCEL FINAL:
**     Cada indicador se exporta como una hoja independiente:

**         num_mujeres_edad_fertil
**         pct_mujeres_edad_fertil
**         ...

** INDICADORES INCLUIDOS HASTA AHORA:

**     BLOQUE 01:
**         06_001_01 | Número de Mujeres en Edad Fértil

**     BLOQUE 02:
**         06_001_02 | Proporción de Mujeres en Edad Fértil
**		...

** AUTORA:
**     Alejandra Arleth Lafuente-Luizaga

** FECHA DE CREACIÓN:
**     06-may-2026

** ÚLTIMA ACTUALIZACIÓN:
**     08-may-2026

** NOTAS IMPORTANTES:
**     - Este script debe ejecutarse completo, no por bloques sueltos.
**     - Si se ejecuta por bloques sueltos, pueden perderse las rutas globales.
**    - El primer indicador usa replace porque crea el Excel desde cero.
**     - Los siguientes indicadores usan sheetreplace para no borrar las demás hojas.
**     - Para agregar nuevos indicadores, copiar la plantilla de cualquier indicador.

*************************************************************************************

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
	di as result "out_dta:      $out_dta"
	di as result "out_xlsx:     $out_xlsx"
	di as result "out_graficos: $out_graficos"
	di as result "master:       $master"
	di as result "============================================================"

* ============================================================
* 0A. RANGOS TEMPORALES POR TIPO DE INDICADOR
* ============================================================

	global yini_esp 2005
	global yfin_esp 2024

	global yini_edades 2012
	global yfin_edades 2024

* ============================================================
* 1. VERIFICAR INSUMOS NECESARIOS
* ============================================================

	di as result "============================================================"
	di as result "VERIFICANDO INSUMOS PARA INDICADORES"
	di as result "============================================================"

	confirm file "$out_dta/pob_grupos_especiales_2005_2024.dta"
	confirm file "$out_dta/pob_edades_simples_long_2012_2024.dta"

	di as result "Insumos encontrados correctamente."

* ============================================================
**# BLOQUE 01: NÚMERO DE MUJERES EN EDAD FÉRTIL
* ============================================================
* Código:
*     06_001_01
*
* Indicador:
*     Número de Mujeres en Edad Fértil
*
* Fórmula:
*     valor = mef
*
* Numerador:
*     mef
*
* Denominador:
*     No aplica.
*
* Fuente:
*     _out/Bases DTA/pob_grupos_especiales_2005_20XX.dta
*
* Universo:
*     Bolivia y departamentos.
*
* Consideraciones:
*     - Este indicador se reporta como número absoluto.
*     - Es el primer indicador del archivo Excel, por eso usa replace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 1A. Abrir base fuente
* ------------------------------------------------------------

	use "$out_dta/pob_grupos_especiales_2005_2024.dta", clear

	keep if inrange(anio, $yini_esp, $yfin_esp)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

* ------------------------------------------------------------
* 1B. Control de duplicados
* ------------------------------------------------------------

	duplicates tag anio codigo_geo, gen(dup_key)

	count if dup_key > 0

	if r(N) > 0 {
		di as error "ERROR: Hay duplicados en MEF por anio + codigo_geo."

		preserve
			keep if dup_key > 0
			sort anio codigo_geo
			export excel using "$out_aud/auditoria_duplicados_mef.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	drop dup_key

	isid anio codigo_geo

* ------------------------------------------------------------
* 1C. Calcular indicador
* ------------------------------------------------------------

	gen valor = mef

* ------------------------------------------------------------
* 1D. Preparar tabla final
* ------------------------------------------------------------

	keep anio codigo_geo valor

	reshape wide valor, i(anio) j(codigo_geo) string

	rename valor00 BOLIVIA
	rename valor01 CHUQUISACA
	rename valor02 LA_PAZ
	rename valor03 COCHABAMBA
	rename valor04 ORURO
	rename valor05 POTOSI
	rename valor06 TARIJA
	rename valor07 SANTA_CRUZ
	rename valor08 BENI
	rename valor09 PANDO

	order anio BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO

	sort anio

	format BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO %15.0fc

* ------------------------------------------------------------
* 1E. Exportar hoja al Excel final
* ------------------------------------------------------------
* Este es el primer indicador.
* Usamos replace para crear/reemplazar el libro completo.
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("num_mujeres_edad_fertil") firstrow(variables) replace

* ------------------------------------------------------------
* 1F. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_001_01 | Número de Mujeres en Edad Fértil"
	di as result "Hoja: num_mujeres_edad_fertil"
	di as result "============================================================"

	list, clean noobs

* ============================================================
**# BLOQUE 02: PROPORCIÓN DE MUJERES EN EDAD FÉRTIL
* ============================================================
* Código:
*     06_001_02
*
* Indicador:
*     Proporción de Mujeres en Edad Fértil
*
* Fórmula:
*     valor = mef / mujer_total * 100
*
* Fuente del numerador:
*     _out/Bases DTA/pob_grupos_especiales_2005_20XX.dta
*
* Fuente del denominador:
*     _out/Bases DTA/pob_edades_simples_long_2012_20XX.dta
*
* Rango temporal:
*     2012–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Consideraciones:
*     - Este indicador mide el porcentaje de mujeres en edad fértil
*       sobre el total de mujeres.
*     - No usa población total como denominador.
*     - Como este es el segundo indicador, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 2A. Construir denominador: total de mujeres
* ------------------------------------------------------------

	tempfile mujer_total

	use "$out_dta/pob_edades_simples_long_2012_2024.dta", clear

	keep if inrange(anio, $yini_edades, $yfin_edades)
	keep if sexo == "Mujer"
	keep if grupo_edad == "total"

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

	keep anio codigo_geo poblacion

	rename poblacion mujer_total

	isid anio codigo_geo

	save `mujer_total', replace

* ------------------------------------------------------------
* 2B. Abrir numerador: mujeres en edad fértil
* ------------------------------------------------------------
* IMPORTANTE:
*     El numerador mef NO está en la base de edades simples.
*     Está en pob_grupos_especiales_2005_20XX.dta.
* ------------------------------------------------------------

	use "$out_dta/pob_grupos_especiales_2005_2024.dta", clear

	keep if inrange(anio, $yini_edades, $yfin_edades)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

* ------------------------------------------------------------
* 2C. Control de duplicados
* ------------------------------------------------------------

	duplicates tag anio codigo_geo, gen(dup_key)

	count if dup_key > 0

	if r(N) > 0 {
		di as error "ERROR: Hay duplicados en el indicador 06_001_02."

		preserve
			keep if dup_key > 0
			sort anio codigo_geo
			export excel using "$out_aud/auditoria_duplicados_pct_mujeres_edad_fertil.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	drop dup_key

	isid anio codigo_geo

* ------------------------------------------------------------
* 2D. Unir numerador y denominador
* ------------------------------------------------------------

	merge 1:1 anio codigo_geo using `mujer_total'

	gen merge_mujer_total = _merge

	preserve
		keep if merge_mujer_total != 3
		count

		if r(N) > 0 {
			export excel using "$out_aud/auditoria_merge_pct_mujeres_edad_fertil.xlsx", firstrow(variables) replace
		}
	restore

	keep if merge_mujer_total == 3

	drop _merge

* ------------------------------------------------------------
* 2E. Calcular indicador
* ------------------------------------------------------------

	gen valor = .
	replace valor = (mef / mujer_total) * 100 if mujer_total > 0 & !missing(mef)

	replace valor = round(valor, .01)

* ------------------------------------------------------------
* 2F. Preparar tabla final
* ------------------------------------------------------------

	keep anio codigo_geo valor

	reshape wide valor, i(anio) j(codigo_geo) string

	rename valor00 BOLIVIA
	rename valor01 CHUQUISACA
	rename valor02 LA_PAZ
	rename valor03 COCHABAMBA
	rename valor04 ORURO
	rename valor05 POTOSI
	rename valor06 TARIJA
	rename valor07 SANTA_CRUZ
	rename valor08 BENI
	rename valor09 PANDO

	order anio BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO

	sort anio

	format BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO %9.2f

* ------------------------------------------------------------
* 2G. Exportar hoja al Excel final
* ------------------------------------------------------------
* Como este es el segundo indicador, usamos sheetreplace para
* no borrar la hoja del primer indicador.
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("pct_mujeres_edad_fertil") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 2H. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_001_02 | Proporción de Mujeres en Edad Fértil"
	di as result "Hoja: pct_mujeres_edad_fertil"
	di as result "============================================================"

	list, clean noobs

* ============================================================
**# BLOQUE 03: NÚMERO DE EMBARAZOS
* ============================================================
* Código:
*     06_001_03
*
* Indicador:
*     Número de embarazos
*
* Fórmula:
*     valor = embarazos
*
* Fuente:
*     _out/Bases DTA/pob_grupos_especiales_2005_20XX.dta
*
* Rango temporal:
*     2005–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Consideraciones:
*     - Este indicador se reporta como número absoluto.
*     - Se filtra solo nivel nacional y departamental.
*     - Como es un indicador posterior al primero, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 3A. Abrir base fuente
* ------------------------------------------------------------

	use "$out_dta/pob_grupos_especiales_2005_2024.dta", clear

	keep if inrange(anio, 2005, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

* ------------------------------------------------------------
* 3B. Control de duplicados
* ------------------------------------------------------------

	duplicates tag anio codigo_geo, gen(dup_key)

	count if dup_key > 0

	if r(N) > 0 {
		di as error "ERROR: Hay duplicados en embarazos por anio + codigo_geo."

		preserve
			keep if dup_key > 0
			sort anio codigo_geo
			export excel using "$out_aud/auditoria_duplicados_embarazos.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	drop dup_key

	isid anio codigo_geo

* ------------------------------------------------------------
* 3C. Calcular indicador
* ------------------------------------------------------------

	gen valor = embarazos

* ------------------------------------------------------------
* 3D. Preparar tabla final
* ------------------------------------------------------------

	keep anio codigo_geo valor

	reshape wide valor, i(anio) j(codigo_geo) string

	rename valor00 BOLIVIA
	rename valor01 CHUQUISACA
	rename valor02 LA_PAZ
	rename valor03 COCHABAMBA
	rename valor04 ORURO
	rename valor05 POTOSI
	rename valor06 TARIJA
	rename valor07 SANTA_CRUZ
	rename valor08 BENI
	rename valor09 PANDO

	order anio BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO

	sort anio

	format BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO %15.0fc

* ------------------------------------------------------------
* 3E. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("num_embarazos") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 3F. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_001_03 | Número de embarazos"
	di as result "Hoja: num_embarazos"
	di as result "============================================================"

	list, clean noobs

* ============================================================
**# BLOQUE 04: NÚMERO DE PARTOS
* ============================================================
* Código:
*     06_001_04
*
* Indicador:
*     Número de partos
*
* Fórmula:
*     valor = partos
*
* Fuente:
*     _out/Bases DTA/pob_grupos_especiales_2005_20XX.dta
*
* Rango temporal:
*     2005–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Consideraciones:
*     - Este indicador se reporta como número absoluto.
*     - Se filtra solo nivel nacional y departamental.
*     - Como es un indicador posterior al primero, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 4A. Abrir base fuente
* ------------------------------------------------------------

	use "$out_dta/pob_grupos_especiales_2005_2024.dta", clear

	keep if inrange(anio, 2005, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

* ------------------------------------------------------------
* 4B. Control de duplicados
* ------------------------------------------------------------

	duplicates tag anio codigo_geo, gen(dup_key)

	count if dup_key > 0

	if r(N) > 0 {
		di as error "ERROR: Hay duplicados en partos por anio + codigo_geo."

		preserve
			keep if dup_key > 0
			sort anio codigo_geo
			export excel using "$out_aud/auditoria_duplicados_partos.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	drop dup_key

	isid anio codigo_geo

* ------------------------------------------------------------
* 4C. Calcular indicador
* ------------------------------------------------------------

	gen valor = partos

* ------------------------------------------------------------
* 4D. Preparar tabla final
* ------------------------------------------------------------

	keep anio codigo_geo valor

	reshape wide valor, i(anio) j(codigo_geo) string

	rename valor00 BOLIVIA
	rename valor01 CHUQUISACA
	rename valor02 LA_PAZ
	rename valor03 COCHABAMBA
	rename valor04 ORURO
	rename valor05 POTOSI
	rename valor06 TARIJA
	rename valor07 SANTA_CRUZ
	rename valor08 BENI
	rename valor09 PANDO

	order anio BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO

	sort anio

	format BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO %15.0fc

* ------------------------------------------------------------
* 4E. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("num_partos") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 4F. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_001_04 | Número de partos"
	di as result "Hoja: num_partos"
	di as result "============================================================"

	list, clean noobs

* ============================================================
**# BLOQUE 05: NÚMERO DE ABORTOS
* ============================================================
* Código:
*     06_001_05
*
* Indicador:
*     Número de abortos
*
* Fórmula:
*     valor = abortos
*
* Fuente:
*     _out/Bases DTA/pob_grupos_especiales_2005_20XX.dta
*
* Rango temporal:
*     2005–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Consideraciones:
*     - Este indicador se reporta como número absoluto.
*     - Se filtra solo nivel nacional y departamental.
*     - Como es un indicador posterior al primero, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 5A. Abrir base fuente
* ------------------------------------------------------------

	use "$out_dta/pob_grupos_especiales_2005_2024.dta", clear

	keep if inrange(anio, 2005, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

* ------------------------------------------------------------
* 5B. Control de duplicados
* ------------------------------------------------------------

	duplicates tag anio codigo_geo, gen(dup_key)

	count if dup_key > 0

	if r(N) > 0 {
		di as error "ERROR: Hay duplicados en abortos por anio + codigo_geo."

		preserve
			keep if dup_key > 0
			sort anio codigo_geo
			export excel using "$out_aud/auditoria_duplicados_abortos.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	drop dup_key

	isid anio codigo_geo

* ------------------------------------------------------------
* 5C. Calcular indicador
* ------------------------------------------------------------

	gen valor = abortos

* ------------------------------------------------------------
* 5D. Preparar tabla final
* ------------------------------------------------------------

	keep anio codigo_geo valor

	reshape wide valor, i(anio) j(codigo_geo) string

	rename valor00 BOLIVIA
	rename valor01 CHUQUISACA
	rename valor02 LA_PAZ
	rename valor03 COCHABAMBA
	rename valor04 ORURO
	rename valor05 POTOSI
	rename valor06 TARIJA
	rename valor07 SANTA_CRUZ
	rename valor08 BENI
	rename valor09 PANDO

	order anio BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO

	sort anio

	format BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO %15.0fc

* ------------------------------------------------------------
* 5E. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("num_abortos") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 5F. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_001_05 | Número de abortos"
	di as result "Hoja: num_abortos"
	di as result "============================================================"

	list, clean noobs

* ============================================================
**# BLOQUE 05: NÚMERO DE ABORTOS
* ============================================================
* Código:
*     06_001_05
*
* Indicador:
*     Número de abortos
*
* Fórmula:
*     valor = abortos
*
* Fuente:
*     _out/Bases DTA/pob_grupos_especiales_2005_20XX.dta
*
* Rango temporal:
*     2005–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Consideraciones:
*     - Este indicador se reporta como número absoluto.
*     - Se filtra solo nivel nacional y departamental.
*     - Como es un indicador posterior al primero, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 5A. Abrir base fuente
* ------------------------------------------------------------

	use "$out_dta/pob_grupos_especiales_2005_2024.dta", clear

	keep if inrange(anio, 2005, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

* ------------------------------------------------------------
* 5B. Control de duplicados
* ------------------------------------------------------------

	duplicates tag anio codigo_geo, gen(dup_key)

	count if dup_key > 0

	if r(N) > 0 {
		di as error "ERROR: Hay duplicados en abortos por anio + codigo_geo."

		preserve
			keep if dup_key > 0
			sort anio codigo_geo
			export excel using "$out_aud/auditoria_duplicados_abortos.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	drop dup_key

	isid anio codigo_geo

* ------------------------------------------------------------
* 5C. Calcular indicador
* ------------------------------------------------------------

	gen valor = abortos

* ------------------------------------------------------------
* 5D. Preparar tabla final
* ------------------------------------------------------------

	keep anio codigo_geo valor

	reshape wide valor, i(anio) j(codigo_geo) string

	rename valor00 BOLIVIA
	rename valor01 CHUQUISACA
	rename valor02 LA_PAZ
	rename valor03 COCHABAMBA
	rename valor04 ORURO
	rename valor05 POTOSI
	rename valor06 TARIJA
	rename valor07 SANTA_CRUZ
	rename valor08 BENI
	rename valor09 PANDO

	order anio BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO

	sort anio

	format BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO %15.0fc

* ------------------------------------------------------------
* 5E. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("num_abortos") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 5F. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_001_05 | Número de abortos"
	di as result "Hoja: num_abortos"
	di as result "============================================================"

	list, clean noobs

* ============================================================
**# BLOQUE 06: PORCENTAJE DE PARTOS Y ABORTOS SOBRE EMBARAZOS
* ============================================================
* Código:
*     06_001_06
*
* Indicador:
*     Distribución porcentual de partos y abortos respecto al total
*     reconstruido de embarazos.
*
* Fórmulas:
*     pct_partos  = partos  / (partos + abortos) * 100
*     pct_abortos = abortos / (partos + abortos) * 100
*
* Denominador:
*     embarazos_calc = partos + abortos
*
* Fuente:
*     _out/Bases DTA/pob_grupos_especiales_2005_20XX.dta
*
* Rango temporal:
*     2005–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Consideraciones:
*     - Este indicador muestra la composición porcentual de los embarazos
*       según resultado: partos y abortos.
*     - El denominador se reconstruye como partos + abortos.
*     - No se usa directamente la variable embarazos.
*     - La suma de pct_partos y pct_abortos debe ser igual a 100.
*     - Como es un indicador posterior al primero, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 6A. Abrir base fuente
* ------------------------------------------------------------

	use "$out_dta/pob_grupos_especiales_2005_2024.dta", clear

	keep if inrange(anio, 2005, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

* ------------------------------------------------------------
* 6B. Control de duplicados
* ------------------------------------------------------------

	duplicates tag anio codigo_geo, gen(dup_key)

	count if dup_key > 0

	if r(N) > 0 {
		di as error "ERROR: Hay duplicados en porcentaje de partos/abortos por anio + codigo_geo."

		preserve
			keep if dup_key > 0
			sort anio codigo_geo
			export excel using "$out_aud/auditoria_duplicados_pct_partos_abortos.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	drop dup_key

	isid anio codigo_geo

* ------------------------------------------------------------
* 6C. Calcular denominador e indicadores
* ------------------------------------------------------------

	gen embarazos_calc = partos + abortos

	gen pct_partos = .
	replace pct_partos = (partos / embarazos_calc) * 100 if embarazos_calc > 0 & !missing(partos, abortos)

	gen pct_abortos = .
	replace pct_abortos = (abortos / embarazos_calc) * 100 if embarazos_calc > 0 & !missing(partos, abortos)

	replace pct_partos  = round(pct_partos, .01)
	replace pct_abortos = round(pct_abortos, .01)

* ------------------------------------------------------------
* 6D. Control: la suma debe ser 100
* ------------------------------------------------------------

	gen suma_pct = pct_partos + pct_abortos

	count if !missing(suma_pct) & abs(suma_pct - 100) > .02

	if r(N) > 0 {
		di as error "ADVERTENCIA: Hay casos donde pct_partos + pct_abortos no suma 100."

		preserve
			keep if !missing(suma_pct) & abs(suma_pct - 100) > .02
			sort anio codigo_geo
			export excel using "$out_aud/auditoria_suma_pct_partos_abortos.xlsx", firstrow(variables) replace
		restore
	}

* ------------------------------------------------------------
* 6E. Preparar tabla final
* ------------------------------------------------------------

	keep anio codigo_geo pct_partos pct_abortos

	reshape wide pct_partos pct_abortos, i(anio) j(codigo_geo) string

* ------------------------------------------------------------
* 6F. Renombrar columnas
* ------------------------------------------------------------

	rename pct_partos00  BOLIVIA_PARTOS
	rename pct_abortos00 BOLIVIA_ABORTOS

	rename pct_partos01  CHUQUISACA_PARTOS
	rename pct_abortos01 CHUQUISACA_ABORTOS

	rename pct_partos02  LA_PAZ_PARTOS
	rename pct_abortos02 LA_PAZ_ABORTOS

	rename pct_partos03  COCHABAMBA_PARTOS
	rename pct_abortos03 COCHABAMBA_ABORTOS

	rename pct_partos04  ORURO_PARTOS
	rename pct_abortos04 ORURO_ABORTOS

	rename pct_partos05  POTOSI_PARTOS
	rename pct_abortos05 POTOSI_ABORTOS

	rename pct_partos06  TARIJA_PARTOS
	rename pct_abortos06 TARIJA_ABORTOS

	rename pct_partos07  SANTA_CRUZ_PARTOS
	rename pct_abortos07 SANTA_CRUZ_ABORTOS

	rename pct_partos08  BENI_PARTOS
	rename pct_abortos08 BENI_ABORTOS

	rename pct_partos09  PANDO_PARTOS
	rename pct_abortos09 PANDO_ABORTOS

* ------------------------------------------------------------
* 6G. Ordenar columnas
* ------------------------------------------------------------

	order anio BOLIVIA_PARTOS BOLIVIA_ABORTOS
	order CHUQUISACA_PARTOS CHUQUISACA_ABORTOS, after(BOLIVIA_ABORTOS)
	order LA_PAZ_PARTOS LA_PAZ_ABORTOS, after(CHUQUISACA_ABORTOS)
	order COCHABAMBA_PARTOS COCHABAMBA_ABORTOS, after(LA_PAZ_ABORTOS)
	order ORURO_PARTOS ORURO_ABORTOS, after(COCHABAMBA_ABORTOS)
	order POTOSI_PARTOS POTOSI_ABORTOS, after(ORURO_ABORTOS)
	order TARIJA_PARTOS TARIJA_ABORTOS, after(POTOSI_ABORTOS)
	order SANTA_CRUZ_PARTOS SANTA_CRUZ_ABORTOS, after(TARIJA_ABORTOS)
	order BENI_PARTOS BENI_ABORTOS, after(SANTA_CRUZ_ABORTOS)
	order PANDO_PARTOS PANDO_ABORTOS, after(BENI_ABORTOS)

	sort anio

	format *_PARTOS *_ABORTOS %9.2f

* ------------------------------------------------------------
* 6H. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("pct_partos_abortos") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 6I. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_001_06 | % partos y % abortos sobre embarazos reconstruidos"
	di as result "Hoja: pct_partos_abortos"
	di as result "============================================================"

	list, clean noobs

* ============================================================
**# BLOQUE 07: NÚMERO DE NACIMIENTOS
* ============================================================
* Código:
*     06_001_07
*
* Indicador:
*     Número de nacimientos
*
* Fórmula:
*     valor = nacimientos
*
* Fuente:
*     _out/Bases DTA/pob_grupos_especiales_2005_20XX.dta
*
* Rango temporal:
*     2005–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Consideraciones:
*     - Este indicador se reporta como número absoluto.
*     - Se filtra solo nivel nacional y departamental.
*     - Como es un indicador posterior al primero, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 7A. Abrir base fuente
* ------------------------------------------------------------

	use "$out_dta/pob_grupos_especiales_2005_2024.dta", clear

	keep if inrange(anio, 2005, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

* ------------------------------------------------------------
* 7B. Control de duplicados
* ------------------------------------------------------------

	duplicates tag anio codigo_geo, gen(dup_key)

	count if dup_key > 0

	if r(N) > 0 {
		di as error "ERROR: Hay duplicados en nacimientos por anio + codigo_geo."

		preserve
			keep if dup_key > 0
			sort anio codigo_geo
			export excel using "$out_aud/auditoria_duplicados_nacimientos.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	drop dup_key

	isid anio codigo_geo

* ------------------------------------------------------------
* 7C. Calcular indicador
* ------------------------------------------------------------

	gen valor = nacimientos

* ------------------------------------------------------------
* 7D. Preparar tabla final
* ------------------------------------------------------------

	keep anio codigo_geo valor

	reshape wide valor, i(anio) j(codigo_geo) string

	rename valor00 BOLIVIA
	rename valor01 CHUQUISACA
	rename valor02 LA_PAZ
	rename valor03 COCHABAMBA
	rename valor04 ORURO
	rename valor05 POTOSI
	rename valor06 TARIJA
	rename valor07 SANTA_CRUZ
	rename valor08 BENI
	rename valor09 PANDO

	order anio BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO

	sort anio

	format BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO %15.0fc

* ------------------------------------------------------------
* 7E. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("num_nacimientos") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 7F. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_001_07 | Número de nacimientos"
	di as result "Hoja: num_nacimientos"
	di as result "============================================================"

	list, clean noobs

* ============================================================
**# BLOQUE 08: NÚMERO DE HIJOS NACIDOS VIVOS
* ============================================================
* Código:
*     06_001_08
*
* Indicador:
*     Número de hijos nacidos vivos
*
* Fórmula:
*     valor = hnv
*
* Fuente:
*     _out/Bases DTA/pob_grupos_especiales_2005_20XX.dta
*
* Rango temporal:
*     2005–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Consideraciones:
*     - Este indicador se reporta como número absoluto.
*     - Se filtra solo nivel nacional y departamental.
*     - Como es un indicador posterior al primero, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 8A. Abrir base fuente
* ------------------------------------------------------------

	use "$out_dta/pob_grupos_especiales_2005_2024.dta", clear

	keep if inrange(anio, 2005, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

* ------------------------------------------------------------
* 8B. Control de duplicados
* ------------------------------------------------------------

	duplicates tag anio codigo_geo, gen(dup_key)

	count if dup_key > 0

	if r(N) > 0 {
		di as error "ERROR: Hay duplicados en hijos nacidos vivos por anio + codigo_geo."

		preserve
			keep if dup_key > 0
			sort anio codigo_geo
			export excel using "$out_aud/auditoria_duplicados_hnv.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	drop dup_key

	isid anio codigo_geo

* ------------------------------------------------------------
* 8C. Calcular indicador
* ------------------------------------------------------------

	gen valor = hnv

* ------------------------------------------------------------
* 8D. Preparar tabla final
* ------------------------------------------------------------

	keep anio codigo_geo valor

	reshape wide valor, i(anio) j(codigo_geo) string

	rename valor00 BOLIVIA
	rename valor01 CHUQUISACA
	rename valor02 LA_PAZ
	rename valor03 COCHABAMBA
	rename valor04 ORURO
	rename valor05 POTOSI
	rename valor06 TARIJA
	rename valor07 SANTA_CRUZ
	rename valor08 BENI
	rename valor09 PANDO

	order anio BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO

	sort anio

	format BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO %15.0fc

* ------------------------------------------------------------
* 8E. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("num_hijos_nacidos_vivos") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 8F. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_001_08 | Número de hijos nacidos vivos"
	di as result "Hoja: num_hijos_nacidos_vivos"
	di as result "============================================================"

	list, clean noobs

* ============================================================
**# BLOQUE 09: NÚMERO DE HIJOS NACIDOS MUERTOS
* ============================================================
* Código:
*     06_001_09
*
* Indicador:
*     Número de hijos nacidos muertos
*
* Fórmula:
*     valor = hnm
*
* Fuente:
*     _out/Bases DTA/pob_grupos_especiales_2005_20XX.dta
*
* Rango temporal:
*     2005–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Consideraciones:
*     - Este indicador se reporta como número absoluto.
*     - Se filtra solo nivel nacional y departamental.
*     - Como es un indicador posterior al primero, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 9A. Abrir base fuente
* ------------------------------------------------------------

	use "$out_dta/pob_grupos_especiales_2005_2024.dta", clear

	keep if inrange(anio, 2005, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

* ------------------------------------------------------------
* 9B. Control de duplicados
* ------------------------------------------------------------

	duplicates tag anio codigo_geo, gen(dup_key)

	count if dup_key > 0

	if r(N) > 0 {
		di as error "ERROR: Hay duplicados en hijos nacidos muertos por anio + codigo_geo."

		preserve
			keep if dup_key > 0
			sort anio codigo_geo
			export excel using "$out_aud/auditoria_duplicados_hnm.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	drop dup_key

	isid anio codigo_geo

* ------------------------------------------------------------
* 9C. Calcular indicador
* ------------------------------------------------------------

	gen valor = hnm

* ------------------------------------------------------------
* 9D. Preparar tabla final
* ------------------------------------------------------------

	keep anio codigo_geo valor

	reshape wide valor, i(anio) j(codigo_geo) string

	rename valor00 BOLIVIA
	rename valor01 CHUQUISACA
	rename valor02 LA_PAZ
	rename valor03 COCHABAMBA
	rename valor04 ORURO
	rename valor05 POTOSI
	rename valor06 TARIJA
	rename valor07 SANTA_CRUZ
	rename valor08 BENI
	rename valor09 PANDO

	order anio BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO

	sort anio

	format BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO %15.0fc

* ------------------------------------------------------------
* 9E. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("num_hijos_nacidos_muertos") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 9F. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_001_09 | Número de hijos nacidos muertos"
	di as result "Hoja: num_hijos_nacidos_muertos"
	di as result "============================================================"

	list, clean noobs

* ============================================================
**# BLOQUE 10: DISTRIBUCIÓN DE HIJOS NACIDOS VIVOS Y MUERTOS
* ============================================================
* Código:
*     06_001_10
*
* Indicador:
*     Distribución porcentual de hijos nacidos vivos y muertos
*
* Fórmulas:
*     pct_hnv = hnv / (hnv + hnm) * 100
*     pct_hnm = hnm / (hnv + hnm) * 100
*
* Denominador:
*     nacimientos_calc = hnv + hnm
*
* Fuente:
*     _out/Bases DTA/pob_grupos_especiales_2005_20XX.dta
*
* Rango temporal:
*     2005–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Nota metodológica:
*     Para este indicador de distribución porcentual, el denominador
*     se construye como la suma de los componentes:
*
*         hijos nacidos vivos + hijos nacidos muertos.
*
*     Esta decisión garantiza consistencia interna entre las categorías
*     y permite que la distribución sume 100%.
* ============================================================
capture restore
* ------------------------------------------------------------
* 10A. Abrir base fuente
* ------------------------------------------------------------

	use "$out_dta/pob_grupos_especiales_2005_2024.dta", clear

	keep if inrange(anio, 2005, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

* ------------------------------------------------------------
* 10B. Control de duplicados
* ------------------------------------------------------------

	duplicates tag anio codigo_geo, gen(dup_key)

	count if dup_key > 0

	if r(N) > 0 {
		di as error "ERROR: Hay duplicados en distribución HNV/HNM por anio + codigo_geo."

		preserve
			keep if dup_key > 0
			sort anio codigo_geo
			export excel using "$out_aud/auditoria_duplicados_dist_hnv_hnm.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	drop dup_key

	isid anio codigo_geo

* ------------------------------------------------------------
* 10C. Calcular denominador reconstruido
* ------------------------------------------------------------

	gen nacimientos_calc = hnv + hnm

* ------------------------------------------------------------
* 10D. Calcular porcentajes
* ------------------------------------------------------------

	gen pct_hnv = .
	replace pct_hnv = (hnv / nacimientos_calc) * 100 if nacimientos_calc > 0 & !missing(hnv, hnm)

	gen pct_hnm = .
	replace pct_hnm = (hnm / nacimientos_calc) * 100 if nacimientos_calc > 0 & !missing(hnv, hnm)

	replace pct_hnv = round(pct_hnv, .01)
	replace pct_hnm = round(pct_hnm, .01)

* ------------------------------------------------------------
* 10E. Control: la suma debe ser 100
* ------------------------------------------------------------

	gen suma_pct = pct_hnv + pct_hnm

	count if !missing(suma_pct) & abs(suma_pct - 100) > .02

	if r(N) > 0 {
		di as error "ADVERTENCIA: Hay casos donde pct_hnv + pct_hnm no suma 100."

		preserve
			keep if !missing(suma_pct) & abs(suma_pct - 100) > .02
			sort anio codigo_geo
			export excel using "$out_aud/auditoria_suma_dist_hnv_hnm.xlsx", firstrow(variables) replace
		restore
	}

* ------------------------------------------------------------
* 10F. Preparar tabla final
* ------------------------------------------------------------

	keep anio codigo_geo pct_hnv pct_hnm

	reshape wide pct_hnv pct_hnm, i(anio) j(codigo_geo) string

* ------------------------------------------------------------
* 10G. Renombrar columnas
* ------------------------------------------------------------

	rename pct_hnv00 BOLIVIA_VIVOS
	rename pct_hnm00 BOLIVIA_MUERTOS

	rename pct_hnv01 CHUQUISACA_VIVOS
	rename pct_hnm01 CHUQUISACA_MUERTOS

	rename pct_hnv02 LA_PAZ_VIVOS
	rename pct_hnm02 LA_PAZ_MUERTOS

	rename pct_hnv03 COCHABAMBA_VIVOS
	rename pct_hnm03 COCHABAMBA_MUERTOS

	rename pct_hnv04 ORURO_VIVOS
	rename pct_hnm04 ORURO_MUERTOS

	rename pct_hnv05 POTOSI_VIVOS
	rename pct_hnm05 POTOSI_MUERTOS

	rename pct_hnv06 TARIJA_VIVOS
	rename pct_hnm06 TARIJA_MUERTOS

	rename pct_hnv07 SANTA_CRUZ_VIVOS
	rename pct_hnm07 SANTA_CRUZ_MUERTOS

	rename pct_hnv08 BENI_VIVOS
	rename pct_hnm08 BENI_MUERTOS

	rename pct_hnv09 PANDO_VIVOS
	rename pct_hnm09 PANDO_MUERTOS

* ------------------------------------------------------------
* 10H. Ordenar columnas
* ------------------------------------------------------------

	order anio BOLIVIA_VIVOS BOLIVIA_MUERTOS
	order CHUQUISACA_VIVOS CHUQUISACA_MUERTOS, after(BOLIVIA_MUERTOS)
	order LA_PAZ_VIVOS LA_PAZ_MUERTOS, after(CHUQUISACA_MUERTOS)
	order COCHABAMBA_VIVOS COCHABAMBA_MUERTOS, after(LA_PAZ_MUERTOS)
	order ORURO_VIVOS ORURO_MUERTOS, after(COCHABAMBA_MUERTOS)
	order POTOSI_VIVOS POTOSI_MUERTOS, after(ORURO_MUERTOS)
	order TARIJA_VIVOS TARIJA_MUERTOS, after(POTOSI_MUERTOS)
	order SANTA_CRUZ_VIVOS SANTA_CRUZ_MUERTOS, after(TARIJA_MUERTOS)
	order BENI_VIVOS BENI_MUERTOS, after(SANTA_CRUZ_MUERTOS)
	order PANDO_VIVOS PANDO_MUERTOS, after(BENI_MUERTOS)

	sort anio

	format *_VIVOS *_MUERTOS %9.2f

* ------------------------------------------------------------
* 10I. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("dist_hnv_hnm") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 10J. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_001_10 | Distribución HNV/HNM"
	di as result "Hoja: dist_hnv_hnm"
	di as result "============================================================"

	list, clean noobs

* ============================================================
**# BLOQUE 11: TASA GENERAL DE ABORTO
* ============================================================
*
* Código:
*     06_002_01
*
* Indicador:
*     Tasa general de aborto
*
* Fórmula:
*     valor = abortos / mef * 1000
*
* Fuente:
*     _out/Bases DTA/pob_grupos_especiales_2005_20XX.dta
*
* Rango temporal:
*     2005–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Unidad:
*     Abortos por cada 1.000 mujeres en edad fértil.
*
* Consideraciones:
*     - La tasa se calcula sobre mujeres en edad fértil.
*     - Se multiplica por 1.000.
*     - Como es un indicador posterior al primero, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 11A. Abrir base fuente
* ------------------------------------------------------------

	use "$out_dta/pob_grupos_especiales_2005_2024.dta", clear

	keep if inrange(anio, 2005, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

* ------------------------------------------------------------
* 11B. Control de duplicados
* ------------------------------------------------------------

	duplicates tag anio codigo_geo, gen(dup_key)

	count if dup_key > 0

	if r(N) > 0 {
		di as error "ERROR: Hay duplicados en tasa general de aborto por anio + codigo_geo."

		preserve
			keep if dup_key > 0
			sort anio codigo_geo
			export excel using "$out_aud/auditoria_duplicados_tasa_aborto_general.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	drop dup_key

	isid anio codigo_geo

* ------------------------------------------------------------
* 11C. Calcular indicador
* ------------------------------------------------------------

	gen valor = .

	replace valor = (abortos / mef) * 1000 if mef > 0 & !missing(abortos, mef)

	replace valor = round(valor, .01)

* ------------------------------------------------------------
* 11D. Preparar tabla final
* ------------------------------------------------------------

	keep anio codigo_geo valor

	reshape wide valor, i(anio) j(codigo_geo) string

	rename valor00 BOLIVIA
	rename valor01 CHUQUISACA
	rename valor02 LA_PAZ
	rename valor03 COCHABAMBA
	rename valor04 ORURO
	rename valor05 POTOSI
	rename valor06 TARIJA
	rename valor07 SANTA_CRUZ
	rename valor08 BENI
	rename valor09 PANDO

	order anio BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO

	sort anio

	format BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO %9.2f

* ------------------------------------------------------------
* 11E. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("tasa_aborto_general") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 11F. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_001_11 | Tasa general de aborto"
	di as result "Unidad: abortos por cada 1.000 mujeres en edad fértil"
	di as result "Hoja: tasa_aborto_general"
	di as result "============================================================"

	list, clean noobs

* ============================================================
**# BLOQUE 12: TASA GENERAL DE FECUNDIDAD
* ============================================================
*
* Código:
*     06_002_02
*
* Indicador:
*     Tasa general de fecundidad
*
* Fórmula:
*     valor = hnv / mef * 1000
*
* Fuente:
*     _out/Bases DTA/pob_grupos_especiales_2005_20XX.dta
*
* Rango temporal:
*     2005–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Unidad:
*     Hijos nacidos vivos por cada 1.000 mujeres en edad fértil.
*
* Consideraciones:
*     - La tasa se calcula usando hijos nacidos vivos como numerador.
*     - El denominador es el total de mujeres en edad fértil.
*     - Se multiplica por 1.000.
*     - Se filtra solo nivel nacional y departamental.
*     - Como es un indicador posterior al primero, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 12A. Abrir base fuente
* ------------------------------------------------------------

	use "$out_dta/pob_grupos_especiales_2005_2024.dta", clear

	keep if inrange(anio, 2005, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

* ------------------------------------------------------------
* 12B. Control de duplicados
* ------------------------------------------------------------

	duplicates tag anio codigo_geo, gen(dup_key)

	count if dup_key > 0

	if r(N) > 0 {
		di as error "ERROR: Hay duplicados en tasa general de fecundidad por anio + codigo_geo."

		preserve
			keep if dup_key > 0
			sort anio codigo_geo
			export excel using "$out_aud/auditoria_duplicados_tasa_fecundidad_general.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	drop dup_key

	isid anio codigo_geo

* ------------------------------------------------------------
* 12C. Calcular indicador
* ------------------------------------------------------------

	gen valor = .

	replace valor = (hnv / mef) * 1000 if mef > 0 & !missing(hnv, mef)

	replace valor = round(valor, .01)

* ------------------------------------------------------------
* 12D. Preparar tabla final
* ------------------------------------------------------------

	keep anio codigo_geo valor

	reshape wide valor, i(anio) j(codigo_geo) string

	rename valor00 BOLIVIA
	rename valor01 CHUQUISACA
	rename valor02 LA_PAZ
	rename valor03 COCHABAMBA
	rename valor04 ORURO
	rename valor05 POTOSI
	rename valor06 TARIJA
	rename valor07 SANTA_CRUZ
	rename valor08 BENI
	rename valor09 PANDO

	order anio BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO

	sort anio

	format BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO %9.2f

* ------------------------------------------------------------
* 12E. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("tasa_fecundidad_general") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 12F. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_001_12 | Tasa general de fecundidad"
	di as result "Unidad: hijos nacidos vivos por cada 1.000 mujeres en edad fértil"
	di as result "Hoja: tasa_fecundidad_general"
	di as result "============================================================"

	list, clean noobs

* ============================================================
**# BLOQUE 13: TASA DE MORTALIDAD FETAL
* ============================================================
* Código:
*     06_002_03
*
* Indicador:
*     Tasa de mortalidad fetal
*
* Fórmula:
*     valor = hnm / (hnv + hnm) * 1000
*
* Fuente:
*     _out/Bases DTA/pob_grupos_especiales_2005_20XX.dta
*
* Rango temporal:
*     2005–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Unidad:
*     Hijos nacidos muertos por cada 1.000 nacimientos reconstruidos.
*
* Consideraciones:
*     - La tasa se calcula dividiendo hijos nacidos muertos entre
*       la suma de hijos nacidos vivos e hijos nacidos muertos.
*     - Se filtra solo nivel nacional y departamental.
*     - Como es un indicador posterior al primero, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 13A. Abrir base fuente
* ------------------------------------------------------------

	use "$out_dta/pob_grupos_especiales_2005_2024.dta", clear

	keep if inrange(anio, 2005, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

* ------------------------------------------------------------
* 13B. Control de duplicados
* ------------------------------------------------------------

	duplicates tag anio codigo_geo, gen(dup_key)

	count if dup_key > 0

	if r(N) > 0 {
		di as error "ERROR: Hay duplicados en tasa de mortalidad fetal por anio + codigo_geo."

		preserve
			keep if dup_key > 0
			sort anio codigo_geo
			export excel using "$out_aud/auditoria_duplicados_tasa_mortalidad_fetal.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	drop dup_key

	isid anio codigo_geo

* ------------------------------------------------------------
* 13C. Construir denominador
* ------------------------------------------------------------

	gen nacimientos_calc = hnv + hnm

* ------------------------------------------------------------
* 13D. Control de denominador
* ------------------------------------------------------------

	count if missing(nacimientos_calc) | nacimientos_calc <= 0

	if r(N) > 0 {
		di as error "ADVERTENCIA: Hay observaciones con denominador faltante o menor/igual a cero."

		preserve
			keep if missing(nacimientos_calc) | nacimientos_calc <= 0
			sort anio codigo_geo
			export excel using "$out_aud/auditoria_denominador_tasa_mortalidad_fetal.xlsx", firstrow(variables) replace
		restore
	}

* ------------------------------------------------------------
* 13E. Calcular indicador
* ------------------------------------------------------------

	gen valor = .

	replace valor = (hnm / nacimientos_calc) * 1000 if nacimientos_calc > 0 & !missing(hnm, hnv)

	replace valor = round(valor, .01)

* ------------------------------------------------------------
* 13F. Preparar tabla final
* ------------------------------------------------------------

	keep anio codigo_geo valor

	reshape wide valor, i(anio) j(codigo_geo) string

	rename valor00 BOLIVIA
	rename valor01 CHUQUISACA
	rename valor02 LA_PAZ
	rename valor03 COCHABAMBA
	rename valor04 ORURO
	rename valor05 POTOSI
	rename valor06 TARIJA
	rename valor07 SANTA_CRUZ
	rename valor08 BENI
	rename valor09 PANDO

	order anio BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO

	sort anio

	format BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO %9.2f

* ------------------------------------------------------------
* 13G. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("tasa_mortalidad_fetal") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 13H. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_001_13 | Tasa de mortalidad fetal"
	di as result "Unidad: hijos nacidos muertos por cada 1.000 nacimientos reconstruidos"
	di as result "Hoja: tasa_mortalidad_fetal"
	di as result "============================================================"

	list, clean noobs

* ============================================================
**# BLOQUE 14: TASA ESTIMADA DE EMBARAZO
* ============================================================
* Código:
*     06_002_04
*
* Indicador:
*     Tasa estimada de embarazo
*
* Fórmula:
*     valor = (hnv + hnm + abortos) / mef * 1000
*
* Numerador:
*     Hijos nacidos vivos + hijos nacidos muertos + abortos
*
* Denominador:
*     Mujeres en edad fértil
*
* Fuente:
*     _out/Bases DTA/pob_grupos_especiales_2005_202XX.dta
*
* Rango temporal:
*     2005–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Unidad:
*     Embarazos estimados por cada 1.000 mujeres en edad fértil.
*
* Consideraciones:
*     - El numerador se reconstruye como la suma de:
*           hijos nacidos vivos + hijos nacidos muertos + abortos.
*     - Se filtra solo nivel nacional y departamental.
*     - Como es un indicador posterior al primero, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 14A. Abrir base fuente
* ------------------------------------------------------------

	use "$out_dta/pob_grupos_especiales_2005_2024.dta", clear

	keep if inrange(anio, 2005, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

* ------------------------------------------------------------
* 14B. Control de duplicados
* ------------------------------------------------------------

	duplicates tag anio codigo_geo, gen(dup_key)

	count if dup_key > 0

	if r(N) > 0 {
		di as error "ERROR: Hay duplicados en tasa estimada de embarazo por anio + codigo_geo."

		preserve
			keep if dup_key > 0
			sort anio codigo_geo
			export excel using "$out_aud/auditoria_duplicados_tasa_embarazo_estimada.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	drop dup_key

	isid anio codigo_geo

* ------------------------------------------------------------
* 14C. Construir numerador
* ------------------------------------------------------------

	gen embarazos_est = hnv + hnm + abortos

* ------------------------------------------------------------
* 14D. Control de denominador y numerador
* ------------------------------------------------------------

	count if missing(mef) | mef <= 0

	if r(N) > 0 {
		di as error "ADVERTENCIA: Hay observaciones con MEF faltante o menor/igual a cero."

		preserve
			keep if missing(mef) | mef <= 0
			sort anio codigo_geo
			export excel using "$out_aud/auditoria_denominador_tasa_embarazo_estimada.xlsx", firstrow(variables) replace
		restore
	}

	count if missing(embarazos_est)

	if r(N) > 0 {
		di as error "ADVERTENCIA: Hay observaciones con numerador faltante en tasa estimada de embarazo."

		preserve
			keep if missing(embarazos_est)
			sort anio codigo_geo
			export excel using "$out_aud/auditoria_numerador_tasa_embarazo_estimada.xlsx", firstrow(variables) replace
		restore
	}

* ------------------------------------------------------------
* 14E. Calcular indicador
* ------------------------------------------------------------

	gen valor = .

	replace valor = (embarazos_est / mef) * 1000 if mef > 0 & !missing(embarazos_est, mef)

	replace valor = round(valor, .01)

* ------------------------------------------------------------
* 14F. Preparar tabla final
* ------------------------------------------------------------

	keep anio codigo_geo valor

	reshape wide valor, i(anio) j(codigo_geo) string

	rename valor00 BOLIVIA
	rename valor01 CHUQUISACA
	rename valor02 LA_PAZ
	rename valor03 COCHABAMBA
	rename valor04 ORURO
	rename valor05 POTOSI
	rename valor06 TARIJA
	rename valor07 SANTA_CRUZ
	rename valor08 BENI
	rename valor09 PANDO

	order anio BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO

	sort anio

	format BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO %9.2f

* ------------------------------------------------------------
* 14G. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("tasa_embarazo_estimada") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 14H. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_001_14 | Tasa estimada de embarazo"
	di as result "Unidad: embarazos estimados por cada 1.000 mujeres en edad fértil"
	di as result "Hoja: tasa_embarazo_estimada"
	di as result "============================================================"

	list, clean noobs

* ============================================================
**# BLOQUE 15: TASA BRUTA DE NATALIDAD
* ============================================================
* Código:
*     06_003_01
*
* Indicador:
*     Tasa bruta de natalidad
*
* Fórmula:
*     valor = hnv / pob_total * 1000
*
* Numerador:
*     Hijos nacidos vivos
*
* Denominador:
*     Población total extraída desde la última columna con dato
*     del Excel de población general por grupos de edad en salud.
*
* Fuente del numerador:
*     _out/Bases DTA/pob_grupos_especiales_2005_20XX.dta
*
* Fuente del denominador:
*     _out/Bases DTA/pob_total_salud_2005_20XX.dta
*
* Rango temporal:
*     2005–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Unidad:
*     Hijos nacidos vivos por cada 1.000 habitantes.
*
* Consideraciones:
*     - La población total se toma desde la base auxiliar
*       pob_total_salud_2005_20XX.dta.
*     - Esta base auxiliar extrae la población desde la última columna
*       con dato de cada fila geográfica, equivalente a "Total general".
*     - Se filtra solo nivel nacional y departamental.
*     - Como es un indicador posterior al primero, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 15A. Construir denominador: población total
* ------------------------------------------------------------

	tempfile pob_total

	use "$out_dta/pob_total_salud_2005_2024.dta", clear

	keep if inrange(anio, 2005, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

	keep anio codigo_geo pob_total

	isid anio codigo_geo

	count if missing(pob_total) | pob_total <= 0

	if r(N) > 0 {
		di as error "ERROR: Hay población total faltante o menor/igual a cero."

		preserve
			keep if missing(pob_total) | pob_total <= 0
			sort anio codigo_geo
			export excel using "$out_aud/auditoria_pob_total_tasa_bruta_natalidad.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	save `pob_total', replace

* ------------------------------------------------------------
* 15B. Abrir numerador: hijos nacidos vivos
* ------------------------------------------------------------

	use "$out_dta/pob_grupos_especiales_2005_2024.dta", clear

	keep if inrange(anio, 2005, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

* ------------------------------------------------------------
* 15C. Control de duplicados
* ------------------------------------------------------------

	duplicates tag anio codigo_geo, gen(dup_key)

	count if dup_key > 0

	if r(N) > 0 {
		di as error "ERROR: Hay duplicados en tasa bruta de natalidad por anio + codigo_geo."

		preserve
			keep if dup_key > 0
			sort anio codigo_geo
			export excel using "$out_aud/auditoria_duplicados_tasa_bruta_natalidad.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	drop dup_key

	isid anio codigo_geo

* ------------------------------------------------------------
* 15D. Unir numerador y denominador
* ------------------------------------------------------------

	merge 1:1 anio codigo_geo using `pob_total'

	tab _merge

	preserve
		keep if _merge != 3
		count

		if r(N) > 0 {
			export excel using "$out_aud/auditoria_merge_tasa_bruta_natalidad.xlsx", firstrow(variables) replace
		}
	restore

	keep if _merge == 3
	drop _merge

* ------------------------------------------------------------
* 15E. Calcular indicador
* ------------------------------------------------------------

	gen valor = .

	replace valor = (hnv / pob_total) * 1000 if pob_total > 0 & !missing(hnv, pob_total)

	replace valor = round(valor, .01)

* ------------------------------------------------------------
* 15F. Preparar tabla final
* ------------------------------------------------------------

	keep anio codigo_geo valor

	reshape wide valor, i(anio) j(codigo_geo) string

	rename valor00 BOLIVIA
	rename valor01 CHUQUISACA
	rename valor02 LA_PAZ
	rename valor03 COCHABAMBA
	rename valor04 ORURO
	rename valor05 POTOSI
	rename valor06 TARIJA
	rename valor07 SANTA_CRUZ
	rename valor08 BENI
	rename valor09 PANDO

	order anio BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO

	sort anio

	format BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO %9.2f

* ------------------------------------------------------------
* 15G. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("tasa_bruta_natalidad") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 15H. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_001_15 | Tasa bruta de natalidad"
	di as result "Unidad: hijos nacidos vivos por cada 1.000 habitantes"
	di as result "Hoja: tasa_bruta_natalidad"
	di as result "============================================================"

	list, clean noobs

* ============================================================
**# BLOQUE 16: ÍNDICE DE MASCULINIDAD
* ============================================================
* Código:
*     06_003_02
*
* Indicador:
*     Índice de masculinidad
*
* Fórmula:
*     valor = hombres / mujeres * 100
*
* Numerador:
*     población total de hombres
*
* Denominador:
*     población total de mujeres
*
* Fuente:
*     _out/Bases DTA/pob_edades_simples_long_2012_20XX.dta
*
* Rango temporal:
*     2012–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Unidad:
*     Hombres por cada 100 mujeres.
*
* Consideraciones:
*     - Este indicador compara la población masculina total con la
*       población femenina total.
*     - Se filtra solo nivel nacional y departamental.
*     - Como la base limpia por sexo está disponible desde 2012,
*       el indicador se calcula para 2012–20XX.
*     - Como es un indicador posterior al primero, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 16A. Construir denominador: población total de mujeres
* ------------------------------------------------------------

	tempfile mujeres_total

	use "$out_dta/pob_edades_simples_long_2012_2024.dta", clear

	keep if inrange(anio, 2012, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

	* Mantener población total de mujeres.
	keep if sexo == "Mujer"
	keep if grupo_edad == "total"

	keep anio codigo_geo poblacion

	rename poblacion mujeres_total

	isid anio codigo_geo

	save `mujeres_total', replace

* ------------------------------------------------------------
* 16B. Construir numerador: población total de hombres
* ------------------------------------------------------------

	tempfile hombres_total

	use "$out_dta/pob_edades_simples_long_2012_2024.dta", clear

	keep if inrange(anio, 2012, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

	* Mantener población total de hombres.
	keep if sexo == "Hombre"
	keep if grupo_edad == "total"

	keep anio codigo_geo poblacion

	rename poblacion hombres_total

	isid anio codigo_geo

	save `hombres_total', replace

* ------------------------------------------------------------
* 16C. Unir hombres y mujeres
* ------------------------------------------------------------

	use `hombres_total', clear

	merge 1:1 anio codigo_geo using `mujeres_total'

	tab _merge

	preserve
		keep if _merge != 3
		count

		if r(N) > 0 {
			export excel using "$out_aud/auditoria_merge_indice_masculinidad.xlsx", firstrow(variables) replace
		}
	restore

	keep if _merge == 3
	drop _merge

* ------------------------------------------------------------
* 16D. Calcular indicador
* ------------------------------------------------------------

	gen valor = .

	replace valor = (hombres_total / mujeres_total) * 100 if mujeres_total > 0 & !missing(hombres_total, mujeres_total)

	replace valor = round(valor, .01)

* ------------------------------------------------------------
* 16E. Preparar tabla final
* ------------------------------------------------------------

	keep anio codigo_geo valor

	reshape wide valor, i(anio) j(codigo_geo) string

	rename valor00 BOLIVIA
	rename valor01 CHUQUISACA
	rename valor02 LA_PAZ
	rename valor03 COCHABAMBA
	rename valor04 ORURO
	rename valor05 POTOSI
	rename valor06 TARIJA
	rename valor07 SANTA_CRUZ
	rename valor08 BENI
	rename valor09 PANDO

	order anio BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO

	sort anio

	format BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO %9.2f

* ------------------------------------------------------------
* 16F. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("indice_masculinidad") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 16G. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_003_02 | Índice de masculinidad"
	di as result "Unidad: hombres por cada 100 mujeres"
	di as result "Hoja: indice_masculinidad"
	di as result "============================================================"

	list, clean noobs
	
* ============================================================
**# BLOQUE 17: TASA DE CRECIMIENTO POBLACIONAL
* ============================================================
* Código:
*     06_003_03
*
* Indicador:
*     Tasa de crecimiento poblacional
*
* Fórmula:
*     valor = ((pob_total_t - pob_total_t_1) / pob_total_t_1) * 100
*
* Numerador:
*     pob_total_t - pob_total_t_1
*
* Denominador:
*     pob_total_t_1
*
* Fuente:
*     _out/Bases DTA/pob_total_salud_2005_20XX.dta
*
* Rango temporal:
*     2006–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Unidad:
*     Porcentaje anual de crecimiento poblacional.
*
* Consideraciones:
*     - La tasa compara la población total de un año con la población
*       total del año anterior.
*     - Como se requiere población rezagada, el primer año calculable
*       es 2006.
*     - La población total proviene de la base auxiliar construida
*       desde la última columna con dato del Excel de salud,
*       equivalente a "Total general".
*     - Se filtra solo nivel nacional y departamental.
*     - Como es un indicador posterior al primero, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 17A. Abrir base fuente
* ------------------------------------------------------------

	use "$out_dta/pob_total_salud_2005_2024.dta", clear

	keep if inrange(anio, 2005, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

* ------------------------------------------------------------
* 17B. Estandarizar código geográfico
* ------------------------------------------------------------

	capture confirm numeric variable codigo_geo

	if _rc == 0 {
		gen str6 codigo_geo_str = string(codigo_geo, "%02.0f")
		drop codigo_geo
		rename codigo_geo_str codigo_geo
	}

	replace codigo_geo = strtrim(codigo_geo)

* ------------------------------------------------------------
* 17C. Control de duplicados
* ------------------------------------------------------------

	duplicates tag anio codigo_geo, gen(dup_key)

	count if dup_key > 0

	if r(N) > 0 {
		di as error "ERROR: Hay duplicados en tasa de crecimiento poblacional por anio + codigo_geo."

		preserve
			keep if dup_key > 0
			sort anio codigo_geo
			export excel using "$out_aud/auditoria_duplicados_tasa_crecimiento_poblacional.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	drop dup_key

	isid anio codigo_geo

* ------------------------------------------------------------
* 17D. Calcular población rezagada
* ------------------------------------------------------------

	sort codigo_geo anio

	by codigo_geo: gen pob_total_lag = pob_total[_n-1]

	* Control: asegurar que el rezago corresponde al año anterior.
	by codigo_geo: gen anio_lag = anio[_n-1]

	replace pob_total_lag = . if anio_lag != anio - 1

* ------------------------------------------------------------
* 17E. Calcular indicador
* ------------------------------------------------------------

	gen valor = .

	replace valor = ((pob_total - pob_total_lag) / pob_total_lag) * 100 if pob_total_lag > 0 & !missing(pob_total, pob_total_lag)

	replace valor = round(valor, .01)

	* Como 2005 no tiene año anterior en la base, se reporta desde 2006.
	* Dado el corte de 2024, este año no tendrá datos sino en adelante.
	keep if inrange(anio, 2006, 2024)

* ------------------------------------------------------------
* 17E.1. Corte metodológico para gráficos
* ------------------------------------------------------------
* Se deja 2012 y 2024 sin datos para generar un corte visual en la serie.
* Esto es útil si existe un cambio de fuente, estructura o criterio
* de cálculo que no se desea graficar como una continuidad artificial.
*
* En Excel, el valor missing de Stata se exporta como celda vacía.
* ------------------------------------------------------------

replace valor = . if anio == 2012
replace valor = . if anio == 2024

* ------------------------------------------------------------
* 17F. Preparar tabla final
* ------------------------------------------------------------

	keep anio codigo_geo valor

	reshape wide valor, i(anio) j(codigo_geo) string

	rename valor00 BOLIVIA
	rename valor01 CHUQUISACA
	rename valor02 LA_PAZ
	rename valor03 COCHABAMBA
	rename valor04 ORURO
	rename valor05 POTOSI
	rename valor06 TARIJA
	rename valor07 SANTA_CRUZ
	rename valor08 BENI
	rename valor09 PANDO

	order anio BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO

	sort anio

	format BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO %9.2f

* ------------------------------------------------------------
* 17G. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("tasa_crecimiento_pob") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 17H. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_003_03 | Tasa de crecimiento poblacional"
	di as result "Unidad: porcentaje anual"
	di as result "Hoja: tasa_crecimiento_pob"
	di as result "============================================================"

	list, clean noobs

* ============================================================
**# BLOQUE 18: DISTRIBUCIÓN DE LA POBLACIÓN SEGÚN GRUPOS ETARIOS
* ============================================================
* Código:
*     06_003_04
*
* Indicador:
*     Distribución de la población según grupos etarios
*
* Fórmulas:
*     pct_0_4    = pob_0_4 / pob_total * 100
*     pct_5_14   = pob_5_14 / pob_total * 100
*     pct_15_64  = pob_15_64 / pob_total * 100
*     pct_65_mas = pob_65_mas / pob_total * 100
*
* Numeradores:
*     población de 0 a 4 años
*     población de 5 a 14 años
*     población de 15 a 64 años
*     población de 65 años y más
*
* Denominador:
*     población total
*
* Fuentes:
*     2005–2011:
*         _out/Bases DTA/pob_edad_salud_2005_20XX.dta
*
*     2012–20XX:
*         _out/Bases DTA/pob_edades_simples_long_2012_20XX.dta
*
* Rango temporal:
*     2005–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Unidad:
*     Porcentaje de la población total.
*
* Consideraciones:
*     - Para 2005–2011 se usa la base de grupos de edad en salud.
*     - Para 2012–20XX se reconstruyen los grupos desde edades simples
*       y grupos quinquenales.
*     - Como es un indicador posterior al primero, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 18A. Construir grupos etarios 2005–2011
* ------------------------------------------------------------

	tempfile grupos_2005_2011

	use "$out_dta/pob_edad_salud_2005_2024.dta", clear

	keep if inrange(anio, 2005, 2011)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

	* Estandarizar codigo_geo como texto.
	capture confirm numeric variable codigo_geo

	if _rc == 0 {
		gen str6 codigo_geo_str = string(codigo_geo, "%02.0f")
		drop codigo_geo
		rename codigo_geo_str codigo_geo
	}

	replace codigo_geo = strtrim(codigo_geo)

	* En la estructura 2005–2011:
	*     grupo_0_4 = 0-4 años
	*     col_i     = 5-14 años
	*     col_j     = 15-64 años
	*     col_k     = 65 años y más
	*     col_l     = Total general

	gen double pob_0_4    = grupo_0_4
	gen double pob_5_14   = col_i
	gen double pob_15_64  = col_j
	gen double pob_65_mas = col_k
	gen double pob_total  = col_l

	keep anio codigo_geo nombre_geo nivel_geo pob_0_4 pob_5_14 pob_15_64 pob_65_mas pob_total

	isid anio codigo_geo

	save `grupos_2005_2011', replace

* ------------------------------------------------------------
* 18B. Construir grupos etarios 2012–20XX
* ------------------------------------------------------------

	tempfile grupos_2012_2024

	use "$out_dta/pob_edades_simples_long_2012_2024.dta", clear

	keep if inrange(anio, 2012, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

	* Usar población total por sexo.
	keep if sexo == "Total"

	* Estandarizar codigo_geo como texto.
	capture confirm numeric variable codigo_geo

	if _rc == 0 {
		gen str6 codigo_geo_str = string(codigo_geo, "%02.0f")
		drop codigo_geo
		rename codigo_geo_str codigo_geo
	}

	replace codigo_geo = strtrim(codigo_geo)

	* Asegurar que grupo_edad sea texto.
	capture confirm string variable grupo_edad

	if _rc != 0 {
		tostring grupo_edad, replace
	}

	replace grupo_edad = strtrim(grupo_edad)

	gen edad_simple = real(grupo_edad)

	* 0-4 años:
	*     edades simples 0 a 4.
	gen double pob_0_4 = 0
	replace pob_0_4 = poblacion if inrange(edad_simple, 0, 4)

	* 5-14 años:
	*     edades simples 5 a 14.
	gen double pob_5_14 = 0
	replace pob_5_14 = poblacion if inrange(edad_simple, 5, 14)

	* 15-64 años:
	*     edades simples 15 a 29
	*     + grupos quinquenales 30_34 a 60_64.
	gen double pob_15_64 = 0
	replace pob_15_64 = poblacion if inrange(edad_simple, 15, 29)
	replace pob_15_64 = poblacion if regexm(grupo_edad, "^(30_34|35_39|40_44|45_49|50_54|55_59|60_64)$")

	* 65 años y más:
	*     grupos quinquenales 65_69, 70_74, 75_79 y 80_mas.
	gen double pob_65_mas = 0
	replace pob_65_mas = poblacion if regexm(grupo_edad, "^(65_69|70_74|75_79|80_mas)$")

	* Población total:
	*     grupo_edad == total.
	gen double pob_total = 0
	replace pob_total = poblacion if grupo_edad == "total"

	collapse (sum) pob_0_4 pob_5_14 pob_15_64 pob_65_mas pob_total, by(anio codigo_geo nombre_geo nivel_geo)

	isid anio codigo_geo

	save `grupos_2012_2024', replace

* ------------------------------------------------------------
* 18C. Unir ambos periodos
* ------------------------------------------------------------

	use `grupos_2005_2011', clear

	append using `grupos_2012_2024'

	sort anio codigo_geo

	isid anio codigo_geo

* ------------------------------------------------------------
* 18D. Control de consistencia
* ------------------------------------------------------------
* La población total debe ser igual a:
*
*     pob_total = pob_0_4 + pob_5_14 + pob_15_64 + pob_65_mas
*
* Si no coincide, el indicador se detiene.
* ------------------------------------------------------------

	gen double pob_total_suma = pob_0_4 + pob_5_14 + pob_15_64 + pob_65_mas

	gen double dif_abs = pob_total - pob_total_suma

	gen double dif_pct = .
	replace dif_pct = (dif_abs / pob_total) * 100 if pob_total > 0 & !missing(dif_abs, pob_total)

	format pob_0_4 pob_5_14 pob_15_64 pob_65_mas pob_total pob_total_suma dif_abs %15.2fc
	format dif_pct %12.8f

	count if !missing(dif_abs) & abs(dif_abs) > 0.5

	if r(N) > 0 {
		di as error "ERROR: La población total no coincide con la suma de grupos etarios."

		preserve
			keep if !missing(dif_abs) & abs(dif_abs) > 0.5
			sort anio codigo_geo
			export excel using "$out_aud/error_dist_grupos_etarios_total_no_coincide.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	count if missing(pob_total) | pob_total <= 0

	if r(N) > 0 {
		di as error "ERROR: Hay población total faltante o menor/igual a cero."

		preserve
			keep if missing(pob_total) | pob_total <= 0
			sort anio codigo_geo
			export excel using "$out_aud/error_pob_total_dist_grupos_etarios.xlsx", firstrow(variables) replace
		restore

		error 459
	}

* ------------------------------------------------------------
* 18E. Calcular distribución porcentual
* ------------------------------------------------------------

	gen pct_0_4 = .
	replace pct_0_4 = (pob_0_4 / pob_total) * 100 if pob_total > 0 & !missing(pob_0_4, pob_total)

	gen pct_5_14 = .
	replace pct_5_14 = (pob_5_14 / pob_total) * 100 if pob_total > 0 & !missing(pob_5_14, pob_total)

	gen pct_15_64 = .
	replace pct_15_64 = (pob_15_64 / pob_total) * 100 if pob_total > 0 & !missing(pob_15_64, pob_total)

	gen pct_65_mas = .
	replace pct_65_mas = (pob_65_mas / pob_total) * 100 if pob_total > 0 & !missing(pob_65_mas, pob_total)

	replace pct_0_4    = round(pct_0_4, .01)
	replace pct_5_14   = round(pct_5_14, .01)
	replace pct_15_64  = round(pct_15_64, .01)
	replace pct_65_mas = round(pct_65_mas, .01)

	gen suma_pct = pct_0_4 + pct_5_14 + pct_15_64 + pct_65_mas

	count if !missing(suma_pct) & abs(suma_pct - 100) > .05

	if r(N) > 0 {
		di as error "ERROR: Hay casos donde la distribución porcentual no suma 100."

		preserve
			keep if !missing(suma_pct) & abs(suma_pct - 100) > .05
			sort anio codigo_geo
			export excel using "$out_aud/error_suma_pct_grupos_etarios.xlsx", firstrow(variables) replace
		restore

		error 459
	}

* ------------------------------------------------------------
* 18F. Preparar tabla final
* ------------------------------------------------------------
* Nota:
*     No se exporta columna total, porque siempre equivale a 100.
* ------------------------------------------------------------

	keep anio codigo_geo pct_0_4 pct_5_14 pct_15_64 pct_65_mas

	reshape wide pct_0_4 pct_5_14 pct_15_64 pct_65_mas, i(anio) j(codigo_geo) string

	rename pct_0_400 BOLIVIA_0_4
	rename pct_5_1400 BOLIVIA_5_14
	rename pct_15_6400 BOLIVIA_15_64
	rename pct_65_mas00 BOLIVIA_65MAS

	rename pct_0_401 CHUQUISACA_0_4
	rename pct_5_1401 CHUQUISACA_5_14
	rename pct_15_6401 CHUQUISACA_15_64
	rename pct_65_mas01 CHUQUISACA_65MAS

	rename pct_0_402 LA_PAZ_0_4
	rename pct_5_1402 LA_PAZ_5_14
	rename pct_15_6402 LA_PAZ_15_64
	rename pct_65_mas02 LA_PAZ_65MAS

	rename pct_0_403 COCHABAMBA_0_4
	rename pct_5_1403 COCHABAMBA_5_14
	rename pct_15_6403 COCHABAMBA_15_64
	rename pct_65_mas03 COCHABAMBA_65MAS

	rename pct_0_404 ORURO_0_4
	rename pct_5_1404 ORURO_5_14
	rename pct_15_6404 ORURO_15_64
	rename pct_65_mas04 ORURO_65MAS

	rename pct_0_405 POTOSI_0_4
	rename pct_5_1405 POTOSI_5_14
	rename pct_15_6405 POTOSI_15_64
	rename pct_65_mas05 POTOSI_65MAS

	rename pct_0_406 TARIJA_0_4
	rename pct_5_1406 TARIJA_5_14
	rename pct_15_6406 TARIJA_15_64
	rename pct_65_mas06 TARIJA_65MAS

	rename pct_0_407 SANTA_CRUZ_0_4
	rename pct_5_1407 SANTA_CRUZ_5_14
	rename pct_15_6407 SANTA_CRUZ_15_64
	rename pct_65_mas07 SANTA_CRUZ_65MAS

	rename pct_0_408 BENI_0_4
	rename pct_5_1408 BENI_5_14
	rename pct_15_6408 BENI_15_64
	rename pct_65_mas08 BENI_65MAS

	rename pct_0_409 PANDO_0_4
	rename pct_5_1409 PANDO_5_14
	rename pct_15_6409 PANDO_15_64
	rename pct_65_mas09 PANDO_65MAS

	order anio, first

	sort anio

	format *_0_4 *_5_14 *_15_64 *_65MAS %9.2f

* ------------------------------------------------------------
* 18G. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("dist_grupos_etarios") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 18H. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_003_04 | Distribución de la población según grupos etarios"
	di as result "Unidad: porcentaje de la población total"
	di as result "Hoja: dist_grupos_etarios"
	di as result "============================================================"

	list, clean noobs
	
* ============================================================
**# BLOQUE 19: ÍNDICE DE DEPENDENCIA POTENCIAL
* ============================================================
*
* Código:
*     06_003_05
*
* Indicador:
*     Índice de dependencia potencial
*
* Fórmula:
*     idp = ((pob_0_14 + pob_65_mas) / pob_15_64) * 100
*
* Numerador:
*     población de 0 a 14 años + población de 65 años y más
*
* Denominador:
*     población de 15 a 64 años
*
* Fuentes:
*     2005–2011:
*         _out/Bases DTA/pob_edad_salud_2005_20XX.dta
*
*     2012–20XX:
*         _out/Bases DTA/pob_edades_simples_long_2012_20XX.dta
*
* Rango temporal:
*     2005–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Unidad:
*     Personas potencialmente dependientes por cada 100 personas
*     de 15 a 64 años.
*
* Consideraciones:
*     - Para 2005–2011 se usa la base de grupos de edad en salud.
*     - Para 2012–20XX se reconstruyen los grupos desde edades simples
*       y grupos quinquenales.
*     - Como es un indicador posterior al primero, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 19A. Construir grupos etarios 2005–2011
* ------------------------------------------------------------

	tempfile grupos_2005_2011_idp

	use "$out_dta/pob_edad_salud_2005_2024.dta", clear

	keep if inrange(anio, 2005, 2011)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

	* Estandarizar codigo_geo como texto.
	capture confirm numeric variable codigo_geo

	if _rc == 0 {
		gen str6 codigo_geo_str = string(codigo_geo, "%02.0f")
		drop codigo_geo
		rename codigo_geo_str codigo_geo
	}

	replace codigo_geo = strtrim(codigo_geo)

	* En la estructura 2005–2011:
	*     grupo_0_4 = 0-4 años
	*     col_i     = 5-14 años
	*     col_j     = 15-64 años
	*     col_k     = 65 años y más
	*     col_l     = Total general

	gen double pob_0_14   = grupo_0_4 + col_i
	gen double pob_15_64  = col_j
	gen double pob_65_mas = col_k
	gen double pob_total  = col_l

	keep anio codigo_geo nombre_geo nivel_geo pob_0_14 pob_15_64 pob_65_mas pob_total

	isid anio codigo_geo

	save `grupos_2005_2011_idp', replace

* ------------------------------------------------------------
* 19B. Construir grupos etarios 2012–20XX
* ------------------------------------------------------------

	tempfile grupos_2012_2024_idp

	use "$out_dta/pob_edades_simples_long_2012_2024.dta", clear

	keep if inrange(anio, 2012, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

	* Usar población total por sexo.
	keep if sexo == "Total"

	* Estandarizar codigo_geo como texto.
	capture confirm numeric variable codigo_geo

	if _rc == 0 {
		gen str6 codigo_geo_str = string(codigo_geo, "%02.0f")
		drop codigo_geo
		rename codigo_geo_str codigo_geo
	}

	replace codigo_geo = strtrim(codigo_geo)

	* Asegurar que grupo_edad sea texto.
	capture confirm string variable grupo_edad

	if _rc != 0 {
		tostring grupo_edad, replace
	}

	replace grupo_edad = strtrim(grupo_edad)

	gen edad_simple = real(grupo_edad)

	* 0-14 años:
	*     edades simples 0 a 14.
	gen double pob_0_14 = 0
	replace pob_0_14 = poblacion if inrange(edad_simple, 0, 14)

	* 15-64 años:
	*     edades simples 15 a 29
	*     + grupos quinquenales 30_34 a 60_64.
	gen double pob_15_64 = 0
	replace pob_15_64 = poblacion if inrange(edad_simple, 15, 29)
	replace pob_15_64 = poblacion if regexm(grupo_edad, "^(30_34|35_39|40_44|45_49|50_54|55_59|60_64)$")

	* 65 años y más:
	*     grupos quinquenales 65_69, 70_74, 75_79 y 80_mas.
	gen double pob_65_mas = 0
	replace pob_65_mas = poblacion if regexm(grupo_edad, "^(65_69|70_74|75_79|80_mas)$")

	* Población total:
	*     grupo_edad == total.
	gen double pob_total = 0
	replace pob_total = poblacion if grupo_edad == "total"

	collapse (sum) pob_0_14 pob_15_64 pob_65_mas pob_total, by(anio codigo_geo nombre_geo nivel_geo)

	isid anio codigo_geo

	save `grupos_2012_2024_idp', replace

* ------------------------------------------------------------
* 19C. Unir ambos periodos
* ------------------------------------------------------------

	use `grupos_2005_2011_idp', clear

	append using `grupos_2012_2024_idp'

	sort anio codigo_geo

	isid anio codigo_geo

* ------------------------------------------------------------
* 19D. Control de consistencia
* ------------------------------------------------------------
* La población total debe ser igual a:
*
*     pob_total = pob_0_14 + pob_15_64 + pob_65_mas
*
* Si no coincide, el indicador se detiene.
* ------------------------------------------------------------

	gen double pob_total_suma = pob_0_14 + pob_15_64 + pob_65_mas

	gen double dif_abs = pob_total - pob_total_suma

	gen double dif_pct = .
	replace dif_pct = (dif_abs / pob_total) * 100 if pob_total > 0 & !missing(dif_abs, pob_total)

	format pob_0_14 pob_15_64 pob_65_mas pob_total pob_total_suma dif_abs %15.2fc
	format dif_pct %12.8f

	count if !missing(dif_abs) & abs(dif_abs) > 0.5

	if r(N) > 0 {
		di as error "ERROR: La población total no coincide con la suma de grupos IDP."

		preserve
			keep if !missing(dif_abs) & abs(dif_abs) > 0.5
			sort anio codigo_geo
			export excel using "$out_aud/error_idp_total_no_coincide.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	* Control del denominador.
	count if missing(pob_15_64) | pob_15_64 <= 0

	if r(N) > 0 {
		di as error "ERROR: Hay población de 15 a 64 años faltante o menor/igual a cero."

		preserve
			keep if missing(pob_15_64) | pob_15_64 <= 0
			sort anio codigo_geo
			export excel using "$out_aud/error_denominador_idp.xlsx", firstrow(variables) replace
		restore

		error 459
	}

* ------------------------------------------------------------
* 19E. Calcular indicador
* ------------------------------------------------------------
* Índice de dependencia potencial:
*
*     IDP = ((población 0–14 años + población 65 años y más)
*            / población 15–64 años) * 100
*
* Interpretación:
*     Personas potencialmente dependientes por cada 100 personas
*     en edad potencialmente activa.
* ------------------------------------------------------------

	gen idp = .

	replace idp = ((pob_0_14 + pob_65_mas) / pob_15_64) * 100 if pob_15_64 > 0 & !missing(pob_0_14, pob_65_mas, pob_15_64)

	replace idp = round(idp, .01)

* ------------------------------------------------------------
* 19F. Preparar tabla final
* ------------------------------------------------------------
* Nota:
*     No se exportan los componentes poblacionales.
*     Solo se exporta el índice de dependencia potencial.
* ------------------------------------------------------------

	keep anio codigo_geo idp

	reshape wide idp, i(anio) j(codigo_geo) string

	rename idp00 BOLIVIA
	rename idp01 CHUQUISACA
	rename idp02 LA_PAZ
	rename idp03 COCHABAMBA
	rename idp04 ORURO
	rename idp05 POTOSI
	rename idp06 TARIJA
	rename idp07 SANTA_CRUZ
	rename idp08 BENI
	rename idp09 PANDO

	order anio BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO

	sort anio

	format BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO %9.2f

* ------------------------------------------------------------
* 19G. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("indice_dependencia_pot") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 19H. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_003_05 | Índice de dependencia potencial"
	di as result "Unidad: personas dependientes por cada 100 personas de 15–64 años"
	di as result "Hoja: indice_dependencia_pot"
	di as result "============================================================"

	list, clean noobs

* ============================================================
**# BLOQUE 20: ÍNDICE DE ENVEJECIMIENTO
* ============================================================
*
* Código:
*     06_003_06
*
* Indicador:
*     Índice de envejecimiento
*
* Fórmula:
*     indice_envejecimiento = (pob_65_mas / pob_0_14) * 100
*
* Numerador:
*     población de 65 años y más
*
* Denominador:
*     población de 0 a 14 años
*
* Fuentes:
*     2005–2011:
*         _out/Bases DTA/pob_edad_salud_2005_20XX.dta
*
*     2012–20XX:
*         _out/Bases DTA/pob_edades_simples_long_2012_20XX.dta
*
* Rango temporal:
*     2005–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Unidad:
*     Personas de 65 años y más por cada 100 personas de 0 a 14 años.
*
* Consideraciones:
*     - Para 2005–2011 se usa la base de grupos de edad en salud.
*     - Para 2012–20XX se reconstruyen los grupos desde edades simples
*       y grupos quinquenales.
*     - Como es un indicador posterior al primero, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 20A. Construir grupos etarios 2005–2011
* ------------------------------------------------------------

	tempfile grupos_2005_2011_env

	use "$out_dta/pob_edad_salud_2005_2024.dta", clear

	keep if inrange(anio, 2005, 2011)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

	* Estandarizar codigo_geo como texto.
	capture confirm numeric variable codigo_geo

	if _rc == 0 {
		gen str6 codigo_geo_str = string(codigo_geo, "%02.0f")
		drop codigo_geo
		rename codigo_geo_str codigo_geo
	}

	replace codigo_geo = strtrim(codigo_geo)

	* En la estructura 2005–2011:
	*     grupo_0_4 = 0-4 años
	*     col_i     = 5-14 años
	*     col_j     = 15-64 años
	*     col_k     = 65 años y más
	*     col_l     = Total general

	gen double pob_0_14   = grupo_0_4 + col_i
	gen double pob_15_64  = col_j
	gen double pob_65_mas = col_k
	gen double pob_total  = col_l

	keep anio codigo_geo nombre_geo nivel_geo pob_0_14 pob_15_64 pob_65_mas pob_total

	isid anio codigo_geo

	save `grupos_2005_2011_env', replace

* ------------------------------------------------------------
* 20B. Construir grupos etarios 2012–20XX
* ------------------------------------------------------------

	tempfile grupos_2012_2024_env

	use "$out_dta/pob_edades_simples_long_2012_2024.dta", clear

	keep if inrange(anio, 2012, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

	* Usar población total por sexo.
	keep if sexo == "Total"

	* Estandarizar codigo_geo como texto.
	capture confirm numeric variable codigo_geo

	if _rc == 0 {
		gen str6 codigo_geo_str = string(codigo_geo, "%02.0f")
		drop codigo_geo
		rename codigo_geo_str codigo_geo
	}

	replace codigo_geo = strtrim(codigo_geo)

	* Asegurar que grupo_edad sea texto.
	capture confirm string variable grupo_edad

	if _rc != 0 {
		tostring grupo_edad, replace
	}

	replace grupo_edad = strtrim(grupo_edad)

	gen edad_simple = real(grupo_edad)

	* 0-14 años:
	*     edades simples 0 a 14.
	gen double pob_0_14 = 0
	replace pob_0_14 = poblacion if inrange(edad_simple, 0, 14)

	* 15-64 años:
	*     edades simples 15 a 29
	*     + grupos quinquenales 30_34 a 60_64.
	gen double pob_15_64 = 0
	replace pob_15_64 = poblacion if inrange(edad_simple, 15, 29)
	replace pob_15_64 = poblacion if regexm(grupo_edad, "^(30_34|35_39|40_44|45_49|50_54|55_59|60_64)$")

	* 65 años y más:
	*     grupos quinquenales 65_69, 70_74, 75_79 y 80_mas.
	gen double pob_65_mas = 0
	replace pob_65_mas = poblacion if regexm(grupo_edad, "^(65_69|70_74|75_79|80_mas)$")

	* Población total:
	*     grupo_edad == total.
	gen double pob_total = 0
	replace pob_total = poblacion if grupo_edad == "total"

	collapse (sum) pob_0_14 pob_15_64 pob_65_mas pob_total, by(anio codigo_geo nombre_geo nivel_geo)

	isid anio codigo_geo

	save `grupos_2012_2024_env', replace

* ------------------------------------------------------------
* 20C. Unir ambos periodos
* ------------------------------------------------------------

	use `grupos_2005_2011_env', clear

	append using `grupos_2012_2024_env'

	sort anio codigo_geo

	isid anio codigo_geo

* ------------------------------------------------------------
* 20D. Control de consistencia
* ------------------------------------------------------------
* La población total debe ser igual a:
*
*     pob_total = pob_0_14 + pob_15_64 + pob_65_mas
*
* Si no coincide, el indicador se detiene.
* ------------------------------------------------------------

	gen double pob_total_suma = pob_0_14 + pob_15_64 + pob_65_mas

	gen double dif_abs = pob_total - pob_total_suma

	gen double dif_pct = .
	replace dif_pct = (dif_abs / pob_total) * 100 if pob_total > 0 & !missing(dif_abs, pob_total)

	format pob_0_14 pob_15_64 pob_65_mas pob_total pob_total_suma dif_abs %15.2fc
	format dif_pct %12.8f

	count if !missing(dif_abs) & abs(dif_abs) > 0.5

	if r(N) > 0 {
		di as error "ERROR: La población total no coincide con la suma de grupos para índice de envejecimiento."

		preserve
			keep if !missing(dif_abs) & abs(dif_abs) > 0.5
			sort anio codigo_geo
			export excel using "$out_aud/error_indice_envejecimiento_total_no_coincide.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	* Control del denominador.
	count if missing(pob_0_14) | pob_0_14 <= 0

	if r(N) > 0 {
		di as error "ERROR: Hay población de 0 a 14 años faltante o menor/igual a cero."

		preserve
			keep if missing(pob_0_14) | pob_0_14 <= 0
			sort anio codigo_geo
			export excel using "$out_aud/error_denominador_indice_envejecimiento.xlsx", firstrow(variables) replace
		restore

		error 459
	}

* ------------------------------------------------------------
* 20E. Calcular indicador
* ------------------------------------------------------------
* Índice de envejecimiento:
*
*     IE = (población de 65 años y más / población de 0 a 14 años) * 100
*
* Interpretación:
*     Personas de 65 años y más por cada 100 personas de 0 a 14 años.
* ------------------------------------------------------------

	gen indice_envejecimiento = .

	replace indice_envejecimiento = (pob_65_mas / pob_0_14) * 100 if pob_0_14 > 0 & !missing(pob_65_mas, pob_0_14)

	replace indice_envejecimiento = round(indice_envejecimiento, .01)

* ------------------------------------------------------------
* 20F. Preparar tabla final
* ------------------------------------------------------------
* Nota:
*     No se exportan los componentes poblacionales.
*     Solo se exporta el índice de envejecimiento.
* ------------------------------------------------------------

	keep anio codigo_geo indice_envejecimiento

	reshape wide indice_envejecimiento, i(anio) j(codigo_geo) string

	rename indice_envejecimiento00 BOLIVIA
	rename indice_envejecimiento01 CHUQUISACA
	rename indice_envejecimiento02 LA_PAZ
	rename indice_envejecimiento03 COCHABAMBA
	rename indice_envejecimiento04 ORURO
	rename indice_envejecimiento05 POTOSI
	rename indice_envejecimiento06 TARIJA
	rename indice_envejecimiento07 SANTA_CRUZ
	rename indice_envejecimiento08 BENI
	rename indice_envejecimiento09 PANDO

	order anio BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO

	sort anio

	format BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO %9.2f

* ------------------------------------------------------------
* 20G. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("indice_envejecimiento") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 20H. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_003_06 | Índice de envejecimiento"
	di as result "Unidad: personas de 65 años y más por cada 100 personas de 0–14 años"
	di as result "Hoja: indice_envejecimiento"
	di as result "============================================================"

	list, clean noobs

* ============================================================
**# BLOQUE 21: ÍNDICE DE DEPENDENCIA JUVENIL
* ============================================================
*
* Código:
*     06_003_07
*
* Indicador:
*     Índice de dependencia juvenil
*
* Fórmula:
*     idp_jovenes = (pob_0_14 / pob_15_64) * 100
*
* Numerador:
*     población de 0 a 14 años
*
* Denominador:
*     población de 15 a 64 años
*
* Fuentes:
*     2005–2011:
*         _out/Bases DTA/pob_edad_salud_2005_20XX.dta
*
*     2012–20XX:
*         _out/Bases DTA/pob_edades_simples_long_2012_20XX.dta
*
* Rango temporal:
*     2005–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Unidad:
*     Personas de 0 a 14 años por cada 100 personas de 15 a 64 años.
*
* Consideraciones:
*     - Como es un indicador posterior al primero, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 21A. Construir grupos etarios 2005–2011
* ------------------------------------------------------------

	tempfile grupos_2005_2011_juvenil

	use "$out_dta/pob_edad_salud_2005_2024.dta", clear

	keep if inrange(anio, 2005, 2011)

	keep if inlist(nivel_geo, 0, 1)

	capture confirm numeric variable codigo_geo

	if _rc == 0 {
		gen str6 codigo_geo_str = string(codigo_geo, "%02.0f")
		drop codigo_geo
		rename codigo_geo_str codigo_geo
	}

	replace codigo_geo = strtrim(codigo_geo)

	gen double pob_0_14   = grupo_0_4 + col_i
	gen double pob_15_64  = col_j
	gen double pob_65_mas = col_k
	gen double pob_total  = col_l

	keep anio codigo_geo nombre_geo nivel_geo pob_0_14 pob_15_64 pob_65_mas pob_total

	isid anio codigo_geo

	save `grupos_2005_2011_juvenil', replace

* ------------------------------------------------------------
* 21B. Construir grupos etarios 2012–20XX
* ------------------------------------------------------------

	tempfile grupos_2012_2024_juvenil

	use "$out_dta/pob_edades_simples_long_2012_2024.dta", clear

	keep if inrange(anio, 2012, 2024)

	keep if inlist(nivel_geo, 0, 1)

	keep if sexo == "Total"

	capture confirm numeric variable codigo_geo

	if _rc == 0 {
		gen str6 codigo_geo_str = string(codigo_geo, "%02.0f")
		drop codigo_geo
		rename codigo_geo_str codigo_geo
	}

	replace codigo_geo = strtrim(codigo_geo)

	capture confirm string variable grupo_edad

	if _rc != 0 {
		tostring grupo_edad, replace
	}

	replace grupo_edad = strtrim(grupo_edad)

	gen edad_simple = real(grupo_edad)

	gen double pob_0_14 = 0
	replace pob_0_14 = poblacion if inrange(edad_simple, 0, 14)

	gen double pob_15_64 = 0
	replace pob_15_64 = poblacion if inrange(edad_simple, 15, 29)
	replace pob_15_64 = poblacion if regexm(grupo_edad, "^(30_34|35_39|40_44|45_49|50_54|55_59|60_64)$")

	gen double pob_65_mas = 0
	replace pob_65_mas = poblacion if regexm(grupo_edad, "^(65_69|70_74|75_79|80_mas)$")

	gen double pob_total = 0
	replace pob_total = poblacion if grupo_edad == "total"

	collapse (sum) pob_0_14 pob_15_64 pob_65_mas pob_total, by(anio codigo_geo nombre_geo nivel_geo)

	isid anio codigo_geo

	save `grupos_2012_2024_juvenil', replace

* ------------------------------------------------------------
* 21C. Unir ambos periodos
* ------------------------------------------------------------

	use `grupos_2005_2011_juvenil', clear

	append using `grupos_2012_2024_juvenil'

	sort anio codigo_geo

	isid anio codigo_geo

* ------------------------------------------------------------
* 21D. Control de consistencia
* ------------------------------------------------------------

	gen double pob_total_suma = pob_0_14 + pob_15_64 + pob_65_mas

	gen double dif_abs = pob_total - pob_total_suma

	gen double dif_pct = .
	replace dif_pct = (dif_abs / pob_total) * 100 if pob_total > 0 & !missing(dif_abs, pob_total)

	format pob_0_14 pob_15_64 pob_65_mas pob_total pob_total_suma dif_abs %15.2fc
	format dif_pct %12.8f

	count if !missing(dif_abs) & abs(dif_abs) > 0.5

	if r(N) > 0 {
		di as error "ERROR: La población total no coincide con la suma de grupos para dependencia juvenil."

		preserve
			keep if !missing(dif_abs) & abs(dif_abs) > 0.5
			sort anio codigo_geo
			export excel using "$out_aud/error_dependencia_juvenil_total_no_coincide.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	count if missing(pob_15_64) | pob_15_64 <= 0

	if r(N) > 0 {
		di as error "ERROR: Hay población de 15 a 64 años faltante o menor/igual a cero."

		preserve
			keep if missing(pob_15_64) | pob_15_64 <= 0
			sort anio codigo_geo
			export excel using "$out_aud/error_denominador_dependencia_juvenil.xlsx", firstrow(variables) replace
		restore

		error 459
	}

* ------------------------------------------------------------
* 21E. Calcular indicador
* ------------------------------------------------------------

	gen idp_jovenes = .

	replace idp_jovenes = (pob_0_14 / pob_15_64) * 100 if pob_15_64 > 0 & !missing(pob_0_14, pob_15_64)

	replace idp_jovenes = round(idp_jovenes, .01)

* ------------------------------------------------------------
* 21F. Preparar tabla final
* ------------------------------------------------------------

	keep anio codigo_geo idp_jovenes

	reshape wide idp_jovenes, i(anio) j(codigo_geo) string

	rename idp_jovenes00 BOLIVIA
	rename idp_jovenes01 CHUQUISACA
	rename idp_jovenes02 LA_PAZ
	rename idp_jovenes03 COCHABAMBA
	rename idp_jovenes04 ORURO
	rename idp_jovenes05 POTOSI
	rename idp_jovenes06 TARIJA
	rename idp_jovenes07 SANTA_CRUZ
	rename idp_jovenes08 BENI
	rename idp_jovenes09 PANDO

	order anio BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO

	sort anio

	format BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO %9.2f

* ------------------------------------------------------------
* 21G. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("indice_dependencia_juvenil") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 21H. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_003_07 | Índice de dependencia juvenil"
	di as result "Unidad: personas de 0–14 años por cada 100 personas de 15–64 años"
	di as result "Hoja: indice_dependencia_juvenil"
	di as result "============================================================"

	list, clean noobs

* ============================================================
**# BLOQUE 22: ÍNDICE DE DEPENDENCIA SENIL
* ============================================================
* Código:
*     06_003_08
*
* Indicador:
*     Índice de dependencia senil
*
* Fórmula:
*     idp_ancianos = (pob_65_mas / pob_15_64) * 100
*
* Numerador:
*     población de 65 años y más
*
* Denominador:
*     población de 15 a 64 años
*
* Fuentes:
*     2005–2011:
*         _out/Bases DTA/pob_edad_salud_2005_20XX.dta
*
*     2012–20XX:
*         _out/Bases DTA/pob_edades_simples_long_2012_20XX.dta
*
* Rango temporal:
*     2005–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Unidad:
*     Personas de 65 años y más por cada 100 personas de 15 a 64 años.
*
* Consideraciones:
*     - Como es un indicador posterior al primero, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 22A. Construir grupos etarios 2005–2011
* ------------------------------------------------------------

	tempfile grupos_2005_2011_senil

	use "$out_dta/pob_edad_salud_2005_2024.dta", clear

	keep if inrange(anio, 2005, 2011)

	keep if inlist(nivel_geo, 0, 1)

	capture confirm numeric variable codigo_geo

	if _rc == 0 {
		gen str6 codigo_geo_str = string(codigo_geo, "%02.0f")
		drop codigo_geo
		rename codigo_geo_str codigo_geo
	}

	replace codigo_geo = strtrim(codigo_geo)

	gen double pob_0_14   = grupo_0_4 + col_i
	gen double pob_15_64  = col_j
	gen double pob_65_mas = col_k
	gen double pob_total  = col_l

	keep anio codigo_geo nombre_geo nivel_geo pob_0_14 pob_15_64 pob_65_mas pob_total

	isid anio codigo_geo

	save `grupos_2005_2011_senil', replace

* ------------------------------------------------------------
* 22B. Construir grupos etarios 2012–20XX
* ------------------------------------------------------------

	tempfile grupos_2012_2024_senil

	use "$out_dta/pob_edades_simples_long_2012_2024.dta", clear

	keep if inrange(anio, 2012, 2024)

	keep if inlist(nivel_geo, 0, 1)

	keep if sexo == "Total"

	capture confirm numeric variable codigo_geo

	if _rc == 0 {
		gen str6 codigo_geo_str = string(codigo_geo, "%02.0f")
		drop codigo_geo
		rename codigo_geo_str codigo_geo
	}

	replace codigo_geo = strtrim(codigo_geo)

	capture confirm string variable grupo_edad

	if _rc != 0 {
		tostring grupo_edad, replace
	}

	replace grupo_edad = strtrim(grupo_edad)

	gen edad_simple = real(grupo_edad)

	gen double pob_0_14 = 0
	replace pob_0_14 = poblacion if inrange(edad_simple, 0, 14)

	gen double pob_15_64 = 0
	replace pob_15_64 = poblacion if inrange(edad_simple, 15, 29)
	replace pob_15_64 = poblacion if regexm(grupo_edad, "^(30_34|35_39|40_44|45_49|50_54|55_59|60_64)$")

	gen double pob_65_mas = 0
	replace pob_65_mas = poblacion if regexm(grupo_edad, "^(65_69|70_74|75_79|80_mas)$")

	gen double pob_total = 0
	replace pob_total = poblacion if grupo_edad == "total"

	collapse (sum) pob_0_14 pob_15_64 pob_65_mas pob_total, by(anio codigo_geo nombre_geo nivel_geo)

	isid anio codigo_geo

	save `grupos_2012_2024_senil', replace

* ------------------------------------------------------------
* 22C. Unir ambos periodos
* ------------------------------------------------------------

	use `grupos_2005_2011_senil', clear

	append using `grupos_2012_2024_senil'

	sort anio codigo_geo

	isid anio codigo_geo

* ------------------------------------------------------------
* 22D. Control de consistencia
* ------------------------------------------------------------

	gen double pob_total_suma = pob_0_14 + pob_15_64 + pob_65_mas

	gen double dif_abs = pob_total - pob_total_suma

	gen double dif_pct = .
	replace dif_pct = (dif_abs / pob_total) * 100 if pob_total > 0 & !missing(dif_abs, pob_total)

	format pob_0_14 pob_15_64 pob_65_mas pob_total pob_total_suma dif_abs %15.2fc
	format dif_pct %12.8f

	count if !missing(dif_abs) & abs(dif_abs) > 0.5

	if r(N) > 0 {
		di as error "ERROR: La población total no coincide con la suma de grupos para dependencia senil."

		preserve
			keep if !missing(dif_abs) & abs(dif_abs) > 0.5
			sort anio codigo_geo
			export excel using "$out_aud/error_dependencia_senil_total_no_coincide.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	count if missing(pob_15_64) | pob_15_64 <= 0

	if r(N) > 0 {
		di as error "ERROR: Hay población de 15 a 64 años faltante o menor/igual a cero."

		preserve
			keep if missing(pob_15_64) | pob_15_64 <= 0
			sort anio codigo_geo
			export excel using "$out_aud/error_denominador_dependencia_senil.xlsx", firstrow(variables) replace
		restore

		error 459
	}

* ------------------------------------------------------------
* 22E. Calcular indicador
* ------------------------------------------------------------

	gen idp_ancianos = .

	replace idp_ancianos = (pob_65_mas / pob_15_64) * 100 if pob_15_64 > 0 & !missing(pob_65_mas, pob_15_64)

	replace idp_ancianos = round(idp_ancianos, .01)

* ------------------------------------------------------------
* 22F. Preparar tabla final
* ------------------------------------------------------------

	keep anio codigo_geo idp_ancianos

	reshape wide idp_ancianos, i(anio) j(codigo_geo) string

	rename idp_ancianos00 BOLIVIA
	rename idp_ancianos01 CHUQUISACA
	rename idp_ancianos02 LA_PAZ
	rename idp_ancianos03 COCHABAMBA
	rename idp_ancianos04 ORURO
	rename idp_ancianos05 POTOSI
	rename idp_ancianos06 TARIJA
	rename idp_ancianos07 SANTA_CRUZ
	rename idp_ancianos08 BENI
	rename idp_ancianos09 PANDO

	order anio BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO

	sort anio

	format BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO %9.2f

* ------------------------------------------------------------
* 22G. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("indice_dependencia_senil") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 22H. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_003_08 | Índice de dependencia senil"
	di as result "Unidad: personas de 65 años y más por cada 100 personas de 15–64 años"
	di as result "Hoja: indice_dependencia_senil"
	di as result "============================================================"

	list, clean noobs

* ============================================================
**# BLOQUE 23: EDAD MEDIANA DE LA POBLACIÓN
* ============================================================
*
* Código:
*     06_003_09
*
* Indicador:
*     Edad mediana de la población
*
* Fórmula:
*     edad_mediana = L + ((N/2 - F_ant) / f_mediana) * amplitud
*
* Donde:
*     L          = límite inferior del grupo donde cae la mediana
*     N          = población total
*     F_ant      = frecuencia acumulada antes del grupo mediano
*     f_mediana  = población del grupo mediano
*     amplitud   = amplitud del grupo mediano
*
* Fuente:
*     _out/Bases DTA/pob_edades_simples_long_2012_20XX.dta
*
* Rango temporal:
*     2012–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Unidad:
*     Años de edad.
*
* Consideraciones:
*     - Se calcula desde 2012 porque desde ese año la base tiene
*       edades simples de 0 a 29 años y grupos quinquenales desde
*       30_34 hasta 80_mas.
*     - No se calcula para 2005–2011 porque la base disponible agrupa
*       la población adulta en intervalos demasiado amplios.
*     - La edad mediana se calcula usando interpolación dentro del
*       grupo etario donde se alcanza el 50% de la población acumulada.
*     - Como es un indicador posterior al primero, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 23A. Abrir base fuente
* ------------------------------------------------------------

	use "$out_dta/pob_edades_simples_long_2012_2024.dta", clear

	keep if inrange(anio, 2012, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

	* Usar población total por sexo.
	keep if sexo == "Total"

	* Estandarizar codigo_geo como texto.
	capture confirm numeric variable codigo_geo

	if _rc == 0 {
		gen str6 codigo_geo_str = string(codigo_geo, "%02.0f")
		drop codigo_geo
		rename codigo_geo_str codigo_geo
	}

	replace codigo_geo = strtrim(codigo_geo)

	* Asegurar que grupo_edad sea texto.
	capture confirm string variable grupo_edad

	if _rc != 0 {
		tostring grupo_edad, replace
	}

	replace grupo_edad = strtrim(grupo_edad)

* ------------------------------------------------------------
* 23B. Preparar intervalos etarios
* ------------------------------------------------------------
* Para edades simples:
*     grupo_edad = "0", "1", ..., "29"
*     límite inferior = edad
*     amplitud = 1
*
* Para grupos quinquenales:
*     30_34, 35_39, ..., 75_79
*     límite inferior = 30, 35, ..., 75
*     amplitud = 5
*
* Para 80_mas:
*     Se identifica como grupo abierto.
*     La mediana no debería caer en este grupo; si cayera, se deja
*     como missing por no tener límite superior definido.
* ------------------------------------------------------------

	gen edad_simple = real(grupo_edad)

	gen double edad_inf = .
	gen double amplitud = .

	replace edad_inf = edad_simple if !missing(edad_simple)
	replace amplitud = 1 if !missing(edad_simple)

	replace edad_inf = 30 if grupo_edad == "30_34"
	replace edad_inf = 35 if grupo_edad == "35_39"
	replace edad_inf = 40 if grupo_edad == "40_44"
	replace edad_inf = 45 if grupo_edad == "45_49"
	replace edad_inf = 50 if grupo_edad == "50_54"
	replace edad_inf = 55 if grupo_edad == "55_59"
	replace edad_inf = 60 if grupo_edad == "60_64"
	replace edad_inf = 65 if grupo_edad == "65_69"
	replace edad_inf = 70 if grupo_edad == "70_74"
	replace edad_inf = 75 if grupo_edad == "75_79"

	replace amplitud = 5 if regexm(grupo_edad, "^(30_34|35_39|40_44|45_49|50_54|55_59|60_64|65_69|70_74|75_79)$")

	replace edad_inf = 80 if grupo_edad == "80_mas"
	replace amplitud = . if grupo_edad == "80_mas"

	* Excluir la fila total del cálculo acumulado.
	drop if grupo_edad == "total"

	* Control de intervalos faltantes.
	count if missing(edad_inf)

	if r(N) > 0 {
		di as error "ERROR: Hay grupos de edad sin límite inferior identificado."

		preserve
			keep if missing(edad_inf)
			sort anio codigo_geo grupo_edad
			export excel using "$out_aud/error_grupos_sin_edad_inf_mediana.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	count if missing(amplitud) & grupo_edad != "80_mas"

	if r(N) > 0 {
		di as error "ERROR: Hay grupos de edad sin amplitud identificada."

		preserve
			keep if missing(amplitud) & grupo_edad != "80_mas"
			sort anio codigo_geo grupo_edad
			export excel using "$out_aud/error_grupos_sin_amplitud_mediana.xlsx", firstrow(variables) replace
		restore

		error 459
	}

* ------------------------------------------------------------
* 23C. Ordenar y calcular acumulados
* ------------------------------------------------------------

	sort anio codigo_geo edad_inf

	by anio codigo_geo: gen double pob_acum = sum(poblacion)

	by anio codigo_geo: gen double pob_total = pob_acum[_N]

	by anio codigo_geo: gen double mitad_pob = pob_total / 2

	by anio codigo_geo: gen double pob_acum_ant = pob_acum - poblacion

* ------------------------------------------------------------
* 23D. Identificar grupo mediano
* ------------------------------------------------------------

	gen byte es_grupo_mediano = 0

	replace es_grupo_mediano = 1 if pob_acum >= mitad_pob & pob_acum_ant < mitad_pob

	count if es_grupo_mediano == 1

	di as result "Grupos medianos identificados: " r(N)

	* Debe haber un grupo mediano por año y territorio.
	bysort anio codigo_geo: egen n_grupos_medianos = total(es_grupo_mediano)

	count if n_grupos_medianos != 1

	if r(N) > 0 {
		di as error "ERROR: Hay año-territorio sin grupo mediano único."

		preserve
			keep if n_grupos_medianos != 1
			sort anio codigo_geo edad_inf
			export excel using "$out_aud/error_grupo_mediano_no_unico.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	* Control: si la mediana cae en 80_mas, no se calcula.
	count if es_grupo_mediano == 1 & grupo_edad == "80_mas"

	if r(N) > 0 {
		di as error "ERROR: La mediana cae en el grupo abierto 80_mas."

		preserve
			keep if es_grupo_mediano == 1 & grupo_edad == "80_mas"
			sort anio codigo_geo
			export excel using "$out_aud/error_mediana_en_grupo_abierto_80mas.xlsx", firstrow(variables) replace
		restore

		error 459
	}

* ------------------------------------------------------------
* 23E. Calcular edad mediana
* ------------------------------------------------------------

	gen double edad_mediana = .

	replace edad_mediana = edad_inf + ((mitad_pob - pob_acum_ant) / poblacion) * amplitud if es_grupo_mediano == 1 & poblacion > 0 & !missing(edad_inf, amplitud, mitad_pob, pob_acum_ant, poblacion)

	replace edad_mediana = round(edad_mediana, .01)

* ------------------------------------------------------------
* 23F. Preparar tabla final
* ------------------------------------------------------------

	keep if es_grupo_mediano == 1

	keep anio codigo_geo edad_mediana

	reshape wide edad_mediana, i(anio) j(codigo_geo) string

	rename edad_mediana00 BOLIVIA
	rename edad_mediana01 CHUQUISACA
	rename edad_mediana02 LA_PAZ
	rename edad_mediana03 COCHABAMBA
	rename edad_mediana04 ORURO
	rename edad_mediana05 POTOSI
	rename edad_mediana06 TARIJA
	rename edad_mediana07 SANTA_CRUZ
	rename edad_mediana08 BENI
	rename edad_mediana09 PANDO

	order anio BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO

	sort anio

	format BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO %9.2f

* ------------------------------------------------------------
* 23G. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("edad_mediana_poblacion") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 23H. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_003_09 | Edad mediana de la población"
	di as result "Unidad: años de edad"
	di as result "Hoja: edad_mediana_poblacion"
	di as result "============================================================"

	list, clean noobs

* ============================================================
**# BLOQUE 24: EDAD MEDIA DE LA POBLACIÓN
* ============================================================
* Código:
*     06_003_10
*
* Indicador:
*     Edad media de la población
*
* Fórmula:
*     edad_media = sum(poblacion * edad_referencia) / poblacion_total
*
* Numerador:
*     suma ponderada de edades:
*         poblacion * edad_referencia
*
* Denominador:
*     población total
*
* Fuente:
*     _out/Bases DTA/pob_edades_simples_long_2012_20XX.dta
*
* Rango temporal:
*     2012–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Unidad:
*     Años de edad.
*
* Consideraciones:
*     - Se calcula desde 2012 porque desde ese año la base tiene
*       edades simples de 0 a 29 años y grupos quinquenales desde
*       30_34 hasta 80_mas.
*     - No se calcula para 2005–2011 porque la base disponible agrupa
*       la población adulta en intervalos demasiado amplios.
*     - Para edades simples se usa como edad de referencia:
*           edad + 0.5
*       Por ejemplo:
*           grupo 0 = 0.5 años
*           grupo 1 = 1.5 años
*     - Para grupos quinquenales se usa el punto medio:
*           30_34 = 32.5
*           35_39 = 37.5
*           etc.
*     - Para el grupo abierto 80_mas se usa 82.5 como aproximación.
*     - Como es un indicador posterior al primero, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 24A. Abrir base fuente
* ------------------------------------------------------------

	use "$out_dta/pob_edades_simples_long_2012_2024.dta", clear

	keep if inrange(anio, 2012, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

	* Usar población total por sexo.
	keep if sexo == "Total"

	* Estandarizar codigo_geo como texto.
	capture confirm numeric variable codigo_geo

	if _rc == 0 {
		gen str6 codigo_geo_str = string(codigo_geo, "%02.0f")
		drop codigo_geo
		rename codigo_geo_str codigo_geo
	}

	replace codigo_geo = strtrim(codigo_geo)

	* Asegurar que grupo_edad sea texto.
	capture confirm string variable grupo_edad

	if _rc != 0 {
		tostring grupo_edad, replace
	}

	replace grupo_edad = strtrim(grupo_edad)

* ------------------------------------------------------------
* 24B. Definir edad de referencia
* ------------------------------------------------------------
* Para edades simples:
*     edad_ref = edad + 0.5
*
* Para grupos quinquenales:
*     edad_ref = punto medio del intervalo
*
* Para 80_mas:
*     edad_ref = 82.5
* ------------------------------------------------------------

	gen edad_simple = real(grupo_edad)

	gen double edad_ref = .

	* Edades simples 0 a 29.
	replace edad_ref = edad_simple + 0.5 if !missing(edad_simple)

	* Grupos quinquenales.
	replace edad_ref = 32.5 if grupo_edad == "30_34"
	replace edad_ref = 37.5 if grupo_edad == "35_39"
	replace edad_ref = 42.5 if grupo_edad == "40_44"
	replace edad_ref = 47.5 if grupo_edad == "45_49"
	replace edad_ref = 52.5 if grupo_edad == "50_54"
	replace edad_ref = 57.5 if grupo_edad == "55_59"
	replace edad_ref = 62.5 if grupo_edad == "60_64"
	replace edad_ref = 67.5 if grupo_edad == "65_69"
	replace edad_ref = 72.5 if grupo_edad == "70_74"
	replace edad_ref = 77.5 if grupo_edad == "75_79"

	* Grupo abierto.
	replace edad_ref = 82.5 if grupo_edad == "80_mas"

	* Excluir fila total del cálculo.
	drop if grupo_edad == "total"

* ------------------------------------------------------------
* 24C. Controles previos
* ------------------------------------------------------------

	count if missing(edad_ref)

	if r(N) > 0 {
		di as error "ERROR: Hay grupos de edad sin edad de referencia."

		preserve
			keep if missing(edad_ref)
			sort anio codigo_geo grupo_edad
			export excel using "$out_aud/error_grupos_sin_edad_ref_edad_media.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	count if missing(poblacion) | poblacion < 0

	if r(N) > 0 {
		di as error "ERROR: Hay población faltante o negativa."

		preserve
			keep if missing(poblacion) | poblacion < 0
			sort anio codigo_geo grupo_edad
			export excel using "$out_aud/error_poblacion_edad_media.xlsx", firstrow(variables) replace
		restore

		error 459
	}

* ------------------------------------------------------------
* 24D. Calcular edad media
* ------------------------------------------------------------

	gen double pob_x_edad = poblacion * edad_ref

	collapse (sum) pob_total = poblacion suma_pob_x_edad = pob_x_edad, by(anio codigo_geo nombre_geo nivel_geo)

	gen double edad_media = .

	replace edad_media = suma_pob_x_edad / pob_total if pob_total > 0 & !missing(suma_pob_x_edad, pob_total)

	replace edad_media = round(edad_media, .01)

* ------------------------------------------------------------
* 24E. Controles del indicador
* ------------------------------------------------------------

	count if missing(edad_media)

	if r(N) > 0 {
		di as error "ERROR: Hay edad media faltante."

		preserve
			keep if missing(edad_media)
			sort anio codigo_geo
			export excel using "$out_aud/error_edad_media_faltante.xlsx", firstrow(variables) replace
		restore

		error 459
	}

* ------------------------------------------------------------
* 24F. Preparar tabla final
* ------------------------------------------------------------

	keep anio codigo_geo edad_media

	reshape wide edad_media, i(anio) j(codigo_geo) string

	rename edad_media00 BOLIVIA
	rename edad_media01 CHUQUISACA
	rename edad_media02 LA_PAZ
	rename edad_media03 COCHABAMBA
	rename edad_media04 ORURO
	rename edad_media05 POTOSI
	rename edad_media06 TARIJA
	rename edad_media07 SANTA_CRUZ
	rename edad_media08 BENI
	rename edad_media09 PANDO

	order anio BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO

	sort anio

	format BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO %9.2f

* ------------------------------------------------------------
* 24G. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("edad_media_poblacion") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 24H. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_003_10 | Edad media de la población"
	di as result "Unidad: años de edad"
	di as result "Hoja: edad_media_poblacion"
	di as result "============================================================"

	list, clean noobs

* ============================================================
**# BLOQUE 25: ÍNDICE DE REEMPLAZAMIENTO DE LA POBLACIÓN
*            EN EDAD ACTIVA
* ============================================================
* Código:
*     06_003_11
*
* Indicador:
*     Índice de reemplazamiento de la población en edad activa
*
* Sigla:
*     IRPEA
*
* Fórmula:
*     irpea = (pob_60_64 / pob_15_19) * 100
*
* Numerador:
*     población de 60 a 64 años
*
* Denominador:
*     población de 15 a 19 años
*
* Fuente:
*     _out/Bases DTA/pob_edades_simples_long_2012_20XX.dta
*
* Rango temporal:
*     2012–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Unidad:
*     Personas de 60 a 64 años por cada 100 personas de 15 a 19 años.
*
* Consideraciones:
*     - El grupo 15-19 se construye sumando las edades simples:
*           15 + 16 + 17 + 18 + 19
*     - El grupo 60-64 se toma directamente de grupo_edad == "60_64".
*     - No se calcula para 2005–2011 porque la base disponible no
*       contiene estos grupos específicos.
*     - Como es un indicador posterior al primero, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 25A. Abrir base fuente
* ------------------------------------------------------------

	use "$out_dta/pob_edades_simples_long_2012_2024.dta", clear

	keep if inrange(anio, 2012, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

	* Usar población total por sexo.
	keep if sexo == "Total"

	* Estandarizar codigo_geo como texto.
	capture confirm numeric variable codigo_geo

	if _rc == 0 {
		gen str6 codigo_geo_str = string(codigo_geo, "%02.0f")
		drop codigo_geo
		rename codigo_geo_str codigo_geo
	}

	replace codigo_geo = strtrim(codigo_geo)

	* Asegurar que grupo_edad sea texto.
	capture confirm string variable grupo_edad

	if _rc != 0 {
		tostring grupo_edad, replace
	}

	replace grupo_edad = strtrim(grupo_edad)

* ------------------------------------------------------------
* 25B. Construir numerador y denominador
* ------------------------------------------------------------

	gen edad_simple = real(grupo_edad)

	* Denominador:
	*     población de 15 a 19 años.
	gen double pob_15_19 = 0
	replace pob_15_19 = poblacion if inrange(edad_simple, 15, 19)

	* Numerador:
	*     población de 60 a 64 años.
	gen double pob_60_64 = 0
	replace pob_60_64 = poblacion if grupo_edad == "60_64"

* ------------------------------------------------------------
* 25C. Colapsar a año + territorio
* ------------------------------------------------------------

	collapse (sum) pob_15_19 pob_60_64, by(anio codigo_geo nombre_geo nivel_geo)

	isid anio codigo_geo

* ------------------------------------------------------------
* 25D. Controles previos
* ------------------------------------------------------------

	count if missing(pob_15_19) | pob_15_19 <= 0

	if r(N) > 0 {
		di as error "ERROR: Hay población de 15 a 19 años faltante o menor/igual a cero."

		preserve
			keep if missing(pob_15_19) | pob_15_19 <= 0
			sort anio codigo_geo
			export excel using "$out_aud/error_denominador_irpea.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	count if missing(pob_60_64) | pob_60_64 < 0

	if r(N) > 0 {
		di as error "ERROR: Hay población de 60 a 64 años faltante o negativa."

		preserve
			keep if missing(pob_60_64) | pob_60_64 < 0
			sort anio codigo_geo
			export excel using "$out_aud/error_numerador_irpea.xlsx", firstrow(variables) replace
		restore

		error 459
	}

* ------------------------------------------------------------
* 25E. Calcular indicador
* ------------------------------------------------------------
* Índice de reemplazamiento de la población en edad activa:
*
*     IRPEA = población 60-64 / población 15-19 * 100
*
* Interpretación:
*     Personas próximas a salir de la edad activa por cada 100
*     personas próximas a ingresar a la edad activa.
* ------------------------------------------------------------

	gen irpea = .

	replace irpea = (pob_60_64 / pob_15_19) * 100 if pob_15_19 > 0 & !missing(pob_60_64, pob_15_19)

	replace irpea = round(irpea, .01)

* ------------------------------------------------------------
* 25F. Preparar tabla final
* ------------------------------------------------------------

	keep anio codigo_geo irpea

	reshape wide irpea, i(anio) j(codigo_geo) string

	rename irpea00 BOLIVIA
	rename irpea01 CHUQUISACA
	rename irpea02 LA_PAZ
	rename irpea03 COCHABAMBA
	rename irpea04 ORURO
	rename irpea05 POTOSI
	rename irpea06 TARIJA
	rename irpea07 SANTA_CRUZ
	rename irpea08 BENI
	rename irpea09 PANDO

	order anio BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO

	sort anio

	format BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO %9.2f

* ------------------------------------------------------------
* 25G. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("indice_reemplazamiento_pea") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 25H. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_003_11 | Índice de reemplazamiento de la población en edad activa"
	di as result "Unidad: personas de 60–64 por cada 100 personas de 15–19 años"
	di as result "Hoja: indice_reemplazamiento_pea"
	di as result "============================================================"

	list, clean noobs
	
* ============================================================
**# BLOQUE 26: ÍNDICE DE REEMPLAZAMIENTO DE LA POBLACIÓN
*            EN EDAD ACTIVA, POR SEXO
* ============================================================
* Código:
*     06_003_11.1
*
* Indicador:
*     Índice de reemplazamiento de la población en edad activa,
*     por sexo
*
* Sigla:
*     IRPEA_SEXO
*
* Fórmula:
*     irpea_sexo = (pob_60_64 / pob_15_19) * 100
*
* Numerador:
*     población de 60 a 64 años, por sexo
*
* Denominador:
*     población de 15 a 19 años, por sexo
*
* Fuente:
*     _out/Bases DTA/pob_edades_simples_long_2012_20XX.dta
*
* Rango temporal:
*     2012–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Desagregación:
*     Hombre
*     Mujer
*
* Unidad:
*     Personas de 60 a 64 años por cada 100 personas de 15 a 19 años,
*     dentro de cada sexo.
*
* Consideraciones:
*     - El grupo 15-19 se construye sumando las edades simples:
*           15 + 16 + 17 + 18 + 19
*     - El grupo 60-64 se toma directamente de grupo_edad == "60_64".
*     - No se calcula para 2005–2011 porque la base disponible no
*       contiene estos grupos específicos por sexo.
*     - No se incluye "Total", porque el indicador total ya fue
*       calculado en 06_003_11.
*     - Como es un indicador posterior al primero, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 26A. Abrir base fuente
* ------------------------------------------------------------

	use "$out_dta/pob_edades_simples_long_2012_2024.dta", clear

	keep if inrange(anio, 2012, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

	* Mantener solo hombres y mujeres.
	keep if inlist(sexo, "Hombre", "Mujer")

	* Estandarizar codigo_geo como texto.
	capture confirm numeric variable codigo_geo

	if _rc == 0 {
		gen str6 codigo_geo_str = string(codigo_geo, "%02.0f")
		drop codigo_geo
		rename codigo_geo_str codigo_geo
	}

	replace codigo_geo = strtrim(codigo_geo)

	* Asegurar que grupo_edad sea texto.
	capture confirm string variable grupo_edad

	if _rc != 0 {
		tostring grupo_edad, replace
	}

	replace grupo_edad = strtrim(grupo_edad)

* ------------------------------------------------------------
* 26B. Construir numerador y denominador por sexo
* ------------------------------------------------------------

	gen edad_simple = real(grupo_edad)

	* Denominador:
	*     población de 15 a 19 años.
	gen double pob_15_19 = 0
	replace pob_15_19 = poblacion if inrange(edad_simple, 15, 19)

	* Numerador:
	*     población de 60 a 64 años.
	gen double pob_60_64 = 0
	replace pob_60_64 = poblacion if grupo_edad == "60_64"

* ------------------------------------------------------------
* 26C. Colapsar a año + territorio + sexo
* ------------------------------------------------------------

	collapse (sum) pob_15_19 pob_60_64, by(anio codigo_geo nombre_geo nivel_geo sexo)

	isid anio codigo_geo sexo

* ------------------------------------------------------------
* 26D. Controles previos
* ------------------------------------------------------------

	count if missing(pob_15_19) | pob_15_19 <= 0

	if r(N) > 0 {
		di as error "ERROR: Hay población de 15 a 19 años faltante o menor/igual a cero por sexo."

		preserve
			keep if missing(pob_15_19) | pob_15_19 <= 0
			sort anio codigo_geo sexo
			export excel using "$out_aud/error_denominador_irpea_sexo.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	count if missing(pob_60_64) | pob_60_64 < 0

	if r(N) > 0 {
		di as error "ERROR: Hay población de 60 a 64 años faltante o negativa por sexo."

		preserve
			keep if missing(pob_60_64) | pob_60_64 < 0
			sort anio codigo_geo sexo
			export excel using "$out_aud/error_numerador_irpea_sexo.xlsx", firstrow(variables) replace
		restore

		error 459
	}

* ------------------------------------------------------------
* 26E. Calcular indicador
* ------------------------------------------------------------

	gen irpea_sexo = .

	replace irpea_sexo = (pob_60_64 / pob_15_19) * 100 if pob_15_19 > 0 & !missing(pob_60_64, pob_15_19)

	replace irpea_sexo = round(irpea_sexo, .01)

* ------------------------------------------------------------
* 26F. Preparar tabla final
* ------------------------------------------------------------

	gen sexo_col = ""

	replace sexo_col = "HOMBRE" if sexo == "Hombre"
	replace sexo_col = "MUJER"  if sexo == "Mujer"

	gen geo_sexo = codigo_geo + "_" + sexo_col

	keep anio geo_sexo irpea_sexo

	reshape wide irpea_sexo, i(anio) j(geo_sexo) string

	rename irpea_sexo00_HOMBRE BOLIVIA_HOMBRE
	rename irpea_sexo00_MUJER  BOLIVIA_MUJER

	rename irpea_sexo01_HOMBRE CHUQUISACA_HOMBRE
	rename irpea_sexo01_MUJER  CHUQUISACA_MUJER

	rename irpea_sexo02_HOMBRE LA_PAZ_HOMBRE
	rename irpea_sexo02_MUJER  LA_PAZ_MUJER

	rename irpea_sexo03_HOMBRE COCHABAMBA_HOMBRE
	rename irpea_sexo03_MUJER  COCHABAMBA_MUJER

	rename irpea_sexo04_HOMBRE ORURO_HOMBRE
	rename irpea_sexo04_MUJER  ORURO_MUJER

	rename irpea_sexo05_HOMBRE POTOSI_HOMBRE
	rename irpea_sexo05_MUJER  POTOSI_MUJER

	rename irpea_sexo06_HOMBRE TARIJA_HOMBRE
	rename irpea_sexo06_MUJER  TARIJA_MUJER

	rename irpea_sexo07_HOMBRE SANTA_CRUZ_HOMBRE
	rename irpea_sexo07_MUJER  SANTA_CRUZ_MUJER

	rename irpea_sexo08_HOMBRE BENI_HOMBRE
	rename irpea_sexo08_MUJER  BENI_MUJER

	rename irpea_sexo09_HOMBRE PANDO_HOMBRE
	rename irpea_sexo09_MUJER  PANDO_MUJER

	order anio ///
		BOLIVIA_HOMBRE BOLIVIA_MUJER ///
		CHUQUISACA_HOMBRE CHUQUISACA_MUJER ///
		LA_PAZ_HOMBRE LA_PAZ_MUJER ///
		COCHABAMBA_HOMBRE COCHABAMBA_MUJER ///
		ORURO_HOMBRE ORURO_MUJER ///
		POTOSI_HOMBRE POTOSI_MUJER ///
		TARIJA_HOMBRE TARIJA_MUJER ///
		SANTA_CRUZ_HOMBRE SANTA_CRUZ_MUJER ///
		BENI_HOMBRE BENI_MUJER ///
		PANDO_HOMBRE PANDO_MUJER

	sort anio

	format *_HOMBRE *_MUJER %9.2f

* ------------------------------------------------------------
* 26G. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("irpea_sexo") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 26H. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_003_11.1 | IRPEA por sexo"
	di as result "Unidad: personas de 60–64 por cada 100 personas de 15–19 años"
	di as result "Hoja: irpea_sexo"
	di as result "============================================================"

	list, clean noobs

* ============================================================
**# BLOQUE 27: ÍNDICE DE ESTRUCTURA DE LA POBLACIÓN
*            EN EDAD ACTIVA
* ============================================================
* Código:
*     06_003_12
*
* Indicador:
*     Índice de estructura de la población en edad activa
*
* Sigla:
*     IEPEA
*
* Fórmula:
*     iepea = (pob_40_64 / pob_15_39) * 100
*
* Numerador:
*     población de 40 a 64 años
*
* Denominador:
*     población de 15 a 39 años
*
* Fuente:
*     _out/Bases DTA/pob_edades_simples_long_2012_20XX.dta
*
* Rango temporal:
*     2012–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Unidad:
*     Personas de 40 a 64 años por cada 100 personas de 15 a 39 años.
*
* Consideraciones:
*     - Se calcula desde 2012 porque desde ese año la base tiene
*       edades simples y grupos quinquenales suficientes.
*     - El grupo 15-39 se construye como:
*           edades simples 15 a 29
*           + grupos 30_34 y 35_39
*     - El grupo 40-64 se construye como:
*           grupos 40_44, 45_49, 50_54, 55_59 y 60_64
*     - No se calcula para 2005–2011 porque la base disponible no
*       contiene estos grupos específicos.
*     - Como es un indicador posterior al primero, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 27A. Abrir base fuente
* ------------------------------------------------------------

	use "$out_dta/pob_edades_simples_long_2012_2024.dta", clear

	keep if inrange(anio, 2012, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

	* Usar población total por sexo.
	keep if sexo == "Total"

	* Estandarizar codigo_geo como texto.
	capture confirm numeric variable codigo_geo

	if _rc == 0 {
		gen str6 codigo_geo_str = string(codigo_geo, "%02.0f")
		drop codigo_geo
		rename codigo_geo_str codigo_geo
	}

	replace codigo_geo = strtrim(codigo_geo)

	* Asegurar que grupo_edad sea texto.
	capture confirm string variable grupo_edad

	if _rc != 0 {
		tostring grupo_edad, replace
	}

	replace grupo_edad = strtrim(grupo_edad)

* ------------------------------------------------------------
* 27B. Construir numerador y denominador
* ------------------------------------------------------------

	gen edad_simple = real(grupo_edad)

	* Denominador:
	*     población de 15 a 39 años.
	gen double pob_15_39 = 0

	replace pob_15_39 = poblacion if inrange(edad_simple, 15, 29)

	replace pob_15_39 = poblacion if regexm(grupo_edad, "^(30_34|35_39)$")


	* Numerador:
	*     población de 40 a 64 años.
	gen double pob_40_64 = 0

	replace pob_40_64 = poblacion if regexm(grupo_edad, "^(40_44|45_49|50_54|55_59|60_64)$")

* ------------------------------------------------------------
* 27C. Colapsar a año + territorio
* ------------------------------------------------------------

	collapse (sum) pob_15_39 pob_40_64, by(anio codigo_geo nombre_geo nivel_geo)

	isid anio codigo_geo

* ------------------------------------------------------------
* 27D. Controles previos
* ------------------------------------------------------------

	count if missing(pob_15_39) | pob_15_39 <= 0

	if r(N) > 0 {
		di as error "ERROR: Hay población de 15 a 39 años faltante o menor/igual a cero."

		preserve
			keep if missing(pob_15_39) | pob_15_39 <= 0
			sort anio codigo_geo
			export excel using "$out_aud/error_denominador_iepea.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	count if missing(pob_40_64) | pob_40_64 < 0

	if r(N) > 0 {
		di as error "ERROR: Hay población de 40 a 64 años faltante o negativa."

		preserve
			keep if missing(pob_40_64) | pob_40_64 < 0
			sort anio codigo_geo
			export excel using "$out_aud/error_numerador_iepea.xlsx", firstrow(variables) replace
		restore

		error 459
	}

* ------------------------------------------------------------
* 27E. Calcular indicador
* ------------------------------------------------------------
* Índice de estructura de la población en edad activa:
*
*     IEPEA = población 40-64 / población 15-39 * 100
*
* Interpretación:
*     Personas de 40 a 64 años por cada 100 personas de 15 a 39 años.
* ------------------------------------------------------------

	gen iepea = .

	replace iepea = (pob_40_64 / pob_15_39) * 100 if pob_15_39 > 0 & !missing(pob_40_64, pob_15_39)

	replace iepea = round(iepea, .01)

* ------------------------------------------------------------
* 27F. Preparar tabla final
* ------------------------------------------------------------

	keep anio codigo_geo iepea

	reshape wide iepea, i(anio) j(codigo_geo) string

	rename iepea00 BOLIVIA
	rename iepea01 CHUQUISACA
	rename iepea02 LA_PAZ
	rename iepea03 COCHABAMBA
	rename iepea04 ORURO
	rename iepea05 POTOSI
	rename iepea06 TARIJA
	rename iepea07 SANTA_CRUZ
	rename iepea08 BENI
	rename iepea09 PANDO

	order anio BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO

	sort anio

	format BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO %9.2f

* ------------------------------------------------------------
* 27G. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("indice_estructura_pea") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 27H. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_003_12 | Índice de estructura de la población en edad activa"
	di as result "Unidad: personas de 40–64 por cada 100 personas de 15–39 años"
	di as result "Hoja: indice_estructura_pea"
	di as result "============================================================"

	list, clean noobs

* ============================================================
**# BLOQUE 28: ÍNDICE DE ESTRUCTURA DE LA POBLACIÓN
*            EN EDAD ACTIVA, POR SEXO
* ============================================================
*
* Código:
*     06_003_12.1
*
* Indicador:
*     Índice de estructura de la población en edad activa,
*     por sexo
*
* Sigla:
*     IEPEA_SEXO
*
* Fórmula:
*     iepea_sexo = (pob_40_64 / pob_15_39) * 100
*
* Numerador:
*     población de 40 a 64 años, por sexo
*
* Denominador:
*     población de 15 a 39 años, por sexo
*
* Fuente:
*     _out/Bases DTA/pob_edades_simples_long_2012_20XX.dta
*
* Rango temporal:
*     2012–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Desagregación:
*     Hombre
*     Mujer
*
* Unidad:
*     Personas de 40 a 64 años por cada 100 personas de 15 a 39 años,
*     dentro de cada sexo.
*
* Consideraciones:
*     - El grupo 15-39 se construye como:
*           edades simples 15 a 29
*           + grupos 30_34 y 35_39.
*     - El grupo 40-64 se construye como:
*           grupos 40_44, 45_49, 50_54, 55_59 y 60_64.
*     - No se calcula para 2005–2011 porque la base disponible no
*       contiene estos grupos específicos por sexo.
*     - No se incluye "Total", porque el indicador total ya fue
*       calculado en 06_003_12.
*     - Como es un indicador posterior al primero, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 28A. Abrir base fuente
* ------------------------------------------------------------

	use "$out_dta/pob_edades_simples_long_2012_2024.dta", clear

	keep if inrange(anio, 2012, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

	* Mantener solo hombres y mujeres.
	keep if inlist(sexo, "Hombre", "Mujer")

	* Estandarizar codigo_geo como texto.
	capture confirm numeric variable codigo_geo

	if _rc == 0 {
		gen str6 codigo_geo_str = string(codigo_geo, "%02.0f")
		drop codigo_geo
		rename codigo_geo_str codigo_geo
	}

	replace codigo_geo = strtrim(codigo_geo)

	* Asegurar que grupo_edad sea texto.
	capture confirm string variable grupo_edad

	if _rc != 0 {
		tostring grupo_edad, replace
	}

	replace grupo_edad = strtrim(grupo_edad)

* ------------------------------------------------------------
* 28B. Construir numerador y denominador por sexo
* ------------------------------------------------------------

	gen edad_simple = real(grupo_edad)

	* Denominador:
	*     población de 15 a 39 años.
	gen double pob_15_39 = 0

	replace pob_15_39 = poblacion if inrange(edad_simple, 15, 29)

	replace pob_15_39 = poblacion if regexm(grupo_edad, "^(30_34|35_39)$")


	* Numerador:
	*     población de 40 a 64 años.
	gen double pob_40_64 = 0

	replace pob_40_64 = poblacion if regexm(grupo_edad, "^(40_44|45_49|50_54|55_59|60_64)$")

* ------------------------------------------------------------
* 28C. Colapsar a año + territorio + sexo
* ------------------------------------------------------------

	collapse (sum) pob_15_39 pob_40_64, by(anio codigo_geo nombre_geo nivel_geo sexo)

	isid anio codigo_geo sexo

* ------------------------------------------------------------
* 28D. Controles previos
* ------------------------------------------------------------

	count if missing(pob_15_39) | pob_15_39 <= 0

	if r(N) > 0 {
		di as error "ERROR: Hay población de 15 a 39 años faltante o menor/igual a cero por sexo."

		preserve
			keep if missing(pob_15_39) | pob_15_39 <= 0
			sort anio codigo_geo sexo
			export excel using "$out_aud/error_denominador_iepea_sexo.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	count if missing(pob_40_64) | pob_40_64 < 0

	if r(N) > 0 {
		di as error "ERROR: Hay población de 40 a 64 años faltante o negativa por sexo."

		preserve
			keep if missing(pob_40_64) | pob_40_64 < 0
			sort anio codigo_geo sexo
			export excel using "$out_aud/error_numerador_iepea_sexo.xlsx", firstrow(variables) replace
		restore

		error 459
	}

* ------------------------------------------------------------
* 28E. Calcular indicador
* ------------------------------------------------------------

	gen iepea_sexo = .

	replace iepea_sexo = (pob_40_64 / pob_15_39) * 100 if pob_15_39 > 0 & !missing(pob_40_64, pob_15_39)

	replace iepea_sexo = round(iepea_sexo, .01)

* ------------------------------------------------------------
* 28F. Preparar tabla final
* ------------------------------------------------------------

	gen sexo_col = ""

	replace sexo_col = "HOMBRE" if sexo == "Hombre"
	replace sexo_col = "MUJER"  if sexo == "Mujer"

	gen geo_sexo = codigo_geo + "_" + sexo_col

	keep anio geo_sexo iepea_sexo

	reshape wide iepea_sexo, i(anio) j(geo_sexo) string

	rename iepea_sexo00_HOMBRE BOLIVIA_HOMBRE
	rename iepea_sexo00_MUJER  BOLIVIA_MUJER

	rename iepea_sexo01_HOMBRE CHUQUISACA_HOMBRE
	rename iepea_sexo01_MUJER  CHUQUISACA_MUJER

	rename iepea_sexo02_HOMBRE LA_PAZ_HOMBRE
	rename iepea_sexo02_MUJER  LA_PAZ_MUJER

	rename iepea_sexo03_HOMBRE COCHABAMBA_HOMBRE
	rename iepea_sexo03_MUJER  COCHABAMBA_MUJER

	rename iepea_sexo04_HOMBRE ORURO_HOMBRE
	rename iepea_sexo04_MUJER  ORURO_MUJER

	rename iepea_sexo05_HOMBRE POTOSI_HOMBRE
	rename iepea_sexo05_MUJER  POTOSI_MUJER

	rename iepea_sexo06_HOMBRE TARIJA_HOMBRE
	rename iepea_sexo06_MUJER  TARIJA_MUJER

	rename iepea_sexo07_HOMBRE SANTA_CRUZ_HOMBRE
	rename iepea_sexo07_MUJER  SANTA_CRUZ_MUJER

	rename iepea_sexo08_HOMBRE BENI_HOMBRE
	rename iepea_sexo08_MUJER  BENI_MUJER

	rename iepea_sexo09_HOMBRE PANDO_HOMBRE
	rename iepea_sexo09_MUJER  PANDO_MUJER

	order anio, first

	sort anio

	format *_HOMBRE *_MUJER %9.2f

* ------------------------------------------------------------
* 28G. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("iepea_sexo") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 28H. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_003_12.1 | IEPEA por sexo"
	di as result "Unidad: personas de 40–64 por cada 100 personas de 15–39 años"
	di as result "Hoja: iepea_sexo"
	di as result "============================================================"

	list, clean noobs
	
* ============================================================
**# BLOQUE 29: ÍNDICE DE MATERNIDAD
* ============================================================
*
* Código:
*     06_003_13
*
* Indicador:
*     Índice de maternidad
*
* Fórmula:
*     indice_maternidad = (pob_0_4 / mef) * 100
*
* Numerador:
*     población de 0 a 4 años
*
* Denominador:
*     mujeres en edad fértil, 15 a 49 años
*
* Fuentes:
*     Numerador 2005–2011:
*         _out/Bases DTA/pob_edad_salud_2005_20XX.dta
*
*     Numerador 2012–20XX:
*         _out/Bases DTA/pob_edades_simples_long_2012_20XX.dta
*
*     Denominador 2005–20XX:
*         _out/Bases DTA/pob_grupos_especiales_2005_20XX.dta
*
* Rango temporal:
*     2005–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Unidad:
*     Niños y niñas de 0 a 4 años por cada 100 mujeres de 15 a 49 años.
*
* Consideraciones:
*     - El numerador corresponde a la población total de 0 a 4 años.
*     - El denominador corresponde a mujeres en edad fértil: mef.
*     - Para 2005–2011 se toma grupo_0_4 directamente.
*     - Para 2012–20XX se suma grupo_edad 0, 1, 2, 3 y 4.
*     - Como es un indicador posterior al primero, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 29A. Construir numerador: población 0-4 años, 2005–2011
* ------------------------------------------------------------

	tempfile pob_0_4_2005_2011

	use "$out_dta/pob_edad_salud_2005_2024.dta", clear

	keep if inrange(anio, 2005, 2011)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

	* Estandarizar codigo_geo como texto.
	capture confirm numeric variable codigo_geo

	if _rc == 0 {
		gen str6 codigo_geo_str = string(codigo_geo, "%02.0f")
		drop codigo_geo
		rename codigo_geo_str codigo_geo
	}

	replace codigo_geo = strtrim(codigo_geo)

	gen double pob_0_4 = grupo_0_4

	keep anio codigo_geo nombre_geo nivel_geo pob_0_4

	isid anio codigo_geo

	save `pob_0_4_2005_2011', replace

* ------------------------------------------------------------
* 29B. Construir numerador: población 0-4 años, 2012–20XX
* ------------------------------------------------------------

	tempfile pob_0_4_2012_2024

	use "$out_dta/pob_edades_simples_long_2012_2024.dta", clear

	keep if inrange(anio, 2012, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

	* Usar población total por sexo.
	keep if sexo == "Total"

	* Estandarizar codigo_geo como texto.
	capture confirm numeric variable codigo_geo

	if _rc == 0 {
		gen str6 codigo_geo_str = string(codigo_geo, "%02.0f")
		drop codigo_geo
		rename codigo_geo_str codigo_geo
	}

	replace codigo_geo = strtrim(codigo_geo)

	* Asegurar que grupo_edad sea texto.
	capture confirm string variable grupo_edad

	if _rc != 0 {
		tostring grupo_edad, replace
	}

	replace grupo_edad = strtrim(grupo_edad)

	gen edad_simple = real(grupo_edad)

	gen double pob_0_4 = 0
	replace pob_0_4 = poblacion if inrange(edad_simple, 0, 4)

	collapse (sum) pob_0_4, by(anio codigo_geo nombre_geo nivel_geo)

	isid anio codigo_geo

	save `pob_0_4_2012_2024', replace

* ------------------------------------------------------------
* 29C. Unir numerador 2005–20XX
* ------------------------------------------------------------

	tempfile pob_0_4_2005_2024

	use `pob_0_4_2005_2011', clear

	append using `pob_0_4_2012_2024'

	sort anio codigo_geo

	isid anio codigo_geo

	save `pob_0_4_2005_2024', replace

* ------------------------------------------------------------
* 29D. Preparar denominador: mujeres de 15 a 49 años
* ------------------------------------------------------------

	tempfile mef_2005_2024

	use "$out_dta/pob_grupos_especiales_2005_2024.dta", clear

	keep if inrange(anio, 2005, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

	* Estandarizar codigo_geo como texto.
	capture confirm numeric variable codigo_geo

	if _rc == 0 {
		gen str6 codigo_geo_str = string(codigo_geo, "%02.0f")
		drop codigo_geo
		rename codigo_geo_str codigo_geo
	}

	replace codigo_geo = strtrim(codigo_geo)

	* Control de duplicados.
	duplicates tag anio codigo_geo, gen(dup_key)

	count if dup_key > 0

	if r(N) > 0 {
		di as error "ERROR: Hay duplicados en MEF por anio + codigo_geo."

		preserve
			keep if dup_key > 0
			sort anio codigo_geo
			export excel using "$out_aud/auditoria_duplicados_mef_indice_maternidad.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	drop dup_key

	isid anio codigo_geo

	keep anio codigo_geo mef

	save `mef_2005_2024', replace

* ------------------------------------------------------------
* 29E. Unir numerador y denominador
* ------------------------------------------------------------

	use `pob_0_4_2005_2024', clear

	merge 1:1 anio codigo_geo using `mef_2005_2024'

	tab _merge

	count if _merge != 3

	if r(N) > 0 {
		di as error "ERROR: Hay observaciones sin emparejar entre población 0-4 y MEF."

		preserve
			keep if _merge != 3
			sort anio codigo_geo
			export excel using "$out_aud/error_merge_indice_maternidad.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	drop _merge

	isid anio codigo_geo

* ------------------------------------------------------------
* 29F. Controles previos
* ------------------------------------------------------------

	count if missing(pob_0_4) | pob_0_4 < 0

	if r(N) > 0 {
		di as error "ERROR: Hay población 0-4 faltante o negativa."

		preserve
			keep if missing(pob_0_4) | pob_0_4 < 0
			sort anio codigo_geo
			export excel using "$out_aud/error_numerador_indice_maternidad.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	count if missing(mef) | mef <= 0

	if r(N) > 0 {
		di as error "ERROR: Hay MEF faltante o menor/igual a cero."

		preserve
			keep if missing(mef) | mef <= 0
			sort anio codigo_geo
			export excel using "$out_aud/error_denominador_indice_maternidad.xlsx", firstrow(variables) replace
		restore

		error 459
	}

* ------------------------------------------------------------
* 29G. Calcular indicador
* ------------------------------------------------------------
* Índice de maternidad:
*
*     IM = población 0-4 / mujeres 15-49 * 100
*
* Interpretación:
*     Niños y niñas de 0 a 4 años por cada 100 mujeres
*     en edad fértil.
* ------------------------------------------------------------

	gen indice_maternidad = .

	replace indice_maternidad = (pob_0_4 / mef) * 100 if mef > 0 & !missing(pob_0_4, mef)

	replace indice_maternidad = round(indice_maternidad, .01)

* ------------------------------------------------------------
* 29H. Preparar tabla final
* ------------------------------------------------------------

	keep anio codigo_geo indice_maternidad

	reshape wide indice_maternidad, i(anio) j(codigo_geo) string

	rename indice_maternidad00 BOLIVIA
	rename indice_maternidad01 CHUQUISACA
	rename indice_maternidad02 LA_PAZ
	rename indice_maternidad03 COCHABAMBA
	rename indice_maternidad04 ORURO
	rename indice_maternidad05 POTOSI
	rename indice_maternidad06 TARIJA
	rename indice_maternidad07 SANTA_CRUZ
	rename indice_maternidad08 BENI
	rename indice_maternidad09 PANDO

	order anio BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO

	sort anio

	format BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO %9.2f

* ------------------------------------------------------------
* 29I. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("indice_maternidad") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 29J. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_003_13 | Índice de maternidad"
	di as result "Unidad: población 0-4 por cada 100 mujeres de 15-49 años"
	di as result "Hoja: indice_maternidad"
	di as result "============================================================"

	list, clean noobs

* ============================================================
* BLOQUE 30: ÍNDICE DE FRIZ
* ============================================================
* Código:
*     06_003_14
*
* Indicador:
*     Índice de Friz
*
* Fórmula:
*     indice_friz = (pob_0_19 / pob_30_49) * 100
*
* Numerador:
*     población de 0 a 19 años
*
* Denominador:
*     población de 30 a 49 años
*
* Fuente:
*     _out/Bases DTA/pob_edades_simples_long_2012_20XX.dta
*
* Rango temporal:
*     2012–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Unidad:
*     Personas de 0 a 19 años por cada 100 personas de 30 a 49 años.
*
* Consideraciones:
*     - Se calcula desde 2012 porque desde ese año la base tiene
*       edades simples y grupos quinquenales suficientes.
*     - El grupo 0-19 se construye sumando edades simples:
*           0 + 1 + ... + 19
*     - El grupo 30-49 se construye sumando:
*           30_34 + 35_39 + 40_44 + 45_49
*     - No se calcula para 2005–2011 porque la base disponible no
*       contiene estos grupos específicos.
*     - Como es un indicador posterior al primero, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 30A. Abrir base fuente
* ------------------------------------------------------------

	use "$out_dta/pob_edades_simples_long_2012_2024.dta", clear

	keep if inrange(anio, 2012, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

	* Usar población total por sexo.
	keep if sexo == "Total"

	* Estandarizar codigo_geo como texto.
	capture confirm numeric variable codigo_geo

	if _rc == 0 {
		gen str6 codigo_geo_str = string(codigo_geo, "%02.0f")
		drop codigo_geo
		rename codigo_geo_str codigo_geo
	}

	replace codigo_geo = strtrim(codigo_geo)

	* Asegurar que grupo_edad sea texto.
	capture confirm string variable grupo_edad

	if _rc != 0 {
		tostring grupo_edad, replace
	}

	replace grupo_edad = strtrim(grupo_edad)

* ------------------------------------------------------------
* 30B. Construir numerador y denominador
* ------------------------------------------------------------

	gen edad_simple = real(grupo_edad)

	* Numerador:
	*     población de 0 a 19 años.
	gen double pob_0_19 = 0

	replace pob_0_19 = poblacion if inrange(edad_simple, 0, 19)


	* Denominador:
	*     población de 30 a 49 años.
	gen double pob_30_49 = 0

	replace pob_30_49 = poblacion if regexm(grupo_edad, "^(30_34|35_39|40_44|45_49)$")

* ------------------------------------------------------------
* 30C. Colapsar a año + territorio
* ------------------------------------------------------------

	collapse (sum) pob_0_19 pob_30_49, by(anio codigo_geo nombre_geo nivel_geo)

	isid anio codigo_geo

* ------------------------------------------------------------
* 30D. Controles previos
* ------------------------------------------------------------

	count if missing(pob_0_19) | pob_0_19 < 0

	if r(N) > 0 {
		di as error "ERROR: Hay población de 0 a 19 años faltante o negativa."

		preserve
			keep if missing(pob_0_19) | pob_0_19 < 0
			sort anio codigo_geo
			export excel using "$out_aud/error_numerador_indice_friz.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	count if missing(pob_30_49) | pob_30_49 <= 0

	if r(N) > 0 {
		di as error "ERROR: Hay población de 30 a 49 años faltante o menor/igual a cero."

		preserve
			keep if missing(pob_30_49) | pob_30_49 <= 0
			sort anio codigo_geo
			export excel using "$out_aud/error_denominador_indice_friz.xlsx", firstrow(variables) replace
		restore

		error 459
	}

* ------------------------------------------------------------
* 30E. Calcular indicador
* ------------------------------------------------------------
* Índice de Friz:
*
*     IF = población 0-19 / población 30-49 * 100
*
* Interpretación:
*     Personas de 0 a 19 años por cada 100 personas de 30 a 49 años.
* ------------------------------------------------------------

	gen indice_friz = .

	replace indice_friz = (pob_0_19 / pob_30_49) * 100 if pob_30_49 > 0 & !missing(pob_0_19, pob_30_49)

	replace indice_friz = round(indice_friz, .01)

* ------------------------------------------------------------
* 30F. Preparar tabla final
* ------------------------------------------------------------

	keep anio codigo_geo indice_friz

	reshape wide indice_friz, i(anio) j(codigo_geo) string

	rename indice_friz00 BOLIVIA
	rename indice_friz01 CHUQUISACA
	rename indice_friz02 LA_PAZ
	rename indice_friz03 COCHABAMBA
	rename indice_friz04 ORURO
	rename indice_friz05 POTOSI
	rename indice_friz06 TARIJA
	rename indice_friz07 SANTA_CRUZ
	rename indice_friz08 BENI
	rename indice_friz09 PANDO

	order anio BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO

	sort anio

	format BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO %9.2f

* ------------------------------------------------------------
* 30G. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("indice_friz") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 30H. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_003_14 | Índice de Friz"
	di as result "Unidad: personas de 0–19 por cada 100 personas de 30–49 años"
	di as result "Hoja: indice_friz"
	di as result "============================================================"

	list, clean noobs

* ============================================================
**# BLOQUE 31: ÍNDICE DE SUNDBÄRG
* ============================================================
* Código:
*     06_003_15
*
* Indicador:
*     Índice de Sundbärg
*
* Fórmulas:
*     is_jovenes = (pob_0_14 / pob_15_49) * 100
*     is_mayores = (pob_50_mas / pob_15_49) * 100
*
* Numeradores:
*     IS jóvenes:
*         población de 0 a 14 años
*
*     IS mayores:
*         población de 50 años y más
*
* Denominador:
*     población de 15 a 49 años
*
* Fuente:
*     _out/Bases DTA/pob_edades_simples_long_2012_20XX.dta
*
* Rango temporal:
*     2012–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Unidad:
*     Personas por cada 100 personas de 15 a 49 años.
*
* Consideraciones:
*     - Se calcula desde 2012 porque desde ese año la base tiene
*       edades simples y grupos quinquenales suficientes.
*     - El grupo 0-14 se construye sumando edades simples 0 a 14.
*     - El grupo 15-49 se construye como:
*           edades simples 15 a 29
*           + grupos 30_34, 35_39, 40_44 y 45_49.
*     - El grupo 50+ se construye como:
*           50_54, 55_59, 60_64, 65_69, 70_74, 75_79 y 80_mas.
*     - No se calcula para 2005–2011 porque la base disponible no
*       contiene estos grupos específicos.
*     - Se exportan ambos componentes: IS jóvenes e IS mayores.
* ============================================================
capture restore
* ------------------------------------------------------------
* 31A. Abrir base fuente
* ------------------------------------------------------------

	use "$out_dta/pob_edades_simples_long_2012_2024.dta", clear

	keep if inrange(anio, 2012, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

	* Usar población total por sexo.
	keep if sexo == "Total"

	* Estandarizar codigo_geo como texto.
	capture confirm numeric variable codigo_geo

	if _rc == 0 {
		gen str6 codigo_geo_str = string(codigo_geo, "%02.0f")
		drop codigo_geo
		rename codigo_geo_str codigo_geo
	}

	replace codigo_geo = strtrim(codigo_geo)

	* Asegurar que grupo_edad sea texto.
	capture confirm string variable grupo_edad

	if _rc != 0 {
		tostring grupo_edad, replace
	}

	replace grupo_edad = strtrim(grupo_edad)

* ------------------------------------------------------------
* 31B. Construir grupos etarios del índice
* ------------------------------------------------------------

	gen edad_simple = real(grupo_edad)

	* Numerador IS jóvenes:
	*     población de 0 a 14 años.
	gen double pob_0_14 = 0
	replace pob_0_14 = poblacion if inrange(edad_simple, 0, 14)

	* Denominador común:
	*     población de 15 a 49 años.
	gen double pob_15_49 = 0
	replace pob_15_49 = poblacion if inrange(edad_simple, 15, 29)
	replace pob_15_49 = poblacion if regexm(grupo_edad, "^(30_34|35_39|40_44|45_49)$")

	* Numerador IS mayores:
	*     población de 50 años y más.
	gen double pob_50_mas = 0
	replace pob_50_mas = poblacion if regexm(grupo_edad, "^(50_54|55_59|60_64|65_69|70_74|75_79|80_mas)$")

* ------------------------------------------------------------
* 31C. Colapsar a año + territorio
* ------------------------------------------------------------

	collapse (sum) pob_0_14 pob_15_49 pob_50_mas, by(anio codigo_geo nombre_geo nivel_geo)

	isid anio codigo_geo

* ------------------------------------------------------------
* 31D. Controles previos
* ------------------------------------------------------------

	count if missing(pob_0_14) | pob_0_14 < 0

	if r(N) > 0 {
		di as error "ERROR: Hay población 0-14 faltante o negativa."

		preserve
			keep if missing(pob_0_14) | pob_0_14 < 0
			sort anio codigo_geo
			export excel using "$out_aud/error_numerador_sundbarg_jovenes.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	count if missing(pob_15_49) | pob_15_49 <= 0

	if r(N) > 0 {
		di as error "ERROR: Hay población 15-49 faltante o menor/igual a cero."

		preserve
			keep if missing(pob_15_49) | pob_15_49 <= 0
			sort anio codigo_geo
			export excel using "$out_aud/error_denominador_sundbarg.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	count if missing(pob_50_mas) | pob_50_mas < 0

	if r(N) > 0 {
		di as error "ERROR: Hay población 50+ faltante o negativa."

		preserve
			keep if missing(pob_50_mas) | pob_50_mas < 0
			sort anio codigo_geo
			export excel using "$out_aud/error_numerador_sundbarg_mayores.xlsx", firstrow(variables) replace
		restore

		error 459
	}

* ------------------------------------------------------------
* 31E. Calcular indicadores
* ------------------------------------------------------------

	gen is_jovenes = .
	replace is_jovenes = (pob_0_14 / pob_15_49) * 100 if pob_15_49 > 0 & !missing(pob_0_14, pob_15_49)
	replace is_jovenes = round(is_jovenes, .01)

	gen is_mayores = .
	replace is_mayores = (pob_50_mas / pob_15_49) * 100 if pob_15_49 > 0 & !missing(pob_50_mas, pob_15_49)
	replace is_mayores = round(is_mayores, .01)

* ------------------------------------------------------------
* 31F. Preparar tabla final
* ------------------------------------------------------------

	keep anio codigo_geo is_jovenes is_mayores

	reshape wide is_jovenes is_mayores, i(anio) j(codigo_geo) string

	rename is_jovenes00 BOLIVIA_IS_JOVENES
	rename is_mayores00 BOLIVIA_IS_MAYORES

	rename is_jovenes01 CHUQUISACA_IS_JOVENES
	rename is_mayores01 CHUQUISACA_IS_MAYORES

	rename is_jovenes02 LA_PAZ_IS_JOVENES
	rename is_mayores02 LA_PAZ_IS_MAYORES

	rename is_jovenes03 COCHABAMBA_IS_JOVENES
	rename is_mayores03 COCHABAMBA_IS_MAYORES

	rename is_jovenes04 ORURO_IS_JOVENES
	rename is_mayores04 ORURO_IS_MAYORES

	rename is_jovenes05 POTOSI_IS_JOVENES
	rename is_mayores05 POTOSI_IS_MAYORES

	rename is_jovenes06 TARIJA_IS_JOVENES
	rename is_mayores06 TARIJA_IS_MAYORES

	rename is_jovenes07 SANTA_CRUZ_IS_JOVENES
	rename is_mayores07 SANTA_CRUZ_IS_MAYORES

	rename is_jovenes08 BENI_IS_JOVENES
	rename is_mayores08 BENI_IS_MAYORES

	rename is_jovenes09 PANDO_IS_JOVENES
	rename is_mayores09 PANDO_IS_MAYORES

	order anio, first

	sort anio

	format *_IS_JOVENES *_IS_MAYORES %9.2f

* ------------------------------------------------------------
* 31G. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("indice_sundbarg") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 31H. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_003_15 | Índice de Sundbärg"
	di as result "Unidad: personas por cada 100 personas de 15–49 años"
	di as result "Hoja: indice_sundbarg"
	di as result "============================================================"

	list, clean noobs

* ============================================================
**# BLOQUE 32 CORREGIDO: ÍNDICE DE BURGDÖFER
* ============================================================
* Código:
*     06_003_16
*
* Indicador:
*     Índice de Burgdöfer
*
* Fórmula corregida:
*     indice_burgdofer = (pob_5_14 / pob_45_64) * 100
*
* Numerador:
*     población de 5 a 14 años
*
* Denominador:
*     población de 45 a 64 años
*
* Fuente:
*     _out/Bases DTA/pob_edades_simples_long_2012_20XX.dta
*
* Rango temporal:
*     2012–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Unidad:
*     Personas de 5 a 14 años por cada 100 personas de 45 a 64 años.
*
* Consideraciones:
*     - Se calcula desde 2012 porque desde ese año la base tiene
*       edades simples y grupos quinquenales suficientes.
*     - El grupo 5-14 se construye sumando edades simples:
*           5 + 6 + ... + 14.
*     - El grupo 45-64 se construye sumando:
*           45_49 + 50_54 + 55_59 + 60_64.
*     - No se calcula para 2005–2011 porque la base disponible no
*       contiene el grupo 45-64 separado.
*     - Como es un indicador posterior al primero, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 32A. Abrir base fuente
* ------------------------------------------------------------

	use "$out_dta/pob_edades_simples_long_2012_2024.dta", clear

	keep if inrange(anio, 2012, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

	* Usar población total por sexo.
	keep if sexo == "Total"

	* Estandarizar codigo_geo como texto.
	capture confirm numeric variable codigo_geo

	if _rc == 0 {
		gen str6 codigo_geo_str = string(codigo_geo, "%02.0f")
		drop codigo_geo
		rename codigo_geo_str codigo_geo
	}

	replace codigo_geo = strtrim(codigo_geo)

	* Asegurar que grupo_edad sea texto.
	capture confirm string variable grupo_edad

	if _rc != 0 {
		tostring grupo_edad, replace
	}

	replace grupo_edad = strtrim(grupo_edad)

* ------------------------------------------------------------
* 32B. Construir numerador y denominador
* ------------------------------------------------------------

	gen edad_simple = real(grupo_edad)

	* Numerador:
	*     población de 5 a 14 años.
	gen double pob_5_14 = 0

	replace pob_5_14 = poblacion if inrange(edad_simple, 5, 14)


	* Denominador:
	*     población de 45 a 64 años.
	gen double pob_45_64 = 0

	replace pob_45_64 = poblacion if regexm(grupo_edad, "^(45_49|50_54|55_59|60_64)$")

* ------------------------------------------------------------
* 32C. Colapsar a año + territorio
* ------------------------------------------------------------

	collapse (sum) pob_5_14 pob_45_64, by(anio codigo_geo nombre_geo nivel_geo)

	isid anio codigo_geo

* ------------------------------------------------------------
* 32D. Controles previos
* ------------------------------------------------------------

	count if missing(pob_5_14) | pob_5_14 < 0

	if r(N) > 0 {
		di as error "ERROR: Hay población de 5 a 14 años faltante o negativa."

		preserve
			keep if missing(pob_5_14) | pob_5_14 < 0
			sort anio codigo_geo
			export excel using "$out_aud/error_numerador_indice_burgdofer.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	count if missing(pob_45_64) | pob_45_64 <= 0

	if r(N) > 0 {
		di as error "ERROR: Hay población de 45 a 64 años faltante o menor/igual a cero."

		preserve
			keep if missing(pob_45_64) | pob_45_64 <= 0
			sort anio codigo_geo
			export excel using "$out_aud/error_denominador_indice_burgdofer.xlsx", firstrow(variables) replace
		restore

		error 459
	}

* ------------------------------------------------------------
* 32E. Calcular indicador
* ------------------------------------------------------------
* Índice de Burgdöfer:
*
*     IB = población 5-14 / población 45-64 * 100
*
* Interpretación:
*     Personas de 5 a 14 años por cada 100 personas de 45 a 64 años.
* ------------------------------------------------------------

	gen indice_burgdofer = .

	replace indice_burgdofer = (pob_5_14 / pob_45_64) * 100 if pob_45_64 > 0 & !missing(pob_5_14, pob_45_64)

	replace indice_burgdofer = round(indice_burgdofer, .01)

* ------------------------------------------------------------
* 32F. Preparar tabla final
* ------------------------------------------------------------

	keep anio codigo_geo indice_burgdofer

	reshape wide indice_burgdofer, i(anio) j(codigo_geo) string

	rename indice_burgdofer00 BOLIVIA
	rename indice_burgdofer01 CHUQUISACA
	rename indice_burgdofer02 LA_PAZ
	rename indice_burgdofer03 COCHABAMBA
	rename indice_burgdofer04 ORURO
	rename indice_burgdofer05 POTOSI
	rename indice_burgdofer06 TARIJA
	rename indice_burgdofer07 SANTA_CRUZ
	rename indice_burgdofer08 BENI
	rename indice_burgdofer09 PANDO

	order anio BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO

	sort anio

	format BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO %9.2f

* ------------------------------------------------------------
* 32G. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("indice_burgdofer") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 32H. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_003_16 | Índice de Burgdöfer"
	di as result "Unidad: personas de 5–14 por cada 100 personas de 45–64 años"
	di as result "Hoja: indice_burgdofer"
	di as result "============================================================"

	list, clean noobs

* ============================================================
**# BLOQUE 33: ÍNDICE GENERACIONAL DE ANCIANOS
* ============================================================
* Código:
*     06_003_17
*
* Indicador:
*     Índice generacional de ancianos
*
* Fórmula:
*     indice_generacional_ancianos = (pob_35_64 / pob_65_mas) * 100
*
* Numerador:
*     población de 35 a 64 años
*
* Denominador:
*     población de 65 años y más
*
* Fuente:
*     _out/Bases DTA/pob_edades_simples_long_2012_20XX.dta
*
* Rango temporal:
*     2012–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Unidad:
*     Personas de 35 a 64 años por cada 100 personas de 65 años y más.
*
* Consideraciones:
*     - Se calcula desde 2012 porque desde ese año la base tiene
*       grupos quinquenales suficientes.
*     - El grupo 35-64 se construye como:
*           35_39, 40_44, 45_49, 50_54, 55_59 y 60_64.
*     - El grupo 65+ se construye como:
*           65_69, 70_74, 75_79 y 80_mas.
*     - No se calcula para 2005–2011 porque la base disponible no
*       contiene el grupo 35-64 separado.
*     - Como es un indicador posterior al primero, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 33A. Abrir base fuente
* ------------------------------------------------------------

	use "$out_dta/pob_edades_simples_long_2012_2024.dta", clear

	keep if inrange(anio, 2012, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

	* Usar población total por sexo.
	keep if sexo == "Total"

	* Estandarizar codigo_geo como texto.
	capture confirm numeric variable codigo_geo

	if _rc == 0 {
		gen str6 codigo_geo_str = string(codigo_geo, "%02.0f")
		drop codigo_geo
		rename codigo_geo_str codigo_geo
	}

	replace codigo_geo = strtrim(codigo_geo)

	* Asegurar que grupo_edad sea texto.
	capture confirm string variable grupo_edad

	if _rc != 0 {
		tostring grupo_edad, replace
	}

	replace grupo_edad = strtrim(grupo_edad)

* ------------------------------------------------------------
* 33B. Construir numerador y denominador
* ------------------------------------------------------------

	* Numerador:
	*     población de 35 a 64 años.
	gen double pob_35_64 = 0

	replace pob_35_64 = poblacion if regexm(grupo_edad, "^(35_39|40_44|45_49|50_54|55_59|60_64)$")


	* Denominador:
	*     población de 65 años y más.
	gen double pob_65_mas = 0

	replace pob_65_mas = poblacion if regexm(grupo_edad, "^(65_69|70_74|75_79|80_mas)$")

* ------------------------------------------------------------
* 33C. Colapsar a año + territorio
* ------------------------------------------------------------

	collapse (sum) pob_35_64 pob_65_mas, by(anio codigo_geo nombre_geo nivel_geo)

	isid anio codigo_geo

* ------------------------------------------------------------
* 33D. Controles previos
* ------------------------------------------------------------

	count if missing(pob_35_64) | pob_35_64 < 0

	if r(N) > 0 {
		di as error "ERROR: Hay población de 35 a 64 años faltante o negativa."

		preserve
			keep if missing(pob_35_64) | pob_35_64 < 0
			sort anio codigo_geo
			export excel using "$out_aud/error_numerador_indice_generacional_ancianos.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	count if missing(pob_65_mas) | pob_65_mas <= 0

	if r(N) > 0 {
		di as error "ERROR: Hay población de 65 años y más faltante o menor/igual a cero."

		preserve
			keep if missing(pob_65_mas) | pob_65_mas <= 0
			sort anio codigo_geo
			export excel using "$out_aud/error_denominador_indice_generacional_ancianos.xlsx", firstrow(variables) replace
		restore

		error 459
	}

* ------------------------------------------------------------
* 33E. Calcular indicador
* ------------------------------------------------------------
* Índice generacional de ancianos:
*
*     IGA = población 35-64 / población 65+ * 100
*
* Interpretación:
*     Personas de 35 a 64 años por cada 100 personas de 65 años y más.
* ------------------------------------------------------------

	gen indice_generacional_ancianos = .

	replace indice_generacional_ancianos = (pob_35_64 / pob_65_mas) * 100 if pob_65_mas > 0 & !missing(pob_35_64, pob_65_mas)

	replace indice_generacional_ancianos = round(indice_generacional_ancianos, .01)

* ------------------------------------------------------------
* 33F. Preparar tabla final
* ------------------------------------------------------------

	keep anio codigo_geo indice_generacional_ancianos

	reshape wide indice_generacional_ancianos, i(anio) j(codigo_geo) string

	rename indice_generacional_ancianos00 BOLIVIA
	rename indice_generacional_ancianos01 CHUQUISACA
	rename indice_generacional_ancianos02 LA_PAZ
	rename indice_generacional_ancianos03 COCHABAMBA
	rename indice_generacional_ancianos04 ORURO
	rename indice_generacional_ancianos05 POTOSI
	rename indice_generacional_ancianos06 TARIJA
	rename indice_generacional_ancianos07 SANTA_CRUZ
	rename indice_generacional_ancianos08 BENI
	rename indice_generacional_ancianos09 PANDO

	order anio BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO

	sort anio

	format BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO %9.2f

* ------------------------------------------------------------
* 33G. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("indice_generacional_ancianos") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 33H. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_003_17 | Índice generacional de ancianos"
	di as result "Unidad: personas de 35–64 por cada 100 personas de 65 años y más"
	di as result "Hoja: indice_generacional_ancianos"
	di as result "============================================================"

	list, clean noobs

* ============================================================
**# BLOQUE 34: ÍNDICE DE VEJEZ
* ============================================================
* Código:
*     06_003_18
*
* Indicador:
*     Índice de vejez
*
* Fórmula:
*     indice_vejez = (pob_65_mas / pob_total) * 100
*
* Numerador:
*     población de 65 años y más
*
* Denominador:
*     población total
*
* Fuentes:
*     2005–2011:
*         _out/Bases DTA/pob_edad_salud_2005_20XX.dta
*
*     2012–20XX:
*         _out/Bases DTA/pob_edades_simples_long_2012_20XX.dta
*
* Rango temporal:
*     2005–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Unidad:
*     Personas de 65 años y más como porcentaje de la población total.
*
* Consideraciones:
*     - Para 2005–2011:
*           col_k = población de 65 años y más
*           col_l = población total / Total general
*     - Para 2012–20XX:
*           65+ se construye con 65_69, 70_74, 75_79 y 80_mas.
*           La población total se toma de grupo_edad == "total".
*     - Como es un indicador posterior al primero, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 34A. Construir indicador base 2005–2011
* ------------------------------------------------------------

	tempfile vejez_2005_2011

	use "$out_dta/pob_edad_salud_2005_2024.dta", clear

	keep if inrange(anio, 2005, 2011)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

	* Estandarizar codigo_geo como texto.
	capture confirm numeric variable codigo_geo

	if _rc == 0 {
		gen str6 codigo_geo_str = string(codigo_geo, "%02.0f")
		drop codigo_geo
		rename codigo_geo_str codigo_geo
	}

	replace codigo_geo = strtrim(codigo_geo)

	gen double pob_65_mas = col_k
	gen double pob_total  = col_l

	keep anio codigo_geo nombre_geo nivel_geo pob_65_mas pob_total

	isid anio codigo_geo

	save `vejez_2005_2011', replace

* ------------------------------------------------------------
* 34B. Construir indicador base 2012–20XX
* ------------------------------------------------------------

	tempfile vejez_2012_2024

	use "$out_dta/pob_edades_simples_long_2012_2024.dta", clear

	keep if inrange(anio, 2012, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

	* Usar población total por sexo.
	keep if sexo == "Total"

	* Estandarizar codigo_geo como texto.
	capture confirm numeric variable codigo_geo

	if _rc == 0 {
		gen str6 codigo_geo_str = string(codigo_geo, "%02.0f")
		drop codigo_geo
		rename codigo_geo_str codigo_geo
	}

	replace codigo_geo = strtrim(codigo_geo)

	* Asegurar que grupo_edad sea texto.
	capture confirm string variable grupo_edad

	if _rc != 0 {
		tostring grupo_edad, replace
	}

	replace grupo_edad = strtrim(grupo_edad)

	* Numerador:
	*     población de 65 años y más.
	gen double pob_65_mas = 0
	replace pob_65_mas = poblacion if regexm(grupo_edad, "^(65_69|70_74|75_79|80_mas)$")

	* Denominador:
	*     población total.
	gen double pob_total = 0
	replace pob_total = poblacion if grupo_edad == "total"

	collapse (sum) pob_65_mas pob_total, by(anio codigo_geo nombre_geo nivel_geo)

	isid anio codigo_geo

	save `vejez_2012_2024', replace

* ------------------------------------------------------------
* 34C. Unir periodos
* ------------------------------------------------------------

	use `vejez_2005_2011', clear

	append using `vejez_2012_2024'

	sort anio codigo_geo

	isid anio codigo_geo

* ------------------------------------------------------------
* 34D. Controles previos
* ------------------------------------------------------------

	count if missing(pob_65_mas) | pob_65_mas < 0

	if r(N) > 0 {
		di as error "ERROR: Hay población de 65 años y más faltante o negativa."

		preserve
			keep if missing(pob_65_mas) | pob_65_mas < 0
			sort anio codigo_geo
			export excel using "$out_aud/error_numerador_indice_vejez.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	count if missing(pob_total) | pob_total <= 0

	if r(N) > 0 {
		di as error "ERROR: Hay población total faltante o menor/igual a cero."

		preserve
			keep if missing(pob_total) | pob_total <= 0
			sort anio codigo_geo
			export excel using "$out_aud/error_denominador_indice_vejez.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	count if pob_65_mas > pob_total & !missing(pob_65_mas, pob_total)

	if r(N) > 0 {
		di as error "ERROR: Hay casos donde la población 65+ supera la población total."

		preserve
			keep if pob_65_mas > pob_total & !missing(pob_65_mas, pob_total)
			sort anio codigo_geo
			export excel using "$out_aud/error_65mas_mayor_total_indice_vejez.xlsx", firstrow(variables) replace
		restore

		error 459
	}

* ------------------------------------------------------------
* 34E. Calcular indicador
* ------------------------------------------------------------
* Índice de vejez:
*
*     IV = población 65+ / población total * 100
*
* Interpretación:
*     Porcentaje de la población total que tiene 65 años o más.
* ------------------------------------------------------------

	gen indice_vejez = .

	replace indice_vejez = (pob_65_mas / pob_total) * 100 if pob_total > 0 & !missing(pob_65_mas, pob_total)

	replace indice_vejez = round(indice_vejez, .01)

* ------------------------------------------------------------
* 34F. Preparar tabla final
* ------------------------------------------------------------

	keep anio codigo_geo indice_vejez

	reshape wide indice_vejez, i(anio) j(codigo_geo) string

	rename indice_vejez00 BOLIVIA
	rename indice_vejez01 CHUQUISACA
	rename indice_vejez02 LA_PAZ
	rename indice_vejez03 COCHABAMBA
	rename indice_vejez04 ORURO
	rename indice_vejez05 POTOSI
	rename indice_vejez06 TARIJA
	rename indice_vejez07 SANTA_CRUZ
	rename indice_vejez08 BENI
	rename indice_vejez09 PANDO

	order anio BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO

	sort anio

	format BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO %9.2f

* ------------------------------------------------------------
* 34G. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("indice_vejez") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 34H. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_003_18 | Índice de vejez"
	di as result "Unidad: porcentaje de la población total"
	di as result "Hoja: indice_vejez"
	di as result "============================================================"

	list, clean noobs

* ============================================================
**# BLOQUE 35: ÍNDICE DE SOBREENVEJECIMIENTO
* ============================================================
* Código:
*     06_003_19
*
* Indicador:
*     Índice de sobreenvejecimiento
*
* Fórmula:
*     indice_sobreenvejecimiento = (pob_80_mas / pob_65_mas) * 100
*
* Numerador:
*     población de 80 años y más
*
* Denominador:
*     población de 65 años y más
*
* Fuente:
*     _out/Bases DTA/pob_edades_simples_long_2012_20XX.dta
*
* Rango temporal:
*     2012–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Unidad:
*     Personas de 80 años y más como porcentaje de la población
*     de 65 años y más.
*
* Consideraciones:
*     - Se calcula desde 2012 porque desde ese año la base tiene
*       el grupo 80_mas.
*     - El grupo 65+ se construye como:
*           65_69 + 70_74 + 75_79 + 80_mas.
*     - El grupo 80+ se toma directamente de:
*           grupo_edad == "80_mas".
*     - No se calcula para 2005–2011 porque la base disponible no
*       contiene 80+ separado.
*     - Como es un indicador posterior al primero, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 35A. Abrir base fuente
* ------------------------------------------------------------

	use "$out_dta/pob_edades_simples_long_2012_2024.dta", clear

	keep if inrange(anio, 2012, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

	* Usar población total por sexo.
	keep if sexo == "Total"

	* Estandarizar codigo_geo como texto.
	capture confirm numeric variable codigo_geo

	if _rc == 0 {
		gen str6 codigo_geo_str = string(codigo_geo, "%02.0f")
		drop codigo_geo
		rename codigo_geo_str codigo_geo
	}

	replace codigo_geo = strtrim(codigo_geo)

	* Asegurar que grupo_edad sea texto.
	capture confirm string variable grupo_edad

	if _rc != 0 {
		tostring grupo_edad, replace
	}

	replace grupo_edad = strtrim(grupo_edad)

* ------------------------------------------------------------
* 35B. Construir numerador y denominador
* ------------------------------------------------------------

	* Numerador:
	*     población de 80 años y más.
	gen double pob_80_mas = 0

	replace pob_80_mas = poblacion if grupo_edad == "80_mas"


	* Denominador:
	*     población de 65 años y más.
	gen double pob_65_mas = 0

	replace pob_65_mas = poblacion if regexm(grupo_edad, "^(65_69|70_74|75_79|80_mas)$")

* ------------------------------------------------------------
* 35C. Colapsar a año + territorio
* ------------------------------------------------------------

	collapse (sum) pob_80_mas pob_65_mas, by(anio codigo_geo nombre_geo nivel_geo)

	isid anio codigo_geo

* ------------------------------------------------------------
* 35D. Controles previos
* ------------------------------------------------------------

	count if missing(pob_80_mas) | pob_80_mas < 0

	if r(N) > 0 {
		di as error "ERROR: Hay población de 80 años y más faltante o negativa."

		preserve
			keep if missing(pob_80_mas) | pob_80_mas < 0
			sort anio codigo_geo
			export excel using "$out_aud/error_numerador_indice_sobreenvejecimiento.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	count if missing(pob_65_mas) | pob_65_mas <= 0

	if r(N) > 0 {
		di as error "ERROR: Hay población de 65 años y más faltante o menor/igual a cero."

		preserve
			keep if missing(pob_65_mas) | pob_65_mas <= 0
			sort anio codigo_geo
			export excel using "$out_aud/error_denominador_indice_sobreenvejecimiento.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	count if pob_80_mas > pob_65_mas & !missing(pob_80_mas, pob_65_mas)

	if r(N) > 0 {
		di as error "ERROR: Hay casos donde la población 80+ supera la población 65+."

		preserve
			keep if pob_80_mas > pob_65_mas & !missing(pob_80_mas, pob_65_mas)
			sort anio codigo_geo
			export excel using "$out_aud/error_80mas_mayor_65mas_sobreenvejecimiento.xlsx", firstrow(variables) replace
		restore

		error 459
	}

* ------------------------------------------------------------
* 35E. Calcular indicador
* ------------------------------------------------------------
* Índice de sobreenvejecimiento:
*
*     ISO = población 80+ / población 65+ * 100
*
* Interpretación:
*     Porcentaje de la población adulta mayor de 65 años y más
*     que tiene 80 años o más.
* ------------------------------------------------------------

	gen indice_sobreenvejecimiento = .

	replace indice_sobreenvejecimiento = (pob_80_mas / pob_65_mas) * 100 if pob_65_mas > 0 & !missing(pob_80_mas, pob_65_mas)

	replace indice_sobreenvejecimiento = round(indice_sobreenvejecimiento, .01)

* ------------------------------------------------------------
* 35F. Preparar tabla final
* ------------------------------------------------------------

	keep anio codigo_geo indice_sobreenvejecimiento

	reshape wide indice_sobreenvejecimiento, i(anio) j(codigo_geo) string

	rename indice_sobreenvejecimiento00 BOLIVIA
	rename indice_sobreenvejecimiento01 CHUQUISACA
	rename indice_sobreenvejecimiento02 LA_PAZ
	rename indice_sobreenvejecimiento03 COCHABAMBA
	rename indice_sobreenvejecimiento04 ORURO
	rename indice_sobreenvejecimiento05 POTOSI
	rename indice_sobreenvejecimiento06 TARIJA
	rename indice_sobreenvejecimiento07 SANTA_CRUZ
	rename indice_sobreenvejecimiento08 BENI
	rename indice_sobreenvejecimiento09 PANDO

	order anio BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO

	sort anio

	format BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO %9.2f

* ------------------------------------------------------------
* 35G. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("indice_sobreenvejecimiento") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 35H. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_003_19 | Índice de sobreenvejecimiento"
	di as result "Unidad: porcentaje de la población de 65 años y más"
	di as result "Hoja: indice_sobreenvejecimiento"
	di as result "============================================================"

	list, clean noobs

* ============================================================
**# BLOQUE 36: ÍNDICE DE JUVENTUD
* ============================================================
* Código:
*     06_003_20
*
* Indicador:
*     Índice de juventud
*
* Fórmula:
*     indice_juventud = (pob_15_29 / pob_total) * 100
*
* Numerador:
*     población de 15 a 29 años
*
* Denominador:
*     población total
*
* Fuente:
*     _out/Bases DTA/pob_edades_simples_long_2012_20XX.dta
*
* Rango temporal:
*     2012–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Unidad:
*     Porcentaje de la población total.
*
* Consideraciones:
*     - Se calcula desde 2012 porque desde ese año la base tiene
*       edades simples suficientes para construir 15–29 años.
*     - El grupo 15–29 se construye sumando edades simples:
*           15 + 16 + ... + 29
*     - La población total se toma de grupo_edad == "total".
*     - No se calcula para 2005–2011 porque la base disponible no
*       contiene el grupo 15–29 separado.
*     - Como es un indicador posterior al primero, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 36A. Abrir base fuente
* ------------------------------------------------------------

	use "$out_dta/pob_edades_simples_long_2012_2024.dta", clear

	keep if inrange(anio, 2012, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

	* Usar población total por sexo.
	keep if sexo == "Total"

	* Estandarizar codigo_geo como texto.
	capture confirm numeric variable codigo_geo

	if _rc == 0 {
		gen str6 codigo_geo_str = string(codigo_geo, "%02.0f")
		drop codigo_geo
		rename codigo_geo_str codigo_geo
	}

	replace codigo_geo = strtrim(codigo_geo)

	* Asegurar que grupo_edad sea texto.
	capture confirm string variable grupo_edad

	if _rc != 0 {
		tostring grupo_edad, replace
	}

	replace grupo_edad = strtrim(grupo_edad)

* ------------------------------------------------------------
* 36B. Construir numerador y denominador
* ------------------------------------------------------------

	gen edad_simple = real(grupo_edad)

	* Numerador:
	*     población de 15 a 29 años.
	gen double pob_15_29 = 0
	replace pob_15_29 = poblacion if inrange(edad_simple, 15, 29)

	* Denominador:
	*     población total.
	gen double pob_total = 0
	replace pob_total = poblacion if grupo_edad == "total"

* ------------------------------------------------------------
* 36C. Colapsar a año + territorio
* ------------------------------------------------------------

	collapse (sum) pob_15_29 pob_total, by(anio codigo_geo nombre_geo nivel_geo)

	isid anio codigo_geo

* ------------------------------------------------------------
* 36D. Controles previos
* ------------------------------------------------------------

	count if missing(pob_15_29) | pob_15_29 < 0

	if r(N) > 0 {
		di as error "ERROR: Hay población de 15 a 29 años faltante o negativa."

		preserve
			keep if missing(pob_15_29) | pob_15_29 < 0
			sort anio codigo_geo
			export excel using "$out_aud/error_numerador_indice_juventud.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	count if missing(pob_total) | pob_total <= 0

	if r(N) > 0 {
		di as error "ERROR: Hay población total faltante o menor/igual a cero."

		preserve
			keep if missing(pob_total) | pob_total <= 0
			sort anio codigo_geo
			export excel using "$out_aud/error_denominador_indice_juventud.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	count if pob_15_29 > pob_total & !missing(pob_15_29, pob_total)

	if r(N) > 0 {
		di as error "ERROR: Hay casos donde la población 15–29 supera la población total."

		preserve
			keep if pob_15_29 > pob_total & !missing(pob_15_29, pob_total)
			sort anio codigo_geo
			export excel using "$out_aud/error_15_29_mayor_total_indice_juventud.xlsx", firstrow(variables) replace
		restore

		error 459
	}

* ------------------------------------------------------------
* 36E. Calcular indicador
* ------------------------------------------------------------
* Índice de juventud:
*
*     IJ = población 15–29 / población total * 100
*
* Interpretación:
*     Porcentaje de la población total que tiene entre 15 y 29 años.
* ------------------------------------------------------------

	gen indice_juventud = .

	replace indice_juventud = (pob_15_29 / pob_total) * 100 if pob_total > 0 & !missing(pob_15_29, pob_total)

	replace indice_juventud = round(indice_juventud, .01)

* ------------------------------------------------------------
* 36F. Preparar tabla final
* ------------------------------------------------------------

	keep anio codigo_geo indice_juventud

	reshape wide indice_juventud, i(anio) j(codigo_geo) string

	rename indice_juventud00 BOLIVIA
	rename indice_juventud01 CHUQUISACA
	rename indice_juventud02 LA_PAZ
	rename indice_juventud03 COCHABAMBA
	rename indice_juventud04 ORURO
	rename indice_juventud05 POTOSI
	rename indice_juventud06 TARIJA
	rename indice_juventud07 SANTA_CRUZ
	rename indice_juventud08 BENI
	rename indice_juventud09 PANDO

	order anio BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO

	sort anio

	format BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO %9.2f

* ------------------------------------------------------------
* 36G. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("indice_juventud") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 36H. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_003_20 | Índice de juventud"
	di as result "Unidad: porcentaje de la población total"
	di as result "Hoja: indice_juventud"
	di as result "============================================================"

	list, clean noobs

* ============================================================
**# BLOQUE 36: ÍNDICE DE INFANCIA
* ============================================================
* Código:
*     06_003_21
*
* Indicador:
*     Índice de infancia
*
* Fórmula:
*     indice_infancia = (pob_0_14 / pob_total) * 100
*
* Numerador:
*     población de 0 a 14 años
*
* Denominador:
*     población total
*
* Fuentes:
*     2005–2011:
*         _out/Bases DTA/pob_edad_salud_2005_20XX.dta
*
*     2012–20XX:
*         _out/Bases DTA/pob_edades_simples_long_2012_20XX.dta
*
* Rango temporal:
*     2005–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Unidad:
*     Porcentaje de la población total.
*
* Consideraciones:
*     - Para 2005–2011:
*           grupo_0_4 = población de 0 a 4 años
*           col_i     = población de 5 a 14 años
*           col_l     = población total / Total general
*     - Para 2012–20XX:
*           0-14 se construye sumando edades simples 0 a 14.
*           La población total se toma de grupo_edad == "total".
*     - Como es un indicador posterior al primero, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 36A. Construir indicador base 2005–2011
* ------------------------------------------------------------

	tempfile infancia_2005_2011

	use "$out_dta/pob_edad_salud_2005_2024.dta", clear

	keep if inrange(anio, 2005, 2011)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

	* Estandarizar codigo_geo como texto.
	capture confirm numeric variable codigo_geo

	if _rc == 0 {
		gen str6 codigo_geo_str = string(codigo_geo, "%02.0f")
		drop codigo_geo
		rename codigo_geo_str codigo_geo
	}

	replace codigo_geo = strtrim(codigo_geo)

	gen double pob_0_14  = grupo_0_4 + col_i
	gen double pob_total = col_l

	keep anio codigo_geo nombre_geo nivel_geo pob_0_14 pob_total

	isid anio codigo_geo

	save `infancia_2005_2011', replace

* ------------------------------------------------------------
* 36B. Construir indicador base 2012–20XX
* ------------------------------------------------------------

	tempfile infancia_2012_2024

	use "$out_dta/pob_edades_simples_long_2012_2024.dta", clear

	keep if inrange(anio, 2012, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

	* Usar población total por sexo.
	keep if sexo == "Total"

	* Estandarizar codigo_geo como texto.
	capture confirm numeric variable codigo_geo

	if _rc == 0 {
		gen str6 codigo_geo_str = string(codigo_geo, "%02.0f")
		drop codigo_geo
		rename codigo_geo_str codigo_geo
	}

	replace codigo_geo = strtrim(codigo_geo)

	* Asegurar que grupo_edad sea texto.
	capture confirm string variable grupo_edad

	if _rc != 0 {
		tostring grupo_edad, replace
	}

	replace grupo_edad = strtrim(grupo_edad)

	gen edad_simple = real(grupo_edad)

	* Numerador:
	*     población de 0 a 14 años.
	gen double pob_0_14 = 0
	replace pob_0_14 = poblacion if inrange(edad_simple, 0, 14)

	* Denominador:
	*     población total.
	gen double pob_total = 0
	replace pob_total = poblacion if grupo_edad == "total"

	collapse (sum) pob_0_14 pob_total, by(anio codigo_geo nombre_geo nivel_geo)

	isid anio codigo_geo

	save `infancia_2012_2024', replace

* ------------------------------------------------------------
* 36C. Unir periodos
* ------------------------------------------------------------

	use `infancia_2005_2011', clear

	append using `infancia_2012_2024'

	sort anio codigo_geo

	isid anio codigo_geo

* ------------------------------------------------------------
* 36D. Controles previos
* ------------------------------------------------------------

	count if missing(pob_0_14) | pob_0_14 < 0

	if r(N) > 0 {
		di as error "ERROR: Hay población de 0 a 14 años faltante o negativa."

		preserve
			keep if missing(pob_0_14) | pob_0_14 < 0
			sort anio codigo_geo
			export excel using "$out_aud/error_numerador_indice_infancia.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	count if missing(pob_total) | pob_total <= 0

	if r(N) > 0 {
		di as error "ERROR: Hay población total faltante o menor/igual a cero."

		preserve
			keep if missing(pob_total) | pob_total <= 0
			sort anio codigo_geo
			export excel using "$out_aud/error_denominador_indice_infancia.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	count if pob_0_14 > pob_total & !missing(pob_0_14, pob_total)

	if r(N) > 0 {
		di as error "ERROR: Hay casos donde la población 0–14 supera la población total."

		preserve
			keep if pob_0_14 > pob_total & !missing(pob_0_14, pob_total)
			sort anio codigo_geo
			export excel using "$out_aud/error_0_14_mayor_total_indice_infancia.xlsx", firstrow(variables) replace
		restore

		error 459
	}

* ------------------------------------------------------------
* 36E. Calcular indicador
* ------------------------------------------------------------
* Índice de infancia:
*
*     II = población 0–14 / población total * 100
*
* Interpretación:
*     Porcentaje de la población total que tiene entre 0 y 14 años.
* ------------------------------------------------------------

	gen indice_infancia = .

	replace indice_infancia = (pob_0_14 / pob_total) * 100 if pob_total > 0 & !missing(pob_0_14, pob_total)

	replace indice_infancia = round(indice_infancia, .01)

* ------------------------------------------------------------
* 36F. Preparar tabla final
* ------------------------------------------------------------

	keep anio codigo_geo indice_infancia

	reshape wide indice_infancia, i(anio) j(codigo_geo) string

	rename indice_infancia00 BOLIVIA
	rename indice_infancia01 CHUQUISACA
	rename indice_infancia02 LA_PAZ
	rename indice_infancia03 COCHABAMBA
	rename indice_infancia04 ORURO
	rename indice_infancia05 POTOSI
	rename indice_infancia06 TARIJA
	rename indice_infancia07 SANTA_CRUZ
	rename indice_infancia08 BENI
	rename indice_infancia09 PANDO

	order anio BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO

	sort anio

	format BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO %9.2f

* ------------------------------------------------------------
* 36G. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("indice_infancia") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 36H. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_003_21 | Índice de infancia"
	di as result "Fórmula usada: 0–14 / población total * 100"
	di as result "Hoja: indice_infancia"
	di as result "============================================================"

	list, clean noobs

* ============================================================
**# BLOQUE 37: ÍNDICE DE POTENCIALIDAD DE REPRODUCCIÓN FEMENINA
* ============================================================
* Código:
*     06_003_22
*
* Indicador:
*     Índice de potencialidad de reproducción femenina
*
* Sigla:
*     IPRF
*
* Fórmula:
*     iprf = (mujeres_20_34 / mujeres_35_49) * 100
*
* Numerador:
*     mujeres de 20 a 34 años
*
* Denominador:
*     mujeres de 35 a 49 años
*
* Fuente:
*     _out/Bases DTA/pob_edades_simples_long_2012_20XX.dta
*
* Rango temporal:
*     2012–20XX
*
* Universo:
*     Bolivia y departamentos.
*
* Unidad:
*     Mujeres de 20 a 34 años por cada 100 mujeres de 35 a 49 años.
*
* Consideraciones:
*     - Se calcula desde 2012 porque desde ese año la base tiene
*       edades simples y grupos quinquenales por sexo.
*     - El grupo 20-34 se construye como:
*           edades simples 20 a 29
*           + grupo 30_34.
*     - El grupo 35-49 se construye como:
*           35_39 + 40_44 + 45_49.
*     - Solo se usa sexo == "Mujer".
*     - No se calcula para 2005–2011 porque la base disponible no
*       contiene estos grupos específicos por sexo.
*     - Como es un indicador posterior al primero, se exporta con sheetreplace.
* ============================================================
capture restore
* ------------------------------------------------------------
* 37A. Abrir base fuente
* ------------------------------------------------------------

	use "$out_dta/pob_edades_simples_long_2012_2024.dta", clear

	keep if inrange(anio, 2012, 2024)

	* Mantener solo Bolivia y departamentos.
	keep if inlist(nivel_geo, 0, 1)

	* Mantener solo mujeres.
	keep if sexo == "Mujer"

	* Estandarizar codigo_geo como texto.
	capture confirm numeric variable codigo_geo

	if _rc == 0 {
		gen str6 codigo_geo_str = string(codigo_geo, "%02.0f")
		drop codigo_geo
		rename codigo_geo_str codigo_geo
	}

	replace codigo_geo = strtrim(codigo_geo)

	* Asegurar que grupo_edad sea texto.
	capture confirm string variable grupo_edad

	if _rc != 0 {
		tostring grupo_edad, replace
	}

	replace grupo_edad = strtrim(grupo_edad)

* ------------------------------------------------------------
* 37B. Construir numerador y denominador
* ------------------------------------------------------------

	gen edad_simple = real(grupo_edad)

	* Numerador:
	*     mujeres de 20 a 34 años.
	gen double mujeres_20_34 = 0

	replace mujeres_20_34 = poblacion if inrange(edad_simple, 20, 29)

	replace mujeres_20_34 = poblacion if grupo_edad == "30_34"


	* Denominador:
	*     mujeres de 35 a 49 años.
	gen double mujeres_35_49 = 0

	replace mujeres_35_49 = poblacion if regexm(grupo_edad, "^(35_39|40_44|45_49)$")

* ------------------------------------------------------------
* 37C. Colapsar a año + territorio
* ------------------------------------------------------------

	collapse (sum) mujeres_20_34 mujeres_35_49, by(anio codigo_geo nombre_geo nivel_geo)

	isid anio codigo_geo

* ------------------------------------------------------------
* 37D. Controles previos
* ------------------------------------------------------------

	count if missing(mujeres_20_34) | mujeres_20_34 < 0

	if r(N) > 0 {
		di as error "ERROR: Hay mujeres de 20 a 34 años faltantes o negativas."

		preserve
			keep if missing(mujeres_20_34) | mujeres_20_34 < 0
			sort anio codigo_geo
			export excel using "$out_aud/error_numerador_iprf.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	count if missing(mujeres_35_49) | mujeres_35_49 <= 0

	if r(N) > 0 {
		di as error "ERROR: Hay mujeres de 35 a 49 años faltantes o menor/igual a cero."

		preserve
			keep if missing(mujeres_35_49) | mujeres_35_49 <= 0
			sort anio codigo_geo
			export excel using "$out_aud/error_denominador_iprf.xlsx", firstrow(variables) replace
		restore

		error 459
	}

* ------------------------------------------------------------
* 37E. Calcular indicador
* ------------------------------------------------------------
* Índice de potencialidad de reproducción femenina:
*
*     IPRF = mujeres 20-34 / mujeres 35-49 * 100
*
* Interpretación:
*     Mujeres de 20 a 34 años por cada 100 mujeres de 35 a 49 años.
* ------------------------------------------------------------

	gen iprf = .

	replace iprf = (mujeres_20_34 / mujeres_35_49) * 100 if mujeres_35_49 > 0 & !missing(mujeres_20_34, mujeres_35_49)

	replace iprf = round(iprf, .01)

* ------------------------------------------------------------
* 37F. Preparar tabla final
* ------------------------------------------------------------

	keep anio codigo_geo iprf

	reshape wide iprf, i(anio) j(codigo_geo) string

	rename iprf00 BOLIVIA
	rename iprf01 CHUQUISACA
	rename iprf02 LA_PAZ
	rename iprf03 COCHABAMBA
	rename iprf04 ORURO
	rename iprf05 POTOSI
	rename iprf06 TARIJA
	rename iprf07 SANTA_CRUZ
	rename iprf08 BENI
	rename iprf09 PANDO

	order anio BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO

	sort anio

	format BOLIVIA CHUQUISACA LA_PAZ COCHABAMBA ORURO POTOSI TARIJA SANTA_CRUZ BENI PANDO %9.2f

* ------------------------------------------------------------
* 37G. Exportar hoja al Excel final
* ------------------------------------------------------------

	cd "$out_xlsx"

	export excel using "indicadores_demografia_ssr_2005_2024.xlsx", sheet("indice_pot_reprod_fem") firstrow(variables) sheetreplace

* ------------------------------------------------------------
* 37H. Control en pantalla
* ------------------------------------------------------------

	di as result "============================================================"
	di as result "INDICADOR EXPORTADO:"
	di as result "06_003_22 | Índice de potencialidad de reproducción femenina"
	di as result "Unidad: mujeres de 20–34 por cada 100 mujeres de 35–49 años"
	di as result "Hoja: indice_pot_reprod_fem"
	di as result "============================================================"

	list, clean noobs

* ============================================================
* CONTROL FINAL DEL SCRIPT
* ============================================================

	di as result "============================================================"
	di as result "SCRIPT 03 FINALIZADO"
	di as result "Archivo Excel actualizado:"
	di as result "$out_xlsx/indicadores_demografia_ssr_2005_2024.xlsx"
	di as result "============================================================"
	
	
	