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
**     02_limpiar_edades_hombre_mujer.do

** OBJETIVO GENERAL:
**     Limpiar la base raw de edades simples y quinquenales, identificando correctamente
**     las filas de sexo y transformando la base desde formato ancho hacia formato largo.

** OBJETIVOS ESPECÍFICOS:
**     1. Leer la base raw importada por 01_importar_proyecciones.do.
**     2. Identificar correctamente filas geográficas, filas de mujeres y filas de hombres.
**     3. Corregir el problema de mayúsculas en algunos años:
**            1.Mujer / 2.Hombre
**            1.MUJER / 2.HOMBRE
**     4. Heredar código geográfico, nombre geográfico y nivel geográfico a las filas
**        de sexo.
**     5. Renombrar columnas de edad simple y grupos quinquenales.
**     6. Transformar la base de formato ancho a formato largo.
**     7. Guardar la base limpia en:
**            _out\Bases DTA
**     8. Exportar auditorías en:
**            _out\Auditorías
**        solo si se detectan problemas.

** PROBLEMA QUE RESUELVE:
**     La base Excel de edades simples y quinquenales tiene una estructura jerárquica:

**         00. BOLIVIA
**         1.Mujer
**         2.Hombre
**         01.CHUQUISACA
**         1.Mujer
**         2.Hombre

**     En algunos años, las filas de sexo aparecen en mayúsculas:

**         1.MUJER
**         2.HOMBRE

**     Por eso el script normaliza el texto con upper() antes de clasificar sexo.

** ENTRADA:
**     _out\Bases DTA\pob_edades_simples_raw_2012_20XX.dta

** SALIDA:
**     _out\Bases DTA\pob_edades_simples_long_2012_20XX.dta

** VARIABLES PRINCIPALES DE SALIDA:
**     anio
**     codigo_geo
**     nombre_geo
**     nivel_geo
**     sexo
**     grupo_edad
**     grupo_edad_label
**     edad_min
**     edad_max
**     orden_grupo_edad
**     tipo_grupo_edad
**     poblacion

** AUTORA:
**     Alejandra Arleth Lafuente-Luizaga

** FECHA DE CREACIÓN:
**     06-may-2026

** ÚLTIMA ACTUALIZACIÓN:
**     06-may-2026

** CONTROLES ESPERADOS:
**     count if missing(grupo_edad_label) = 0
**    count if missing(tipo_grupo_edad)  = 0
**     count if missing(nivel_geo)        = 0

** NOTAS:
**     - Este script no calcula indicadores.
**     - Este script no exporta indicadores a Excel.
**     - Este script solo genera una base limpia necesaria para calcular indicadores.

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
	di as result "out_dta:      $out_dta"
	di as result "out_xlsx:     $out_xlsx"
	di as result "out_graficos: $out_graficos"
	di as result "master:       $master"
	di as result "============================================================"

* ============================================================
* 0.A. PARÁMETROS DEL SCRIPT
* ============================================================

	local yini_edades 2012
	local yfin_edades 2024

* ============================================================
* 1. VERIFICAR Y ABRIR BASE RAW
* ============================================================
*
* Esta base debe haber sido creada por:
*     01_importar_proyecciones.do
*
* Archivo esperado:
*     _out\Bases DTA\pob_edades_simples_raw_2012_20XX.dta
* ============================================================

	confirm file "$out_dta\pob_edades_simples_raw_`yini_edades'_`yfin_edades'.dta"

	use "$out_dta\pob_edades_simples_raw_`yini_edades'_`yfin_edades'.dta", clear

	replace geo_raw = strtrim(geo_raw)
	drop if missing(geo_raw)

	gen orden_original = _n

* ============================================================
* 2. IDENTIFICAR FILAS DE SEXO
* ============================================================
* Se crea una versión normalizada de geo_raw para detectar:
*     1.Mujer
*     2.Hombre
*     1.MUJER
*     2.HOMBRE
*
* Resultado:
*     sexo = Total, Mujer u Hombre
* ============================================================

	gen geo_norm = upper(strtrim(geo_raw))

	gen es_sexo = inlist(geo_norm, "1.MUJER", "2.HOMBRE")

	gen sexo = "Total"
	replace sexo = "Mujer"  if geo_norm == "1.MUJER"
	replace sexo = "Hombre" if geo_norm == "2.HOMBRE"


* ============================================================
* 3. EXTRAER CÓDIGO Y NOMBRE GEOGRÁFICO
* ============================================================
* Esta extracción solo se hace para filas geográficas.
* Las filas de sexo heredarán la información geográfica en el paso siguiente.
* ============================================================

	gen codigo_geo = ""
	replace codigo_geo = regexs(1) if es_sexo == 0 & regexm(geo_raw, "^([0-9]+)")

	gen nombre_geo = ""
	replace nombre_geo = strtrim(regexr(geo_raw, "^[0-9]+\.?[ ]*", "")) if es_sexo == 0


* ============================================================
* 4. CLASIFICAR NIVEL GEOGRÁFICO
* ============================================================
* Reglas:
*     00      = Nacional
*     01-09   = Departamental
*     4 díg.  = Provincial
*     6 díg.  = Municipal
* ============================================================

	gen nivel_geo = .

	replace nivel_geo = 0 if codigo_geo == "00"
	replace nivel_geo = 1 if strlen(codigo_geo) == 2 & codigo_geo != "00"
	replace nivel_geo = 2 if strlen(codigo_geo) == 4
	replace nivel_geo = 3 if strlen(codigo_geo) == 6

	label define nivel_geo 0 "Nacional" 1 "Departamental" 2 "Provincial" 3 "Municipal", replace
	label values nivel_geo nivel_geo


* ============================================================
* 5. HEREDAR IDENTIFICADORES GEOGRÁFICOS EN FILAS DE SEXO
* ============================================================
* Las filas 1.Mujer y 2.Hombre no traen código geográfico propio.
* Por eso se copia el código, nombre y nivel desde la fila geográfica anterior.
* ============================================================

	sort anio orden_original

	by anio (orden_original): replace codigo_geo = codigo_geo[_n-1] if es_sexo == 1 & codigo_geo == ""
	by anio (orden_original): replace nombre_geo = nombre_geo[_n-1] if es_sexo == 1 & nombre_geo == ""
	by anio (orden_original): replace nivel_geo  = nivel_geo[_n-1]  if es_sexo == 1 & missing(nivel_geo)


* ============================================================
* 5A. CONTROLES DE IDENTIFICACIÓN GEOGRÁFICA
* ============================================================

	count if missing(codigo_geo) | codigo_geo == ""

	if r(N) > 0 {
		di as error "ERROR: Hay filas sin codigo_geo después de heredar identificadores."

		preserve
			keep if missing(codigo_geo) | codigo_geo == ""
			export excel using "$out_aud\auditoria_missing_codigo_geo_edades.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	count if missing(nivel_geo)

	if r(N) > 0 {
		di as error "ERROR: Hay filas sin nivel_geo después de heredar identificadores."

		preserve
			keep if missing(nivel_geo)
			export excel using "$out_aud\auditoria_missing_nivel_geo_edades.xlsx", firstrow(variables) replace
		restore

		error 459
	}

* ============================================================
* 6. RENOMBRAR COLUMNAS DE EDAD
* ============================================================
* La base viene en formato ancho:
*     B  = población de 0 años
*     C  = población de 1 año
*     ...
*     AE = población de 29 años
*     AF = población de 30-34 años
*     ...
*     AP = población de 80 años y más
*     AQ = total general
*
* Se renombran como:
*     pob_0
*     pob_1
*     ...
*     pob_30_34
*     ...
*     pob_80_mas
*     pob_total
* ============================================================

	local edad = 0

	foreach v in B C D E F G H I J K L M N O P Q R S T U V W X Y Z AA AB AC AD AE {
		rename `v' pob_`edad'
		local edad = `edad' + 1
	}

	rename AF pob_30_34
	rename AG pob_35_39
	rename AH pob_40_44
	rename AI pob_45_49
	rename AJ pob_50_54
	rename AK pob_55_59
	rename AL pob_60_64
	rename AM pob_65_69
	rename AN pob_70_74
	rename AO pob_75_79
	rename AP pob_80_mas
	rename AQ pob_total

* ============================================================
* 7. TRANSFORMAR A FORMATO LARGO
* ============================================================
* De:
*     una fila con muchas columnas de edad
*
* A:
*     una fila por año, código geográfico, sexo y grupo de edad.
* ============================================================

	gen id_row = _n

	reshape long pob_, i(id_row) j(grupo_edad) string

	rename pob_ poblacion

* ============================================================
* 8. ETIQUETAR GRUPOS DE EDAD
* ============================================================

	gen grupo_edad_label = ""

	replace grupo_edad_label = grupo_edad + " años" if regexm(grupo_edad, "^[0-9]+$")
	replace grupo_edad_label = subinstr(grupo_edad, "_", "-", .) + " años" if regexm(grupo_edad, "^[0-9]+_[0-9]+$")
	replace grupo_edad_label = "80 años y más" if grupo_edad == "80_mas"
	replace grupo_edad_label = "Total general" if grupo_edad == "total"

* ============================================================
* 9. CREAR EDAD MÍNIMA, EDAD MÁXIMA Y ORDEN
* ============================================================

	gen edad_min = .
	gen edad_max = .

	replace edad_min = real(grupo_edad) if regexm(grupo_edad, "^[0-9]+$")
	replace edad_max = edad_min if regexm(grupo_edad, "^[0-9]+$")

	replace edad_min = real(regexs(1)) if regexm(grupo_edad, "^([0-9]+)_([0-9]+)$")
	replace edad_max = real(regexs(2)) if regexm(grupo_edad, "^([0-9]+)_([0-9]+)$")

	replace edad_min = 80 if grupo_edad == "80_mas"
	replace edad_max = .  if grupo_edad == "80_mas"

	gen orden_grupo_edad = edad_min
	replace orden_grupo_edad = 999 if grupo_edad == "total"

* ============================================================
* 10. CLASIFICAR TIPO DE GRUPO DE EDAD
* ============================================================

	gen tipo_grupo_edad = ""

	replace tipo_grupo_edad = "edad_simple" if regexm(grupo_edad, "^[0-9]+$")
	replace tipo_grupo_edad = "grupo_quinquenal" if regexm(grupo_edad, "^[0-9]+_[0-9]+$") | grupo_edad == "80_mas"
	replace tipo_grupo_edad = "total" if grupo_edad == "total"

* ============================================================
* 11. CONTROLES FINALES DE CALIDAD
* ============================================================

	count if missing(grupo_edad_label) | grupo_edad_label == ""

	if r(N) > 0 {
		di as error "ERROR: Hay filas sin grupo_edad_label."

		preserve
			keep if missing(grupo_edad_label) | grupo_edad_label == ""
			export excel using "$out_aud\auditoria_missing_grupo_edad_label.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	count if missing(tipo_grupo_edad) | tipo_grupo_edad == ""

	if r(N) > 0 {
		di as error "ERROR: Hay filas sin tipo_grupo_edad."

		preserve
			keep if missing(tipo_grupo_edad) | tipo_grupo_edad == ""
			export excel using "$out_aud\auditoria_missing_tipo_grupo_edad.xlsx", firstrow(variables) replace
		restore

		error 459
	}

	count if missing(nivel_geo)

	if r(N) > 0 {
		di as error "ERROR: Hay filas sin nivel_geo en la base final."

		preserve
			keep if missing(nivel_geo)
			export excel using "$out_aud\auditoria_missing_nivel_geo_final.xlsx", firstrow(variables) replace
		restore

		error 459
	}

* ============================================================
* 12. ORDENAR Y GUARDAR BASE LIMPIA
* ============================================================

	drop geo_norm

	order anio codigo_geo nombre_geo nivel_geo sexo grupo_edad grupo_edad_label edad_min edad_max orden_grupo_edad tipo_grupo_edad poblacion geo_raw orden_original id_row es_sexo

	sort anio codigo_geo sexo orden_grupo_edad

	compress

	save "$out_dta\pob_edades_simples_long_`yini_edades'_`yfin_edades'.dta", replace

* ============================================================
* 13. CONTROLES MÍNIMOS EN PANTALLA
* ============================================================

	di as result "============================================================"
	di as result "CONTROL DE LIMPIEZA: EDADES SIMPLES Y QUINQUENALES"
	di as result "============================================================"

	di as text "Distribución por sexo:"
	tab sexo

	di as text "Nivel geográfico por sexo:"
	tab nivel_geo sexo, missing

	di as text "Tipo de grupo de edad:"
	tab tipo_grupo_edad, missing

	di as text "Conteos esperados de faltantes:"
	count if missing(grupo_edad_label)
	count if missing(tipo_grupo_edad)
	count if missing(nivel_geo)

	di as result "============================================================"
	di as result "BASE LIMPIA GUARDADA EN:"
	di as result "$out_dta\pob_edades_simples_long_`yini_edades'_`yfin_edades'.dta"
	di as result "FIN DEL SCRIPT 02_limpiar_edades_hombre_mujer.do"
	di as result "============================================================"