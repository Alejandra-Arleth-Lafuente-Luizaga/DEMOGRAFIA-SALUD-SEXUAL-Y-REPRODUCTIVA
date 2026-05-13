clear all
set more off

////////////////////////////////////////////////////////////////////////////////
//                                                                            //
//          INSTITUTO DE INVESTIGACIONES SOCIO-ECONÓMICAS IISEC               //
//                                                                            //
////////////////////////////////////////////////////////////////////////////////

**  PROYECTO:                  INDICADORES DE DEMOGRAFÍA, SALUD SEXUAL Y REPRODUCTIVA
**  IDENTIFICADOR:             grap_demografía_salud_sexual_y_reproductiva.do
**  AUTOR:                     ALEJANDRA ARLETH LAFUENTE LUIZAGA
**  FECHA DE CREACIÓN:         06.06.2024
**  ÚLTIMO EN MODIFICAR:       ALEJANDRA ARLETH LAFUENTE LUIZAGA
**  FECHA DE MODIFICACIÓN:     08.05.2026


* ============================================================
* 0. DIRECTORIOS DEL PROYECTO
* ============================================================
*
* Este script:
*     - Lee los Excel individuales para gráficos desde:
*           !EH_ARMONIZADA/_salud
*
*     - Exporta las imágenes a:
*           INDICADORES/6_SALUD/_gph/DemografíaSaludSexualyReproductiva
*
* Nota:
*     Se usa "/" en lugar de "\" para evitar problemas con rutas.
* ============================================================

* ============================================================
* 0. DIRECTORIOS PARA GRÁFICOS
* ============================================================

	clear all
	set more off

	global path "G:/Unidades compartidas/1_INDICADORES_ODSB_2022/01COPIA_IISEC_ODSB"

	* Verificar ruta raíz
	capture dir "$path/*"

	if _rc != 0 {
		di as error "ERROR: Stata no puede leer la ruta raíz:"
		di as error "$path"
		error 601
	}

	* Directorios principales
	global work "$path/INDICADORES/6_SALUD"

	global in   "$path/!EH_ARMONIZADA/_salud"
	global do   "$work/_do"
	global gph  "$work/_gph/DemografíaSaludSexualyReproductiva"
	global xlsx "$work/_xlsx"
	global out  "$work/_out"

	* Crear carpetas si no existen
	capture mkdir "$work/_gph"
	capture mkdir "$gph"

	* Verificar insumos
	capture dir "$in/*"

	if _rc != 0 {
		di as error "ERROR: Stata no puede leer la carpeta de insumos gráficos:"
		di as error "$in"
		error 601
	}

	* Verificar carpeta de salida
	capture dir "$gph/*"

	if _rc != 0 {
		di as error "ERROR: Stata no puede leer o crear la carpeta de gráficos:"
		di as error "$gph"
		error 601
	}

	di as result "============================================================"
	di as result "DIRECTORIOS PARA GRÁFICOS CONFIGURADOS"
	di as result "path: $path"
	di as result "in:   $in"
	di as result "gph:  $gph"
	di as result "============================================================"

* ============================================================
* 1. FORMATO DE IMÁGENES
* ============================================================

	graph set window fontface "Arial Narrow"

	grstyle init
	grstyle set legend 5, nobox 
	grstyle set graphsize 1200 900
	grstyle set size 15pt: heading
	grstyle set size 15pt: subheading axis_title 
	grstyle set size 15pt: tick_label 
	grstyle set size 15pt: text_option

* ============================================================
* 2. NOTAS, FUENTES Y EJES
* ============================================================

	global nota1  "{bf:[1] Nota:} Las proyecciones de datos para el período 2005-2011 se basan en el Censo Nacional de Población y Vivienda (CNPV) 2001."
	global nota11 "{bf:[1] Nota:} De 2012 a 2023, las proyecciones se basan en la Revisión 2020. Para el caso de 2024 en adelante, corresponden a la Revisión 2025."
	global nota2  "{bf:[2] Nota:} De 2012 a 2023, las proyecciones se basan en la Revisión 2020. Para el caso de 2024 en adelante, corresponden a la Revisión 2025."
	global nota21 "{bf:[2] Nota:} Las proyecciones de población son elaboradas con base a información sobre los componentes demográficos de censos y encuestas de demografía y salud."
	global nota3  "{bf:[3] Nota:} Las proyecciones de población son elaboradas con base a información sobre los componentes demográficos de censos y encuestas de demografía y salud."

	global fuente  "{bf:Fuente:} Ministerio de Educación, Ministerio de Salud y Deportes e Instituto Nacional de Estadísticas (INE)."
	global fuente1 "{bf:Fuente:} Elaboración Observatorio IISEC-UCB."
	global fuente2 "{bf:Fuente:} Elaboración Observatorio IISEC-UCB en base a datos del ME, MSyD e INE."

	global eje    "Número de Mujeres en Edad Fértil (MEF)"
	global eje1   "Porcentaje (%)"    
	global eje2   "Número de embarazos"
	global eje3   "Número de partos"
	global eje4   "Número de abortos"
	global eje5   "Número de nacimientos"
	global eje6   "Número de Hijos Nacidos Vivos"
	global eje7   "Número de Hijos Nacidos Muertos"
	global eje8   "Número de abortos por cada 1,000 Mujeres en Edad Fértil"    
	global eje9   "Número de Hijos Nacidos Vivos por cada 1,000 Mujeres en Edad Fértil"    
	global eje10  "Número de defunciones fetales por cada 1,000 Hijos Nacidos Vivos"
	global eje11  "Número de embarazos por cada 1,000 Mujeres en Edad Fértil"
	global eje12  "Número de Hijos Nacidos Vivos por cada 1,000 habitantes"
	global eje13  "Número de hombres por cada 100 mujeres"
	global eje14  "Número de personas dependientes por cada 100 personas en edad activa"
	global eje15  "Número de adultos mayores por cada 100 niños y jóvenes"
	global eje16  "Número de jóvenes dependientes por cada 100 personas en edad activa"
	global eje17  "Número de adultos mayores por cada 100 personas en edad activa"
	global eje18  "Número de jóvenes por cada 100 personas mayores de 65 y más años"
	global eje19  "Edad mediana en años"
	global eje20  "Edad media en años"
	global eje21  "Número de personas que se retiran por cada 100 personas que ingresan"
	global eje22  "Número de personas maduras por cada 100 personas jóvenes"
	global eje23  "Número de niños pequeños de 0 a 4 años por cada 100 Mujeres en Edad Fértil"
	global eje24  "Número de personas de 0 a 19 años por cada 100 personas de 30 a 49 años"
	global eje25  "Número de personas jóvenes y mayores por cada 100 personas de 15 a 49 años"
	global eje26  "Número de personas de 35 a 64 años por cada 100 personas de 65 y más años"
	global eje27  "Número de personas mayores de 80 años por cada 100 mayores de 65 años"
	global eje28  "Número de mujeres de 20 a 34 años por cada 100 mujeres entre 35 a 49 años"
	global eje29  "Tasa de crecimiento poblacional"
	global eje30  "Número de personas de 65 años y más por cada 100 habitantes"
	global eje31  "Número de personas de 80 años o más por cada 100 personas de 65 años o más"
	global eje32  "Número de personas de 5 a 14 años por cada 100 personas de 45 a 64 años"
	global eje33  "Número de personas de 15 a 29 años por cada 100 habitantes"
	global eje34  "Número de personas menores a 15 años por cada 100 habitantes"
	
* ============================================================
**# 6.1.1. Número de Mujeres en Edad Fértil (MEF)
* Código: 06_001_01
* ============================================================

* ------------------------------------------------------------
* 6.1.1.1. Nacional
* ------------------------------------------------------------

	import excel "$in/06_001_01.xlsx", sheet("_in") firstrow clear

	set scheme white_tableau

	format %15.0fc Bolivia
	
	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Bolivia_2005_2011 = Bolivia if Año <= 2011
	gen Bolivia_2012_2023 = Bolivia if inrange(Año, 2012, 2023)
	gen Bolivia_2024_2026 = Bolivia if Año >= 2024

	* Etiquetas de datos solo para el gráfico nacional.
	gen Bolivia_lbl_2005_2011 = string(Bolivia, "%15.0fc") if Año <= 2011
	gen Bolivia_lbl_2012_2023 = string(Bolivia, "%15.0fc") if inrange(Año, 2012, 2023)
	gen Bolivia_lbl_2024_2026 = string(Bolivia, "%15.0fc") if Año >= 2024

	twoway ///
		(connected Bolivia_2005_2011 Año, ///
			mlabel(Bolivia_lbl_2005_2011) ///
			mlabposition(12) ///
			mlabsize(*.75) ///
			msymbol(circle)) ///
		(connected Bolivia_2012_2023 Año, ///
			mlabel(Bolivia_lbl_2012_2023) ///
			mlabposition(12) ///
			mlabsize(*.75) ///
			msymbol(circle)) ///
		(connected Bolivia_2024_2026 Año, ///
			mlabel(Bolivia_lbl_2024_2026) ///
			mlabposition(12) ///
			mlabsize(*.75) ///
			msymbol(circle)), ///
		xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		xlab(2005(1)2024, angle(45)) ///
		ylab(2000000(250000)3500000, format(%15.0fc)) ///
		xtitle("") ///
		ytitle("$eje") ///
		note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
		legend(off)

	gr export "$gph/06_001_01_01.png", replace

* ------------------------------------------------------------
* 6.1.1.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_001_01.xlsx", sheet("_in1") firstrow clear

	set scheme white_viridis

	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"

	label values Departamento_order Departamento_order_lbl
	
	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Valor_2005_2011 = Valor if Año <= 2011
	gen Valor_2012_2023 = Valor if inrange(Año, 2012, 2023)
	gen Valor_2024_2026 = Valor if Año >= 2024

	twoway ///
		(connected Valor_2005_2011 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Chuquisaca", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Chuquisaca", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "LaPaz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "LaPaz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Cochabamba", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Cochabamba", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Oruro", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Oruro", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Potosí", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Potosí", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Tarija", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Tarija", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "SantaCruz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "SantaCruz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Beni", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Beni", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Pando", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Pando", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)), ///
		by(Departamento_order, note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) legend(off)) ///
		xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		ylab(0(500000)1000000, format(%15.0fc)) ///
		xlab(2005(1)2024, angle(90)) ///
		xtitle("") ///
		ytitle("$eje") ///
		title("", box color(black) span)

	gr export "$gph/06_001_01_02.png", replace

* ============================================================
**# 6.1.2. Proporción de Mujeres en Edad Fértil (MEF)
* Código: 06_001_02
* ============================================================

* ------------------------------------------------------------
* 6.1.2.1. Nacional
* ------------------------------------------------------------	
	
	import excel "$in/06_001_02.xlsx", sheet("_in") firstrow clear
	
	set scheme white_tableau
	
	format %4.2fc Bolivia

	* Separar tramos:
	* 2012-2023 y 2024 en adelante.
	gen Bolivia_2012_2023 = Bolivia if inrange(Año, 2012, 2023)
	gen Bolivia_2024_2026 = Bolivia if Año >= 2024

	* Etiquetas de datos solo para el gráfico nacional.
	gen Bolivia_lbl_2012_2023 = string(Bolivia, "%4.2fc") if inrange(Año, 2012, 2023)
	gen Bolivia_lbl_2024_2026 = string(Bolivia, "%4.2fc") if Año >= 2024

	twoway ///
		(connected Bolivia_2012_2023 Año, ///
			mlabel(Bolivia_lbl_2012_2023) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2024_2026 Año, ///
			mlabel(Bolivia_lbl_2024_2026) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)), ///
		xline(2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		xlab(2012(1)2024, angle(45)) ///
		ylab(40(4)60, format(%9.0fc)) ///
		xtitle("") ///
		ytitle("$eje1") ///
		note("$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
		legend(off)

	gr export "$gph/06_001_02_01.png", replace

* ------------------------------------------------------------
* 6.1.2.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_001_02.xlsx", sheet("_in1") firstrow clear

	set scheme white_viridis

	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"
	label values Departamento_order Departamento_order_lbl

	* Separar tramos para generar el corte visual
	gen Valor_2012_2023 = Valor if inrange(Año, 2012, 2023)
	gen Valor_2024_2026 = Valor if Año >= 2024

	twoway ///
		(connected Valor_2012_2023 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Chuquisaca", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "LaPaz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Cochabamba", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Oruro", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Potosí", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Tarija", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "SantaCruz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Beni", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Pando", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)), ///
		by(Departamento_order, note("$nota2" "$nota3" "$fuente2", size(vsmall) span) legend(off)) ///
		xline(2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		ylab(40(10)60, format(%9.0fc)) ///
		xlab(2012(1)2024, angle(45)) ///
		xtitle("") ///
		ytitle("$eje1") ///
		title("", box color(black) span)

	gr export "$gph/06_001_02_02.png", replace
	
* ============================================================
**# 6.1.3. Número de embarazos
* Código: 06_001_03
* ============================================================

* ------------------------------------------------------------
* 6.1.2.1. Nacional
* ------------------------------------------------------------	
	
	import excel "$in/06_001_03.xlsx", sheet("_in") firstrow clear

	set scheme white_tableau

	format %9.0fc Bolivia

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Bolivia_2005_2011 = Bolivia if Año <= 2011
	gen Bolivia_2012_2023 = Bolivia if inrange(Año, 2012, 2023)
	gen Bolivia_2024_2026 = Bolivia if Año >= 2024

	* Etiquetas de datos solo para el gráfico nacional.
	gen Bolivia_lbl_2005_2011 = string(Bolivia, "%15.0fc") if Año <= 2011
	gen Bolivia_lbl_2012_2023 = string(Bolivia, "%15.0fc") if inrange(Año, 2012, 2023)
	gen Bolivia_lbl_2024_2026 = string(Bolivia, "%15.0fc") if Año >= 2024

	twoway ///
	(connected Bolivia_2005_2011 Año, ///
		mlabel(Bolivia_lbl_2005_2011) ///
		mlabp(12) ///
		mlabs(*.90) ///
		msymbol(circle)) ///
	(connected Bolivia_2012_2023 Año, ///
		mlabel(Bolivia_lbl_2012_2023) ///
		mlabp(12) ///
		mlabs(*.90) ///
		msymbol(circle)) ///
	(connected Bolivia_2024_2026 Año, ///
		mlabel(Bolivia_lbl_2024_2026) ///
		mlabposition(12) ///
		mlabsize(*.90) ///
		msymbol(circle)), ///
	xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
	xlab(2005(1)2024, angle(45)) ///
	ylab(160000(40000)360000, format(%9.0fc)) ///
	xtitle("") ///
	ytitle("$eje2") ///
	note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
	legend(off)

	gr export "$gph/06_001_03_01.png", replace
	
* ------------------------------------------------------------
* 6.1.3.2. Departamental
* ------------------------------------------------------------
	
	import excel "$in/06_001_03.xlsx", sheet("_in1") firstrow clear

	set scheme white_viridis

	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"

	label values Departamento_order Departamento_order_lbl

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Valor_2005_2011 = Valor if Año <= 2011
	gen Valor_2012_2023 = Valor if inrange(Año, 2012, 2023)
	gen Valor_2024_2026 = Valor if Año >= 2024

	twoway ///
		(connected Valor_2005_2011 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Chuquisaca", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Chuquisaca", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		///
		(connected Valor_2005_2011 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "LaPaz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "LaPaz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		///
		(connected Valor_2005_2011 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Cochabamba", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Cochabamba", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		///
		(connected Valor_2005_2011 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Oruro", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Oruro", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		///
		(connected Valor_2005_2011 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Potosí", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Potosí", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		///
		(connected Valor_2005_2011 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Tarija", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Tarija", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		///
		(connected Valor_2005_2011 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "SantaCruz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "SantaCruz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		///
		(connected Valor_2005_2011 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Beni", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Beni", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		///
		(connected Valor_2005_2011 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Pando", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Pando", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)), ///
		by(Departamento_order, note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) legend(off)) ///
		xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		ylab(0(50000)100000, format(%9.0fc)) ///
		xlab(2005(1)2024, angle(90)) ///
		xtitle("") ///
		ytitle("$eje2") ///
		title("", box color(black) span)

	gr export "$gph/06_001_03_02.png", replace
	
* ============================================================
* 6.1.4. Número de partos
* Código: 06_001_04
* ============================================================
	
* ------------------------------------------------------------
* 6.1.4.1. Nacional
* ------------------------------------------------------------
	
	import excel "$in/06_001_04.xlsx", sheet("_in") firstrow clear
	
	set scheme white_tableau

	format %9.0fc Bolivia

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Bolivia_2005_2011 = Bolivia if Año <= 2011
	gen Bolivia_2012_2023 = Bolivia if inrange(Año, 2012, 2023)
	gen Bolivia_2024_2026 = Bolivia if Año >= 2024

	* Etiquetas de datos solo para el gráfico nacional
	gen Bolivia_lbl_2005_2011 = string(Bolivia_2005_2011, "%9.0fc")
	gen Bolivia_lbl_2012_2023 = string(Bolivia_2012_2023, "%9.0fc")
	gen Bolivia_lbl_2024_2026 = string(Bolivia, "%15.0fc") if Año >= 2024

	twoway ///
	(connected Bolivia_2005_2011 Año, ///
		mlabel(Bolivia_lbl_2005_2011) ///
		mlabp(12) ///
		mlabs(*.90) ///
		msymbol(circle)) ///
	(connected Bolivia_2012_2023 Año, ///
		mlabel(Bolivia_lbl_2012_2023) ///
		mlabp(12) ///
		mlabs(*.90) ///
		msymbol(circle)) ///
	(connected Bolivia_2024_2026 Año, ///
		mlabel(Bolivia_lbl_2024_2026) ///
		mlabposition(12) ///
		mlabsize(*.90) ///
		msymbol(circle)), ///
	xline(2011.5 2023.5, lpattern(dash) lcolor(gs10)) ///
	xlab(2005(1)2024, angle(45)) ///
	ylab(100000(40000)300000, format(%9.0fc)) ///
	xtitle("") ///
	ytitle("$eje3") ///
	note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
	legend(off)

	gr export "$gph/06_001_04_01.png", replace
	
* ------------------------------------------------------------
* 6.1.4.2. Departamental
* ------------------------------------------------------------
	
	import excel "$in/06_001_04.xlsx", sheet("_in1") firstrow clear

	set scheme white_viridis
	
	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"
	label values Departamento_order Departamento_order_lbl

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Valor_2005_2011 = Valor if Año <= 2011
	gen Valor_2012_2023 = Valor if inrange(Año, 2012, 2023)
	gen Valor_2024_2026 = Valor if Año >= 2024

	twoway ///
	(connected Valor_2005_2011 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
	(connected Valor_2012_2023 Año if Departamento == "Chuquisaca", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
	(connected Valor_2024_2026 Año if Departamento == "Chuquisaca", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
	(connected Valor_2005_2011 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
	(connected Valor_2012_2023 Año if Departamento == "LaPaz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
	(connected Valor_2024_2026 Año if Departamento == "LaPaz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
	(connected Valor_2005_2011 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
	(connected Valor_2012_2023 Año if Departamento == "Cochabamba", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
	(connected Valor_2024_2026 Año if Departamento == "Cochabamba", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
	(connected Valor_2005_2011 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
	(connected Valor_2012_2023 Año if Departamento == "Oruro", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
	(connected Valor_2024_2026 Año if Departamento == "Oruro", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
	(connected Valor_2005_2011 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
	(connected Valor_2012_2023 Año if Departamento == "Potosí", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
	(connected Valor_2024_2026 Año if Departamento == "Potosí", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
	(connected Valor_2005_2011 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
	(connected Valor_2012_2023 Año if Departamento == "Tarija", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
	(connected Valor_2024_2026 Año if Departamento == "Tarija", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
	(connected Valor_2005_2011 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
	(connected Valor_2012_2023 Año if Departamento == "SantaCruz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
	(connected Valor_2024_2026 Año if Departamento == "SantaCruz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
	(connected Valor_2005_2011 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
	(connected Valor_2012_2023 Año if Departamento == "Beni", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
	(connected Valor_2024_2026 Año if Departamento == "Beni", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
	(connected Valor_2005_2011 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
	(connected Valor_2012_2023 Año if Departamento == "Pando", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
	(connected Valor_2024_2026 Año if Departamento == "Pando", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)), ///
    by(Departamento_order, note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) legend(off)) ///
    xline(2011.5 2023.5, lpattern(dash) lcolor(gs10)) ///
    ylab(0(50000)100000, format(%9.0fc)) ///
    xlab(2005(1)2024, angle(90)) ///
    xtitle("") ///
    ytitle("$eje3") ///
    title("", box color(black) span)

	gr export "$gph/06_001_04_02.png", replace
	
* ============================================================
**# 6.1.5. Número de abortos
* Código: 06_001_05
* ============================================================

* ------------------------------------------------------------
* 6.1.5.1. Nacional
* ------------------------------------------------------------

	import excel "$in/06_001_05.xlsx", sheet("_in") firstrow clear

	set scheme white_tableau

	format %9.0fc Bolivia

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Bolivia_2005_2011 = Bolivia if Año <= 2011
	gen Bolivia_2012_2023 = Bolivia if inrange(Año, 2012, 2023)
	gen Bolivia_2024_2026 = Bolivia if Año >= 2024

	* Etiquetas de datos solo para el gráfico nacional
	gen Bolivia_lbl_2005_2011 = string(Bolivia_2005_2011, "%9.0fc")
	gen Bolivia_lbl_2012_2023 = string(Bolivia, "%15.0fc") if inrange(Año, 2012, 2023)
	gen Bolivia_lbl_2024_2026 = string(Bolivia, "%15.0fc") if Año >= 2024

	twoway ///
	(connected Bolivia_2005_2011 Año, ///
		mlabel(Bolivia_lbl_2005_2011) ///
		mlabp(12) ///
		mlabs(*.90) ///
		msymbol(circle)) ///
	(connected Bolivia_2012_2023 Año, ///
		mlabel(Bolivia_lbl_2012_2023) ///
		mlabp(12) ///
		mlabs(*.90) ///
		msymbol(circle)) ///
	(connected Bolivia_2024_2026 Año, ///
		mlabel(Bolivia_lbl_2024_2026) ///
		mlabposition(12) ///
		mlabsize(*.90) ///
		msymbol(circle)), ///
	xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
	xlab(2005(1)2024, angle(45)) ///
	ylab(1000(9000)46000, format(%9.0fc)) ///
	xtitle("") ///
	ytitle("$eje4") ///
	note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
	legend(off)

	gr export "$gph/06_001_05_01.png", replace

* ------------------------------------------------------------
* 6.1.5.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_001_05.xlsx", sheet("_in1") firstrow clear

	set scheme white_viridis
	
	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"
	label values Departamento_order Departamento_order_lbl

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Valor_2005_2011 = Valor if Año <= 2011
	gen Valor_2012_2023 = Valor if inrange(Año, 2012, 2023)
	gen Valor_2024_2026 = Valor if Año >= 2024

	twoway ///
	(connected Valor_2005_2011 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
	(connected Valor_2012_2023 Año if Departamento == "Chuquisaca", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
	(connected Valor_2024_2026 Año if Departamento == "Chuquisaca", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
	(connected Valor_2005_2011 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
	(connected Valor_2012_2023 Año if Departamento == "LaPaz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
	(connected Valor_2024_2026 Año if Departamento == "LaPaz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
	(connected Valor_2005_2011 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
	(connected Valor_2012_2023 Año if Departamento == "Cochabamba", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
	(connected Valor_2024_2026 Año if Departamento == "Cochabamba", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
	(connected Valor_2005_2011 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
	(connected Valor_2012_2023 Año if Departamento == "Oruro", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
	(connected Valor_2024_2026 Año if Departamento == "Oruro", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
	(connected Valor_2005_2011 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
	(connected Valor_2012_2023 Año if Departamento == "Potosí", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
	(connected Valor_2024_2026 Año if Departamento == "Potosí", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
	(connected Valor_2005_2011 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
	(connected Valor_2012_2023 Año if Departamento == "Tarija", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
	(connected Valor_2024_2026 Año if Departamento == "Tarija", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
	(connected Valor_2005_2011 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
	(connected Valor_2012_2023 Año if Departamento == "SantaCruz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
	(connected Valor_2024_2026 Año if Departamento == "SantaCruz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
	(connected Valor_2005_2011 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
	(connected Valor_2012_2023 Año if Departamento == "Beni", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
	(connected Valor_2024_2026 Año if Departamento == "Beni", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
	(connected Valor_2005_2011 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
	(connected Valor_2012_2023 Año if Departamento == "Pando", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
	(connected Valor_2024_2026 Año if Departamento == "Pando", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)), ///
    by(Departamento_order, note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) legend(off)) ///
    xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
    ylab(0(7500)15000, format(%9.0fc)) ///
    xlab(2005(1)2024, angle(90)) ///
    xtitle("") ///
    ytitle("$eje4") ///
    title("", box color(black) span)

	gr export "$gph/06_001_05_02.png", replace
	
* ============================================================
**# 6.1.6. Proporción de Partos y Abortos sobre total de Embarazos
* Código: 06_001_06
* ============================================================

* ------------------------------------------------------------
* 6.1.6.1. Nacional
* ------------------------------------------------------------

	import excel "$in/06_001_06.xlsx", sheet("_in") firstrow clear

	set scheme white_tableau

	graph bar (asis) Valor, ///
		over(Tipo) ///
		over(Año, label(angle(45))) ///
		stack asyvars percent ///
		ytitle("$eje1") ///
		b1title("") ///
		ylabel(0(20)100, format(%9.0f)) ///
		blabel(bar, position(center) size(small) format(%9.0f)) ///
		legend(label(1 "Abortos") label(2 "Partos") ///
			   position(6) cols(2) size(small)) ///
		note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span)

	gr export "$gph/06_001_06_01.png", replace

* ------------------------------------------------------------
* 6.1.6.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_001_06.xlsx", sheet("_in1") firstrow clear

	set scheme white_viridis

	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"
	label values Departamento_order Departamento_order_lbl

	capture graph drop g1
	capture graph drop g2

	* Panel 1: Chuquisaca a Tarija
	graph bar (asis) Valor if Departamento_order <= 6, ///
		over(Tipo) ///
		over(Año, label(nolabel)) ///
		stack asyvars ///
		ytitle("") ///
		ylabel(0(50)100, format(%9.0f)) ///
		by(Departamento_order, ///
			title("") ///
			col(3) ///
			note("") ///
			legend(off) ///
			xrescale) ///
		name(g1, replace)

	* Panel 2: Santa Cruz, Beni y Pando
	graph bar (asis) Valor if Departamento_order >= 7, ///
		over(Tipo) ///
		over(Año, label(angle(90))) ///
		stack asyvars ///
		ytitle("") ///
		ylabel(0(50)100, format(%9.0f)) ///
		legend(label(1 "Abortos") label(2 "Partos") cols(2)) ///
		by(Departamento_order, ///
			title("") ///
			col(3) ///
			note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
			xrescale) ///
		name(g2, replace)

	graph combine g1 g2, cols(1) imargin(0 0 0 0)

	graph save "06_001_06_02.gph", replace
	gr export "$gph/06_001_06_02.png", replace
	
* ============================================================
**# 6.1.7. Número de nacimientos
* Código: 06_001_07
* ============================================================

* ------------------------------------------------------------
* 6.1.7.1. Nacional
* ------------------------------------------------------------

	import excel "$in/06_001_07.xlsx", sheet("_in") firstrow clear
	
	set scheme white_tableau

	format %15.0fc Bolivia

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Bolivia_2005_2011 = Bolivia if Año <= 2011
	gen Bolivia_2012_2023 = Bolivia if inrange(Año, 2012, 2023)
	gen Bolivia_2024_2026 = Bolivia if Año >= 2024

	* Etiquetas de datos solo para el gráfico nacional.
	gen Bolivia_lbl_2005_2011 = string(Bolivia, "%15.0fc") if Año <= 2011
	gen Bolivia_lbl_2012_2023 = string(Bolivia, "%15.0fc") if inrange(Año, 2012, 2023)
	gen Bolivia_lbl_2024_2026 = string(Bolivia, "%15.0fc") if Año >= 2024

	twoway ///
		(connected Bolivia_2005_2011 Año, ///
			mlabel(Bolivia_lbl_2005_2011) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2012_2023 Año, ///
			mlabel(Bolivia_lbl_2012_2023) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2024_2026 Año, ///
			mlabel(Bolivia_lbl_2024_2026) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)), ///
		xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		xlab(2005(1)2024, angle(45)) ///
		xline(2011.5, lcolor(gs10) lpattern(dash)) ///
		ylab(100000(40000)300000, format(%15.0fc)) ///
		xtitle("") ///
		ytitle("$eje5") ///
		note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
		legend(off)

	gr export "$gph/06_001_07_01.png", replace

* ------------------------------------------------------------
* 6.1.7.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_001_07.xlsx", sheet("_in1") firstrow clear

	set scheme white_viridis
	
	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"

	label values Departamento_order Departamento_order_lbl

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024.
	gen Valor_2005_2011 = Valor if Año <= 2011
	gen Valor_2012_2023 = Valor if inrange(Año, 2012, 2023)
	gen Valor_2024_2026 = Valor if Año >= 2024

	twoway ///
		(connected Valor_2005_2011 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Chuquisaca", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Chuquisaca", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "LaPaz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "LaPaz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Cochabamba", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Cochabamba", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Oruro", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Oruro", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Potosí", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Potosí", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Tarija", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Tarija", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "SantaCruz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "SantaCruz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Beni", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Beni", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Pando", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Pando", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)), ///
		by(Departamento_order, ///
			note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
			legend(off)) ///
		xline(2011.5 2023.5, lcolor(gs10) lpattern(dash)) ///
		ylab(0(45000)90000, format(%15.0fc)) ///
		xlab(2005(1)2024, angle(90)) ///
		xtitle("") ///
		ytitle("$eje5") ///
		title("", box color(black) span)

	gr export "$gph/06_001_07_02.png", replace
	
* ============================================================
**# 6.1.8. Número de Hijos Nacidos Vivos
* Código: 06_001_08
* ============================================================

* ------------------------------------------------------------
* 6.1.8.1. Nacional
* ------------------------------------------------------------

	import excel "$in/06_001_08.xlsx", sheet("_in") firstrow clear
	
	set scheme white_tableau

	format %15.0fc Bolivia

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Bolivia_2005_2011 = Bolivia if Año <= 2011
	gen Bolivia_2012_2023 = Bolivia if inrange(Año, 2012, 2023)
	gen Bolivia_2024_2026 = Bolivia if Año >= 2024

	* Etiquetas de datos solo para el gráfico nacional.
	gen Bolivia_lbl_2005_2011 = string(Bolivia, "%15.0fc") if Año <= 2011
	gen Bolivia_lbl_2012_2023 = string(Bolivia, "%15.0fc") if inrange(Año, 2012, 2023)
	gen Bolivia_lbl_2024_2026 = string(Bolivia, "%15.0fc") if Año >= 2024

	twoway ///
		(connected Bolivia_2005_2011 Año, ///
			mlabel(Bolivia_lbl_2005_2011) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2012_2023 Año, ///
			mlabel(Bolivia_lbl_2012_2023) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2024_2026 Año, ///
			mlabel(Bolivia_lbl_2024_2026) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)), ///
		xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		xlab(2005(1)2024, angle(45)) ///
		ylab(100000(40000)300000, format(%15.0fc)) ///
		xtitle("") ///
		ytitle("$eje6") ///
		note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
		legend(off)

	gr export "$gph/06_001_08_01.png", replace

* ------------------------------------------------------------
* 6.1.8.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_001_08.xlsx", sheet("_in1") firstrow clear

	set scheme white_viridis
	
	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"

	label values Departamento_order Departamento_order_lbl

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024.
	gen Valor_2005_2011 = Valor if Año <= 2011
	gen Valor_2012_2023 = Valor if inrange(Año, 2012, 2023)
	gen Valor_2024_2026 = Valor if Año >= 2024

	twoway ///
		(connected Valor_2005_2011 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Chuquisaca", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Chuquisaca", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "LaPaz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "LaPaz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Cochabamba", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Cochabamba", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Oruro", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Oruro", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Potosí", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Potosí", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Tarija", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Tarija", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "SantaCruz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "SantaCruz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Beni", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Beni", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Pando", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Pando", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)), ///
		by(Departamento_order, ///
			note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
			legend(off)) ///
		xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		ylab(0(45000)90000, format(%15.0fc)) ///
		xlab(2005(1)2024, angle(90)) ///
		xtitle("") ///
		ytitle("$eje6") ///
		title("", box color(black) span)

	gr export "$gph/06_001_08_02.png", replace
	
* ============================================================
**# 6.1.9. Número de Hijos Nacidos Muertos
* Código: 06_001_09
* ============================================================

* ------------------------------------------------------------
* 6.1.9.1. Nacional
* ------------------------------------------------------------

	import excel "$in/06_001_09.xlsx", sheet("_in") firstrow clear
	
	set scheme white_tableau

	format %15.0fc Bolivia

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Bolivia_2005_2011 = Bolivia if Año <= 2011
	gen Bolivia_2012_2023 = Bolivia if inrange(Año, 2012, 2023)
	gen Bolivia_2024_2026 = Bolivia if Año >= 2024

	* Etiquetas de datos solo para el gráfico nacional.
	gen Bolivia_lbl_2005_2011 = string(Bolivia, "%15.0fc") if Año <= 2011
	gen Bolivia_lbl_2012_2023 = string(Bolivia, "%15.0fc") if inrange(Año, 2012, 2023)
	gen Bolivia_lbl_2024_2026 = string(Bolivia, "%15.0fc") if Año >= 2024

	twoway ///
		(connected Bolivia_2005_2011 Año, ///
			mlabel(Bolivia_lbl_2005_2011) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2012_2023 Año, ///
			mlabel(Bolivia_lbl_2012_2023) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2024_2026 Año, ///
			mlabel(Bolivia_lbl_2024_2026) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)), ///
		xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		xlab(2005(1)2024, angle(45)) ///
		ylab(0(1000)5000, format(%15.0fc)) ///
		xtitle("") ///
		ytitle("$eje7") ///
		note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
		legend(off)

	gr export "$gph/06_001_09_01.png", replace

* ------------------------------------------------------------
* 6.1.9.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_001_09.xlsx", sheet("_in1") firstrow clear

	set scheme white_viridis
	
	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"

	label values Departamento_order Departamento_order_lbl

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Valor_2005_2011 = Valor if Año <= 2011
	gen Valor_2012_2023 = Valor if inrange(Año, 2012, 2023)
	gen Valor_2024_2026 = Valor if Año >= 2024

	twoway ///
		(connected Valor_2005_2011 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Chuquisaca", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Chuquisaca", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "LaPaz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "LaPaz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Cochabamba", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Cochabamba", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Oruro", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Oruro", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Potosí", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Potosí", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Tarija", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Tarija", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "SantaCruz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "SantaCruz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Beni", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Beni", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Pando", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Pando", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)), ///
		by(Departamento_order, ///
			note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
			legend(off)) ///
		xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		ylab(0(1000)2000, format(%15.0fc)) ///
		xlab(2005(1)2024, angle(90)) ///
		xtitle("") ///
		ytitle("$eje7") ///
		title("", box color(black) span)

	gr export "$gph/06_001_09_02.png", replace
	
* ============================================================
**# 6.1.10. Proporción de HNV e HNM sobre total de Nacimientos
* Código: 06_001_10
* ============================================================

* ------------------------------------------------------------
* 6.1.10.1. Nacional
* ------------------------------------------------------------

	import excel "$in/06_001_10.xlsx", sheet("_in") firstrow clear
	
	set scheme white_tableau

	graph bar (asis) Valor, ///
		over(Tipo) ///
		over(Año, label(angle(45))) ///
		stack asyvars percent ///
		ytitle("$eje1") ///
		b1title("") ///
		ylabel(0(20)100, format(%9.0f)) ///
		blabel(bar, position(center) size(small) format(%9.0f)) ///
		legend(label(1 "Hijos Nacidos Muertos") ///
			   label(2 "Hijos Nacidos Vivos") ///
			   position(6) cols(2) size(small)) ///
		note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span)

	gr export "$gph/06_001_10_01.png", replace

* ------------------------------------------------------------
* 6.1.10.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_001_10.xlsx", sheet("_in1") firstrow clear
	
	set scheme white_viridis
	
	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"
	label values Departamento_order Departamento_order_lbl

	capture graph drop g1
	capture graph drop g2

	* Panel 1: Chuquisaca a Tarija
	graph bar (asis) Valor if Departamento_order <= 6, ///
		over(Tipo) ///
		over(Año, label(nolabel)) ///
		stack asyvars ///
		ytitle("") ///
		b1title("") ///
		ylabel(0(50)100, format(%9.0f)) ///
		by(Departamento_order, ///
			title("") ///
			col(3) ///
			note("") ///
			legend(off) ///
			xrescale) ///
		name(g1, replace)

	* Panel 2: Santa Cruz, Beni y Pando
	graph bar (asis) Valor if Departamento_order >= 7, ///
		over(Tipo) ///
		over(Año, label(angle(90))) ///
		stack asyvars ///
		ytitle("") ///
		b1title("") ///
		ylabel(0(50)100, format(%9.0f)) ///
		legend(label(1 "Hijos Nacidos Muertos") ///
			   label(2 "Hijos Nacidos Vivos") ///
			   cols(2)) ///
		by(Departamento_order, ///
			title("") ///
			col(3) ///
			note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
			xrescale) ///
		name(g2, replace)

	graph combine g1 g2, cols(1) imargin(0 0 0 0)

	graph save "$gph/06_001_10_02.gph", replace
	gr export "$gph/06_001_10_02.png", replace
	
* ============================================================
**# 6.2.1. Tasa de Aborto General (TAG)
* Código: 06_002_01
* ============================================================

* ------------------------------------------------------------
* 6.2.1.1. Nacional
* ------------------------------------------------------------

	import excel "$in/06_002_01.xlsx", sheet("_in") firstrow clear
	
	set scheme white_tableau

	format %4.2fc Bolivia

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Bolivia_2005_2011 = Bolivia if Año <= 2011
	gen Bolivia_2012_2023 = Bolivia if inrange(Año, 2012, 2023)
	gen Bolivia_2024_2026 = Bolivia if Año >= 2024

	* Etiquetas de datos solo para el gráfico nacional.
	gen Bolivia_lbl_2005_2011 = string(Bolivia, "%15.0fc") if Año <= 2011
	gen Bolivia_lbl_2012_2023 = string(Bolivia, "%15.0fc") if inrange(Año, 2012, 2023)
	gen Bolivia_lbl_2024_2026 = string(Bolivia, "%15.0fc") if Año >= 2024

	twoway ///
		(connected Bolivia_2005_2011 Año, ///
			mlabel(Bolivia_lbl_2005_2011) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2012_2023 Año, ///
			mlabel(Bolivia_lbl_2012_2023) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2024_2026 Año, ///
			mlabel(Bolivia_lbl_2024_2026) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)), ///
		xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		xlab(2005(1)2024, angle(45)) ///
		ylab(0(4)20, format(%9.0fc)) ///
		xtitle("") ///
		ytitle("$eje8") ///
		note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
		legend(off)

	gr export "$gph/06_002_01_01.png", replace

* ------------------------------------------------------------
* 6.2.1.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_002_01.xlsx", sheet("_in1") firstrow clear

	set scheme white_viridis
	
	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"

	label values Departamento_order Departamento_order_lbl

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Valor_2005_2011 = Valor if Año <= 2011
	gen Valor_2012_2023 = Valor if inrange(Año, 2012, 2023)
	gen Valor_2024_2026 = Valor if Año >= 2024

	twoway ///
		(connected Valor_2005_2011 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Chuquisaca", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Chuquisaca", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "LaPaz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "LaPaz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Cochabamba", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Cochabamba", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Oruro", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Oruro", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Potosí", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Potosí", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Tarija", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Tarija", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "SantaCruz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "SantaCruz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Beni", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Beni", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Pando", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Pando", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)), ///
		by(Departamento_order, ///
			note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
			legend(off)) ///
		xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		ylab(0(20)40, format(%9.0fc)) ///
		xlab(2005(1)2024, angle(90)) ///
		xtitle("") ///
		ytitle("$eje8") ///
		title("", box color(black) span)

	gr export "$gph/06_002_01_02.png", replace	
	
* ============================================================
**# 6.2.2. Tasa de Fecundidad General (TFG)
* Código: 06_002_02
* ============================================================

* ------------------------------------------------------------
* 6.2.2.1. Nacional
* ------------------------------------------------------------

	import excel "$in/06_002_02.xlsx", sheet("_in") firstrow clear
	
	set scheme white_tableau

	format %4.2fc Bolivia

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Bolivia_2005_2011 = Bolivia if Año <= 2011
	gen Bolivia_2012_2023 = Bolivia if inrange(Año, 2012, 2023)
	gen Bolivia_2024_2026 = Bolivia if Año >= 2024

	* Etiquetas de datos solo para el gráfico nacional.
	gen Bolivia_lbl_2005_2011 = string(Bolivia, "%15.0fc") if Año <= 2011
	gen Bolivia_lbl_2012_2023 = string(Bolivia, "%15.0fc") if inrange(Año, 2012, 2023)
	gen Bolivia_lbl_2024_2026 = string(Bolivia, "%15.0fc") if Año >= 2024

	twoway ///
		(connected Bolivia_2005_2011 Año, ///
			mlabel(Bolivia_lbl_2005_2011) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2012_2023 Año, ///
			mlabel(Bolivia_lbl_2012_2023) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2024_2026 Año, ///
			mlabel(Bolivia_lbl_2024_2026) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)), ///
		xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		xlab(2005(1)2024, angle(45)) ///
		ylab(0(30)150, format(%9.0fc)) ///
		xtitle("") ///
		ytitle("$eje9") ///
		note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
		legend(off)

	gr export "$gph/06_002_02_01.png", replace

* ------------------------------------------------------------
* 6.2.2.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_002_02.xlsx", sheet("_in1") firstrow clear
	
	set scheme white_viridis
	
	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"

	label values Departamento_order Departamento_order_lbl

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Valor_2005_2011 = Valor if Año <= 2011
	gen Valor_2012_2023 = Valor if inrange(Año, 2012, 2023)
	gen Valor_2024_2026 = Valor if Año >= 2024

	twoway ///
		(connected Valor_2005_2011 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Chuquisaca", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Chuquisaca", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "LaPaz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "LaPaz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Cochabamba", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Cochabamba", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Oruro", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Oruro", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Potosí", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Potosí", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Tarija", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Tarija", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "SantaCruz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "SantaCruz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Beni", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Beni", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Pando", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Pando", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)), ///
		by(Departamento_order, ///
			note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
			legend(off)) ///
		xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		ylab(0(90)180, format(%9.0fc)) ///
		xlab(2005(1)2024, angle(90)) ///
		xtitle("") ///
		ytitle("$eje9") ///
		title("", box color(black) span)

	gr export "$gph/06_002_02_02.png", replace

* ============================================================
**# 6.2.3. Tasa de Mortalidad Fetal (TMF)
* Código: 06_002_03
* ============================================================

* ------------------------------------------------------------
* 6.2.3.1. Nacional
* ------------------------------------------------------------

	import excel "$in/06_002_03.xlsx", sheet("_in") firstrow clear
	
	set scheme white_tableau

	format %4.2fc Bolivia

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Bolivia_2005_2011 = Bolivia if Año <= 2011
	gen Bolivia_2012_2023 = Bolivia if inrange(Año, 2012, 2023)
	gen Bolivia_2024_2026 = Bolivia if Año >= 2024

	* Etiquetas de datos solo para el gráfico nacional.
	gen Bolivia_lbl_2005_2011 = string(Bolivia, "%15.0fc") if Año <= 2011
	gen Bolivia_lbl_2012_2023 = string(Bolivia, "%15.0fc") if inrange(Año, 2012, 2023)
	gen Bolivia_lbl_2024_2026 = string(Bolivia, "%15.0fc") if Año >= 2024

	twoway ///
		(connected Bolivia_2005_2011 Año, ///
			mlabel(Bolivia_lbl_2005_2011) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2012_2023 Año, ///
			mlabel(Bolivia_lbl_2012_2023) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2024_2026 Año, ///
			mlabel(Bolivia_lbl_2024_2026) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)), ///
		xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		xlab(2005(1)2024, angle(45)) ///
		ylab(0(4)20, format(%9.0fc)) ///
		xtitle("") ///
		ytitle("$eje10") ///
		note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
		legend(off)

	gr export "$gph/06_002_03_01.png", replace

* ------------------------------------------------------------
* 6.2.3.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_002_03.xlsx", sheet("_in1") firstrow clear
	
	set scheme white_viridis
	
	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"

	label values Departamento_order Departamento_order_lbl

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Valor_2005_2011 = Valor if Año <= 2011
	gen Valor_2012_2023 = Valor if inrange(Año, 2012, 2023)
	gen Valor_2024_2026 = Valor if Año >= 2024

	twoway ///
		(connected Valor_2005_2011 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Chuquisaca", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Chuquisaca", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "LaPaz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "LaPaz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Cochabamba", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Cochabamba", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Oruro", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Oruro", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Potosí", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Potosí", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Tarija", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Tarija", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "SantaCruz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "SantaCruz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Beni", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Beni", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Pando", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Pando", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)), ///
		by(Departamento_order, ///
			note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
			legend(off)) ///
		xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		ylab(0(15)30, format(%9.0fc)) ///
		xlab(2005(1)2024, angle(90)) ///
		xtitle("") ///
		ytitle("$eje10") ///
		title("", box color(black) span)

	gr export "$gph/06_002_03_02.png", replace
	
* ============================================================
**# 6.2.4. Tasa Estimada de Embarazo (TEE)
* Código: 06_002_04
* ============================================================

* ------------------------------------------------------------
* 6.2.4.1. Nacional
* ------------------------------------------------------------

	import excel "$in/06_002_04.xlsx", sheet("_in") firstrow clear
	
	set scheme white_tableau

	format %4.2fc Bolivia

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Bolivia_2005_2011 = Bolivia if Año <= 2011
	gen Bolivia_2012_2023 = Bolivia if inrange(Año, 2012, 2023)
	gen Bolivia_2024_2026 = Bolivia if Año >= 2024

	* Etiquetas de datos solo para el gráfico nacional.
	gen Bolivia_lbl_2005_2011 = string(Bolivia, "%15.0fc") if Año <= 2011
	gen Bolivia_lbl_2012_2023 = string(Bolivia, "%15.0fc") if inrange(Año, 2012, 2023)
	gen Bolivia_lbl_2024_2026 = string(Bolivia, "%15.0fc") if Año >= 2024

	twoway ///
		(connected Bolivia_2005_2011 Año, ///
			mlabel(Bolivia_lbl_2005_2011) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2012_2023 Año, ///
			mlabel(Bolivia_lbl_2012_2023) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2024_2026 Año, ///
			mlabel(Bolivia_lbl_2024_2026) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)), ///
		xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		xlab(2005(1)2024, angle(45)) ///
		ylab(0(30)150, format(%9.0fc)) ///
		xtitle("") ///
		ytitle("$eje11") ///
		note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
		legend(off)

	gr export "$gph/06_002_04_01.png", replace

* ------------------------------------------------------------
* 6.2.4.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_002_04.xlsx", sheet("_in1") firstrow clear
	
	set scheme white_viridis
	
	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"

	label values Departamento_order Departamento_order_lbl

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Valor_2005_2011 = Valor if Año <= 2011
	gen Valor_2012_2023 = Valor if inrange(Año, 2012, 2023)
	gen Valor_2024_2026 = Valor if Año >= 2024

	twoway ///
		(connected Valor_2005_2011 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Chuquisaca", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Chuquisaca", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "LaPaz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "LaPaz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Cochabamba", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Cochabamba", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Oruro", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Oruro", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Potosí", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Potosí", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Tarija", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Tarija", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "SantaCruz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "SantaCruz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Beni", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Beni", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Pando", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Pando", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)), ///
		by(Departamento_order, ///
			note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
			legend(off)) ///
		xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		ylab(0(100)200, format(%9.0fc)) ///
		xlab(2005(1)2024, angle(90)) ///
		xtitle("") ///
		ytitle("$eje11") ///
		title("", box color(black) span)

	gr export "$gph/06_002_04_02.png", replace

* ============================================================
**# 6.3.1. Tasa Bruta de Natalidad (TBN)
* Código: 06_003_01
* ============================================================

* ------------------------------------------------------------
* 6.3.1.1. Nacional
* ------------------------------------------------------------

	import excel "$in/06_003_01.xlsx", sheet("_in") firstrow clear
	
	set scheme white_tableau

	format %4.2fc Bolivia

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Bolivia_2005_2011 = Bolivia if Año <= 2011
	gen Bolivia_2012_2023 = Bolivia if inrange(Año, 2012, 2023)
	gen Bolivia_2024_2026 = Bolivia if Año >= 2024

	* Etiquetas de datos solo para el gráfico nacional.
	gen Bolivia_lbl_2005_2011 = string(Bolivia, "%15.0fc") if Año <= 2011
	gen Bolivia_lbl_2012_2023 = string(Bolivia, "%15.0fc") if inrange(Año, 2012, 2023)
	gen Bolivia_lbl_2024_2026 = string(Bolivia, "%15.0fc") if Año >= 2024

	twoway ///
		(connected Bolivia_2005_2011 Año, ///
			mlabel(Bolivia_lbl_2005_2011) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2012_2023 Año, ///
			mlabel(Bolivia_lbl_2012_2023) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2024_2026 Año, ///
			mlabel(Bolivia_lbl_2024_2026) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)), ///
		xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		xlab(2005(1)2024, angle(45)) ///
		ylab(0(7)35, format(%9.0fc)) ///
		xtitle("") ///
		ytitle("$eje12") ///
		note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
		legend(off)

	gr export "$gph/06_003_01_01.png", replace

* ------------------------------------------------------------
* 6.3.1.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_003_01.xlsx", sheet("_in1") firstrow clear
	
	set scheme white_viridis
	
	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"

	label values Departamento_order Departamento_order_lbl

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Valor_2005_2011 = Valor if Año <= 2011
	gen Valor_2012_2023 = Valor if inrange(Año, 2012, 2023)
	gen Valor_2024_2026 = Valor if Año >= 2024

	twoway ///
		(connected Valor_2005_2011 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Chuquisaca", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Chuquisaca", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "LaPaz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "LaPaz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Cochabamba", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Cochabamba", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Oruro", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Oruro", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Potosí", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Potosí", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Tarija", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Tarija", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "SantaCruz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "SantaCruz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Beni", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Beni", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Pando", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Pando", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)), ///
		by(Departamento_order, ///
			note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
			legend(off)) ///
		xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		ylab(0(20)40, format(%9.0fc)) ///
		xlab(2005(1)2024, angle(90)) ///
		xtitle("") ///
		ytitle("$eje12") ///
		title("", box color(black) span)

	gr export "$gph/06_003_01_02.png", replace
	
* ============================================================
**# 6.3.2. Índice de Masculinidad (IM)
* Código: 06_003_02
* ============================================================

* ------------------------------------------------------------
* 6.3.2.1. Nacional
* ------------------------------------------------------------

	import excel "$in/06_003_02.xlsx", sheet("_in") firstrow clear
	
	set scheme white_tableau
	
	format %4.2fc Bolivia

	* Separar dos tramos:
	* 2012-2023 y 2024 en adelante.
	gen Bolivia_2012_2023 = Bolivia if inrange(Año, 2012, 2023)
	gen Bolivia_2024_2026 = Bolivia if Año >= 2024

	* Etiquetas de datos solo para el gráfico nacional.
	gen Bolivia_lbl_2012_2023 = string(Bolivia, "%4.2fc") if inrange(Año, 2012, 2023)
	gen Bolivia_lbl_2024_2026 = string(Bolivia, "%4.2fc") if Año >= 2024
	
	twoway ///
		(connected Bolivia_2012_2023 Año, ///
			mlabel(Bolivia_lbl_2012_2023) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2024_2026 Año, ///
			mlabel(Bolivia_lbl_2024_2026) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)), ///
		xline(2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		xlab(2012(1)2024, angle(45)) ///
		ylab(95(2)105, format(%9.0fc)) ///
		xtitle("") ///
		ytitle("$eje13") ///
		note("$nota11" "$nota21" "$fuente2", size(vsmall) span) ///
		legend(off)

	gr export "$gph/06_003_02_01.png", replace

* ------------------------------------------------------------
* 6.3.2.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_003_02.xlsx", sheet("_in1") firstrow clear
	
	set scheme white_viridis
	
	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"

	label values Departamento_order Departamento_order_lbl

	* Separar los dos tramos:
	* 2012-2023 y 2024 en adelante.
	gen Valor_2012_2023 = Valor if inrange(Año, 2012, 2023)
	gen Valor_2024_2026 = Valor if Año >= 2024

	twoway ///
		(connected Valor_2012_2023 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Chuquisaca", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "LaPaz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Cochabamba", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Oruro", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Potosí", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Tarija", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "SantaCruz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Beni", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Pando", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)), ///
		by(Departamento_order, ///
			note("$nota11" "$nota21" "$fuente2", size(vsmall) span) ///
			legend(off)) ///
		ylab(90(20)130, format(%9.0fc)) ///
		xlab(2012(1)2024, angle(45)) ///
		xtitle("") ///
		ytitle("$eje13") ///
		title("", box color(black) span)

	gr export "$gph/06_003_02_02.png", replace
	
* ============================================================
**# 6.3.3. Tasa de Crecimiento Poblacional (TCP)
* Código: 06_003_03
* ============================================================

* ------------------------------------------------------------
* 6.3.3.1. Nacional
* ------------------------------------------------------------

	import excel "$in/06_003_03.xlsx", sheet("_in") firstrow clear
	
	set scheme white_tableau

	format %4.2fc Bolivia

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Bolivia_2005_2011 = Bolivia if Año <= 2011
	gen Bolivia_2012_2023 = Bolivia if inrange(Año, 2012, 2023)
	gen Bolivia_2024_2026 = Bolivia if Año >= 2024

	* Etiquetas de datos solo para el gráfico nacional.
	gen Bolivia_lbl_2005_2011 = string(Bolivia, "%4.2fc") if Año <= 2011
	gen Bolivia_lbl_2012_2023 = string(Bolivia, "%4.2fc") if inrange(Año, 2012, 2023)
	gen Bolivia_lbl_2024_2026 = string(Bolivia, "%4.2fc") if Año >= 2024

	twoway ///
		(connected Bolivia_2005_2011 Año, ///
			mlabel(Bolivia_lbl_2005_2011) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2012_2023 Año, ///
			mlabel(Bolivia_lbl_2012_2023) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2024_2026 Año, ///
			mlabel(Bolivia_lbl_2024_2026) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)), ///
		xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		xlab(2005(1)2024, angle(45)) ///
		ylab(0(1)5, format(%9.2fc)) ///
		xtitle("") ///
		ytitle("$eje29") ///
		note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
		legend(off)

	gr export "$gph/06_003_03_01.png", replace

* ------------------------------------------------------------
* 6.3.3.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_003_03.xlsx", sheet("_in1") firstrow clear
	
	set scheme white_viridis
	
	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"

	label values Departamento_order Departamento_order_lbl

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Valor_2005_2011 = Valor if Año <= 2011
	gen Valor_2012_2023 = Valor if inrange(Año, 2012, 2023)
	gen Valor_2024_2026 = Valor if Año >= 2024

	twoway ///
		(connected Valor_2005_2011 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Chuquisaca", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Chuquisaca", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "LaPaz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "LaPaz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Cochabamba", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Cochabamba", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Oruro", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Oruro", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Potosí", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Potosí", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Tarija", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Tarija", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "SantaCruz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "SantaCruz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Beni", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Beni", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Pando", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Pando", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)), ///
		by(Departamento_order, ///
			note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
			legend(off)) ///
		xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		ylab(0(2.5)5, format(%9.2fc)) ///
		xlab(2005(1)2024, angle(90)) ///
		xtitle("") ///
		ytitle("$eje29") ///
		title("", box color(black) span)

	gr export "$gph/06_003_03_02.png", replace
	
* ============================================================
**# 6.3.4. Distribución de la Población por Grupos Etarios (DPGE)
* Código: 06_003_04
* ============================================================

* ------------------------------------------------------------
* 6.3.4.1. Nacional
* ------------------------------------------------------------

	import excel "$in/06_003_04.xlsx", sheet("_in") firstrow clear

	set scheme white_tableau

	graph bar (asis) Valor, ///
		over(Tipo) ///
		over(Año, label(angle(45))) ///
		stack asyvars percent ///
		ytitle("$eje1") ///
		b1title("") ///
		ylabel(0(20)100, format(%9.0f)) ///
		blabel(bar, position(center) size(small) format(%9.0f)) ///
		legend(label(1 "0-4 años") ///
			   label(2 "5-14 años") ///
			   label(3 "15-64 años") ///
			   label(4 "65 años y más") ///
			   position(6) cols(4) size(small)) ///
		note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span)

	gr export "$gph/06_003_04_01.png", replace

* ------------------------------------------------------------
* 6.3.4.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_003_04.xlsx", sheet("_in1") firstrow clear
	
	set scheme white_viridis
	
	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"

	label values Departamento_order Departamento_order_lbl

	capture graph drop g1
	capture graph drop g2

	* Panel 1: Chuquisaca a Tarija
	graph bar (asis) Valor if Departamento_order <= 6, ///
		over(Tipo) ///
		over(Año, label(nolabel)) ///
		stack asyvars ///
		ytitle("") ///
		b1title("") ///
		ylabel(0(50)100, format(%9.0f)) ///
		by(Departamento_order, ///
			title("") ///
			col(3) ///
			note("") ///
			legend(off) ///
			xrescale) ///
		name(g1, replace)

	* Panel 2: Santa Cruz, Beni y Pando
	graph bar (asis) Valor if Departamento_order >= 7, ///
		over(Tipo) ///
		over(Año, label(angle(90))) ///
		stack asyvars ///
		ytitle("") ///
		b1title("") ///
		ylabel(0(50)100, format(%9.0f)) ///
		legend(label(1 "0-4 años") ///
			   label(2 "5-14 años") ///
			   label(3 "15-64 años") ///
			   label(4 "65 años y más") ///
			   cols(4)) ///
		by(Departamento_order, ///
			title("") ///
			col(3) ///
			note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
			xrescale) ///
		name(g2, replace)

	graph combine g1 g2, cols(1) imargin(0 0 0 0)

	graph save "$gph/06_003_04_02.gph", replace
	gr export "$gph/06_003_04_02.png", replace
	
* ============================================================
**# 6.3.5. Índice de Dependencia Potencial (IDP)
* Código: 06_003_05
* ============================================================

* ------------------------------------------------------------
* 6.3.5.1. Nacional
* ------------------------------------------------------------

	import excel "$in/06_003_05.xlsx", sheet("_in") firstrow clear
	
	set scheme white_tableau

	format %4.2fc Bolivia

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Bolivia_2005_2011 = Bolivia if Año <= 2011
	gen Bolivia_2012_2023 = Bolivia if inrange(Año, 2012, 2023)
	gen Bolivia_2024_2026 = Bolivia if Año >= 2024

	* Etiquetas de datos solo para el gráfico nacional.
	gen Bolivia_lbl_2005_2011 = string(Bolivia, "%4.2fc") if Año <= 2011
	gen Bolivia_lbl_2012_2023 = string(Bolivia, "%4.2fc") if inrange(Año, 2012, 2023)
	gen Bolivia_lbl_2024_2026 = string(Bolivia, "%4.2fc") if Año >= 2024

	twoway ///
		(connected Bolivia_2005_2011 Año, ///
			mlabel(Bolivia_lbl_2005_2011) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2012_2023 Año, ///
			mlabel(Bolivia_lbl_2012_2023) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2024_2026 Año, ///
			mlabel(Bolivia_lbl_2024_2026) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)), ///
		xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		xlab(2005(1)2024, angle(45)) ///
		ylab(40(10)90, format(%9.0fc)) ///
		xtitle("") ///
		ytitle("$eje14") ///
		note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
		legend(off)

	gr export "$gph/06_003_05_01.png", replace

* ------------------------------------------------------------
* 6.3.5.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_003_05.xlsx", sheet("_in1") firstrow clear

	set scheme white_viridis
	
	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"

	label values Departamento_order Departamento_order_lbl

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Valor_2005_2011 = Valor if Año <= 2011
	gen Valor_2012_2023 = Valor if inrange(Año, 2012, 2023)
	gen Valor_2024_2026 = Valor if Año >= 2024

	twoway ///
		(connected Valor_2005_2011 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Chuquisaca", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Chuquisaca", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "LaPaz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "LaPaz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Cochabamba", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Cochabamba", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Oruro", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Oruro", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Potosí", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Potosí", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Tarija", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Tarija", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "SantaCruz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "SantaCruz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Beni", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Beni", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Pando", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Pando", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)), ///
		by(Departamento_order, ///
			note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
			legend(off)) ///
		xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		ylab(40(30)100, format(%9.0fc)) ///
		xlab(2005(1)2024, angle(90)) ///
		xtitle("") ///
		ytitle("$eje14") ///
		title("", box color(black) span)

	gr export "$gph/06_003_05_02.png", replace
	
* ============================================================
**# 6.3.6. Índice de Envejecimiento (IE)
* Código: 06_003_06
* ============================================================

* ------------------------------------------------------------
* 6.3.6.1. Nacional
* ------------------------------------------------------------

	import excel "$in/06_003_06.xlsx", sheet("_in") firstrow clear
	
	set scheme white_tableau

	format %4.2fc Bolivia

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Bolivia_2005_2011 = Bolivia if Año <= 2011
	gen Bolivia_2012_2023 = Bolivia if inrange(Año, 2012, 2023)
	gen Bolivia_2024_2026 = Bolivia if Año >= 2024

	* Etiquetas de datos solo para el gráfico nacional.
	gen Bolivia_lbl_2005_2011 = string(Bolivia, "%4.2fc") if Año <= 2011
	gen Bolivia_lbl_2012_2023 = string(Bolivia, "%4.2fc") if inrange(Año, 2012, 2023)
	gen Bolivia_lbl_2024_2026 = string(Bolivia, "%4.2fc") if Año >= 2024

	twoway ///
		(connected Bolivia_2005_2011 Año, ///
			mlabel(Bolivia_lbl_2005_2011) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2012_2023 Año, ///
			mlabel(Bolivia_lbl_2012_2023) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2024_2026 Año, ///
			mlabel(Bolivia_lbl_2024_2026) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)), ///
		xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		xlab(2005(1)2024, angle(45)) ///
		ylab(0(7)35, format(%9.0fc)) ///
		xtitle("") ///
		ytitle("$eje15") ///
		note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
		legend(off)

	gr export "$gph/06_003_06_01.png", replace

* ------------------------------------------------------------
* 6.3.6.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_003_06.xlsx", sheet("_in1") firstrow clear

	set scheme white_viridis
	
	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"

	label values Departamento_order Departamento_order_lbl

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Valor_2005_2011 = Valor if Año <= 2011
	gen Valor_2012_2023 = Valor if inrange(Año, 2012, 2023)
	gen Valor_2024_2026 = Valor if Año >= 2024

	twoway ///
		(connected Valor_2005_2011 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Chuquisaca", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Chuquisaca", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "LaPaz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "LaPaz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Cochabamba", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Cochabamba", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Oruro", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Oruro", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Potosí", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Potosí", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Tarija", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Tarija", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "SantaCruz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "SantaCruz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Beni", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Beni", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Pando", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Pando", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)), ///
		by(Departamento_order, ///
			note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
			legend(off)) ///
		xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		ylab(0(25)50, format(%9.0fc)) ///
		xlab(2005(1)2024, angle(90)) ///
		xtitle("") ///
		ytitle("$eje15") ///
		title("", box color(black) span)

	gr export "$gph/06_003_06_02.png", replace
	
* ============================================================
**# 6.3.7. Índice de Dependencia Juvenil (IDJ)
* Código: 06_003_07
* ============================================================

* ------------------------------------------------------------
* 6.3.7.1. Nacional
* ------------------------------------------------------------

	import excel "$in/06_003_07.xlsx", sheet("_in") firstrow clear
	
	set scheme white_tableau

	format %4.2fc Bolivia

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Bolivia_2005_2011 = Bolivia if Año <= 2011
	gen Bolivia_2012_2023 = Bolivia if inrange(Año, 2012, 2023)
	gen Bolivia_2024_2026 = Bolivia if Año >= 2024

	* Etiquetas de datos solo para el gráfico nacional.
	gen Bolivia_lbl_2005_2011 = string(Bolivia, "%4.2fc") if Año <= 2011
	gen Bolivia_lbl_2012_2023 = string(Bolivia, "%4.2fc") if inrange(Año, 2012, 2023)
	gen Bolivia_lbl_2024_2026 = string(Bolivia, "%4.2fc") if Año >= 2024

	twoway ///
		(connected Bolivia_2005_2011 Año, ///
			mlabel(Bolivia_lbl_2005_2011) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2012_2023 Año, ///
			mlabel(Bolivia_lbl_2012_2023) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2024_2026 Año, ///
			mlabel(Bolivia_lbl_2024_2026) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)), ///
		xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		xlab(2005(1)2024, angle(45)) ///
		ylab(30(10)80, format(%9.0fc)) ///
		xtitle("") ///
		ytitle("$eje16") ///
		note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
		legend(off)

	gr export "$gph/06_003_07_01.png", replace

* ------------------------------------------------------------
* 6.3.7.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_003_07.xlsx", sheet("_in1") firstrow clear

	set scheme white_viridis
	
	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"

	label values Departamento_order Departamento_order_lbl

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Valor_2005_2011 = Valor if Año <= 2011
	gen Valor_2012_2023 = Valor if inrange(Año, 2012, 2023)
	gen Valor_2024_2026 = Valor if Año >= 2024

	twoway ///
		(connected Valor_2005_2011 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Chuquisaca", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Chuquisaca", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "LaPaz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "LaPaz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Cochabamba", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Cochabamba", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Oruro", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Oruro", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Potosí", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Potosí", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Tarija", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Tarija", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "SantaCruz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "SantaCruz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Beni", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Beni", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Pando", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Pando", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)), ///
		by(Departamento_order, ///
			note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
			legend(off)) ///
		xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		ylab(20(35)90, format(%9.0fc)) ///
		xlab(2005(1)2024, angle(90)) ///
		xtitle("") ///
		ytitle("$eje16") ///
		title("", box color(black) span)

	gr export "$gph/06_003_07_02.png", replace
	
* ============================================================
**# 6.3.8. Índice de Dependencia Senil (IDS)
* Código: 06_003_08
* ============================================================

* ------------------------------------------------------------
* 6.3.8.1. Nacional
* ------------------------------------------------------------

	import excel "$in/06_003_08.xlsx", sheet("_in") firstrow clear
	
	set scheme white_tableau

	format %4.2fc Bolivia

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Bolivia_2005_2011 = Bolivia if Año <= 2011
	gen Bolivia_2012_2023 = Bolivia if inrange(Año, 2012, 2023)
	gen Bolivia_2024_2026 = Bolivia if Año >= 2024

	* Etiquetas de datos solo para el gráfico nacional.
	gen Bolivia_lbl_2005_2011 = string(Bolivia, "%4.2fc") if Año <= 2011
	gen Bolivia_lbl_2012_2023 = string(Bolivia, "%4.2fc") if inrange(Año, 2012, 2023)
	gen Bolivia_lbl_2024_2026 = string(Bolivia, "%4.2fc") if Año >= 2024

	twoway ///
		(connected Bolivia_2005_2011 Año, ///
			mlabel(Bolivia_lbl_2005_2011) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2012_2023 Año, ///
			mlabel(Bolivia_lbl_2012_2023) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2024_2026 Año, ///
			mlabel(Bolivia_lbl_2024_2026) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)), ///
		xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		xlab(2005(1)2024, angle(45)) ///
		ylab(0(4)20, format(%9.0fc)) ///
		xtitle("") ///
		ytitle("$eje17") ///
		note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
		legend(off)

	gr export "$gph/06_003_08_01.png", replace

* ------------------------------------------------------------
* 6.3.8.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_003_08.xlsx", sheet("_in1") firstrow clear

	set scheme white_viridis
	
	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"

	label values Departamento_order Departamento_order_lbl

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Valor_2005_2011 = Valor if Año <= 2011
	gen Valor_2012_2023 = Valor if inrange(Año, 2012, 2023)
	gen Valor_2024_2026 = Valor if Año >= 2024

	twoway ///
		(connected Valor_2005_2011 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Chuquisaca", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Chuquisaca", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "LaPaz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "LaPaz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Cochabamba", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Cochabamba", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Oruro", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Oruro", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Potosí", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Potosí", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Tarija", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Tarija", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "SantaCruz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "SantaCruz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Beni", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Beni", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Pando", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Pando", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)), ///
		by(Departamento_order, ///
			note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
			legend(off)) ///
		xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		ylab(0(10)20, format(%9.0fc)) ///
		xlab(2005(1)2024, angle(90)) ///
		xtitle("") ///
		ytitle("$eje17") ///
		title("", box color(black) span)

	gr export "$gph/06_003_08_02.png", replace
	
* ============================================================
**# 6.3.9. Edad Mediana de la Población (EMDP)
* Código: 06_003_09
* ============================================================

* ------------------------------------------------------------
* 6.3.9.1. Nacional
* ------------------------------------------------------------

	import excel "$in/06_003_09.xlsx", sheet("_in") firstrow clear
	
	set scheme white_tableau

	format %4.2fc Bolivia

	* Separar los dos tramos:
	* 2012-2023 y 2024 en adelante.
	gen Bolivia_2012_2023 = Bolivia if inrange(Año, 2012, 2023)
	gen Bolivia_2024_2026 = Bolivia if Año >= 2024

	* Etiquetas de datos solo para el gráfico nacional.
	gen Bolivia_lbl_2012_2023 = string(Bolivia, "%4.2fc") if inrange(Año, 2012, 2023)
	gen Bolivia_lbl_2024_2026 = string(Bolivia, "%4.2fc") if Año >= 2024

	twoway ///
		(connected Bolivia_2012_2023 Año, ///
			mlabel(Bolivia_lbl_2012_2023) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2024_2026 Año, ///
			mlabel(Bolivia_lbl_2024_2026) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)), ///
		xline(2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		xlab(2012(1)2024, angle(45)) ///
		ylab(20(2)30, format(%9.0fc)) ///
		xtitle("") ///
		ytitle("$eje19") ///
		note("$nota11" "$nota21" "$fuente2", size(vsmall) span) ///
		legend(off)

	gr export "$gph/06_003_09_01.png", replace

* ------------------------------------------------------------
* 6.3.9.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_003_09.xlsx", sheet("_in1") firstrow clear

	set scheme white_viridis
	
	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"

	label values Departamento_order Departamento_order_lbl
	
	* Separar los dos tramos:
	* 2012-2023 y 2024 en adelante.
	gen Valor_2012_2023 = Valor if inrange(Año, 2012, 2023)
	gen Valor_2024_2026 = Valor if Año >= 2024

	twoway ///
		(connected Valor_2012_2023 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Chuquisaca", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "LaPaz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Cochabamba", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Oruro", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Potosí", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Tarija", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "SantaCruz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Beni", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Pando", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)), ///
		by(Departamento_order, ///
			note("$nota11" "$nota21" "$fuente2", size(vsmall) span) ///
			legend(off)) ///
		xline(2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		ylab(10(15)40, format(%9.0fc)) ///
		xlab(2012(1)2024, angle(90)) ///
		xtitle("") ///
		ytitle("$eje19") ///
		title("", box color(black) span)

	gr export "$gph/06_003_09_02.png", replace
	
* ============================================================
**# 6.3.10. Edad Media de la Población (EMP)
* Código: 06_003_10
* ============================================================

* ------------------------------------------------------------
* 6.3.10.1. Nacional
* ------------------------------------------------------------

	import excel "$in/06_003_10.xlsx", sheet("_in") firstrow clear
	
	set scheme white_tableau

	format %4.2fc Bolivia

	* Separar los dos tramos:
	* 2012-2023 y 2024 en adelante.
	gen Bolivia_2012_2023 = Bolivia if inrange(Año, 2012, 2023)
	gen Bolivia_2024_2026 = Bolivia if Año >= 2024

	* Etiquetas de datos solo para el gráfico nacional.
	gen Bolivia_lbl_2012_2023 = string(Bolivia, "%4.2fc") if inrange(Año, 2012, 2023)
	gen Bolivia_lbl_2024_2026 = string(Bolivia, "%4.2fc") if Año >= 2024

	twoway ///
		(connected Bolivia_2012_2023 Año, ///
			mlabel(Bolivia_lbl_2012_2023) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2024_2026 Año, ///
			mlabel(Bolivia_lbl_2024_2026) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)), ///
		xline(2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		xlab(2012(1)2024, angle(45)) ///
		ylab(20(4)40, format(%9.0fc)) ///
		xtitle("") ///
		ytitle("$eje20") ///
		note("$nota11" "$nota21" "$fuente2", size(vsmall) span) ///
		legend(off)

	gr export "$gph/06_003_10_01.png", replace

* ------------------------------------------------------------
* 6.3.10.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_003_10.xlsx", sheet("_in1") firstrow clear

	set scheme white_viridis
	
	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"

	label values Departamento_order Departamento_order_lbl

	* Separar los dos tramos:
	* 2012-2023 y 2024 en adelante.
	gen Valor_2012_2023 = Valor if inrange(Año, 2012, 2023)
	gen Valor_2024_2026 = Valor if Año >= 2024

	twoway ///
		(connected Valor_2012_2023 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Chuquisaca", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "LaPaz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Cochabamba", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Oruro", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Potosí", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Tarija", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "SantaCruz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Beni", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Pando", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)), ///
		by(Departamento_order, ///
			note("$nota11" "$nota21" "$fuente2", size(vsmall) span) ///
			legend(off)) ///
		xline(2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		ylab(20(10)40, format(%9.0fc)) ///
		xlab(2012(1)2024, angle(90)) ///
		xtitle("") ///
		ytitle("$eje20") ///
		title("", box color(black) span)

	gr export "$gph/06_003_10_02.png", replace
	
* ============================================================
**# 6.3.11. Índice de Reemplazamiento de la Población en Edad Activa (IRPEA)
* Código: 06_003_11
* ============================================================

* ------------------------------------------------------------
* 6.3.11.1. Nacional
* ------------------------------------------------------------

	import excel "$in/06_003_11.xlsx", sheet("_in") firstrow clear
	
	set scheme white_tableau

	format %4.2fc Bolivia

	* Separar los dos tramos:
	* 2012-2023 y 2024 en adelante.
	gen Bolivia_2012_2023 = Bolivia if inrange(Año, 2012, 2023)
	gen Bolivia_2024_2026 = Bolivia if Año >= 2024

	* Etiquetas de datos solo para el gráfico nacional.
	gen Bolivia_lbl_2012_2023 = string(Bolivia, "%4.2fc") if inrange(Año, 2012, 2023)
	gen Bolivia_lbl_2024_2026 = string(Bolivia, "%4.2fc") if Año >= 2024

	twoway ///
		(connected Bolivia_2012_2023 Año, ///
			mlabel(Bolivia_lbl_2012_2023) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2024_2026 Año, ///
			mlabel(Bolivia_lbl_2024_2026) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)), ///
		xline(2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		xlab(2012(1)2024, angle(45)) ///
		ylab(20(5)45, format(%9.0fc)) ///
		xtitle("") ///
		ytitle("$eje21") ///
		note("$nota11" "$nota21" "$fuente2", size(vsmall) span) ///
		legend(off)

	gr export "$gph/06_003_11_01.png", replace

* ------------------------------------------------------------
* 6.3.11.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_003_11.xlsx", sheet("_in1") firstrow clear

	set scheme white_viridis
	
	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"

	label values Departamento_order Departamento_order_lbl

	* Separar los dos tramos:
	* 2012-2023 y 2024 en adelante.
	gen Valor_2012_2023 = Valor if inrange(Año, 2012, 2023)
	gen Valor_2024_2026 = Valor if Año >= 2024

	twoway ///
		(connected Valor_2012_2023 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Chuquisaca", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "LaPaz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Cochabamba", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Oruro", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Potosí", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Tarija", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "SantaCruz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Beni", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Pando", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)), ///
		by(Departamento_order, ///
			note("$nota11" "$nota21" "$fuente2", size(vsmall) span) ///
			legend(off)) ///
		xline(2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		ylab(10(25)60, format(%9.0fc)) ///
		xlab(2012(1)2024, angle(90)) ///
		xtitle("") ///
		ytitle("$eje21") ///
		title("", box color(black) span)

	gr export "$gph/06_003_11_02.png", replace
	
* ============================================================
**# 6.3.11a. Índice de Reemplazamiento de la Población en Edad Activa (IRPEA) según sexo
* Código: 06_003_11a
* ============================================================

* ------------------------------------------------------------
* 6.3.11a.1. Nacional
* ------------------------------------------------------------

	import excel "$in/06_003_11a.xlsx", sheet("_in") firstrow clear
	
	set scheme white_tableau

	format %4.2fc Bolivia1
	format %4.2fc Bolivia2

	* Separar dos tramos:
	* 2012-2023 y 2024 en adelante.
	gen Bolivia1_2012_2023 = Bolivia1 if inrange(Año, 2012, 2023)
	gen Bolivia1_2024_2026 = Bolivia1 if Año >= 2024

	gen Bolivia2_2012_2023 = Bolivia2 if inrange(Año, 2012, 2023)
	gen Bolivia2_2024_2026 = Bolivia2 if Año >= 2024

	* Etiquetas de datos solo para el gráfico nacional.
	gen Bolivia1_lbl_2012_2023 = string(Bolivia1, "%4.2fc") if inrange(Año, 2012, 2023)
	gen Bolivia1_lbl_2024_2026 = string(Bolivia1, "%4.2fc") if Año >= 2024

	gen Bolivia2_lbl_2012_2023 = string(Bolivia2, "%4.2fc") if inrange(Año, 2012, 2023)
	gen Bolivia2_lbl_2024_2026 = string(Bolivia2, "%4.2fc") if Año >= 2024

	twoway /// 
		(connected Bolivia1_2012_2023 Año, ///
			mlabel(Bolivia1_lbl_2012_2023) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) /// 
		(connected Bolivia2_2012_2023 Año, ///
			mlabel(Bolivia2_lbl_2012_2023) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia1_2024_2026 Año, ///
			mlabel(Bolivia1_lbl_2024_2026) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) /// 
		(connected Bolivia2_2024_2026 Año, ///
			mlabel(Bolivia2_lbl_2024_2026) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)), /// 
		xline(2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		xlab(2012(1)2024, angle(45)) /// 
		ylab(20(6)50, format(%9.0fc)) /// 
		xtitle("") /// 
		ytitle("$eje21") /// 
		legend(order(1 "IRPEA hombres" 2 "IRPEA mujeres") ///
			   position(6) cols(2) size(small)) /// 
		note("$nota11" "$nota21" "$fuente2", size(vsmall) span)

	gr export "$gph/06_003_11a_01.png", replace	

* ------------------------------------------------------------
* 6.3.11a.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_003_11a.xlsx", sheet("_in1") firstrow clear

	set scheme white_viridis
	
	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"

	label values Departamento_order Departamento_order_lbl

	* Separar dos tramos:
	* 2012-2023 y 2024 en adelante.
	gen Valor1_2012_2023 = Valor1 if inrange(Año, 2012, 2023)
	gen Valor1_2024_2026 = Valor1 if Año >= 2024

	gen Valor2_2012_2023 = Valor2 if inrange(Año, 2012, 2023)
	gen Valor2_2024_2026 = Valor2 if Año >= 2024

	twoway ///
		(connected Valor1_2012_2023 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2012_2023 Año if Departamento == "Chuquisaca", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor1_2024_2026 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2024_2026 Año if Departamento == "Chuquisaca", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		///
		(connected Valor1_2012_2023 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2012_2023 Año if Departamento == "LaPaz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor1_2024_2026 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2024_2026 Año if Departamento == "LaPaz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		///
		(connected Valor1_2012_2023 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2012_2023 Año if Departamento == "Cochabamba", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor1_2024_2026 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2024_2026 Año if Departamento == "Cochabamba", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		///
		(connected Valor1_2012_2023 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2012_2023 Año if Departamento == "Oruro", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor1_2024_2026 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2024_2026 Año if Departamento == "Oruro", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		///
		(connected Valor1_2012_2023 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2012_2023 Año if Departamento == "Potosí", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor1_2024_2026 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2024_2026 Año if Departamento == "Potosí", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		///
		(connected Valor1_2012_2023 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2012_2023 Año if Departamento == "Tarija", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor1_2024_2026 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2024_2026 Año if Departamento == "Tarija", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		///
		(connected Valor1_2012_2023 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2012_2023 Año if Departamento == "SantaCruz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor1_2024_2026 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2024_2026 Año if Departamento == "SantaCruz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		///
		(connected Valor1_2012_2023 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2012_2023 Año if Departamento == "Beni", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor1_2024_2026 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2024_2026 Año if Departamento == "Beni", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		///
		(connected Valor1_2012_2023 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2012_2023 Año if Departamento == "Pando", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor1_2024_2026 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2024_2026 Año if Departamento == "Pando", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)), ///
		by(Departamento_order, ///
			note("$nota11" "$nota21" "$fuente2", size(vsmall) span)) ///
		xline(2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		legend(order(1 "IRPEA hombres" 2 "IRPEA mujeres") ///
			   position(6) cols(2) size(small)) ///
		ylab(10(25)60, format(%9.0fc)) ///
		xlab(2012(1)2024, angle(90)) ///
		xtitle("") ///
		ytitle("$eje21") ///
		title("", box color(black) span)

	gr export "$gph/06_003_11a_02.png", replace
	
* ============================================================
**# 6.3.12. Índice de Estructura de la Población en Edad Activa (IEPEA)
* Código: 06_003_12
* ============================================================

* ------------------------------------------------------------
* 6.3.12.1. Nacional
* ------------------------------------------------------------

	import excel "$in/06_003_12.xlsx", sheet("_in") firstrow clear
	
	set scheme white_tableau

	format %4.2fc Bolivia

	* Separar dos tramos:
	* 2012-2023 y 2024 en adelante.
	gen Bolivia_2012_2023 = Bolivia if inrange(Año, 2012, 2023)
	gen Bolivia_2024_2026 = Bolivia if Año >= 2024

	* Etiquetas de datos solo para el gráfico nacional.
	gen Bolivia_lbl_2012_2023 = string(Bolivia, "%4.2fc") if inrange(Año, 2012, 2023)
	gen Bolivia_lbl_2024_2026 = string(Bolivia, "%4.2fc") if Año >= 2024

	twoway ///
		(connected Bolivia_2012_2023 Año, ///
			mlabel(Bolivia_lbl_2012_2023) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2024_2026 Año, ///
			mlabel(Bolivia_lbl_2024_2026) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)), ///
		xline(2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		xlab(2012(1)2024, angle(45)) ///
		ylab(40(6)70, format(%9.0fc)) ///
		xtitle("") ///
		ytitle("$eje22") ///
		note("$nota11" "$nota21" "$fuente2", size(vsmall) span) ///
		legend(off)

	gr export "$gph/06_003_12_01.png", replace


* ------------------------------------------------------------
* 6.3.12.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_003_12.xlsx", sheet("_in1") firstrow clear

	set scheme white_viridis
	
	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"

	label values Departamento_order Departamento_order_lbl

	* Separar dos tramos:
	* 2012-2023 y 2024 en adelante.
	gen Valor_2012_2023 = Valor if inrange(Año, 2012, 2023)
	gen Valor_2024_2026 = Valor if Año >= 2024

	twoway ///
		(connected Valor_2012_2023 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Chuquisaca", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "LaPaz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Cochabamba", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Oruro", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Potosí", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Tarija", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "SantaCruz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Beni", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Pando", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)), ///
		by(Departamento_order, ///
			note("$nota11" "$nota21" "$fuente2", size(vsmall) span) ///
			legend(off)) ///
		xline(2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		ylab(20(40)100, format(%9.0fc)) ///
		xlab(2012(1)2024, angle(90)) ///
		xtitle("") ///
		ytitle("$eje22") ///
		title("", box color(black) span)

	gr export "$gph/06_003_12_02.png", replace

* ============================================================
**# 6.3.12a. Índice de Estructura de la Población en Edad Activa (IEPEA) según sexo
* Código: 06_003_12a
* ============================================================

* ------------------------------------------------------------
* 6.3.12a.1. Nacional
* ------------------------------------------------------------

	import excel "$in/06_003_12.xlsx", sheet("_in2") firstrow clear
	
	set scheme white_tableau

	format %4.2fc Bolivia1
	format %4.2fc Bolivia2

	* Separar dos tramos:
	* 2012-2023 y 2024 en adelante.
	gen Bolivia1_2012_2023 = Bolivia1 if inrange(Año, 2012, 2023)
	gen Bolivia1_2024_2026 = Bolivia1 if Año >= 2024

	gen Bolivia2_2012_2023 = Bolivia2 if inrange(Año, 2012, 2023)
	gen Bolivia2_2024_2026 = Bolivia2 if Año >= 2024

	* Etiquetas de datos solo para el gráfico nacional.
	gen Bolivia1_lbl_2012_2023 = string(Bolivia1, "%4.2fc") if inrange(Año, 2012, 2023)
	gen Bolivia1_lbl_2024_2026 = string(Bolivia1, "%4.2fc") if Año >= 2024

	gen Bolivia2_lbl_2012_2023 = string(Bolivia2, "%4.2fc") if inrange(Año, 2012, 2023)
	gen Bolivia2_lbl_2024_2026 = string(Bolivia2, "%4.2fc") if Año >= 2024

	twoway ///
		(connected Bolivia1_2012_2023 Año, ///
			mlabel(Bolivia1_lbl_2012_2023) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			lcolor("108 171 44") ///
			mcolor("108 171 44") ///
			msymbol(circle)) ///
		(connected Bolivia2_2012_2023 Año, ///
			mlabel(Bolivia2_lbl_2012_2023) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			lcolor("255 227 20") ///
			mcolor("255 227 20") ///
			msymbol(circle)) ///
		(connected Bolivia1_2024_2026 Año, ///
			mlabel(Bolivia1_lbl_2024_2026) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			lcolor("108 171 44") ///
			mcolor("108 171 44") ///
			msymbol(circle)) ///
		(connected Bolivia2_2024_2026 Año, ///
			mlabel(Bolivia2_lbl_2024_2026) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			lcolor("255 227 20") ///
			mcolor("255 227 20") ///
			msymbol(circle)), ///
		xline(2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		xlab(2012(1)2024, angle(45)) ///
		ylab(30(8)70, format(%9.0fc)) ///
		xtitle("") ///
		ytitle("$eje22") ///
		legend(order(1 "IEPEA hombres" 2 "IEPEA mujeres") ///
			position(6) cols(2) size(small)) ///
		note("$nota11" "$nota21" "$fuente2", size(vsmall) span)

	gr export "$gph/06_003_12_03.png", replace

* ------------------------------------------------------------
* 6.3.12a.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_003_12.xlsx", sheet("_in3") firstrow clear

	set scheme white_viridis
	
	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"

	label values Departamento_order Departamento_order_lbl

	* Separar dos tramos:
	* 2012-2023 y 2024 en adelante.
	gen Valor1_2012_2023 = Valor1 if inrange(Año, 2012, 2023)
	gen Valor1_2024_2026 = Valor1 if Año >= 2024

	gen Valor2_2012_2023 = Valor2 if inrange(Año, 2012, 2023)
	gen Valor2_2024_2026 = Valor2 if Año >= 2024

	twoway ///
		(connected Valor1_2012_2023 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2012_2023 Año if Departamento == "Chuquisaca", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor1_2024_2026 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2024_2026 Año if Departamento == "Chuquisaca", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor1_2012_2023 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2012_2023 Año if Departamento == "LaPaz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor1_2024_2026 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2024_2026 Año if Departamento == "LaPaz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor1_2012_2023 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2012_2023 Año if Departamento == "Cochabamba", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor1_2024_2026 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2024_2026 Año if Departamento == "Cochabamba", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor1_2012_2023 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2012_2023 Año if Departamento == "Oruro", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor1_2024_2026 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2024_2026 Año if Departamento == "Oruro", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor1_2012_2023 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2012_2023 Año if Departamento == "Potosí", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor1_2024_2026 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2024_2026 Año if Departamento == "Potosí", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor1_2012_2023 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2012_2023 Año if Departamento == "Tarija", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor1_2024_2026 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2024_2026 Año if Departamento == "Tarija", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor1_2012_2023 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2012_2023 Año if Departamento == "SantaCruz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor1_2024_2026 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2024_2026 Año if Departamento == "SantaCruz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor1_2012_2023 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2012_2023 Año if Departamento == "Beni", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor1_2024_2026 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2024_2026 Año if Departamento == "Beni", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor1_2012_2023 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2012_2023 Año if Departamento == "Pando", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor1_2024_2026 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2024_2026 Año if Departamento == "Pando", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)), ///
		by(Departamento_order, ///
			note("$nota11" "$nota21" "$fuente2", size(vsmall) span)) ///
		xline(2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		legend(order(1 "IEPEA hombres" 2 "IEPEA mujeres") cols(2) size(small)) ///
		ylab(30(20)70, format(%9.0fc)) ///
		xlab(2012(1)2024, angle(90)) ///
		xtitle("") ///
		ytitle("$eje22") ///
		title("", box color(black) span)

	gr export "$gph/06_003_12_04.png", replace
	
* ============================================================
**# 6.3.13. Índice de Maternidad (IMAT)
* Código: 06_003_13
* ============================================================

* ------------------------------------------------------------
* 6.3.13.1. Nacional
* ------------------------------------------------------------

	import excel "$in/06_003_13.xlsx", sheet("_in") firstrow clear
	
	set scheme white_tableau

	format %4.2fc Bolivia

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Bolivia_2005_2011 = Bolivia if Año <= 2011
	gen Bolivia_2012_2023 = Bolivia if inrange(Año, 2012, 2023)
	gen Bolivia_2024_2026 = Bolivia if Año >= 2024

	* Etiquetas de datos solo para el gráfico nacional.
	gen Bolivia_lbl_2005_2011 = string(Bolivia, "%4.2fc") if Año <= 2011
	gen Bolivia_lbl_2012_2023 = string(Bolivia, "%4.2fc") if inrange(Año, 2012, 2023)
	gen Bolivia_lbl_2024_2026 = string(Bolivia, "%4.2fc") if Año >= 2024

	twoway ///
		(connected Bolivia_2005_2011 Año, ///
			mlabel(Bolivia_lbl_2005_2011) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2012_2023 Año, ///
			mlabel(Bolivia_lbl_2012_2023) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2024_2026 Año, ///
			mlabel(Bolivia_lbl_2024_2026) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)), ///
		xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		xlab(2005(1)2024, angle(45)) ///
		ylab(0(15)75, format(%9.0fc)) ///
		xtitle("") ///
		ytitle("$eje23") ///
		note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
		legend(off)

	gr export "$gph/06_003_13_01.png", replace

* ------------------------------------------------------------
* 6.3.13.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_003_13.xlsx", sheet("_in1") firstrow clear
	
	set scheme white_viridis
	
	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"

	label values Departamento_order Departamento_order_lbl

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Valor_2005_2011 = Valor if Año <= 2011
	gen Valor_2012_2023 = Valor if inrange(Año, 2012, 2023)
	gen Valor_2024_2026 = Valor if Año >= 2024

	twoway ///
		(connected Valor_2005_2011 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Chuquisaca", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Chuquisaca", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "LaPaz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "LaPaz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Cochabamba", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Cochabamba", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Oruro", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Oruro", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Potosí", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Potosí", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Tarija", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Tarija", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "SantaCruz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "SantaCruz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Beni", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Beni", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Pando", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Pando", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)), ///
		by(Departamento_order, ///
			note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
			legend(off)) ///
		xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		ylab(10(40)90, format(%9.0fc)) ///
		xlab(2005(1)2024, angle(90)) ///
		xtitle("") ///
		ytitle("$eje23") ///
		title("", box color(black) span)

	gr export "$gph/06_003_13_02.png", replace

* ============================================================
**# 6.3.14. Índice de Friz (IF)
* Código: 06_003_14
* ============================================================

* ------------------------------------------------------------
* 6.3.14.1. Nacional
* ------------------------------------------------------------

	import excel "$in/06_003_14.xlsx", sheet("_in") firstrow clear
	
	set scheme white_tableau

	format %4.2fc Bolivia

	* Separar los dos tramos:
	* 2012-2023 y 2024 en adelante.
	gen Bolivia_2012_2023 = Bolivia if inrange(Año, 2012, 2023)
	gen Bolivia_2024_2026 = Bolivia if Año >= 2024

	* Etiquetas de datos solo para el gráfico nacional.
	gen Bolivia_lbl_2012_2023 = string(Bolivia, "%4.2fc") if inrange(Año, 2012, 2023)
	gen Bolivia_lbl_2024_2026 = string(Bolivia, "%4.2fc") if Año >= 2024

	twoway ///
		(connected Bolivia_2012_2023 Año, ///
			mlabel(Bolivia_lbl_2012_2023) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2024_2026 Año, ///
			mlabel(Bolivia_lbl_2024_2026) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)), ///
		xline(2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		xlab(2012(1)2024, angle(45)) ///
		ylab(100(24)220, format(%9.0fc)) ///
		xtitle("") ///
		ytitle("$eje24") ///
		note("$nota11" "$nota21" "$fuente2", size(vsmall) span) ///
		legend(off)

	gr export "$gph/06_003_14_01.png", replace

* ------------------------------------------------------------
* 6.3.14.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_003_14.xlsx", sheet("_in1") firstrow clear

	set scheme white_viridis
	
	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"

	label values Departamento_order Departamento_order_lbl

	* Separar los dos tramos:
	* 2012-2023 y 2024 en adelante.
	gen Valor_2012_2023 = Valor if inrange(Año, 2012, 2023)
	gen Valor_2024_2026 = Valor if Año >= 2024

	twoway ///
		(connected Valor_2012_2023 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Chuquisaca", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "LaPaz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Cochabamba", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Oruro", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Potosí", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Tarija", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "SantaCruz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Beni", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Pando", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)), ///
		by(Departamento_order, ///
			note("$nota11" "$nota21" "$fuente2", size(vsmall) span) ///
			legend(off)) ///
		xline(2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		ylab(100(100)300, format(%9.0fc)) ///
		xlab(2012(1)2024, angle(90)) ///
		xtitle("") ///
		ytitle("$eje24") ///
		title("", box color(black) span)

	gr export "$gph/06_003_14_02.png", replace
	
* ============================================================
**# 6.3.15. Índice de Sundbarg (IS)
* Código: 06_003_15
* ============================================================

* ------------------------------------------------------------
* 6.3.15.1. Nacional
* ------------------------------------------------------------

	import excel "$in/06_003_15.xlsx", sheet("_in") firstrow clear
	
	set scheme white_tableau

	format %4.2fc Bolivia1
	format %4.2fc Bolivia2

	* Separar dos tramos:
	* 2012-2023 y 2024 en adelante.
	gen Bolivia1_2012_2023 = Bolivia1 if inrange(Año, 2012, 2023)
	gen Bolivia1_2024_2026 = Bolivia1 if Año >= 2024

	gen Bolivia2_2012_2023 = Bolivia2 if inrange(Año, 2012, 2023)
	gen Bolivia2_2024_2026 = Bolivia2 if Año >= 2024

	* Etiquetas de datos solo para el gráfico nacional.
	gen Bolivia1_lbl_2012_2023 = string(Bolivia1, "%4.2fc") if inrange(Año, 2012, 2023)
	gen Bolivia1_lbl_2024_2026 = string(Bolivia1, "%4.2fc") if Año >= 2024

	gen Bolivia2_lbl_2012_2023 = string(Bolivia2, "%4.2fc") if inrange(Año, 2012, 2023)
	gen Bolivia2_lbl_2024_2026 = string(Bolivia2, "%4.2fc") if Año >= 2024

	twoway /// 
		(connected Bolivia1_2012_2023 Año, ///
			mlabel(Bolivia1_lbl_2012_2023) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			lcolor("108 171 44") ///
			mcolor("108 171 44") ///
			msymbol(circle)) /// 
		(connected Bolivia2_2012_2023 Año, ///
			mlabel(Bolivia2_lbl_2012_2023) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			lcolor("255 227 20") ///
			mcolor("255 227 20") ///
			msymbol(circle)) ///
		(connected Bolivia1_2024_2026 Año, ///
			mlabel(Bolivia1_lbl_2024_2026) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			lcolor("108 171 44") ///
			mcolor("108 171 44") ///
			msymbol(circle)) /// 
		(connected Bolivia2_2024_2026 Año, ///
			mlabel(Bolivia2_lbl_2024_2026) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			lcolor("255 227 20") ///
			mcolor("255 227 20") ///
			msymbol(circle)), /// 
		xline(2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		xlab(2012(1)2024, angle(45)) /// 
		ylab(0(20)100, format(%9.0fc)) /// 
		xtitle("") /// 
		ytitle("$eje25") /// 
		legend(order(1 "IS jóvenes" 2 "IS mayores") ///
			   position(6) cols(2) size(small)) /// 
		note("$nota11" "$nota21" "$fuente2", size(vsmall) span)

	gr export "$gph/06_003_15_01.png", replace	

* ------------------------------------------------------------
* 6.3.15.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_003_15.xlsx", sheet("_in1") firstrow clear

	set scheme white_viridis
	
	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"

	label values Departamento_order Departamento_order_lbl

	* Separar dos tramos:
	* 2012-2023 y 2024 en adelante.
	gen Valor1_2012_2023 = Valor1 if inrange(Año, 2012, 2023)
	gen Valor1_2024_2026 = Valor1 if Año >= 2024

	gen Valor2_2012_2023 = Valor2 if inrange(Año, 2012, 2023)
	gen Valor2_2024_2026 = Valor2 if Año >= 2024

	twoway ///
		(connected Valor1_2012_2023 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2012_2023 Año if Departamento == "Chuquisaca", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor1_2024_2026 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2024_2026 Año if Departamento == "Chuquisaca", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		///
		(connected Valor1_2012_2023 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2012_2023 Año if Departamento == "LaPaz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor1_2024_2026 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2024_2026 Año if Departamento == "LaPaz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		///
		(connected Valor1_2012_2023 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2012_2023 Año if Departamento == "Cochabamba", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor1_2024_2026 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2024_2026 Año if Departamento == "Cochabamba", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		///
		(connected Valor1_2012_2023 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2012_2023 Año if Departamento == "Oruro", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor1_2024_2026 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2024_2026 Año if Departamento == "Oruro", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		///
		(connected Valor1_2012_2023 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2012_2023 Año if Departamento == "Potosí", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor1_2024_2026 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2024_2026 Año if Departamento == "Potosí", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		///
		(connected Valor1_2012_2023 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2012_2023 Año if Departamento == "Tarija", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor1_2024_2026 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2024_2026 Año if Departamento == "Tarija", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		///
		(connected Valor1_2012_2023 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2012_2023 Año if Departamento == "SantaCruz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor1_2024_2026 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2024_2026 Año if Departamento == "SantaCruz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		///
		(connected Valor1_2012_2023 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2012_2023 Año if Departamento == "Beni", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor1_2024_2026 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2024_2026 Año if Departamento == "Beni", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		///
		(connected Valor1_2012_2023 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2012_2023 Año if Departamento == "Pando", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor1_2024_2026 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor2_2024_2026 Año if Departamento == "Pando", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)), ///
		by(Departamento_order, ///
			note("$nota11" "$nota21" "$fuente2", size(vsmall) span)) ///
		xline(2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		legend(order(1 "IS jóvenes" 2 "IS mayores") cols(2) size(small)) ///
		ylab(0(50)100, format(%9.0fc)) ///
		xlab(2012(1)2024, angle(90)) ///
		xtitle("") ///
		ytitle("$eje25") ///
		title("", box color(black) span)

	gr export "$gph/06_003_15_02.png", replace
	
* ============================================================
**# 6.3.16. Índice de Burgdofer (IB)
* Código: 06_003_16
* ============================================================

* ------------------------------------------------------------
* 6.3.16.1. Nacional
* ------------------------------------------------------------

	import excel "$in/06_003_16.xlsx", sheet("_in") firstrow clear
	
	set scheme white_tableau

	format %4.2fc Bolivia

	* Separar dos tramos:
	* 2012-2023 y 2024 en adelante.
	gen Bolivia_2012_2023 = Bolivia if inrange(Año, 2012, 2023)
	gen Bolivia_2024_2026 = Bolivia if Año >= 2024

	* Etiquetas de datos solo para el gráfico nacional.
	gen Bolivia_lbl_2012_2023 = string(Bolivia, "%4.2fc") if inrange(Año, 2012, 2023)
	gen Bolivia_lbl_2024_2026 = string(Bolivia, "%4.2fc") if Año >= 2024

	twoway ///
		(connected Bolivia_2012_2023 Año, ///
			mlabel(Bolivia_lbl_2012_2023) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2024_2026 Año, ///
			mlabel(Bolivia_lbl_2024_2026) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)), ///
		xline(2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		xlab(2012(1)2024, angle(45)) ///
		ylab(90(20)190, format(%9.0fc)) ///
		xtitle("") ///
		ytitle("$eje32") ///
		note("$nota11" "$nota21" "$fuente2", size(vsmall) span) ///
		legend(off)

	gr export "$gph/06_003_16_01.png", replace

* ------------------------------------------------------------
* 6.3.16.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_003_16.xlsx", sheet("_in1") firstrow clear

	set scheme white_viridis
	
	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"

	label values Departamento_order Departamento_order_lbl

	* Separar dos tramos:
	* 2012-2023 y 2024 en adelante.
	gen Valor_2012_2023 = Valor if inrange(Año, 2012, 2023)
	gen Valor_2024_2026 = Valor if Año >= 2024

	twoway ///
		(connected Valor_2012_2023 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Chuquisaca", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "LaPaz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Cochabamba", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Oruro", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Potosí", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Tarija", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "SantaCruz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Beni", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Pando", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)), ///
		by(Departamento_order, ///
			note("$nota11" "$nota21" "$fuente2", size(vsmall) span) ///
			legend(off)) ///
		xline(2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		ylab(0(150)300, format(%9.0fc)) ///
		xlab(2012(1)2024, angle(90)) ///
		xtitle("") ///
		ytitle("$eje32") ///
		title("", box color(black) span)

	gr export "$gph/06_003_16_02.png", replace
	
* ============================================================
**# 6.3.17. Índice Generacional de Ancianos (IGA)
* Código: 06_003_17
* ============================================================

* ------------------------------------------------------------
* 6.3.17.1. Nacional
* ------------------------------------------------------------

	import excel "$in/06_003_17.xlsx", sheet("_in") firstrow clear
	
	set scheme white_tableau

	format %4.2fc Bolivia

	* Separar dos tramos:
	* 2012-2023 y 2024 en adelante.
	gen Bolivia_2012_2023 = Bolivia if inrange(Año, 2012, 2023)
	gen Bolivia_2024_2026 = Bolivia if Año >= 2024

	* Etiquetas de datos solo para el gráfico nacional.
	gen Bolivia_lbl_2012_2023 = string(Bolivia, "%4.2fc") if inrange(Año, 2012, 2023)
	gen Bolivia_lbl_2024_2026 = string(Bolivia, "%4.2fc") if Año >= 2024

	twoway ///
		(connected Bolivia_2012_2023 Año, ///
			mlabel(Bolivia_lbl_2012_2023) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2024_2026 Año, ///
			mlabel(Bolivia_lbl_2024_2026) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)), ///
		xline(2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		xlab(2012(1)2024, angle(45)) ///
		ylab(300(35)475, format(%9.0fc)) ///
		xtitle("") ///
		ytitle("$eje26") ///
		note("$nota11" "$nota21" "$fuente2", size(vsmall) span) ///
		legend(off)

	gr export "$gph/06_003_17_01.png", replace

* ------------------------------------------------------------
* 6.3.17.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_003_17.xlsx", sheet("_in1") firstrow clear

	set scheme white_viridis
	
	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"

	label values Departamento_order Departamento_order_lbl

	* Separar dos tramos:
	* 2012-2023 y 2024 en adelante.
	gen Valor_2012_2023 = Valor if inrange(Año, 2012, 2023)
	gen Valor_2024_2026 = Valor if Año >= 2024

	twoway ///
		(connected Valor_2012_2023 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Chuquisaca", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "LaPaz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Cochabamba", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Oruro", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Potosí", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Tarija", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "SantaCruz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Beni", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Pando", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)), ///
		by(Departamento_order, ///
			note("$nota11" "$nota21" "$fuente2", size(vsmall) span) ///
			legend(off)) ///
		xline(2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		ylab(0(500)1000, format(%9.0fc)) ///
		xlab(2012(1)2024, angle(90)) ///
		xtitle("") ///
		ytitle("$eje26") ///
		title("", box color(black) span)

	gr export "$gph/06_003_17_02.png", replace

* ============================================================
**# 6.3.18. Índice de Vejez (IV)
* Código: 06_003_18
* ============================================================

* ------------------------------------------------------------
* 6.3.18.1. Nacional
* ------------------------------------------------------------

	import excel "$in/06_003_18.xlsx", sheet("_in") firstrow clear
	
	set scheme white_tableau

	format %4.2fc Bolivia

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Bolivia_2005_2011 = Bolivia if Año <= 2011
	gen Bolivia_2012_2023 = Bolivia if inrange(Año, 2012, 2023)
	gen Bolivia_2024_2026 = Bolivia if Año >= 2024

	* Etiquetas de datos solo para el gráfico nacional.
	gen Bolivia_lbl_2005_2011 = string(Bolivia, "%4.2fc") if Año <= 2011
	gen Bolivia_lbl_2012_2023 = string(Bolivia, "%4.2fc") if inrange(Año, 2012, 2023)
	gen Bolivia_lbl_2024_2026 = string(Bolivia, "%4.2fc") if Año >= 2024

	twoway ///
		(connected Bolivia_2005_2011 Año, ///
			mlabel(Bolivia_lbl_2005_2011) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2012_2023 Año, ///
			mlabel(Bolivia_lbl_2012_2023) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2024_2026 Año, ///
			mlabel(Bolivia_lbl_2024_2026) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)), ///
		xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		xlab(2005(1)2024, angle(45)) ///
		ylab(0(2)10, format(%9.2fc)) ///
		xtitle("") ///
		ytitle("$eje30") ///
		note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
		legend(off)

	gr export "$gph/06_003_18_01.png", replace

* ------------------------------------------------------------
* 6.3.18.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_003_18.xlsx", sheet("_in1") firstrow clear

	set scheme white_viridis
	
	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"

	label values Departamento_order Departamento_order_lbl

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Valor_2005_2011 = Valor if Año <= 2011
	gen Valor_2012_2023 = Valor if inrange(Año, 2012, 2023)
	gen Valor_2024_2026 = Valor if Año >= 2024

	twoway ///
		(connected Valor_2005_2011 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Chuquisaca", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Chuquisaca", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "LaPaz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "LaPaz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Cochabamba", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Cochabamba", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Oruro", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Oruro", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Potosí", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Potosí", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Tarija", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Tarija", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "SantaCruz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "SantaCruz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Beni", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Beni", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Pando", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Pando", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)), ///
		by(Departamento_order, ///
			note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
			legend(off)) ///
		xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		ylab(1(4.5)10, format(%9.2fc)) ///
		xlab(2005(1)2024, angle(90)) ///
		xtitle("") ///
		ytitle("$eje30") ///
		title("", box color(black) span)

	gr export "$gph/06_003_18_02.png", replace
	
* ============================================================
**# 6.3.19. Índice de Sobreenvejecimiento (ISE)
* Código: 06_003_19
* ============================================================

* ------------------------------------------------------------
* 6.3.19.1. Nacional
* ------------------------------------------------------------

	import excel "$in/06_003_19.xlsx", sheet("_in") firstrow clear
	
	set scheme white_tableau

	format %4.2fc Bolivia

	* Separar dos tramos:
	* 2012-2023 y 2024 en adelante.
	gen Bolivia_2012_2023 = Bolivia if inrange(Año, 2012, 2023)
	gen Bolivia_2024_2026 = Bolivia if Año >= 2024

	* Etiquetas de datos solo para el gráfico nacional.
	gen Bolivia_lbl_2012_2023 = string(Bolivia, "%4.2fc") if inrange(Año, 2012, 2023)
	gen Bolivia_lbl_2024_2026 = string(Bolivia, "%4.2fc") if Año >= 2024

	twoway ///
		(connected Bolivia_2012_2023 Año, ///
			mlabel(Bolivia_lbl_2012_2023) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2024_2026 Año, ///
			mlabel(Bolivia_lbl_2024_2026) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)), ///
		xline(2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		xlab(2012(1)2024, angle(45)) ///
		ylab(0(6)30, format(%9.0fc)) ///
		xtitle("") ///
		ytitle("$eje31") ///
		note("$nota11" "$nota21" "$fuente2", size(vsmall) span) ///
		legend(off)

	gr export "$gph/06_003_19_01.png", replace

* ------------------------------------------------------------
* 6.3.19.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_003_19.xlsx", sheet("_in1") firstrow clear

	set scheme white_viridis
	
	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"

	label values Departamento_order Departamento_order_lbl

	* Separar dos tramos:
	* 2012-2023 y 2024 en adelante.
	gen Valor_2012_2023 = Valor if inrange(Año, 2012, 2023)
	gen Valor_2024_2026 = Valor if Año >= 2024

	twoway ///
		(connected Valor_2012_2023 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Chuquisaca", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "LaPaz",      lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "LaPaz",      lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Cochabamba", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Oruro",      lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Oruro",      lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Potosí",     lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Potosí",     lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Tarija",     lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Tarija",     lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "SantaCruz",  lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "SantaCruz",  lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Beni",       lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Beni",       lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Pando",      lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Pando",      lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)), ///
		by(Departamento_order, ///
			note("$nota11" "$nota21" "$fuente2", size(vsmall) span) ///
			legend(off)) ///
		xline(2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		ylab(0(15)30, format(%9.0fc)) ///
		xlab(2012(1)2024, angle(90)) ///
		xtitle("") ///
		ytitle("$eje31") ///
		title("", box color(black) span)

	gr export "$gph/06_003_19_02.png", replace
	
* ============================================================
**# 6.3.20. Índice de Juventud (IJ)
* Código: 06_003_20
* ============================================================

* ------------------------------------------------------------
* 6.3.20.1. Nacional
* ------------------------------------------------------------

	import excel "$in/06_003_20.xlsx", sheet("_in") firstrow clear
	
	set scheme white_tableau

	format %4.2fc Bolivia

	* Separar dos tramos:
	* 2012-2023 y 2024 en adelante.
	gen Bolivia_2012_2023 = Bolivia if inrange(Año, 2012, 2023)
	gen Bolivia_2024_2026 = Bolivia if Año >= 2024

	* Etiquetas de datos solo para el gráfico nacional.
	gen Bolivia_lbl_2012_2023 = string(Bolivia, "%4.2fc") if inrange(Año, 2012, 2023)
	gen Bolivia_lbl_2024_2026 = string(Bolivia, "%4.2fc") if Año >= 2024

	twoway ///
		(connected Bolivia_2012_2023 Año, ///
			mlabel(Bolivia_lbl_2012_2023) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2024_2026 Año, ///
			mlabel(Bolivia_lbl_2024_2026) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)), ///
		xline(2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		xlab(2012(1)2024, angle(45)) ///
		ylab(20(2)30, format(%9.2fc)) ///
		xtitle("") ///
		ytitle("$eje33") ///
		note("$nota11" "$nota21" "$fuente2", size(vsmall) span) ///
		legend(off)

	gr export "$gph/06_003_20_01.png", replace

* ------------------------------------------------------------
* 6.3.20.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_003_20.xlsx", sheet("_in1") firstrow clear

	set scheme white_viridis
	
	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"

	label values Departamento_order Departamento_order_lbl

	* Separar dos tramos:
	* 2012-2023 y 2024 en adelante.
	gen Valor_2012_2023 = Valor if inrange(Año, 2012, 2023)
	gen Valor_2024_2026 = Valor if Año >= 2024

	twoway ///
		(connected Valor_2012_2023 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Chuquisaca", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "LaPaz",      lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "LaPaz",      lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Cochabamba", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Oruro",      lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Oruro",      lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Potosí",     lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Potosí",     lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Tarija",     lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Tarija",     lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "SantaCruz",  lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "SantaCruz",  lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Beni",       lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Beni",       lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Pando",      lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Pando",      lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)), ///
		by(Departamento_order, ///
			note("$nota11" "$nota21" "$fuente2", size(vsmall) span) ///
			legend(off)) ///
		xline(2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		ylab(15(10)35, format(%9.2fc)) ///
		xlab(2012(1)2024, angle(90)) ///
		xtitle("") ///
		ytitle("$eje33") ///
		title("", box color(black) span)

	gr export "$gph/06_003_20_02.png", replace
	
* ============================================================
**# 6.3.21. Índice de Infancia (II)
* Código: 06_003_21
* ============================================================

* ------------------------------------------------------------
* 6.3.21.1. Nacional
* ------------------------------------------------------------

	import excel "$in/06_003_21.xlsx", sheet("_in") firstrow clear
	
	set scheme white_tableau

	format %4.2fc Bolivia

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Bolivia_2005_2011 = Bolivia if Año <= 2011
	gen Bolivia_2012_2023 = Bolivia if inrange(Año, 2012, 2023)
	gen Bolivia_2024_2026 = Bolivia if Año >= 2024

	* Etiquetas de datos solo para el gráfico nacional.
	gen Bolivia_lbl_2005_2011 = string(Bolivia, "%4.2fc") if Año <= 2011
	gen Bolivia_lbl_2012_2023 = string(Bolivia, "%4.2fc") if inrange(Año, 2012, 2023)
	gen Bolivia_lbl_2024_2026 = string(Bolivia, "%4.2fc") if Año >= 2024

	twoway ///
		(connected Bolivia_2005_2011 Año, ///
			mlabel(Bolivia_lbl_2005_2011) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2012_2023 Año, ///
			mlabel(Bolivia_lbl_2012_2023) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2024_2026 Año, ///
			mlabel(Bolivia_lbl_2024_2026) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)), ///
		xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		xlab(2005(1)2024, angle(45)) ///
		ylab(20(5)45, format(%9.0fc)) ///
		xtitle("") ///
		ytitle("$eje34") ///
		note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
		legend(off)

	gr export "$gph/06_003_21_01.png", replace

* ------------------------------------------------------------
* 6.3.21.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_003_21.xlsx", sheet("_in1") firstrow clear

	set scheme white_viridis
	
	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"

	label values Departamento_order Departamento_order_lbl

	* Separar los tres tramos:
	* 2005-2011, 2012-2023 y 2024 en adelante.
	gen Valor_2005_2011 = Valor if Año <= 2011
	gen Valor_2012_2023 = Valor if inrange(Año, 2012, 2023)
	gen Valor_2024_2026 = Valor if Año >= 2024

	twoway ///
		(connected Valor_2005_2011 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Chuquisaca", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Chuquisaca", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "LaPaz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "LaPaz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Cochabamba", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Cochabamba", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Oruro", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Oruro", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Potosí", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Potosí", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Tarija", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Tarija", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "SantaCruz", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "SantaCruz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Beni", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Beni", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2005_2011 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Pando", lcolor("255 227 20") mcolor("255 227 20") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Pando", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)), ///
		by(Departamento_order, ///
			note("$nota1" "$nota2" "$nota3" "$fuente2", size(vsmall) span) ///
			legend(off)) ///
		xline(2011.5 2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		ylab(20(15)50, format(%9.0fc)) ///
		xlab(2005(1)2024, angle(90)) ///
		xtitle("") ///
		ytitle("$eje34") ///
		title("", box color(black) span)

	gr export "$gph/06_003_21_02.png", replace

* ============================================================
**# 6.3.22. Índice de Potencialidad de Reproducción Femenina (IPRF)
* Código: 06_003_22
* ============================================================

* ------------------------------------------------------------
* 6.3.22.1. Nacional
* ------------------------------------------------------------

	import excel "$in/06_003_22.xlsx", sheet("_in") firstrow clear
	
	set scheme white_tableau

	format %4.2fc Bolivia

	* Separar dos tramos:
	* 2012-2023 y 2024 en adelante.
	gen Bolivia_2012_2023 = Bolivia if inrange(Año, 2012, 2023)
	gen Bolivia_2024_2026 = Bolivia if Año >= 2024

	* Etiquetas de datos solo para el gráfico nacional.
	gen Bolivia_lbl_2012_2023 = string(Bolivia, "%4.2fc") if inrange(Año, 2012, 2023)
	gen Bolivia_lbl_2024_2026 = string(Bolivia, "%4.2fc") if Año >= 2024

	twoway ///
		(connected Bolivia_2012_2023 Año, ///
			mlabel(Bolivia_lbl_2012_2023) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)) ///
		(connected Bolivia_2024_2026 Año, ///
			mlabel(Bolivia_lbl_2024_2026) ///
			mlabposition(12) ///
			mlabsize(*.90) ///
			msymbol(circle)), ///
		xline(2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		xlab(2012(1)2024, angle(45)) ///
		ylab(120(10)170, format(%9.0fc)) ///
		xtitle("") ///
		ytitle("$eje28") ///
		note("$nota11" "$nota21" "$fuente2", size(vsmall) span) ///
		legend(off)

	gr export "$gph/06_003_22_01.png", replace

* ------------------------------------------------------------
* 6.3.22.2. Departamental
* ------------------------------------------------------------

	import excel "$in/06_003_22.xlsx", sheet("_in1") firstrow clear

	set scheme white_viridis
	
	gen Departamento_order = .
	replace Departamento_order = 1 if Departamento == "Chuquisaca"
	replace Departamento_order = 2 if Departamento == "LaPaz"
	replace Departamento_order = 3 if Departamento == "Cochabamba"
	replace Departamento_order = 4 if Departamento == "Oruro"
	replace Departamento_order = 5 if Departamento == "Potosí"
	replace Departamento_order = 6 if Departamento == "Tarija"
	replace Departamento_order = 7 if Departamento == "SantaCruz"
	replace Departamento_order = 8 if Departamento == "Beni"
	replace Departamento_order = 9 if Departamento == "Pando"

	label define Departamento_order_lbl ///
		1 "Chuquisaca" ///
		2 "La Paz" ///
		3 "Cochabamba" ///
		4 "Oruro" ///
		5 "Potosí" ///
		6 "Tarija" ///
		7 "Santa Cruz" ///
		8 "Beni" ///
		9 "Pando"

	label values Departamento_order Departamento_order_lbl

	* Separar dos tramos:
	* 2012-2023 y 2024 en adelante.
	gen Valor_2012_2023 = Valor if inrange(Año, 2012, 2023)
	gen Valor_2024_2026 = Valor if Año >= 2024

	twoway ///
		(connected Valor_2012_2023 Año if Departamento == "Chuquisaca", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Chuquisaca", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "LaPaz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "LaPaz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Cochabamba", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Cochabamba", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Oruro", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Oruro", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Potosí", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Potosí", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Tarija", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Tarija", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "SantaCruz", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "SantaCruz", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Beni", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Beni", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)) ///
		(connected Valor_2012_2023 Año if Departamento == "Pando", lcolor("108 171 44") mcolor("108 171 44") msymbol(circle)) ///
		(connected Valor_2024_2026 Año if Departamento == "Pando", lcolor("231 126 35") mcolor("231 126 35") msymbol(circle)), ///
		by(Departamento_order, ///
			note("$nota11" "$nota21" "$fuente2", size(vsmall) span) ///
			legend(off)) ///
		xline(2023.5, lcolor(gs10) lpattern(dash) lwidth(thin)) ///
		ylab(110(40)190, format(%9.0fc)) ///
		xlab(2012(1)2024, angle(90)) ///
		xtitle("") ///
		ytitle("$eje28") ///
		title("", box color(black) span)

	gr export "$gph/06_003_22_02.png", replace