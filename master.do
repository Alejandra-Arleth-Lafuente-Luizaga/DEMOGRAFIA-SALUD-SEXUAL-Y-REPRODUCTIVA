/****************************************************************************************
 PROYECTO:
     Automatización de indicadores de Demografía, Salud Sexual y Reproductiva

 ARCHIVO:
     00_master.do

 OBJETIVO:
     Ejecutar todo el flujo del proyecto:
         1. Importar proyecciones.
         2. Limpiar edades simples por sexo.
         3. Calcular indicadores.
         4. Verificar el Excel final.

 AUTORA:
     Alejandra Arleth Lafuente-Luizaga

 FECHA:
     06-may-2026
****************************************************************************************/

* ============================================================
* 0. RUTAS DEL PROYECTO
* ============================================================

global root "G:/Unidades compartidas/1_INDICADORES_ODSB_2022/01COPIA_IISEC_ODSB/INDICADORES/6_SALUD"

global do       "$root/_do"
global out      "$root/_out"
global out_dta  "$out/Bases DTA"
global out_xlsx "$out/Indicadores Excel"
global out_aud  "$out/Auditorías"


* ============================================================
* 1. CREAR CARPETAS SI NO EXISTEN
* ============================================================

cap mkdir "$do"
cap mkdir "$out"
cap mkdir "$out_dta"
cap mkdir "$out_xlsx"
cap mkdir "$out_aud"


* ============================================================
* 2. MOSTRAR RUTAS
* ============================================================

di as result "============================================================"
di as result "INICIO DEL CORRIDO GENERAL"
di as result "============================================================"

di as text "Ruta raíz:"
di as result "$root"

di as text "Scripts:"
di as result "$do"

di as text "Bases DTA:"
di as result "$out_dta"

di as text "Indicadores Excel:"
di as result "$out_xlsx"

di as text "Auditorías:"
di as result "$out_aud"


* ============================================================
* 3. VERIFICAR SCRIPTS
* ============================================================

confirm file "$do/01_importar_proyecciones.do"
confirm file "$do/02_limpiar_edades_hombre_mujer.do"
confirm file "$do/03_calcular_indicadores.do"


* ============================================================
* 4. EJECUTAR SCRIPT 01
* ============================================================

di as result "============================================================"
di as result "EJECUTANDO 01_importar_proyecciones.do"
di as result "============================================================"

do "$do/01_importar_proyecciones.do"

confirm file "$out_dta/pob_grupos_especiales_2005_2024.dta"
confirm file "$out_dta/pob_edad_salud_2005_2024.dta"
confirm file "$out_dta/pob_edades_simples_raw_2012_2024.dta"


* ============================================================
* 5. EJECUTAR SCRIPT 02
* ============================================================

di as result "============================================================"
di as result "EJECUTANDO 02_limpiar_edades_hombre_mujer.do"
di as result "============================================================"

do "$do/02_limpiar_edades_hombre_mujer.do"

confirm file "$out_dta/pob_edades_simples_long_2012_2024.dta"


* ============================================================
* 6. EJECUTAR SCRIPT 03
* ============================================================

di as result "============================================================"
di as result "EJECUTANDO 03_calcular_indicadores.do"
di as result "============================================================"

do "$do/03_calcular_indicadores.do"


* ============================================================
* 7. VERIFICAR EXCEL FINAL
* ============================================================

di as result "============================================================"
di as result "VERIFICANDO EXCEL FINAL"
di as result "============================================================"

local excel_final "G:/Unidades compartidas/1_INDICADORES_ODSB_2022/01COPIA_IISEC_ODSB/INDICADORES/6_SALUD/_out/Indicadores Excel/indicadores_demografia_ssr_2012_2024.xlsx"

confirm file "`excel_final'"

di as result "Excel final encontrado correctamente:"
di as result "`excel_final'"


* ============================================================
* 8. FIN
* ============================================================

di as result "============================================================"
di as result "CORRIDO GENERAL TERMINADO CORRECTAMENTE"
di as result "============================================================"