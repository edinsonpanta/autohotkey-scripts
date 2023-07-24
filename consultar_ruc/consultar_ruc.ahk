; Reto 1 : 
; Reto 2 : Automatizar con Chrome.hk 
; Reto 3 : Selenium con python

#NoEnv
#NoEnv
#Warn
#Persistent

; ============================================== Validar si está instalada la biblioteca COM de AutoHotkey ==============================================
; Instalar la biblioteca COM de AutoHotkey
ComObjMissing := -2147220990
ComObjActive := -2147220991
ComObjError := -2147220992

; Código para instalar la biblioteca COM (Component Object Model) de AutoHotkey
verificar_biblioteca_COM() {

    MsgBox, 36, Biblioteca COM no encontrada, ¿quieres instalarla ahora?
    IfMsgBox, No
    {
        MsgBox, 48, Biblioteca COM requerida, la biblioteca COM es necesaria para que se ejecute este script. Instálalo y vuelve a ejecutar el script.
        ExitApp
    }
    Run, regsvr32 /s "%A_AhkPath%"
    MsgBox, 48, Biblioteca COM instalada, la biblioteca COM se ha instalado.

}

; =========================================================================================================================================
abrir_browser(browser, url_target) {
    browser.Visible := true
    browser.Navigate(url_target)
    return browser
}

consultar_sunat(browser, numero_ruc){

    estado_peticion := "OK"
    fecha_inicio := A_DD . "/" . A_MM . "/" . A_YYYY . " " . A_Hour . ":" . A_Min . ":" . A_Sec

    ; Esperar a que la página se cargue completamente antes de continuar
    While browser.Busy || browser.ReadyState != 4
        Sleep 100

    ; Obtener el elemento de la caja de texto
    ruc_element := browser.document.getElementById("txtRuc")

    ; Ingresar el dato en la caja de texto
    ruc_element.value := numero_ruc

    ; Obtener el elemento del botón de búsqueda
    searchButton := browser.document.getElementById("btnAceptar")

    ; Hacer clic en el botón de búsqueda
    searchButton.Click()

    ;==================== Logica para realizar el scrapping de la pagina =========================
    Sleep 3000

    ; Esperar a que la nueva página se cargue completamente
    While browser.Busy || browser.ReadyState != 4
        Sleep 100

    all_elements := browser.document.getElementsByClassName("list-group-item")

    try {
        ruc_element := all_elements[0]
        tipo_contribuyente_element := all_elements[1]
        tipo_documento_element := all_elements[2]
        nombre_comercial_element := all_elements[3]

        ruc_text := ruc_element.innerText
        tipo_contribuyente_text := tipo_contribuyente_element.innerText
        tipo_documento_text := tipo_documento_element.innerText
        nombre_comercial_text := nombre_comercial_element.innerText

        ruc_value := StrSplit(ruc_text, ":")[2]
        tipo_contribuyente_text := StrSplit(tipo_contribuyente_text, ":")[2]
        tipo_documento_text := StrSplit(tipo_documento_text, ":")[2]
        nombre_comercial_text := StrSplit(nombre_comercial_text, ":")[2]
    } catch {
        ruc_value := "VALOR NO ENCONTRADO"
        tipo_contribuyente_text := "VALOR NO ENCONTRADO"
        tipo_documento_text := "VALOR NO ENCONTRADO"
        nombre_comercial_text := "VALOR NO ENCONTRADO"
        estado_peticion := "NOT OK"
    }

    fecha_fin := A_DD . "/" . A_MM . "/" . A_YYYY . " " . A_Hour . ":" . A_Min . ":" . A_Sec

    detalles_text := StrReplace(numero_ruc, "`r`n", "") . "|" . StrReplace(ruc_value, "`r`n", "") . "|" . StrReplace(tipo_contribuyente_text, "`r`n", "") . "|" . StrReplace(tipo_documento_text, "`r`n", "") . "|" . StrReplace(nombre_comercial_text, "`r`n", "") . "|" . StrReplace(fecha_inicio, "`r`n", "") . "|" . StrReplace(fecha_fin, "`r`n", "") . "|" . StrReplace(estado_peticion, "`r`n", "")
    return detalles_text
}

main(browser, url_target){

    listado_rucs_csv := "listado_rucs.csv"
    listado_rucs_file := FileOpen(listado_rucs_csv, "r")
    if (listado_rucs_file) {

        ; Leer el contenido del archivo línea por línea
        contenido := ""
        while (!listado_rucs_file.AtEOF()) {
            linea := listado_rucs_file.ReadLine()
            contenido .= linea . "\n"
        }

        ; Procesar el contenido del archivo
        filas := StrSplit(Trim(contenido), "\n") ; Dividir el contenido en filas

        cabecera_csv := "ruc|descripcion_ruc|tipo_contribuyente|tipo_documento|nombre_comercial|fecha_inicio|fecha_fin|estado"
        cuerpo_csv := ""

        Loop, % filas.Length() {
            fila := filas[A_Index + 1]
            columnas := StrSplit(fila, "|") ; Dividir la fila en columnas utilizando la coma como delimitador

            ; Acceder a los valores de las columnas
            numero_ruc := columnas[1]
            if(numero_ruc == ""){
                Continue
            }else if(numero_ruc == "ruc"){
                Continue
            }else{
                browser_init := abrir_browser(browser, url_target)
                detalles_sunat := consultar_sunat(browser_init, numero_ruc)
                cuerpo_csv .= detalles_sunat . "`r`n"
            }
        }

        detallles_ruc_file := FileOpen("detalles_ruc.csv", "w")
        full_cuerpo_csv := cabecera_csv . "`r`n" . cuerpo_csv
        detallles_ruc_file.Write(full_cuerpo_csv)
        MsgBox % "El proceso a finalizado...!!"

    } else {
        MsgBox "No se pudo abrir el archivo..."
    }

}

; =================================================== PARTE INICIAL DEL PROCESO ========================================================
; Verificar si la biblioteca COM de AutoHotkey está instalada
verificar_biblioteca_COM()

; Iniciar proceso
browser := ComObjCreate("InternetExplorer.Application")

url_target := "https://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsruc/FrameCriterioBusquedaWeb.jsp"

main(browser, url_target)

ExitApp