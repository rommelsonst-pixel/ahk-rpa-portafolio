#Requires AutoHotkey v2.0
#SingleInstance Force

; ============================================================================
;  DEMO 3: COM - ADJUNTARSE a una instancia YA ABIERTA (GetActiveObject)
;  Diferencia clave:
;    ComObject("Excel.Application")   -> CREA una instancia nueva
;    ComObjActive("Excel.Application") -> se ADJUNTA a la que ya corre,
;                                         sin abrir otra. (En tu build de AHK
;                                         el nombre es ComObjActive, no ComObjectActive)
;  Si tienes 2 Excel abiertos por error, es porque usaste "crear" en vez de
;  "adjuntar". Este demo crea una, y luego se adjunta a ELLA desde otra
;  conexión COM (simulando dos scripts hablando con el mismo Excel).
; ============================================================================

try {
    ; --- 1. SCRIPT "CREADOR": abre Excel y escribe --------------------------
    excel := ComObject("Excel.Application")     ; crea (o toma) una instancia
    excel.Visible := true
    libro := excel.Workbooks.Add()
    hoja := libro.Worksheets.Item(1)
    hoja.Range("A1").Value := "Escrito por el CREADOR"
    hoja.Range("A2").Value := 42

    ; --- 2. SCRIPT "ADJUNTO": se conecta al MISMO Excel que ya está abierto --
    adjunto := ComObjActive("Excel.Application")   ; GetActiveObject
    libroAdjunto := adjunto.Workbooks.Item(1)         ; ve el mismo libro
    hojaAdjunta := libroAdjunto.Worksheets.Item(1)
    leido := hojaAdjunta.Range("A2").Value            ; lee lo que escribió el otro
    hojaAdjunta.Range("A3").Value := "Modificado por el ADJUNTO: " leido

    MsgBox("A2 fue leída por la conexión adjunta = " leido "`nA3 fue escrita por la conexión adjunta.", "Demo GetActiveObject", 64)

    ; --- 3. CERRAR (bloqueado hasta cerrar el MsgBox) -----------------------
    libro.Close(false)
    excel.Quit()
} catch Error as err {
    MsgBox("Error: " err.Message, "Demo GetActiveObject", 16)
}