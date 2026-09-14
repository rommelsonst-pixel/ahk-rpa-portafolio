#Requires AutoHotkey v2.0
#SingleInstance Force

; ============================================================================
;  DEMO 2: COM (Component Object Model)
;  Para qué sirve: hablar con aplicaciones "desde dentro", usando su objeto.
;  Brilla cuando la app EXPONE un objeto COM: Excel, Word, Outlook, Shell...
;  Aquí los conceptos son: ComObject -> objeto base; propiedades (con = para
;  asignar); métodos (con () para llamar).
; ============================================================================

try {
    ; --- 1. CREAR/ABRIR EL OBJETO COM DE EXCEL -----------------------------
    ; ComObject("Excel.Application") abre (o adjunta a) una instancia de Excel.
    ; Con esto tienes a Excel "en tu mano": es como una versión programable.
    excel := ComObject("Excel.Application")

    ; Visible := true  -> mostramos Excel para que veas lo que hace AHK.
    ; (Aquí sí usamos comillas de COM: "propiedad := valor")
    excel.Visible := true

    ; --- 2. LLAMAR MÉTODOS DE EXCEL ---------------------------------------
    ; Workbooks es una colección; .Add() crea un libro nuevo.
    libro := excel.Workbooks.Add()

    ; Worksheets.Item(1) -> la primera hoja.
    hoja := libro.Worksheets.Item(1)

    ; Escribir celdas: Range("A1").Value := "texto/número"
    hoja.Range("A1").Value := "7"
    hoja.Range("A2").Value := "3"
    hoja.Range("A3").Value := "7 + 3"
    hoja.Range("B1").Value := "=A1+A2"   ; fórmula que Excel calcula SOLO

    ; --- 3. LEER DE VUELTA (COM también devuelve valores) ------------------
    resultado := hoja.Range("B1").Value   ; Excel ya calculó la suma
    valorA1 := hoja.Range("A1").Value

    ; --- 4. DAR FORMATO: Rango de celdas como objetos ----------------------
    ; Tomamos OTRO rango y cambiamos propiedades de estilo.
    rango := hoja.Range("A1:B1")
    rango.Bold := true
    rango.Interior.Color := 0xFFFF99      ; color de fondo

    MsgBox("COM: A1 = " valorA1 "`nB1 (fórmula) = " resultado, "Demo COM", 64)

    ; --- 5. CERRAR SIN PREGUNTAR Y SALIR ----------------------------------
    ; Close(false) = no guardar. Quit() = cerrar Excel.
    libro.Close(false)
    excel.Quit()
} catch Error as err {
    MsgBox("Error: " err.Message, "Demo COM", 16)
}