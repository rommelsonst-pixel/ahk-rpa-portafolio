#Requires AutoHotkey v2.0
; --------------------------------------------------------------------------------
; Portafolio RPA - Proyecto 07: Automatización de SAP R/3 RMCA vía Scripting COM
; Qué hace: Se conecta a la sesión activa de SAP, ejecuta la transacción RMCA,
;           procesa datos de cuenta y genera logs de evidencias automatizados.
; --------------------------------------------------------------------------------

TraySetIcon("shell32.dll", 277) ; Icono de engranaje de red

; 1. Conexión al objeto COM nativo de SAP GUI
try {
    ; Captura la aplicación SAP en ejecución
    SapGuiAuto := ComObjGet("SAPGUI")
    AppSap := SapGuiAuto.GetScriptingEngine
    
    ; Obtiene la conexión activa y la sesión indexada (0 por defecto)
    Conexion := AppSap.Children(0)
    Sesion := Conexion.Children(0)
} catch Error as err {
    MsgBox("Error crítico: No se detectó una sesión activa de SAP GUI.`nPor favor, inicia sesión en SAP antes de ejecutar el bot.", "SAP RPA - Error", 16)
    ExitApp
}

; 2. Orquestación del flujo transaccional dentro de SAP RMCA
try {
    ; Bloquea la actualización visual para máxima velocidad (Ejecución silenciosa)
    Sesion.BusyIndicatorCancelable := true
    
    ; Enviar comando directo para ir a la transacción de consulta de cuentas RMCA
    ; Nota: Reemplazar 'FPL9' o el código específico de tu entorno RMCA corporativo
    Sesion.StartTransaction("FPL9") 
    
    ; Captura los campos de texto del formulario usando sus IDs nativos de SAP
    ; Modifica los valores para interactuar con los campos de cuenta contractual
    CamposFormulario := Sesion.FindById("wnd[0]/usr/txtGPART-LOW") ; ID de interlocutor comercial
    CamposFormulario.Text := "20072026" ; Número de cuenta/cliente de prueba
    
    ; Simula la presión de la tecla Enter nativa en SAP para procesar el formulario
    Sesion.FindById("wnd[0]").SendVKey(0) 
    
    ; 3. Extracción de datos del Grid/Tabla de SAP para validación backend
    TablaResultados := Sesion.FindById("wnd[0]/usr/cntlGRID1/shellCONT/shell")
    if (IsObject(TablaResultados)) {
        TotalFilas := TablaResultados.RowCount
        MsgBox("Conexión SAP Exitosa.`nFilas encontradas en Grid RMCA: " TotalFilas, "SAP RPA - Éxito", 64)
    }
    
    ; 4. Captura automática de evidencia de la ventana de SAP
    DirectorioEvidencias := A_ScriptDir "\Evidencias_SAP"
    if !DirExist(DirectorioEvidencias)
        DirCreate(DirectoriosEvidencias)
        
    ; Guarda una captura de pantalla nativa de la ventana actual de SAP
    Sesion.FindById("wnd[0]").HardCopy(DirectorioEvidencias "\Caso_Procesado.png", 2)

} catch Error as err {
    MsgBox("Error durante la automatización del formulario SAP RMCA:`n" err.Message, "SAP RPA - Falla", 16)
}

ExitApp
