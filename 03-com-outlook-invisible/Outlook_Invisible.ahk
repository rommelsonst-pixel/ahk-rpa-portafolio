#Requires AutoHotkey v2.0
#SingleInstance Force

; --- CONFIGURACIÓN DE RUTAS Y ADJUNTOS ---
RutaAdjunto := A_ScriptDir "\Guion_Grupo_NOM_BANCO_FECHA_01.xls"

try {
    ; 1. Conectar con la aplicación de Outlook en segundo plano (Inicia INVISIBLE de fábrica)
    oOutlook := ComObject("Outlook.Application")
    
    ; 2. Crear un elemento de correo nuevo (0 = olMailItem)
    oMail := oOutlook.CreateItem(0)
    
    ; 3. Armar los campos del reporte en la memoria RAM
    oMail.To := "TU_CORREO@correo.com"
    oMail.Subject := "Reporte E2E - AHK v2 (Outlook)"
    oMail.Body := "Hola, adjunto el procesamiento automatizado enviado en total silencio."
    
    ; 4. Inyectar el adjunto si el archivo existe en el disco duro
    if FileExist(RutaAdjunto)
        oMail.Attachments.Add(RutaAdjunto)
        
    ; 5. EL COMANDO MAESTRO: Enviar en segundo plano 
    ; (NUNCA uses oMail.Display() para mantener la invisibilidad absoluta)
    oMail.Send()
    
    ; 6. LIMPIEZA DE RAM EXIGIDA: Destruimos las conexiones para liberar recursos
    oMail := ""
    oOutlook := ""
    
    MsgBox("¡Éxito!`nEl correo fue enviado mediante Outlook de forma 100% invisible.", "QA Automation", 64)

} catch {
    ; En caso de fallo, limpiamos las variables para evitar procesos fantasmas
    oMail := ""
    oOutlook := ""
    MsgBox("Fallo crítico: No se pudo conectar con Outlook o la cuenta no está configurada.", "Error", 16)
}
