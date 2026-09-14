#Requires AutoHotkey v2.0
#SingleInstance Force

; --- CONFIGURACIÓN DE CORREO DIRECTO (Puro Código) ---
CorreoDesde     := "TU_CORREO@gmail.com"
CorreoPara      := "TU_CORREO@gmail.com" ; Te lo puedes mandar a ti mismo para validar
Asunto          := "Reporte E2E - AHK v2 (Metodo CDO/SMTP)"
Cuerpo          := "Hola, adjunto el procesamiento finalizado. Enviado directo por red sin abrir Outlook."
RutaAdjunto     := A_ScriptDir "\Reporte.xls"

; --- CREDENCIALES DE GMAIL ---
ServidorSMTP    := "smtp.gmail.com"
PuertoSMTP      := 465 ; Puerto SSL nativo obligado por Google
UsuarioGmail    := "TU_CORREO@gmail.com"
ContrasenaGmail := "xxxx xxxx xxxx xxxx" ; <--- INYECTA AQUÍ TUS 16 LETRAS AMARILLAS DE GOOGLE

; Prefijo comun de los campos de configuracion CDO
Cfg := "http://schemas.microsoft.com/cdo/configuration/"

try {
    ; 1. Crear los objetos de red CDO invisibles dentro del núcleo de Windows
    objEmail := ComObject("CDO.Message")
    objConfig := ComObject("CDO.Configuration")
    Campos := objConfig.Fields
    
    ; 2. Configurar los parámetros de transporte del túnel SMTP
    Campos.Item(Cfg "sendusing") := 2
    Campos.Item(Cfg "smtpserver") := ServidorSMTP
    Campos.Item(Cfg "smtpserverport") := PuertoSMTP
    Campos.Item(Cfg "smtpauthenticate") := 1 ; Requiere clave
    Campos.Item(Cfg "sendusername") := UsuarioGmail
    Campos.Item(Cfg "sendpassword") := ContrasenaGmail
    Campos.Item(Cfg "smtpusessl") := True ; Forzar cifrado SSL
    Campos.Update()
    
    ; 3. Ensamblar los datos y el archivo Excel en la memoria RAM
    objEmail.Configuration := objConfig
    objEmail.From := CorreoDesde
    objEmail.To := CorreoPara
    objEmail.Subject := Asunto
    objEmail.Textbody := Cuerpo
    
    if FileExist(RutaAdjunto)
        objEmail.AddAttachment(RutaAdjunto)
        
    ; 4. ¡FUEGO! Empujar el correo directo a internet
    objEmail.Send()
    
    MsgBox("¡Método Tanque Completado!`nEl correo y el archivo se enviaron directo por la red en total secreto.", "SMTP Exitoso", 64)

} catch {
    MsgBox("Fallo de Red: No se pudo enviar por SMTP.`nVerifica tu conexión a internet o tu clave de aplicación de Google.", "Error de Envío", 16)
}
