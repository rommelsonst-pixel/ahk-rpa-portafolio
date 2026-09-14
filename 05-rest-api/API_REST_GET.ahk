#Requires AutoHotkey v2.0
#SingleInstance Force

try {
    ; 1. Crear el objeto de red nativo de Windows para peticiones HTTP
    http := ComObject("MSXML2.XMLHTTP.6.0")
    
    ; 2. Definir la URL de la API REST (Esta dirección responde los datos de la red del cliente)
    URL_API := "https://httpbin.org"
    
    ; 3. Abrir el canal de comunicación usando el método GET (Universalmente permitido)
    http.Open("GET", URL_API, false)
    
    ; 4. Inyectar las cabeceras estándar
    http.SetRequestHeader("User-Agent", "Bot_QA_Automation_AHK")
    
    ; 5. ¡FUEGO POR LA RED! Enviar la consulta en completo silencio
    http.Send()
    
    ; 6. Capturar el estatus de la red (200 significa Éxito Total en internet)
    if (http.Status == 200) {
        RespuestaServidor := http.ResponseText
        MsgBox("¡API REST Respondida con Éxito Absoluto (Status 200)!`n`nEl servidor de internet nos devolvió estas cabeceras JSON en silencio de fondo:`n`n" RespuestaServidor, "QA API Automation", 64)
    } else {
        MsgBox("El servidor respondió pero con otro estatus. Status: " http.Status, "Error API", 48)
    }

} catch Error as err {
    MsgBox("Fallo crítico: No se pudo establecer la conexión con la API REST de internet.`n`nDetalles: " err.Message, "Error de Red", 16)
}
