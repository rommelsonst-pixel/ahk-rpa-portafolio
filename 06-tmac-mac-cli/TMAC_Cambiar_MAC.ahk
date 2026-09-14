#Requires AutoHotkey v2.0
#SingleInstance Force

; ============================================================================
;  TMAC v6 - cambio de MAC por CLI (la vía FIABLE)
;
;  Descubrimiento clave: TMAC es una app Delphi que NO expone su árbol UIA
;  (FindFirst/AutomationId fallan). En cambio TMAC trae su propia interfaz
;  de línea de comandos documentada en CLIHelp.txt.
;
;  Comando:
;    TMAC.exe -n "Wi-Fi" -r02 -s -re
;      -n  "Wi-Fi"  = nombre del adaptador de red
;      -r02         = MAC aleatoria PERSISTENTE con primer octeto "02"
;      -s           = modo silencioso (sin cuadros emergentes)
;      -re          = reinicia la conexión para aplicar (puede caer la red ~seg)
;
;  Otras opciones utiles:
;    -r                     aleatoria sin fijar el primer octeto
;    -m 00:11:22:33:44:55   MAC fija (persistente)
;    -nm 00:11:22:33:44:55  MAC no persistente (se restaura al deshabilitar)
;    -ro                    restaurar la MAC original
; ============================================================================

TmacExe   := "C:\Program Files (x86)\Technitium\TMACv6.0\TMAC.exe"
Adaptador := "Wi-Fi"

try {
    codigo := RunWait('"' TmacExe '" -n "' Adaptador '" -r02 -s -re')
    if (codigo = 0) {
        MsgBox("MAC aleatoria aplicada en " Adaptador ".", "TMAC CLI", 64)
    } else {
        MsgBox("TMAC devolvió el código " codigo ".", "TMAC CLI", 48)
    }
} catch Error as err {
    MsgBox("No se pudo ejecutar TMAC: " err.Message, "TMAC CLI", 16)
}