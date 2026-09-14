#Requires AutoHotkey v2.0
#SingleInstance Force
#Include UIA.ahk

; ============================================================================
;  DEMO 1: UIA (UI Automation)
;  Para qué sirve: clics/lecturas INVISIBLES en ventanas que no se dejan
;  controlar con la API normal de AHK. Ideal para apps de terceros, la
;  calculadora, el explorador, etc. NO necesita que la app exponga COM.
;
;  Conceptos básicos del mundo UIA:
;    1. ElementFromHandle  -> obtener la RAÍZ de la ventana (un "árbol")
;    2. FindFirst          -> buscar UN hijo que cumpla una condición
;    3. .Click()           -> disparar el clic sin mover el mouse
; ============================================================================

try {
    ; --- 1. APUNTAR A LA VENTANA (por su proceso real) ---------------------
    ; "ahk_exe" + nombre del exe filtran por proceso (el que viste en
    ; UIATreeInspector: "Calculadora (win32calc.exe)")
    CalcTarget := "ahk_exe win32calc.exe"
    if !WinExist(CalcTarget) {
        Run("win32calc.exe")
        WinWait(CalcTarget)
        Sleep(1500)
    }

    ; --- 2. OBTENER EL "ÁRBOL" UIA DE LA VENTANA ---------------------------
    ; ElementFromHandle devuelve el elemento raíz. A partir de ahí podemos
    ; "bajar" por el árbol buscando botones, textos, etc.
    Raiz := UIA.ElementFromHandle(CalcTarget)

    ; --- 3. BUSCAR ELEMENTOS CON UNA CONDICIÓN (FIND) ----------------------
    ; FindFirst busca el PRIMER elemento que cumpla la condición.
    ; Leemos el AutomationId (el mismo que muestras en UIATreeInspector).
    BotonSiete := Raiz.FindFirst({AutomationId: "137"})
    BotonSumar := Raiz.FindFirst({AutomationId: "93"})
    BotonTres  := Raiz.FindFirst({AutomationId: "133"})
    BotonIgual := Raiz.FindFirst({AutomationId: "121"})

    ; --- 4. CLIC INVISIBLE ------------------------------------------------
    ; FindFirst devuelve "" si no encuentra el elemento.
    if (BotonSiete && BotonSumar && BotonTres && BotonIgual) {
        BotonSiete.Click()        ; presiona 7
        Sleep(400)
        BotonSumar.Click()        ; presiona +
        Sleep(400)
        BotonTres.Click()         ; presiona 3
        Sleep(400)
        BotonIgual.Click()        ; presiona =
        Sleep(400)
        MsgBox("UIA: 7 + 3 presionado sin mover el mouse.", "Demo UIA", 64)
    } else {
        MsgBox("No se encontraron los botones.", "Demo UIA", 48)
    }
} catch Error as err {
    MsgBox("Error: " err.Message, "Demo UIA", 16)
}