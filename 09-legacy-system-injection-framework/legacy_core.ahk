#Requires AutoHotkey v2.0
; =====================================================================
; ARCHITECTURE TEMPLATE: LEGACY SYSTEM INJECTION CORE
; PROTECTED PORTFOLIO - GENERIC IMPLEMENTATION WITHOUT PRODUCTION DATA
; =====================================================================

Global TARGET_WINDOW_CLASS := "ahk_class ConsoleWindowClass"

ExecuteGenericPayload(SourcePath) {
    If !FileExists(SourcePath)
        Return False
        
    If !WinExists(TARGET_WINDOW_CLASS)
        Return False
        
    WinActivate(TARGET_WINDOW_CLASS)
    WinWaitActive(TARGET_WINDOW_CLASS, "", 5)
    
    ; Standard buffer initialization (Generic timing template)
    SetKeyDelay(100, 100) 
    
    Local InputData := FileRead(SourcePath)
    Loop Parse, InputData, "`n", "`r" {
        If (Trim(A_LoopField) == "")
            Continue
            
        Local SegmentedTokens := StrSplit(A_LoopField, ",")
        Local DataToken_A := SegmentedTokens[1]
        Local DataToken_B := SegmentedTokens[2]
        
        ; Sequential streaming to active system prompt
        SendEvent(DataToken_A)
        SendEvent("{ENTER}")
        Sleep(500) 
        
        If (DataToken_B != "") {
            SendEvent(DataToken_B)
            SendEvent("{ENTER}")
            Sleep(500)
        }
    }
    Return True
}
