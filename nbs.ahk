; Adjust to your needs
; ------------------------------------------------------------------------------
scannerName := "CanoScan LiDE 210" ; Title of popup window shown by Windows on button press
waitTime := 8250 ; Wait for scanner to return in ms
scriptPath := A_ScriptDir . "\" . "nbs.ps1"
; ------------------------------------------------------------------------------

DetectHiddenWindows, On
scanning := false
waiting := false
clock := A_NowUTC

Gui, m:New, +AlwaysOnTop -SysMenu, Neat Button Scan
Gui, m:Color, c666666
Gui, m:Font, s46 Q5 cf4f4f4, SegoeUI
Gui, m:Add, Text, w480 vStatus, Scanner
Gui, m:Show, AutoSize

Gosub, Main
return

+Esc::ExitApp

Main:
    SetTimer, UpdateStatus, 100
    Loop {
        WinWait, %scannerName%,, 1
        if (ErrorLevel = 0) {
            WinClose
            Gosub Scan
        }
    }
    return

UpdateStatus:
    if (scanning) {
        if (waiting) {
            statusText := "Waiting.."
            Gui, m:Color, c198a24
        } else {
            statusText := "Scanning.."
            Gui, m:Color, c275ac7
        }
    } else {
        statusText := "Ready.."
        Gui, m:Color, c5a5a5a
    }
    s := A_NowUTC
    s -= clock, Seconds
    GuiControl, m:, Status, % statusText . "  " . s
    return

Scan:
    scanning := true
    clock := A_NowUTC
    RunWait, pwsh.exe %scriptPath%,, Hide
    waiting := true
    Sleep, waitTime
    clock := A_NowUTC
    waiting := false
    scanning := false
    return
