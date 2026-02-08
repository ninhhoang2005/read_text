Func logo($playing=0)
    If $playing=1 Then
        Local $loadingfrm=GUICreate("loading",300,300, -1, -1, $WS_POPUP)
        GUISetBkColor($COLOR_BLUE)
        GuiCtrlCreateLabel("Welcome to VDH Productions", 10, 10, 280, 20)
        GUISetState()
        Sleep(500) ; Reduced from 2000
        Local $welcome_text=GUICtrlCreateLabel("WELCOME", 50, 120, 200, 60)
        GUICtrlSetFont($welcome_text, 30, 700, "Arial")
        GUICtrlSetColor($welcome_text, $COLOR_WHITE)
        SoundPlay("sounds/logo.wav")
        Sleep(1500) ; Reduced from 4000
        GUICtrlDelete($welcome_text)
        GUISetBkColor($COLOR_RED)
        Sleep(500) ; Reduced from 5000
        GUIDelete($loadingfrm)
    EndIf
EndFunc
