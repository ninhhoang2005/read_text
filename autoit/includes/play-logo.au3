Func logo($playing=0)
if $playing=1 Then
$loadingfrm=GUICreate("loading",300,300)
GUISetBkColor($COLOR_BLUE)
GuiCtrlCreateLabel("welcome to vdh productions", 10, 5, 30)
GUISetState()
Sleep(2000)
$welcome_text=GUICtrlCreateLabel("welcome",80,150,150,80)
GUICtrlSetFont($welcome_text,30,700,"Arial")
GUICtrlSetColor($welcome_text,$COLOR_WHITE)
SoundPlay("sounds/logo.wav")
Sleep(4000)
GUICtrlDelete($welcome_text)
GUISetBkColor($COLOR_RED)
sleep(5000)
GUIDelete($loadingfrm)
Else
;start
EndIf
EndFunc
