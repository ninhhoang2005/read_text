#include <GuiConstants.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <ColorConstants.au3>
#include <ComboConstants.au3>
#include <SliderConstants.au3>
#include "includes/play-logo.au3"
logo(1)

Global $g_oSAPI
$g_oSAPI = ObjCreate("SAPI.SpVoice")
If @error Then
    MsgBox(16, "error", "you cannot read by sapi 5. Please try again")
    Exit
EndIf

GuiCreate("ReadTextV2.0(original version)", 300, 400)
GuiSetBkColor($COLOR_BLUE)
GuiCtrlCreateLabel("&enter text", 10, 5)
$entertext = GuiCtrlCreateEdit("", 10, 25, 280, 50)

GuiCtrlCreateLabel("&select voice", 10, 85)
$comboVoice = GuiCtrlCreateCombo("", 10, 105, 280, 20, BitOR($CBS_DROPDOWNLIST, $WS_VSCROLL))
PopulateVoiceComboBox($comboVoice)

GuiCtrlCreateLabel("&Volume", 10, 135)
$sliderVolume = GuiCtrlCreateSlider(10, 155, 330, 30, BitOR($TBS_AUTOTICKS, $WS_TABSTOP))
GUICtrlSetLimit($sliderVolume, 100, 0)
GUICtrlSetData($sliderVolume, 100)

GuiCtrlCreateLabel("&Rate", 10, 195)
$sliderRate = GuiCtrlCreateSlider(10, 215, 330, 30, BitOR($TBS_AUTOTICKS, $WS_TABSTOP))
GUICtrlSetLimit($sliderRate, 10, -10)
GUICtrlSetData($sliderRate, 0)

GuiCtrlCreateLabel("&Pitch", 10, 255)
$sliderPitch = GuiCtrlCreateSlider(10, 275, 330, 30, BitOR($TBS_AUTOTICKS, $WS_TABSTOP))
GUICtrlSetLimit($sliderPitch, 10, -10)
GUICtrlSetData($sliderPitch, 0)

$button = GuiCtrlCreateButton("Read&Text", 40, 150, 280, 30)
$message = GuiCtrlCreateButton("&my message", 200, 180, 80, 50)
$tts = GuiCtrlCreateButton("&Listen text", 50, 320, 230, 60)
$saveAudio = GuiCtrlCreateButton("&Save Audio", 50, 290, 230, 30)

$menu = GuiCtrlCreateMenu("help")
$menu1 = GuiCtrlCreateMenuItem("about...", $menu)
$menu3 = GuiCtrlCreateMenuItem("exit", $menu)
$menu2 = GuiCtrlCreateMenuItem("c&ontribute", $menu)
$SubMenu1 = GuiCtrlCreateMenu("tutorial", $menu)
$menuitem1 = GuiCtrlCreateMenuItem("vietnamese", $SubMenu1)
$menuitem2 = GuiCtrlCreateMenuItem("english", $SubMenu1)
$subMenu2 = GuiCtrlCreateMenu("contact with me", $menu)
$facebook = GuiCtrlCreateMenuItem("Facebook", $subMenu2)
$email = GuiCtrlCreateMenuItem("Email", $subMenu2)
$subMenu3 = GuiCtrlCreateMenu("rules", $menu)
$rules1 = GuiCtrlCreateMenuItem("vietnamese", $subMenu3)
$rules2 = GuiCtrlCreateMenuItem("english", $subMenu3)
GuiCtrlCreateLabel("press the alt key to go the menu help", 20, 380)
GuiSetState()

While 1
    Switch GuiGetMSG()
        Case $GUI_EVENT_CLOSE, $menu3
            SoundPlay("sounds/exit.wav", 1)
            Exit

        Case $rules1
            SoundPlay("sounds/enter.wav")
            Local $virules = "rules\rules_Vietnamese.txt"
            If FileExists($virules) Then
                Run("notepad.exe " & $virules)
            Else
                MsgBox(0, "error", "you cannot read the rules. Please try again")
            EndIf

        Case $rules2
            SoundPlay("sounds/enter.wav")
            Local $enrules = "rules\rules_english.txt"
            If FileExists($enrules) Then
                Run("notepad.exe " & $enrules)
            Else
                MsgBox(0, "error", "you cannot read the rules. Please try again")
            EndIf

        Case $menu1
            SoundPlay("sounds/enter.wav")
            MsgBox(64, "about", "ReadText version: 2.0, by developer vo dinh hung, this is original version. Thanks for using the software.")

        Case $message
            SoundPlay("sounds/message.wav")
            MsgBox(0, "message", "Hello everyone, it's the last time everyone uses ReadText software. Thank you everyone for using my software, this is the final version of the software I developed. I have to stop developing the software because I myself have no ideas for my software. If you have ideas or need to contact, please contact via the following applications: email: vodinhhungtnlg@gmail.com, facebook: Phaolo Vo Dinh Hung")

        Case $facebook
            SoundPlay("sounds/enter.wav")
            ShellExecute("https://www.facebook.com/profile.php?id=100083295244149")

        Case $email
            ShellExecute("https://mail.google.com/mail/u/0/?fs=1&tf=cm&source=mailto&to=vodinhhungtnlg@gmail.com")

        Case $menuitem2
            SoundPlay("sounds/enter.wav")
            Local $cFilePath = "readme\ReadmeEnglish.txt"
            If FileExists($cFilePath) Then
                Run("notepad.exe " & $cFilePath)
            Else
                MsgBox(0, "error", "you cannot read tutorial. Please try again.")
            EndIf

        Case $menuitem1
            SoundPlay("sounds/enter.wav")
            Local $sFilePath = "readme\ReadmeVietnamese.txt"
            If FileExists($sFilePath) Then
                Run("notepad.exe " & $sFilePath)
            Else
                MsgBox(0, "error", "you cannot read tutorial. Please try again.")
            EndIf

        Case $button
            SoundPlay("sounds/enter.wav")
            ReadText()

        Case $tts
            Local $sSelectedVoice = GuiCtrlRead($comboVoice)
            Local $ok = GuiCtrlRead($entertext)
            Local $vol = GuiCtrlRead($sliderVolume)
            Local $rate = GuiCtrlRead($sliderRate)
            Local $pitch = GuiCtrlRead($sliderPitch)

            If Not IsObj($g_oSAPI) Then
                MsgBox(16, "error", "The SAPI.SpVoice object does not exist or has been destroyed. Please restart the application.")
                Exit
            EndIf

            If $sSelectedVoice <> "" Then
                For $oToken In $g_oSAPI.GetVoices()
                    If $oToken.GetDescription() = $sSelectedVoice Then
                        $g_oSAPI.Voice = $oToken
                        ExitLoop
                    EndIf
                Next
            EndIf

            $g_oSAPI.Volume = $vol
            $g_oSAPI.Rate = $rate

            If StringStripWS($ok, 8) <> "" Then
                Local $ssml = '<sapi><pitch middle="' & $pitch & '">' & $ok & '</pitch></sapi>'
                $g_oSAPI.Speak($ssml, 1)
            Else
                SoundPlay("sounds/enter.wav")
                MsgBox(0, "warning", "please enter your text")
            EndIf

        Case $saveAudio
            Local $sSelectedVoice = GuiCtrlRead($comboVoice)
            Local $ok = GuiCtrlRead($entertext)
            Local $vol = GuiCtrlRead($sliderVolume)
            Local $rate = GuiCtrlRead($sliderRate)
            Local $pitch = GuiCtrlRead($sliderPitch)

            If StringStripWS($ok, 8) = "" Then
                SoundPlay("sounds/enter.wav")
                MsgBox(0, "warning", "please enter your text")
                ContinueLoop
            EndIf
            SoundPlay("sounds/enter.wav")
            Local $sFile = FileSaveDialog("Save audio as...", @ScriptDir, "Wave files (*.wav)", 16, "output.wav")
            If @error Or $sFile = "" Then ContinueLoop

            Local $oStream = ObjCreate("SAPI.SpFileStream")
            $oStream.Open($sFile, 3, False)

            If $sSelectedVoice <> "" Then
                For $oToken In $g_oSAPI.GetVoices()
                    If $oToken.GetDescription() = $sSelectedVoice Then
                        $g_oSAPI.Voice = $oToken
                        ExitLoop
                    EndIf
                Next
            EndIf

            $g_oSAPI.Volume = $vol
            $g_oSAPI.Rate = $rate

            $g_oSAPI.AudioOutputStream = $oStream
            Local $ssml = '<sapi><pitch middle="' & $pitch & '">' & $ok & '</pitch></sapi>'
            $g_oSAPI.Speak($ssml)
            $oStream.Close()
            $g_oSAPI.AudioOutputStream = 0
            MsgBox(64, "success", "Audio saved successfully as WAV file.")
exit

        Case $menu2
            SoundPlay("sounds/enter.wav")
            contribute()
    EndSwitch
WEnd

Func ReadText()
    Local $text = GuiCtrlRead($entertext)
    If StringStripWS($text, 8) = "" Then
        SoundPlay("sounds/enter.wav")
        MsgBox(0, "warning", "please enter your text")
        Return
    EndIf
    Local $displayGui = GuiCreate("text", 500, 500)
    Local $document = GUICtrlCreateEdit($text, 20, 20, 450, 400, BitOR($ES_AUTOVSCROLL, $ES_READONLY, $WS_VSCROLL, $WS_TABSTOP))
    Local $btnClose = GUICtrlCreateButton("&Close", 200, 430, 100, 30, $WS_TABSTOP)
    GuiSetState(@SW_SHOW, $displayGui)
    While 1
        Switch GuiGetMSG()
            Case $GUI_EVENT_CLOSE, $btnClose
                GuiDelete($displayGui)
                ExitLoop
        EndSwitch
    WEnd
EndFunc

Func PopulateVoiceComboBox($hCombo)
    If Not IsObj($g_oSAPI) Then Return
    Local $aVoices = $g_oSAPI.GetVoices()
    For $oToken In $aVoices
        GuiCtrlSetData($hCombo, $oToken.GetDescription())
    Next
    If UBound($aVoices) > 0 Then
        GuiCtrlSetData($hCombo, $aVoices.Item(0).GetDescription())
    EndIf
EndFunc

Func contribute()
    $congui = GuiCreate("contribute", 700, 700)
    GuiSetBkColor($COLOR_RED)
    $con = FileRead("contribute.txt")
    Local $conedit = GUICtrlCreateEdit($con, 20, 20, 650, 600, BitOR($ES_AUTOVSCROLL, $ES_READONLY, $WS_VSCROLL, $WS_TABSTOP))
    Local $btnClose = GUICtrlCreateButton("&Close", 300, 630, 100, 30, $WS_TABSTOP)
    GuiSetState(@SW_SHOW, $congui)
    While 1
        Switch GuiGetMSG()
            Case $GUI_EVENT_CLOSE, $btnClose
                GuiDelete($congui)
                ExitLoop
        EndSwitch
    WEnd
EndFunc