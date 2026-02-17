#include <GuiConstants.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <ColorConstants.au3>
#include <ComboConstants.au3>
#include <SliderConstants.au3>
#include <GuiMenu.au3>
#include <WindowsConstants.au3>
#include <Constants.au3>
#include <File.au3>
#include "includes/play-logo.au3"

Global $oErrorHandler = ObjEvent("AutoIt.Error", "_ErrFunc")

Global $sAppVersion = "2.0"
Global Const $SVSFlagsAsync = 1
Global Const $SVSFPurgeBeforeSpeak = 2
Global $isPaused = False
Global $sConfigFile = @ScriptDir & "\ReadText.ini"

Global $bAutoUpdate = False
Global $bAutoClipboard = False

Global $sGoogleVoiceExe = @ScriptDir & "\lib\google_voice.exe"
Global $sSpeakExe = @ScriptDir & "\lib\SpeakToText.exe"

Global $g_aVoiceList[200][2]
Global $g_iVoiceCount = 0

FileChangeDir(@ScriptDir)
logo(1)

Global $g_oSAPI
$g_oSAPI = ObjCreate("SAPI.SpVoice")
If @error Then
    MsgBox(16, "Error", "Cannot initialize SAPI 5. Please check your Windows Speech settings.")
    Exit
EndIf

Global $hGUI = GuiCreate("ReadTextV" & $sAppVersion & "(original version)", 350, 540)
GuiSetBkColor($COLOR_BLUE)

GuiCtrlCreateLabel("&enter text", 10, 5)
$entertext = GuiCtrlCreateEdit("", 10, 25, 280, 50)

GuiCtrlCreateLabel("&select voice", 10, 85)
$comboVoice = GuiCtrlCreateCombo("", 10, 105, 280, 20, BitOR($CBS_DROPDOWNLIST, $WS_VSCROLL))
_PopulateVoiceComboBox($comboVoice)

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
$message = GuiCtrlCreateButton("m&y message", 200, 180, 80, 50)
$saveText = GuiCtrlCreateButton("Save Text, Ctrl+S", 50, 230, 230, 30)
$openText = GuiCtrlCreateButton("Open Text Files, Ctrl+O", 50, 260, 230, 30)
$saveAudio = GuiCtrlCreateButton("Save &Audio", 50, 290, 230, 30)
$tts = GuiCtrlCreateButton("&Listen text", 50, 320, 230, 30)
$btnPause = GuiCtrlCreateButton("Pause", 120, 310, 100, 35)
$btnStop = GuiCtrlCreateButton("Stop", 230, 310, 110, 35)

$btnGetClipboard = GuiCtrlCreateButton("Retrieve text from &clipboard", 50, 355, 230, 30)
$btnSpeakToText = GuiCtrlCreateButton("Speak to Text (Microphone)", 50, 390, 230, 30)
$btnMenuHelp = GuiCtrlCreateButton("&Menu", 50, 430, 230, 30)

$dummyMenu = GuiCtrlCreateDummy()
$contextMenu = GuiCtrlCreateContextMenu($dummyMenu)
$menu1 = GuiCtrlCreateMenuItem("about...", $contextMenu)
$menuUpdate = GuiCtrlCreateMenuItem("checked for &updates", $contextMenu)
$menu2 = GuiCtrlCreateMenuItem("c&ontribute", $contextMenu)
$menuChangelog = GuiCtrlCreateMenuItem("view changelog", $contextMenu)
$SubMenu1 = GuiCtrlCreateMenu("tutorial", $contextMenu)
$menuitem1 = GuiCtrlCreateMenuItem("vietnamese", $SubMenu1)
$menuitem2 = GuiCtrlCreateMenuItem("english", $SubMenu1)

$subMenu2 = GuiCtrlCreateMenu("contact with me", $contextMenu)
$facebook = GuiCtrlCreateMenuItem("Facebook", $subMenu2)
$email = GuiCtrlCreateMenuItem("Email", $subMenu2)

$subMenu3 = GuiCtrlCreateMenu("rules", $contextMenu)
$rules1 = GuiCtrlCreateMenuItem("vietnamese", $subMenu3)
$rules2 = GuiCtrlCreateMenuItem("english", $subMenu3)
$menuSettings = GuiCtrlCreateMenuItem("Settings...	Ctrl+shift+s", $contextMenu)
GuiCtrlCreateMenuItem("", $contextMenu)
$menu3 = GuiCtrlCreateMenuItem("exit", $contextMenu)

_LoadConfig()

GuiSetState()

Local $aAccelKeys[3][2] = [["^s", $saveText], ["^o", $openText], ["^+s", $menuSettings]]
GUISetAccelerators($aAccelKeys)
If $bAutoClipboard Then
    _GetClipboardText(True)
EndIf

If $bAutoUpdate Then
    _CheckGithubUpdate()
EndIf

While 1
    Switch GuiGetMSG()
        Case $GUI_EVENT_CLOSE, $menu3
            _SaveConfig()
            SoundPlay(@ScriptDir & "\sounds\exit.wav", 1)
            Exit

        Case $menuChangelog
            SoundPlay("sounds/enter.wav")
            _ShowChangelog()

        Case $btnPause
            If Not IsGoogleVoice(GuiCtrlRead($comboVoice)) Then
                If $isPaused Then
                    $g_oSAPI.Resume()
                    $isPaused = False
                    GUICtrlSetData($btnPause, "Pause")
                Else
                    $g_oSAPI.Pause()
                    $isPaused = True
                    GUICtrlSetData($btnPause, "Resume")
                EndIf
            EndIf

        Case $btnStop
            $g_oSAPI.Speak("", $SVSFPurgeBeforeSpeak)
            ; Kill process Google
            ProcessClose("google_voice.exe")
            ; Kill process 32-bit bridge
            ProcessClose("cscript.exe")

            $isPaused = False
            GUICtrlSetData($btnPause, "Pause")

        Case $menuSettings
            SoundPlay("sounds/enter.wav")
            _ShowSettings()

        Case $btnGetClipboard
            SoundPlay("sounds/enter.wav")
            _GetClipboardText(False)

        Case $btnSpeakToText
            _PerformSpeakToText()

        Case $btnMenuHelp
             SoundPlay("sounds/enter.wav")
             Local $hMenuHandle = GuiCtrlGetHandle($contextMenu)
             _GUICtrlMenu_TrackPopupMenu($hMenuHandle, $hGUI)

        Case $menuUpdate
            SoundPlay("sounds/enter.wav")
            _CheckGithubUpdate()

        Case $rules1
            SoundPlay("sounds/enter.wav")
            _ShowFileContent("Rules (Vietnamese)", "rules\rules_Vietnamese.txt")

        Case $rules2
            SoundPlay("sounds/enter.wav")
            _ShowFileContent("Rules (English)", "rules\rules_english.txt")

        Case $menu1
            SoundPlay("sounds/enter.wav")
            MsgBox(64, "about", "ReadText version: " & $sAppVersion & ", by developer vo dinh hung.")

        Case $message
            SoundPlay("sounds/message.wav")
            MsgBox(0, "message", "Hello everyone!")

        Case $facebook
            SoundPlay("sounds/enter.wav")
            ShellExecute("https://www.facebook.com/profile.php?id=100083295244149")

        Case $email
            SoundPlay("sounds/enter.wav")
            ShellExecute("https://mail.google.com/mail/u/0/?fs=1&tf=cm&source=mailto&to=vodinhhungtnlg@gmail.com")

        Case $menuitem2
            SoundPlay("sounds/enter.wav")
            _ShowFileContent("Tutorial (English)", "readme\ReadmeEnglish.txt")

        Case $menuitem1
            SoundPlay("sounds/enter.wav")
            _ShowFileContent("Tutorial (Vietnamese)", "readme\ReadmeVietnamese.txt")

        Case $button
            SoundPlay("sounds/enter.wav")
            ReadText()

        Case $tts
            _ActionSpeakOrSave(False)

        Case $saveAudio
            _ActionSpeakOrSave(True)

        Case $saveText
            _SaveTextHotkey()

        Case $openText
            _OpenTextHotkey()

        Case $menu2
            SoundPlay("sounds/enter.wav")
            contribute()
    EndSwitch
WEnd

Func _ActionSpeakOrSave($bIsSave)
    Local $sSelectedName = GuiCtrlRead($comboVoice)
    Local $sText = GuiCtrlRead($entertext)
    Local $iVol = GuiCtrlRead($sliderVolume)
    Local $iRate = GuiCtrlRead($sliderRate)
    Local $iPitch = GuiCtrlRead($sliderPitch)
    Local $sFileSave = ""

    If StringStripWS($sText, 8) = "" Then
        SoundPlay("sounds/enter.wav")
        MsgBox(48, "warning", "please enter your text")
        Return
    EndIf

    If IsGoogleVoice($sSelectedName) Then
        If $bIsSave Then
            SoundPlay("sounds/enter.wav")
            $sFileSave = FileSaveDialog("Save audio as...", @ScriptDir, "WAV Files (*.wav)|MP3 Files (*.mp3)", 16, "output.wav")
             If @error Or $sFileSave = "" Then Return
        EndIf

        If Not FileExists($sGoogleVoiceExe) Then
            MsgBox(16, "Error", "File not found: " & $sGoogleVoiceExe)
            Return
        EndIf

        Local $sLangCode = ($sSelectedName = "Google Vietnamese") ? "vi" : "en"
        Local $sCleanText = StringReplace($sText, '"', "'")
        Local $sCmd = '"' & $sGoogleVoiceExe & '" ' & $sLangCode & ' "' & $sCleanText & '"'

        If $bIsSave Then
            $sCmd &= ' "' & $sFileSave & '"'
            ProgressOn("Saving Audio", "Downloading from Google...", "Please wait")
            Local $pid = Run($sCmd, @ScriptDir, @SW_HIDE)
            ProcessWaitClose($pid)
            ProgressOff()
            If FileExists($sFileSave) Then
                MsgBox(64, "Success", "Audio saved successfully: " & $sFileSave)
            Else
                MsgBox(16, "Error", "Failed to save audio. Check internet connection.")
            EndIf
        Else
            Run($sCmd, @ScriptDir, @SW_HIDE)
        EndIf
        Return
    EndIf

    Local $sTokenID = _GetTokenIDByName($sSelectedName)
    If $sTokenID = "" Then
        MsgBox(16, "Error", "Voice ID could not be found.")
        Return
    EndIf

    Local $bIs32Bit = StringInStr($sTokenID, "WOW6432Node")

    If $bIsSave Then
        SoundPlay("sounds/enter.wav")
        $sFileSave = FileSaveDialog("Save audio as...", @ScriptDir, "WAV Files (*.wav)|MP3 Files (*.mp3)", 16, "output.wav")
        If @error Or $sFileSave = "" Then Return
    EndIf

    If @AutoItX64 And $bIs32Bit Then
        _RunSapiBridge($sTokenID, $sText, $iVol, $iRate, $iPitch, $sFileSave)
    Else
        _RunSapiNative($sTokenID, $sText, $iVol, $iRate, $iPitch, $sFileSave)
    EndIf
EndFunc

Func _RunSapiNative($sID, $sText, $vol, $rate, $pitch, $sFile)
    If Not IsObj($g_oSAPI) Then Return

    Local $oToken = ObjCreate("SAPI.SpObjectToken")
    If Not IsObj($oToken) Then Return

    $oToken.SetId($sID)
    $g_oSAPI.Voice = $oToken
    $g_oSAPI.Volume = $vol
    $g_oSAPI.Rate = $rate

    Local $ssml = '<sapi><pitch middle="' & $pitch & '">' & $sText & '</pitch></sapi>'

    If $sFile <> "" Then
        Local $oStream = ObjCreate("SAPI.SpFileStream")
        $oStream.Open($sFile, 3, False)
        $g_oSAPI.AudioOutputStream = $oStream
        $g_oSAPI.Speak($ssml)
        $oStream.Close()
        $g_oSAPI.AudioOutputStream = 0
        MsgBox(64, "Success", "Audio saved successfully.")
    Else
        $g_oSAPI.Speak($ssml, 1) ; 1 = Async
    EndIf
EndFunc

Func _RunSapiBridge($sID, $sText, $vol, $rate, $pitch, $sFile)
    Local $sTempVBS = @TempDir & "\sapi_bridge.vbs"
    Local $sTempText = @TempDir & "\sapi_content.txt"

    Local $hFile = FileOpen($sTempText, 2 + 32)
    FileWrite($hFile, $sText)
    FileClose($hFile)

    Local $sVBS = 'Option Explicit' & @CRLF
    $sVBS &= 'Dim Voice, Token, Stream, FSO, TS, Text' & @CRLF
    $sVBS &= 'Set FSO = CreateObject("Scripting.FileSystemObject")' & @CRLF
    $sVBS &= 'Set TS = FSO.OpenTextFile("' & $sTempText & '", 1, False, -1)' & @CRLF
    $sVBS &= 'Text = TS.ReadAll' & @CRLF
    $sVBS &= 'TS.Close' & @CRLF
    $sVBS &= 'Set Voice = CreateObject("SAPI.SpVoice")' & @CRLF
    $sVBS &= 'Set Token = CreateObject("SAPI.SpObjectToken")' & @CRLF
    $sVBS &= 'Token.SetId "' & $sID & '"' & @CRLF ; Set ID trực tiếp
    $sVBS &= 'Set Voice.Voice = Token' & @CRLF
    $sVBS &= 'Voice.Volume = ' & $vol & @CRLF
    $sVBS &= 'Voice.Rate = ' & $rate & @CRLF

    Local $sSSML = '<sapi><pitch middle=' & "'" & '" & ' & $pitch & ' & "' & "'" & '>' & '" & Text & "' & '</pitch></sapi>'

    If $sFile <> "" Then
        $sVBS &= 'Set Stream = CreateObject("SAPI.SpFileStream")' & @CRLF
        $sVBS &= 'Stream.Open "' & $sFile & '", 3, False' & @CRLF
        $sVBS &= 'Set Voice.AudioOutputStream = Stream' & @CRLF
        $sVBS &= 'Voice.Speak "' & $sSSML & '"' & @CRLF
        $sVBS &= 'Stream.Close' & @CRLF
    Else
        $sVBS &= 'Voice.Speak "' & $sSSML & '"' & @CRLF
    EndIf

    Local $hVBS = FileOpen($sTempVBS, 2)
    FileWrite($hVBS, $sVBS)
    FileClose($hVBS)

    Local $sCScript = @WindowsDir & "\SysWOW64\cscript.exe"
    If Not FileExists($sCScript) Then
        $sCScript = "cscript.exe" ; Fallback
    EndIf

    If $sFile <> "" Then
        ProgressOn("Saving (32-bit mode)", "Processing...", "Please wait")
        RunWait('"' & $sCScript & '" //Nologo "' & $sTempVBS & '"', @TempDir, @SW_HIDE)
        ProgressOff()
        MsgBox(64, "Success", "Audio saved successfully (via 32-bit bridge).")
    Else
        Run('"' & $sCScript & '" //Nologo "' & $sTempVBS & '"', @TempDir, @SW_HIDE)
    EndIf
EndFunc

Func _PopulateVoiceComboBox($hCombo)
    Local $sList = ""
    $g_iVoiceCount = 0
    Local $aHives[3] = [ _
        "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens", _
        "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Speech\Voices\Tokens", _
        "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Speech\Voices\Tokens" _
    ]

    Local $i, $h, $sKey, $sName, $sID, $sFirstVoice = ""
    For $h = 0 To 2
        $i = 1
        While 1
            $sKey = RegEnumKey($aHives[$h], $i)
            If @error Then ExitLoop

            $sID = $aHives[$h] & "\" & $sKey
            $sName = RegRead($sID, "")

            Local $bExists = False
            For $k = 0 To $g_iVoiceCount - 1
                If $g_aVoiceList[$k][0] = $sName Then $bExists = True
            Next

            If $sName <> "" And Not $bExists Then
                If $sFirstVoice = "" Then $sFirstVoice = $sName
                $g_aVoiceList[$g_iVoiceCount][0] = $sName
                $g_aVoiceList[$g_iVoiceCount][1] = $sID
                $sList &= $sName & "|"
                $g_iVoiceCount += 1
            EndIf
            $i += 1
        WEnd
    Next
    If $sFirstVoice = "" Then $sFirstVoice = "Google English"
    $sList &= "Google English|Google Vietnamese"
    GuiCtrlSetData($hCombo, $sList, $sFirstVoice)
EndFunc

Func _GetTokenIDByName($sName)
    For $i = 0 To $g_iVoiceCount - 1
        If $g_aVoiceList[$i][0] = $sName Then Return $g_aVoiceList[$i][1]
    Next
    Return ""
EndFunc

Func IsGoogleVoice($sName)
    Return ($sName = "Google English" Or $sName = "Google Vietnamese")
EndFunc

Func _PerformSpeakToText()
    If Not FileExists($sSpeakExe) Then
        MsgBox(16, "Error", "File not found: " & $sSpeakExe)
        Return
    EndIf

    Local $sCurrentVoice = GuiCtrlRead($comboVoice)
    Local $sLang = "vi-VN"
    If StringInStr($sCurrentVoice, "English") Or StringInStr($sCurrentVoice, "David") Or StringInStr($sCurrentVoice, "Zira") Then
        $sLang = "en-US"
    EndIf

    SoundPlay("sounds\start.wav", 1)

    ToolTip("Listening... Please speak now.", Default, Default, "Microphone", 1)

    Local $iPID = Run('"' & $sSpeakExe & '" ' & $sLang, @ScriptDir, @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
    Local $bOutput = Binary("")

    While ProcessExists($iPID)
        $bOutput &= StdoutRead($iPID, False, True)
        Sleep(50)
    WEnd

    $bOutput &= StdoutRead($iPID, False, True)
    SoundPlay("sounds\stop.wav", 0)
    ToolTip("")

    Local $sOutput = BinaryToString($bOutput, 4) ; 4 = UTF-8
    $sOutput = StringStripWS($sOutput, 3)

    If $sOutput <> "" Then
        Local $sCurrentText = GuiCtrlRead($entertext)
        If $sCurrentText <> "" Then
            GuiCtrlSetData($entertext, $sCurrentText & " " & $sOutput)
        Else
            GuiCtrlSetData($entertext, $sOutput)
        EndIf
    EndIf
EndFunc

Func _ShowSettings()
    Local $hSettingGUI = GuiCreate("Settings", 350, 200, -1, -1, BitOR($WS_CAPTION, $WS_POPUP, $WS_SYSMENU), -1, $hGUI)
    GuiSetBkColor($COLOR_WHITE)

    Local $chkAutoUpdate = GuiCtrlCreateCheckbox("Automatically checked for updates on startup", 20, 20, 300, 20)
    Local $chkAutoClip = GuiCtrlCreateCheckbox("Automatically retrieve text from clipboard", 20, 50, 300, 20)
    Local $chkStartupWin = GuiCtrlCreateCheckbox("Startup with Windows", 20, 80, 300, 20)

    Local $btnOk = GuiCtrlCreateButton("&OK", 60, 130, 80, 30)
    Local $btnCancel = GuiCtrlCreateButton("&Cancel", 200, 130, 80, 30)

    If $bAutoUpdate Then GUICtrlSetState($chkAutoUpdate, $GUI_CHECKED)
    If $bAutoClipboard Then GUICtrlSetState($chkAutoClip, $GUI_CHECKED)

    Local $sRegKey = "HKCU\Software\Microsoft\Windows\CurrentVersion\Run"
    Local $sAppName = "ReadTextApp"
    If RegRead($sRegKey, $sAppName) = @ScriptFullPath Then
        GUICtrlSetState($chkStartupWin, $GUI_CHECKED)
    EndIf

    GuiSetState(@SW_SHOW, $hSettingGUI)

    While 1
        Switch GuiGetMSG()
            Case $GUI_EVENT_CLOSE, $btnCancel
                GuiDelete($hSettingGUI)
                Return

            Case $btnOk
                $bAutoUpdate = (BitAND(GUICtrlRead($chkAutoUpdate), $GUI_CHECKED) = $GUI_CHECKED)
                $bAutoClipboard = (BitAND(GUICtrlRead($chkAutoClip), $GUI_CHECKED) = $GUI_CHECKED)

                IniWrite($sConfigFile, "Settings", "AutoUpdate", $bAutoUpdate ? "true" : "false")
                IniWrite($sConfigFile, "Settings", "AutoClipboard", $bAutoClipboard ? "true" : "false")

                If BitAND(GUICtrlRead($chkStartupWin), $GUI_CHECKED) = $GUI_CHECKED Then
                    RegWrite($sRegKey, $sAppName, "REG_SZ", @ScriptFullPath)
                Else
                    RegDelete($sRegKey, $sAppName)
                EndIf

                SoundPlay("sounds/enter.wav")
                GuiDelete($hSettingGUI)
                Return
        EndSwitch
    WEnd
EndFunc

Func _GetClipboardText($bSilent)
    Local $sClipText = ClipGet()
    If @error Or StringStripWS($sClipText, 8) = "" Then
        If Not $bSilent Then MsgBox(48, "Warning", "The clipboard is empty! Please copy the text to the clipboard")
    Else
        GUICtrlSetData($entertext, $sClipText)
        If Not $bSilent Then MsgBox(64, "Success", "Text retrieved from clipboard successfully")
    EndIf
EndFunc

Func _CheckGithubUpdate()
    Local $sCheckingText = "Checking for updates..."
    Local $hCheckGUI = GuiCreate("", 300, 80, -1, -1, BitOR($WS_CAPTION, $WS_POPUP), BitOR($WS_EX_TOPMOST, $WS_EX_TOOLWINDOW))
    GuiSetBkColor(0xFFFFFF, $hCheckGUI)
    Local $lblCheck = GuiCtrlCreateLabel($sCheckingText, 10, 25, 280, 30, $ES_CENTER)
    GuiCtrlSetFont($lblCheck, 10, 400, 0, "Arial")
    GuiSetState(@SW_SHOW, $hCheckGUI)
    Sleep(3000)
    GuiDelete($hCheckGUI)

    If Ping("github.com", 2000) = 0 And Ping("google.com", 2000) = 0 Then
         SoundPlay("sounds/update_error.wav")
         MsgBox(48, "Check Update", "No internet connection.")
         Return
    EndIf

    Local $sRepoOwner = "ninhhoang2005"
    Local $sRepoName = "read_text"
    Local $sApiUrl = "https://api.github.com/repos/ninhhoang2005/read_text/releases/latest"

    Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")
    If Not IsObj($oHTTP) Then
        MsgBox(16, "Error", "Cannot create HTTP Object.")
        Return
    EndIf

    $oHTTP.Open("GET", $sApiUrl, False)
    $oHTTP.Send()

    If @error Then
        SoundPlay("sounds/update_error.wav")
        MsgBox(48, "Check Update", "Connection failed. Please check your internet.")
        Return
    EndIf

    If $oHTTP.Status <> 200 Then
        MsgBox(48, "Check Update", "Cannot connect to update server or no release found." & @CRLF & "Status Code: " & $oHTTP.Status)
        Return
    EndIf

    Local $sResponse = $oHTTP.ResponseText
    Local $aMatch = StringRegExp($sResponse, '"tag_name":\s*"([^"]+)"', 3)

    If IsArray($aMatch) Then
        Local $sLatestVersion = $aMatch[0]
        $sLatestVersion = StringReplace($sLatestVersion, "v", "")
        If $sLatestVersion <> $sAppVersion Then
            SoundPlay("sounds/update.wav")
            Local $iMsg = MsgBox(36, "Update Available", "A new version (" & $sLatestVersion & ") is available!" & @CRLF & _
                                     "Your version: " & $sAppVersion & @CRLF & @CRLF & _
                                     "Do you want to download it now?")
            If $iMsg = 6 Then
                Local $downloadtext = "please wait"
                Local $downloadGui = GuiCreate("downloading update", 400, 400, -1, -1)
                GuiSetBkColor($COLOR_WHITE)
                GuiCtrlCreateLabel($downloadtext, 40, 60)
                GuiSetState(@SW_SHOW, $downloadGui)
                Local $sDownloadURL = "https://github.com/ninhhoang2005/read_text/releases/latest/download/read_text.zip"
                Local $sSavePath = @ScriptDir & "\read_text.zip"

                ProgressOn("Downloading Update", "Please wait while downloading...", "0%")
                DllCall("winmm.dll", "int", "PlaySoundW", "wstr", @ScriptDir & "\sounds\updating.wav", "ptr", 0, "dword", 0x0009)
                Local $hDownload = InetGet($sDownloadURL, $sSavePath, 1, 1)

                Do
                    Sleep(100)
                    Local $iBytesRead = InetGetInfo($hDownload, 0)
                    Local $iFileSize = InetGetInfo($hDownload, 1)

                    If $iFileSize > 0 Then
                        Local $iPct = Round(($iBytesRead / $iFileSize) * 100)
                        ProgressSet($iPct, $iPct & "% complete")
                    Else
                        ProgressSet(0, "Connecting...")
                    EndIf

                Until InetGetInfo($hDownload, 2)

                InetClose($hDownload)
                DllCall("winmm.dll", "int", "PlaySoundW", "ptr", 0, "ptr", 0, "dword", 0)
                ProgressOff()
                GuiDelete($downloadGui)
                SoundPlay("sounds/updated.wav")
                MsgBox(64, "Success", "Downloaded successfully!" & @CRLF & "File saved as: " & $sSavePath)
                Run("unzip.bat")
                Exit
            EndIf
        Else
            MsgBox(64, "no update available", "You are using the latest version (" & $sAppVersion & ").")
        EndIf
    Else
        MsgBox(16, "Error", "Could not parse version information.")
    EndIf
EndFunc

Func _ShowFileContent($sTitle, $sFilePath)
    If Not FileExists($sFilePath) Then
        MsgBox(0, "Error", "Cannot find file: " & $sFilePath)
        Return
    EndIf
    Local $sContent = FileRead($sFilePath)
    Local $displayGui = GuiCreate($sTitle, 500, 500)
    Local $document = GUICtrlCreateEdit($sContent, 20, 20, 450, 400, BitOR($ES_AUTOVSCROLL, $ES_READONLY, $WS_VSCROLL, $WS_TABSTOP))
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

Func _OpenTextHotkey()
    SoundPlay("sounds/enter.wav")
    Local $sFile = FileOpenDialog("Open text file...", @ScriptDir, "Text files (*.txt)", 1)
    If @error Or $sFile = "" Then Return
    Local $content = FileRead($sFile)
    GUICtrlSetData($entertext, $content)
EndFunc

Func _SaveTextHotkey()
    SoundPlay("sounds/enter.wav")
    Local $sFile = FileSaveDialog("Save text as...", @ScriptDir, "Text files (*.txt)", 16, "output.txt")
    If @error Or $sFile = "" Then Return
    Local $text = GuiCtrlRead($entertext)
    FileDelete($sFile)
    FileWrite($sFile, $text)
    MsgBox(64, "success", "Text saved successfully.")
EndFunc

Func ReadText()
    Local $text = GuiCtrlRead($entertext)
    If StringStripWS($text, 8) = "" Then
        SoundPlay("sounds/enter.wav")
        MsgBox(48, "warning", "please enter your text")
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

Func contribute()
    Local $congui = GuiCreate("contribute", 700, 700)
    GuiSetBkColor($COLOR_RED)
    Local $con = ""
    If FileExists("contribute.txt") Then $con = FileRead("contribute.txt")
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

Func _LoadConfig()
    If Not FileExists($sConfigFile) Then Return
    Local $sVoice = IniRead($sConfigFile, "Settings", "Voice", "")
    If $sVoice <> "" Then GUICtrlSetData($comboVoice, $sVoice)
    Local $iVol = IniRead($sConfigFile, "Settings", "Volume", 100)
    GUICtrlSetData($sliderVolume, $iVol)
    Local $iRate = IniRead($sConfigFile, "Settings", "Rate", 0)
    GUICtrlSetData($sliderRate, $iRate)
    Local $iPitch = IniRead($sConfigFile, "Settings", "Pitch", 0)
    GUICtrlSetData($sliderPitch, $iPitch)
    $bAutoUpdate = (IniRead($sConfigFile, "Settings", "AutoUpdate", "false") = "true")
    $bAutoClipboard = (IniRead($sConfigFile, "Settings", "AutoClipboard", "false") = "true")
    Local $sLastText = IniRead($sConfigFile, "Data", "LastText", "")
    $sLastText = StringReplace($sLastText, "¶", @CRLF)
    GUICtrlSetData($entertext, $sLastText)
EndFunc

Func _SaveConfig()
    IniWrite($sConfigFile, "Settings", "Voice", GUICtrlRead($comboVoice))
    IniWrite($sConfigFile, "Settings", "Volume", GUICtrlRead($sliderVolume))
    IniWrite($sConfigFile, "Settings", "Rate", GUICtrlRead($sliderRate))
    IniWrite($sConfigFile, "Settings", "Pitch", GUICtrlRead($sliderPitch))
    IniWrite($sConfigFile, "Settings", "AutoUpdate", $bAutoUpdate ? "true" : "false")
    IniWrite($sConfigFile, "Settings", "AutoClipboard", $bAutoClipboard ? "true" : "false")
    Local $sCurrentText = GUICtrlRead($entertext)
    $sCurrentText = StringReplace($sCurrentText, @CRLF, "¶")
    IniWrite($sConfigFile, "Data", "LastText", $sCurrentText)
EndFunc

Func _ShowChangelog()
    Local $sFilePath = @ScriptDir & "\changelog.txt"
    Local $sContent = "No changelog found."

    If FileExists($sFilePath) Then
        $sContent = FileRead($sFilePath)
    EndIf

    Local $hChangelogGUI = GuiCreate("Changelog", 400, 450)
    Local $editChangelog = GUICtrlCreateEdit($sContent, 10, 10, 380, 380, BitOR($ES_AUTOVSCROLL, $ES_READONLY, $WS_VSCROLL, $WS_TABSTOP))
    Local $btnClose = GUICtrlCreateButton("&Close", 150, 400, 100, 30, $WS_TABSTOP)

    GuiSetState(@SW_SHOW, $hChangelogGUI)

    While 1
        Switch GuiGetMSG()
            Case $GUI_EVENT_CLOSE, $btnClose
                GuiDelete($hChangelogGUI)
                ExitLoop
        EndSwitch
    WEnd
EndFunc

Func _ErrFunc()
EndFunc