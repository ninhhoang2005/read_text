Func speak($nvda_text)
If StringInStr($nvda_text, @CRLF) Then
$nvda_text=StringReplace($nvda_text, @CRLF, ". ")
EndIf
If StringInStr($nvda_text, @LF) Then
$nvda_text=StringReplace($nvda_text,@LF,". ")
EndIf
RunWait(@ComSpec& " /c voice_handel.exe "&$nvda_text,@ScriptDir&"\library\nvda_handel",@SW_HIDE)
EndFunc