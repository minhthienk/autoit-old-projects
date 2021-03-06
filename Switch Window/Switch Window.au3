#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

HotKeySet ("^{SPACE}", "Test")

Global $toggle = 'left'




Func Test()
   If $toggle = 'left' Then
	  Send("#^{RIGHT}")
	  $toggle = 'right'
   Else
	  Send("#^{LEFT}")
	  $toggle = 'left'
   EndIf
EndFunc

while 1
   Sleep(100)
WEnd


