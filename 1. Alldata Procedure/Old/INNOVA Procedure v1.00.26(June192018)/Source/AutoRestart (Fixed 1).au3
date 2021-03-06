#include <MsgBoxConstants.au3>


Auto()

Func Auto()
   Local $sFilePath = InputBox ("Input App Path", "Please input the app path!")
;~    Local $sFilePath = "C:\Users\K\Desktop\Procedure Generator v1.00.25.exe"



   ;-----------------------------------------
   ;GET TITLE AND RUN THE APP THE FIRST TIME
   ;Get title
   Global $sTitle = StringRight ($sFilePath, StringLen ($sFilePath) - StringInStr ($sFilePath, "\", 0, -1))
   $sTitle = StringLeft ($sTitle, StringLen ($sTitle) - 4)
   ;Run the app the first time
   Run_App ($sFilePath, $sTitle)


   ;-----------------------------------------
   ;GET KEY AND YMME LINK  AFTER USER INPUTS
   Local $sKey = ""
   Local $sYMME_Link = ""
   Local $sLicense = ""
   Local $iWebVisible = 0
   While 1
	  If Update_Inputs ($sKey, $sYMME_Link, $sLicense, $iWebVisible) = True Then ExitLoop
   WEnd

   ;-----------------------------------------
   ;Automatically turn on the app and put values
   While 1
	  If WinExists ($sTitle & ".exe") = 1 Then
		 MsgBox ($MB_TOPMOST, "CRASH", "CRASH! Restarting?", 10)
		 ConsoleWrite ("Restart Crash" & @CRLF)
		 While WinExists ($sTitle & ".exe") = 1
			ControlClick ($sTitle & ".exe", "", "Button2", "left", 1)
			Sleep (1000)
		 WEnd
		 ;Run the app
		 Run_App ($sFilePath, $sTitle)
		 Set_Inputs ($sKey, $sYMME_Link, $sLicense, $iWebVisible)
	  EndIf

	  If WinExists ("AutoIt Error") = 1 Then
		 MsgBox ($MB_TOPMOST, "CRASH", "ERROR! Restarting?", 10)
		 ConsoleWrite ("Restart Error" & @CRLF)
		 While WinExists ("AutoIt Error") = 1
			ControlClick ("AutoIt Error", "", "Button1", "left", 1)
			Sleep (1000)
		 WEnd
		 ;Run the app
		 Run_App ($sFilePath, $sTitle)
		 Set_Inputs ($sKey, $sYMME_Link, $sLicense, $iWebVisible)
	  EndIf

	  If WinExists ($sTitle) = 0 Then Exit

	  Sleep (100)
   WEnd
EndFunc   ;


Func Set_Inputs ($sKey, $sYMME_Link, $sLicense, $iWebVisible)
   ;Input key
   ControlSetText ($sTitle, "", "Edit1", $sKey)
   Sleep (500)
   ;Press the submit button and wait until the status changes
   ControlClick ($sTitle, "", "Button1")
	  While Get_Status () <> "Standby"
		 If WinExists ($sTitle) = 0 Then Exit
	  WEnd
   Sleep (100)
   ;Switchtab
   ControlClick ($sTitle, "", "SysTabControl321", "left", 1, 168, 10)
   Sleep (100)
   ;Input YMME Link
   ControlSetText ($sTitle, "", "Edit3", $sYMME_Link)
   Sleep (100)
   ;Input License
   ControlCommand ($sTitle, "", "ComboBox2", "SelectString", $sLicense)
   Sleep (100)
   ;Set Show Hide
   If $iWebVisible = 0 Then
	  ControlCommand ($sTitle, "", "Button8", "UnCheck")
   Else
	  ControlCommand ($sTitle, "", "Button8", "Check")
   EndIf
   ;Click begin
   ControlClick ($sTitle, "", "Button3")
   Sleep (100)
EndFunc



Func Run_App ($sFilePath, $sTitle)
	  ; Run App with the window maximized.
	  Local $iPID = Run($sFilePath, "")
	  ; Wait for the Notepad window to appear.
	  WinWait($sTitle, "", 10)
	  ; Wait for 0.5 seconds.
	  Sleep(50)
EndFunc


Func Update_Inputs (ByRef $sKey, ByRef $sYMME_Link, ByRef $sLicense, ByRef $iWebVisible)
   ;Get status
   $bResult = False
   Local $Status = Get_Status ()
   If $Status = "Working" Then
	  $sKey = ControlGetText ($sTitle, "", "Edit1" )
	  $sYMME_Link = ControlGetText ($sTitle, "", "Edit3" )
	  $sLicense = ControlGetText ($sTitle, "", "ComboBox2")
	  $iWebVisible = ControlCommand ($sTitle, "", "Button8", "IsChecked")
	  $bResult = True
   EndIf

   Return $bResult
EndFunc


Func Get_Status ()
   Local $Status = ControlGetText ($sTitle, "", "Static8" )
   Return $Status
EndFunc


