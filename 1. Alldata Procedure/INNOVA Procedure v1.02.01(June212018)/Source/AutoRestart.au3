#include <MsgBoxConstants.au3>


Auto()

Func Auto()
   Local $sFilePath = InputBox ("Input App Path", "Please input the app path!")
   If @error <> 0 Then Exit
   ;-----------------------------------------
   ;GET TITLE AND RUN THE APP THE FIRST TIME
   ;Get title
   Global $sTitle = StringRight ($sFilePath, StringLen ($sFilePath) - StringInStr ($sFilePath, "\", 0, -1))
   $sTitle = StringLeft ($sTitle, StringLen ($sTitle) - 4)
   ;Run the app the first time
   If Run_App ($sFilePath, $sTitle) = False Then
	  MsgBox ($MB_TOPMOST, "Message", "Failed to run the App" & @CRLF & "Please check the APP PATH!")
	  Exit
   EndIf

   ;-----------------------------------------
   ;GET KEY AND YMME LINK  AFTER USER INPUTS
   Local $sKey = ""
   Local $sYMME_Link = ""
   Local $sLicense = ""
   Local $sDTC_Database = ""
   Local $sDTC_MsgID = ""
   Local $iWebVisible = 0
   Local $iImagesDownload = 0

   While 1
	  If WinExists ($sTitle) = 0 Then Exit
	  If Update_Inputs ($sKey, $sYMME_Link, $sLicense, $sDTC_Database, $sDTC_MsgID, $iWebVisible, $iImagesDownload) = True Then ExitLoop
   WEnd
   ;-----------------------------------------
   ;Automatically turn on the app and put values
   While 1
	  Update_Inputs ($sKey, $sYMME_Link, $sLicense, $sDTC_Database, $sDTC_MsgID, $iWebVisible, $iImagesDownload)
	  If WinExists ($sTitle & ".exe") = 1 Then
		 MsgBox ($MB_TOPMOST, "CRASH", "CRASH!" & @CRLF & "Press OK to Restart" & @CRLF & "or the App will automatically be restarted after 10 seconds", 10)
		 ConsoleWrite ("Restart Crash" & @CRLF)
		 While WinExists ($sTitle & ".exe") = 1
			ControlClick ($sTitle & ".exe", "", "Button2", "left", 1)
			Sleep (1000)
		 WEnd
		 ;Run the app
		 If Run_App ($sFilePath, $sTitle) = False Then
			MsgBox ($MB_TOPMOST, "Message", "Failed to run the App" & @CRLF & "Please check the APP PATH!")
			Exit
		 EndIf
		 Set_Inputs ($sKey, $sYMME_Link, $sLicense, $sDTC_Database, $sDTC_MsgID, $iWebVisible, $iImagesDownload)
	  EndIf
	  If WinExists ("AutoIt Error") = 1 Then
		 MsgBox ($MB_TOPMOST, "ERROR", "Unexpected ERROR!" & @CRLF & "Press OK to Restart" & @CRLF & "or the App will automatically be Restarted after 10 seconds", 10)
		 ConsoleWrite ("Restart Error" & @CRLF)
		 While WinExists ("AutoIt Error") = 1
			ControlClick ("AutoIt Error", "", "Button1", "left", 1)
			Sleep (1000)
		 WEnd
		 ;Run the app
		 Run_App ($sFilePath, $sTitle)
		 Set_Inputs ($sKey, $sYMME_Link, $sLicense, $sDTC_Database, $sDTC_MsgID, $iWebVisible, $iImagesDownload)
	  EndIf
	  If WinExists ($sTitle) = 0 Then Exit

	  Sleep (100)
   WEnd
EndFunc   ;


Func Set_Inputs ($sKey, $sYMME_Link, $sLicense, $sDTC_Database, $sDTC_MsgID, $iWebVisible, $iImagesDownload)
   ;Input key
   ControlSetText ($sTitle, "", "Edit1", $sKey)
   Sleep (1000)
   ;Press the submit button and wait until the status changes
   ControlClick ($sTitle, "", "Button1")
	  While Get_Status () <> "Standby"
		 If WinExists ($sTitle) = 0 Then Exit
		 Sleep (100)
	  WEnd
   Sleep (100)
   ;Switchtab
   ControlClick ($sTitle, "", "SysTabControl321", "left", 1, 168, 10)
   Sleep (500)
   ;Input YMME Link
   ControlSetText ($sTitle, "", "Edit3", $sYMME_Link)
   Sleep (200)
   ;Input License
   ControlCommand ($sTitle, "", "ComboBox2", "SelectString", $sLicense)
   Sleep (200)


   ;Input DTC DAtabase
   ControlCommand ($sTitle, "", "ComboBox3", "SelectString", $sDTC_Database)
   Sleep (200)


   ;Input DTC MsgID
   ControlCommand ($sTitle, "", "ComboBox4", "SelectString", $sDTC_MsgID)
   Sleep (200)



   ;Set Show Hide
   If $iWebVisible = 0 Then
	  ControlCommand ($sTitle, "", "Button8", "UnCheck")
   Else
	  ControlCommand ($sTitle, "", "Button8", "Check")
   EndIf
   Sleep (200)

   ;Set Images Download
   If $iImagesDownload = 0 Then
	  ControlCommand ($sTitle, "", "Button10", "UnCheck")
   Else
	  ControlCommand ($sTitle, "", "Button10", "Check")
   EndIf
   Sleep (200)

   ;Click begin
   ControlClick ($sTitle, "", "Button3")
   Sleep (200)
EndFunc



Func Run_App ($sFilePath, $sTitle)
   Local $sResult = False
	  ; Run App with the window maximized.
	  Local $iPID = Run($sFilePath, "")
	  ; Wait for the Notepad window to appear.
	  If WinWait($sTitle, "", 3) <> 0 Then $sResult = True
	  ; Wait for 0.5 seconds.
	  Sleep(500)
   Return $sResult
EndFunc


Func Update_Inputs (ByRef $sKey, ByRef $sYMME_Link, ByRef $sLicense, Byref $sDTC_Database, Byref $sDTC_MsgID, ByRef $iWebVisible, Byref $iImagesDownload)
   ;Get status
   $bResult = False
   Local $Status = Get_Status ()
   If $Status = "Working" Then
	  $sKey = ControlGetText ($sTitle, "", "Edit1" )
	  $sYMME_Link = ControlGetText ($sTitle, "", "Edit3" )
	  $sLicense = ControlGetText ($sTitle, "", "ComboBox2")
	  $sDTC_Database = ControlGetText ($sTitle, "", "ComboBox3")
	  $sDTC_MsgID = ControlGetText ($sTitle, "", "ComboBox4")
	  $iWebVisible = ControlCommand ($sTitle, "", "Button8", "IsChecked")
	  $iImagesDownload = ControlCommand ($sTitle, "", "Button10", "IsChecked")
	  $bResult = True
   EndIf

   Return $bResult
EndFunc


Func Get_Status ()
   Local $Status = ControlGetText ($sTitle, "", "Static7" )
   Return $Status
EndFunc


