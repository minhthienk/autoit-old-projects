#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         Thien Nguyen

 Script Function:
   Copy data from bonbanh.com

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include-once

#include <MsgBoxConstants.au3>
#include <Clipboard.au3>
#include <IE.au3 >
#include <Excel.au3>

#include "Library.au3"
Global $sFilePath = @ScriptDir


; Set Hotkey for the program
HotKeySet("{ESC}", "Autoit_Exit")


Main()

Func Main()
   MsgBox (0, "NOTE", "Click OK" & @CRLF & @CRLF & "PRESS the ESC button to exit the application if needed")

   ;Load links file
   Local $sFileName = 'Vehicle_Links_Sandiego'
   Local  $sFilePath = @ScriptDir
   Local $aLinks = LoadFile ($sFileName, $sFilePath)

   ;Create a text file to save the content from the sites
   Local $sFileName = "Vehicle_Info_Sandiego"
   Local $sFilePath = @ScriptDir

   For $sLink In $aLinks
	  ;$sLink = 'https://www.kengrodyfordorangecounty.com/inventory/new-2018-ford-explorer-xlt-4wd-sport-utility-1fm5k8d88jga56075'
	  Local $sContent = GetHtmlSourceUsingHttpRequest($sLink)
	  ;_ClipBoard_SetData($sContent)
	  Local $sTxt = ''
	  ;Vehicle
	  Local $Vehicle_title = GetItemStringByMark ($sContent, '<h1 class="vehicle-title entry-title">', '</h1>')
	  $sTxt &=  $Vehicle_title & @TAB
	  ;VIN
	  $sTxt &= GetItemStringByMark ($sContent, '<span class="vinstock-number">', '</span>',2,1) & @TAB
	  ;Navigation support?
	  If StringInStr($Vehicle_title, 'with navigation') <> 0 Then
		 $sTxt &= 'Supported' & @TAB
	  Else
		 $sTxt &= 'Not Supported' & @TAB
	  EndIf
	  ;MSRP
	  $sTxt &= '$' & GetItemStringByMark(GetItemStringByMark ($sContent, 'MSRP', '</span>',1,2), '$', @TAB) & @TAB
	  ;Discount
	  $sTxt &= '$' & GetItemStringByMark(GetItemStringByMark ($sContent, 'Ken Grody Discount', '</span>',1,2), '$', @TAB) & @TAB
	  ;Net price
	  $sTxt &= '$' & GetItemStringByMark(GetItemStringByMark ($sContent, 'Net Price', '</div>',2,1), '$', @TAB) & @TAB
	  ;Body style
	  $sTxt &= GetItemStringByMark($sContent, '<li class="list-group-item">Body Style: ', '</li>') & @TAB
	  ;Drivedrain
	  $sTxt &= GetItemStringByMark($sContent, '<li class="list-group-item">Drivetrain: ', '</li>') & @TAB
	  ;Exterior color
	  $sTxt &= GetItemStringByMark(GetItemStringByMark ($sContent, '<dt>Exterior:</dt>', '</dd>') & 'EndMark', '<dd> ', 'EndMark') & @TAB
	  ;Engine
	  $sTxt &= GetItemStringByMark(GetItemStringByMark ($sContent, '<dt>Engine:</dt>', '</dd>') & 'EndMark', '<dd> ', 'EndMark') & @TAB
	  ;Trans
	  $sTxt &= GetItemStringByMark(GetItemStringByMark ($sContent, '<dt>Trans:</dt>', '</dd>') & 'EndMark', '<dd> ', 'EndMark') & @TAB
	  ;Net Horsepower
	  $sTxt &= GetItemStringByMark($sContent, '<li class="list-group-item">SAE Net Horsepower @ RPM: ', '</li>') & @TAB
	  ;Net torque
	  $sTxt &= GetItemStringByMark($sContent, '<li class="list-group-item">SAE Net Torque @ RPM: ', '</li>') & @TAB
	  ;Brake type
	  $sTxt &= GetItemStringByMark($sContent, '<li class="list-group-item">Brake Type: ', '</li>') & @TAB
	  ;ABS
	  $sTxt &= GetItemStringByMark($sContent, '<li class="list-group-item">Brake ABS System: ', '</li>') & @TAB
	  ;Entertainment

	  Local $stemp =  GetItemStringByMark($sContent, '<div id="entertainment_options_contents" class="panel-collapse collapse ">', '</div>',1,4)
	  $stemp = GetItemStringByMark($stemp, '<li class="list-group-item">', '</div>',1,3)
	  $stemp = StringReplace($stemp, '<li class="list-group-item">', '')
	  $stemp = StringRegExpReplace($stemp, '</.+>', '')
	  $stemp = StringReplace($stemp, @TAB, '')
	  $stemp = StringRegExpReplace($stemp, '[\r\n]+', @CRLF)
	  If StringRight($stemp, 2) = @CRLF Then $stemp = StringLeft($stemp, StringLen($stemp) - 2)

	  $sTxt &= '"' & $stemp &  '"' &@TAB

	  ;Link
	  $sTxt &= $sLink & @TAB

	  ;Write data into a file
	  WriteTxtFile ($sFileName, @CRLF & $sTxt, "append")
	  ;WriteTxtFile ($sFileName, @CRLF & $sTxt, "overwrite")

   Next

   MsgBox ($MB_TOPMOST, "", "done")
EndFunc





Func Write_Log ($sText)
   Static Local $bFirst = True
   If $bFirst = True Then
	  Local $hFileOpen = FileOpen(@ScriptDir & "/Log.txt",  $FO_OVERWRITE)
	  FileWrite($hFileOpen, $sText)
	  FileClose($hFileOpen)
	  $bFirst = False
   Else
	  Local $hFileOpen = FileOpen(@ScriptDir & "/Log.txt", $FO_APPEND)
	  FileWrite($hFileOpen, $sText)
	  FileClose($hFileOpen)
   EndIf
EndFunc

