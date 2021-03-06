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

#include "COMErrorHandler.au3"
#include "Library.au3"

Global $sFilePath = @ScriptDir

; Set Hotkey for the program
;~ HotKeySet("{ESC}", "Autoit_Exit")

Main ()

Func Main()
   MsgBox (0, "NOTE", "Click OK" & @CRLF & @CRLF & "PRESS the ESC button to exit the application if needed")
   Local $sMake = "Mercury"
   Local $sLink = "https://autos.mercadolibre.com.mx/mercury/"
   Local $sFileName = "mercadolibre_" & $sMake

   ;Create a text file to save the content from the sites
   Local $sFilePath = @ScriptDir

;~    WriteTxtFile ($sFileName, "This is the collection of " & $sMake & " vehicles on mercadolibre website", "overwrite")
;~    WriteTxtFile ($sFileName, @CRLF & "Vehicle" & @TAB & "Year" & @TAB & "Type" & @TAB & "Engine" & @TAB & "Transmission" & @TAB & "Fuel Type" & @TAB & "Door" & @TAB & "Link", "append")
;~    Exit

   For $i = 1996 To 2018
	  Local $sLinkYear = $sLink & $i
	  Local $iVehicleNum = 1
	  While 1
		 ;Form a page like containing vehicle links
		 Local $sLinkPage = $sLinkYear & '_Desde_' & $iVehicleNum
		 ;Get all vehicle links
		 Local $aData = GetInfoContent ($sLinkPage)
		 ;If their are no vehicle links => Endloop
		 If $aData[0] = 'No information' Then ExitLoop
		 ;Get vehicle data from vehicle link
		 For $Element In $aData
			Local $sContent = GetHtmlSourceUsingHttpRequest($Element)
			Local $sTxt = ''
			$sTxt &= GetItemStringByMark ($sContent, '<meta property="og:title" content="', '"/>') & @TAB
			$sTxt &= GetItemStringByMark(GetItemStringByMark ($sContent, '<strong>Año</strong>', '</span>') & 'EndMark', '<span>', 'EndMark') & @TAB
			$sTxt &= GetItemStringByMark(GetItemStringByMark ($sContent, '<strong>Tipo</strong>', '</span>') & 'EndMark', '<span>', 'EndMark') & @TAB
			$sTxt &= GetItemStringByMark(GetItemStringByMark ($sContent, '<strong>Motor</strong>', '</span>') & 'EndMark', '<span>', 'EndMark') & @TAB
			$sTxt &= GetItemStringByMark(GetItemStringByMark ($sContent, '<strong>Transmisión</strong>', '</span>') & 'EndMark', '<span>', 'EndMark') & @TAB
			$sTxt &= GetItemStringByMark(GetItemStringByMark ($sContent, '<strong>Tipo de combustible</strong>', '</span>') & 'EndMark', '<span>', 'EndMark') & @TAB
			$sTxt &= GetItemStringByMark(GetItemStringByMark ($sContent, '<strong>Puertas</strong>', '</span>') & 'EndMark', '<span>', 'EndMark') & @TAB
			$sTxt &= $Element
			;Write data into a file
			WriteTxtFile ($sFileName, @CRLF & $sTxt, "append")
		 Next
		 $iVehicleNum += 48
	  WEnd
   Next
   MsgBox ($MB_TOPMOST, "", "done")
EndFunc





