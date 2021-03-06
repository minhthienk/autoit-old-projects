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

#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <GuiEdit.au3>
#include <GuiComboBox.au3>
#include <Clipboard.au3>
#include <Inet.au3>
#include <Timers.au3>
#include "Autocats.au3"


;====================================================================================================================
;                  FUNCTION DISCRIPTION:
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func GetInfoContent ($sLink)
   Local $html = GetHtmlSourceUsingHttpRequest ($sLink)
   ;Start mark and end mark used to mark 2 ends of the neccessary content

   Local $sStartMark = '<ol id="searchResults"'
   Local $sEndMark = '::after'

   ;If the mark exists
   If StringInStr ($html, $sStartMark, 0, 1, 1) <> 0 Then
	  ;Get the content
	  Local $iStart = StringInStr ($html, $sStartMark, 0, 1, 1)
	  Local $iEnd = StringInStr ($html, $sEndMark, 0, 1, $iStart)
	  Local $sInfoContent = StringMid ($html, $iStart, $iEnd - $iStart)

	  Local $aRefinedInfo[0]
	  ;Convert string to array
	  Local $aInfoContent = StringSplit ($sInfoContent, @LF, $STR_ENTIRESPLIT)

	  ;Get neccessary string in each line
	  For $i = 1 To $aInfoContent[0]
		 If StringInStr ($aInfoContent[$i], 'item-link item__js-link') <> 0 Then
			ReDim $aRefinedInfo [UBound($aRefinedInfo)+1]
			$aRefinedInfo [UBound($aRefinedInfo)-1] = GetItemStringByMark ($aInfoContent[$i], 'href="', '"')
		 EndIf
	  Next

	  If UBound ($aRefinedInfo) = 0 Then
		 ReDim $aRefinedInfo [1]
		 $aRefinedInfo [0] = 'No information'
	  EndIf

   Else
	  Local $aRefinedInfo [1]
	  $aRefinedInfo [0] = 'No information'
   EndIf


   ;Return an array of the information
   Return $aRefinedInfo
EndFunc






;====================================================================================================================
;                  FUNCTION DISCRIPTION:
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func GetHtmlSourceUsingHttpRequest ($sLink)
   Write_Log (@CRLF & $sLink)
   Sleep (300)
   ;the var $sErrorPosition is used to save the link when the error occurs
   $sErrorPosition = $sLink
   ;Get html source
   Local $html = ''
	  ;Get http object
	  ConsoleWrite ('Begin to get source: ' & $sLink & @CRLF)
	  While 1
		 ConsoleWrite ('   => Create COM')
		 Local $oHTTP = ObjCreate("winhttp.winhttprequest.5.1")
		 ;Check error
		 If @error <> 0 Then
			ConsoleWrite (' => Error: ' & @error & '/ Relink' & @CRLF)
			Sleep (500)
		 Else
			ExitLoop
		 EndIf
	  WEnd

	  While 1
		 ;Get source
		 ConsoleWrite (' => Open source')
		 $oHTTP.Open("GET", $sLink)
		 ConsoleWrite (' => Send Request')
		 $oHTTP.Send()
		 ;Check status to know if the network issue


		 If StringLen($oHTTP.Status) = 0 Then
			ConsoleWrite (' => Object failed/ Relink')
			Sleep (3000)
		 Else
			If $oHTTP.Status = 403 Then
			   ConsoleWrite ('   => Access denied')
			   Local $sText = ('=====================================================================================' & @CRLF & _
						 'Error 403: Access denied'   & @CRLF & _
						 '          The owner of this website has banned your IP address!'   & @CRLF & _
						 '          Current Link: ' & $sLink)
			   Write_Log (@CRLF & @CRLF & $sText)
			   Sleep (30000)

			ElseIf $oHTTP.Status = 200 Then
			   ExitLoop
			Else ;Other status
			   Local $sText = ('=====================================================================================' & @CRLF & _
						 'Error staus: ' &  $oHTTP.Status   & @CRLF  & @CRLF & _
						 '          An unexpected error has occured'   & @CRLF & _
						 '          Current Link: ' & $sLink)
			   Write_Log (@CRLF & @CRLF & $sText)
			   ExitLoop
			EndIf
		 EndIf
	  WEnd

	  ConsoleWrite (' => Status ' & $oHTTP.Status & ' => Get response')
	  $html = $oHTTP.Responsetext
	  ;Release object
	  ConsoleWrite (' => Replease COM')
	  $oHTTP = 0

;~    ;Check content
;~    If StringInStr ($html, 'The page you were looking for was not found') <> 0 Then
;~ 	  $html = ''
;~    EndIf

   ;Refined html source
   $html = StringReplace ($html, '><', '>' & @CRLF & '<')
   $html = StringReplace ($html, '> <', '>' & @CRLF & '<')
   ConsoleWrite (' => Completed' & @CRLF)
   Return $html
EndFunc








;====================================================================================================================
;                  FUNCTION DISCRIPTION:
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func GetItemStringByMark ($sString, $sStartMark, $sEndMark)
   If StringInStr ($sString, $sStartMark, 0, 1, 1) <> 0 Then
	  Local $iStart = StringInStr ($sString, $sStartMark, 0, 1, 1) + StringLen ($sStartMark)
	  Local $iEnd = StringInStr ($sString, $sEndMark, 0, 1, $iStart)
	  Local $sItemString = StringMid ($sString, $iStart, $iEnd - $iStart)
   Else
	  Local $sItemString = ""
   EndIf
   Return $sItemString
EndFunc



;====================================================================================================================
;                  FUNCTION DISCRIPTION: CLOSE ALL IE OBJECT
;                  INPUT			   :
; 				   OUTPUT			   :
;====================================================================================================================
Func Close_All_IE()
   $Proc = "iexplore.exe"
   While ProcessExists($Proc)
      ProcessClose($Proc)
   Wend
EndFunc



;====================================================================================================================
;                  FUNCTION DISCRIPTION: CLOSE ALL IE OBJECT
;                  INPUT			   :
; 				   OUTPUT			   :
;====================================================================================================================
Func WriteTxtFile ($sFileName, $sTxt, $sMode = "append")
   If $sMode = "overwrite" Then
	  Local $hFileOpen = FileOpen ($sFilePath & "\" & $sFileName & ".txt",$FO_OVERWRITE)
   Else
	  Local $hFileOpen = FileOpen ($sFilePath & "\" & $sFileName & ".txt",$FO_APPEND)
   EndIf
   FileWrite($hFileOpen, $sTxt)
   FileClose($hFileOpen)
EndFunc

;====================================================================================================================
;                  FUNCTION DISCRIPTION: LOAD FILE CONTENT
;                  INPUT			   :
; 				   OUTPUT			   :
;====================================================================================================================
Func LoadFile ($sFileName, $sFilePath = @ScriptDir)
   ;Open YMME config file and get data
   Local $hFileOpen = FileOpen ($sFilePath & "\" & $sFileName & ".txt",$FO_READ )
   Local $sFileRead = FileRead($hFileOpen)
   FileClose($hFileOpen)
   ;String => Array
   Local $alConfigData = StringSplit ($sFileRead, @CRLF, $STR_ENTIRESPLIT + $STR_NOCOUNT)
   Return $alConfigData
EndFunc




Func Autoit_Exit ()
   Exit
EndFunc