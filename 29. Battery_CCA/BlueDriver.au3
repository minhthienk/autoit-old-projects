#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include <array.au3>
#include <IE.au3>
#include <clipboard.au3>
#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>
#include <Sound.au3>

HotKeySet('{ESC}', 'Autoit_Exit')


Local $oIE = _IECreate('https://www.autobatteries.com',1,1,1,1)

Local $oYear = _IEGetObjById($oIE, 'product_finder_year')
Local $oMake = _IEGetObjById($oIE, 'product_finder_make')
Local $oModel = _IEGetObjById($oIE, 'product_finder_model')
Local $oEngine = _IEGetObjById($oIE, 'product_finder_engine')





Local $sStartYear = '2019'
Local $sStartMake = 'Ford'
Local $sStartModel = 'yaris'






Local $bCheckStartYear_Flag = False
Local $bCheckStartMake_Flag = False
Local $bCheckStartModel_Flag = False
Local $bContinue_Flag = True



$aYearOption = _IEGetOptions($oIE, $oYear)

For $i = 1 To UBound($aYearOption)
	ConsoleWrite($aYearOption[$i] & @CRLF)
	_IEFormElementSetValue($oYear, $aYearOption[$i])    ;set value to model
	_IECheckLoadDone ('make')
Next


Exit


;====================================================================================================
Func _IEGetOptions (Byref $oIE, Byref $oObject)
    Local $shtml = _IEPropertyGet ($oObject, 'outerhtml')        ;get html of object
    $shtml = StringReplace($shtml, '><', '>' & @CRLF & '<')        ;replace '><' by enter in between to have multi-line string
    $shtml = StringRegExpReplace($shtml, '<.+".+">', '')        ;remove redundant text to get options text only
    $shtml = StringRegExpReplace($shtml, '<.+>', '')            ;remove redundant text to get options text only
    $shtml = StringRegExpReplace($shtml, '<.+>', '')            ;remove redundant text to get options text only
    Local $sAllOptions = _StringRemoveRedundantEnter ($shtml)    ;remove redundant enters
    Local $aOption = StringSplit($sAllOptions, @CRLF,  $STR_ENTIRESPLIT + $STR_NOCOUNT)    ;convert string options to an arrray
    ;return the array of options
    Return $aOption
EndFunc

;====================================================================================================
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

;====================================================================================================
Func _StringRemoveRedundantEnter ($sString)
    ;process the string
    $sString = StringRegExpReplace($sString, '[\r\n]+', @CRLF)    ;replace 2 enters by 1 enter
    $sString = StringRegExpReplace($sString, '[\r\n]+$', '')    ;remove beginning enters
    $sString = StringRegExpReplace($sString, '^[\r\n]+', '')    ;remove ending enters
    ;return processed string
    Return $sString
EndFunc


;====================================================================================================
Func _IECheckLoadDone ($sID_string)
	Local $oObj = _IEGetObjById($oIE, $sID_string)
	Local $oAs = _IETagNameGetCollection ($oObj, 'span')            
	For $oA In $oAs                                                
	    If $oA.GetAttribute("class") = "spinner" Then
	    	Local $oCheckObject = $oA
	        ExitLoop
	    EndIf
	Next
	Sleep(200)
	while 1
		ConsoleWrite('Checking Load ...' & @CRLF)
		If ($oCheckObject.outerhtml = '<span class="spinner"></span>') Then Exitloop
		Sleep(200)
	Wend
	ConsoleWrite('Done Checking Load ...' & @CRLF)
EndFunc


;====================================================================================================
Func Autoit_Exit ()
    Exit
EndFunc













Func PressDown(byref $oObject)
   Local $sPrevious = _IEFormElementGetValue ($oObject)
   Local $time = 0
   While 1

	  _IEAction($oObject, 'focus')
	  Sleep (200)
	  ;_IEAction($oObject, 'click')
	  Sleep (200)
	  ControlSend('Support | BlueDriver - Windows Internet Explorer', '', '', '{DOWN}')
	  ConsoleWrite('Press Down to: ' & _IEFormElementGetValue ($oObject) &@CRLF)
	  Sleep (500)
	  $time += 1


	  If $time = 2 Then Exit


	  If _IEFormElementGetValue ($oObject) <> $sPrevious Then ExitLoop
   WEnd





EndFunc




For $iYear = 2017 To 2018

	  PressDown($oYear)
	  CheckLoad($oMake, 'disabled')
	  ConsoleWrite('Select Year: ' & $iYear & "  " & _IEFormElementGetValue ($oYear) & @CRLF)
	  If $bContinue_Flag = True And _IEFormElementGetValue ($oYear) <> $sStartYear Then ContinueLoop


   StringReplace($oMake.innerhtml, '</option>', '')
   For $iMake = 1 To @extended - 1

	  PressDown($oMake)
	  CheckLoad($oModel, 'disabled')
	  ConsoleWrite('Select Make: ' & $iMake & "  " & _IEFormElementGetValue ($oMake) & @CRLF)
	  If $bContinue_Flag = True And _IEFormElementGetValue ($oMake) <> $sStartMake Then ContinueLoop



	  StringReplace($oModel.innerhtml, '</option>', '')
	  For $iModel = 1 To @extended - 1

		 PressDown($oModel)
		 CheckLoad($oTrim, 'disabled')
		 ConsoleWrite('Select Model: ' & $iModel & "  " & _IEFormElementGetValue ($oModel) & @CRLF)
		 If $bContinue_Flag = True And _IEFormElementGetValue ($oModel) <> $sStartModel Then ContinueLoop

;~ 		 MsgBox (0, '', _IEFormElementGetValue ($oModel))

		 $bContinue_Flag = False

		 If _IEFormElementGetValue ($oMake) = 'tesla' Then ContinueLoop

		 ConsoleWrite('Waiting for Data' & @CRLF)
		 Sleep (550)
		 Local  $iWaitTime = 0
		 While 1
			If $oSupported.innertext <> '' Then ExitLoop
			Sleep(100)
			$iWaitTime += 100
			If $iWaitTime = 10000 Then
			   If StringInStr($oDisclaimer.innertext, 'not compatible with BlueDriver') <> 0 Then ExitLoop
			EndIf

		 WEnd


		 Local $sFilePath = @ScriptDir & '\BlueDriver.txt'
		 Local $sYear = _IEFormElementGetValue ($oYear)

		 Local $sMake = _IEFormElementGetValue ($oMake)
		 $sMake = GetItemStringByMark ($oMake.innerhtml, '"' & $sMake & '">', '</option>')

		 Local $sModel = _IEFormElementGetValue ($oModel)
		 $sModel = GetItemStringByMark ($oModel.innerhtml, '"' & $sModel & '">', '</option>')

		 Local $sTxt =   $sYear & @TAB & $sMake & @TAB & $sModel & @TAB & '"' & $oDisclaimer.innertext & @CRLF & $oSupported.innertext & '"' & @CRLF
		 WriteTxtFile ($sFilePath, $sTxt, "Append")
		 Sleep (500)
	  Next
   Next
Next



MsgBox (0, '', 'Done')





;====================================================================================================================
Func CheckLoad($oObject, $sString)
   ConsoleWrite('                 Waiting for disable appears' & @CRLF)
   Sleep (1000)

   ConsoleWrite('                 Waiting for disable disappears:  ' )
   $time = 0
   While 1
	  If StringInStr($oObject.outerhtml, $sString) = 0 Then ExitLoop
	  Sleep(100)
	  $time += 1
;~ 	  If $time = 7000 Then ExitLoop
	  ConsoleWrite($time & '  ')
   WEnd
   ConsoleWrite (@CRLF)
EndFunc




;====================================================================================================================
Func WriteTxtFile ($sFilePath, $sTxt, $sMode = "append")
   If $sMode = "overwrite" Then
	  Local $hFileOpen = FileOpen ($sFilePath,$FO_OVERWRITE)
   Else
	  Local $hFileOpen = FileOpen ($sFilePath,$FO_APPEND)
   EndIf
   FileWrite($hFileOpen, $sTxt)
   FileClose($hFileOpen)
EndFunc