#cs ----------------------------------------------------------------------------
	This library is to process strings
#ce ----------------------------------------------------------------------------

#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>
#include <Array.au3>
#include <IE.au3>
#include <GUIConstantsEx.au3>
#include <InetConstants.au3>
#include <Clipboard.au3>
#include <StringConstants.au3>

;====================================================================================================
;This function is to remove redundant enter characters in a string just keep 1 enter each line
;Also remove beginning and ending enters
;====================================================================================================
Func StringRemoveRedundantEnter ($sString)
	$sString = StringRegExpReplace($sString, '[\r\n]+', @CRLF)	;replace 2 enters by 1 enter
	$sString = StringRegExpReplace($sString, '[\r\n]+$', '')	;remove beginning enters
	$sString = StringRegExpReplace($sString, '^[\r\n]+', '')	;remove ending enters
	Return $sString
EndFunc



;====================================================================================================
;This function is to create GUI
;====================================================================================================
Func GUIInit()
	Opt('GUIOnEventMode', 1)
	$Form1 = GUICreate('Collector', 500, 200, -1, -1)
	GUISetOnEvent($GUI_EVENT_CLOSE, 'Autoit_Exit')
	GUISetBkColor(0xFFFFFF)
	;-------------------------------------------
	;Create Get Info button
	$Button_Info = GUICtrlCreateButton('Get Files', 200, 30, 100, 50)
	GUICtrlSetOnEvent($Button_Info, 'Set_GetHtml_Flag')
	;-------------------------------------------
	;CREATE GUI NOTIFICATION PLACE
	$Commu_Ctrl = GUICtrlCreateLabel('', 50, 110, 400, 50)
	$CopyRight = GUICtrlCreateLabel('Created by Thien Nguyen', 190, 180, 150, 50)
	;-------------------------------------------
	;Create input
	;$sDelay = GUICtrlCreateInput("", 32, 60, 265, 21)
	;GUICtrlCreateLabel("Input Vehicle Link ", 120, 40, 114, 17)

	;SHOW GUI
	GUISetState(@SW_SHOW)
	GUICtrlSetData ($Commu_Ctrl, 'Press the button to begin to get the documents')
EndFunc




;====================================================================================================
;This function is to get options in a form object and put in to an array
;====================================================================================================
Func _IEGetOptions (Byref $oIE, Byref $oObject)
	Local $shtml = _IEPropertyGet ($oObject, 'outerhtml')		;get html of object
	$shtml = StringReplace($shtml, '><', '>' & @CRLF & '<')		;replace '><' by enter in between to have multi-line string
	$shtml = StringRegExpReplace($shtml, '<.+".+">', '')		;remove redundant text to get options text only
	$shtml = StringRegExpReplace($shtml, '<.+>', '')			;remove redundant text to get options text only
	$shtml = StringRegExpReplace($shtml, '<.+>', '')			;remove redundant text to get options text only
	Local $sAllOptions = _StringRemoveRedundantEnter ($shtml)	;remove redundant enters
	Local $aOption = StringSplit($sAllOptions, @CRLF,  $STR_ENTIRESPLIT + $STR_NOCOUNT)	;convert string options to an arrray
	;return the array of options
	Return $aOption
EndFunc