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
#include "VINsCollector.au3"




Func GetLastPageNum ($sLink)
   ;Get html source
   Local $html = GetHtmlSourceUsingHttpRequest ($sLink)
   Local $iLastPageNum = 0
   $iLastPageNum = GetItemStringByMark(GetItemStringByMark ($html, '<span class="last">' , 'Last'), 'page=', '"')
   If $iLastPageNum = '' Then $iLastPageNum = 1
   Return $iLastPageNum
EndFunc





;====================================================================================================================
;                  FUNCTION DISCRIPTION: LOAD FILE CONTENT
;                  INPUT			   :
; 				   OUTPUT			   :
;====================================================================================================================
Func LoadConfig ($sFileName)
   ;Open YMME config file and get data
   Local $hFileOpen = FileOpen ($sFilePath & "\" & $sFileName & ".txt",$FO_READ )
   Local $sFileRead = FileRead($hFileOpen)
   FileClose($hFileOpen)
   ;Remove redundant @CRLF
   $sFileRead = StringRegExpReplace ($sFileRead, '[\r\n]+', @CRLF)
   $sFileRead = StringRegExpReplace ($sFileRead, '[\r\n]+$', '')
   Return $sFileRead
EndFunc



;====================================================================================================================
;                  FUNCTION DISCRIPTION: LOAD YEAR FROM CONFIG FILE
;                  INPUT			   :
; 				   OUTPUT			   :
;====================================================================================================================
Func LoadYearBegin ()

   GUICtrlSetData ($Commu_Ctrl, 'Loading information ...' & @CRLF & 'Please wait!')
   ;Get years from config file and put into an array

   Local $sYearBegin = StringRegExpReplace ($sConfigData, '(<year> ---)( .+)', '')
   $sYearBegin = StringReplace ($sYearBegin, '<year>', '')
   ;String => Array

   Local $aYearBegin = StringSplit ($sYearBegin, @CRLF, $STR_ENTIRESPLIT + $STR_NOCOUNT)
   ;Remove dupplicates
   $aYearBegin = _ArrayUnique ($aYearBegin, 0, 0, 0, $ARRAYUNIQUE_NOCOUNT)
   ;Sort from less to greater
   _ArraySort ($aYearBegin)

   ;Array => String
   Local $sYears = _ArrayToString ($aYearBegin, '|')
   ;Load years to combo box $Combo_Year_Begin
   GUICtrlSetData($Combo_Year_Begin, "")
   GUICtrlSetData($Combo_Year_Begin, $sYears)
   _GUICtrlComboBox_SetCurSel($Combo_Year_Begin, 0)

   ;Load years to combo box $Combo_Year_End
   GUICtrlSetData($Combo_Year_End, "")
   GUICtrlSetData($Combo_Year_End, $sYears)
   _GUICtrlComboBox_SetCurSel($Combo_Year_End, 0)
   ;----------------------

   YearEndSelected ()
   GUICtrlSetData ($Commu_Ctrl, 'Press <Write Config> to write a new YMME config file' & @CRLF & 'Press <Get VIN> after selecting YMME to begin getting VINs')
EndFunc

Func YearBeginSelected ()

   GUICtrlSetData ($Commu_Ctrl, 'Loading information ...' & @CRLF & 'Please wait!')
   ;Get years from config file and put into an array
   Local $sYearBeginSelected = GUICtrlRead ($Combo_Year_Begin)
   Local $sYearEnd = StringRegExpReplace ($sConfigData, '(<year> ---)( .+)', '')
   $sYearEnd = StringReplace ($sYearEnd, '<year>', '')

   ;String => Array
   Local $aYearEnd = StringSplit ($sYearEnd, @CRLF, $STR_ENTIRESPLIT + $STR_NOCOUNT)
   ;Remove dupplicates
   $aYearEnd = _ArrayUnique ($aYearEnd, 0, 0, 0, $ARRAYUNIQUE_NOCOUNT)
   ;Sort from less to greater
   _ArraySort ($aYearEnd)

   ;Array => String
   $sYears = _ArrayToString ($aYearEnd, '|')
   ;Remove years less than begin
   If StringInStr ($sYears, $sYearBeginSelected - 1) <> 0 Then
	  Local $iPos = StringInStr ($sYears, $sYearBeginSelected - 1) + StringLen ($sYearBeginSelected - 1)
	  $sYears = StringMid($sYears, $iPos + 1, StringLen($sYears) - $iPos + 1)
   EndIf

   ;String => Array
   Local $aYearEnd = StringSplit ($sYears, '|', $STR_ENTIRESPLIT + $STR_NOCOUNT)
   ;-----------------------------------
   ;Decide which line will be top line
   Local $sPreYearEnd = GUICtrlRead ($Combo_Year_End)
   ;Load years to combo box $Combo_Year_End
   GUICtrlSetData($Combo_Year_End, "")
   GUICtrlSetData($Combo_Year_End, $sYears)
   ;Search the position of item in the array
   Local $iTopIndex = _ArraySearch ($aYearEnd, $sPreYearEnd)
   ;Set top item
   If $iTopIndex <> -1 Then
	  _GUICtrlComboBox_SetCurSel($Combo_Year_End, $iTopIndex)
   Else
	  _GUICtrlComboBox_SetCurSel($Combo_Year_End, 0)
   EndIf
   ;----------------------------------
   YearEndSelected ()
   GUICtrlSetData ($Commu_Ctrl, 'Press <Write Config> to write a new YMME config file' & @CRLF & 'Press <Get VIN> after selecting YMME to begin getting VINs')
EndFunc


Func YearEndSelected ()
   GUICtrlSetData ($Commu_Ctrl, 'Loading information ...' & @CRLF & 'Please wait!')
   ;Get data from config
   Local $sMakeList = ""
   ;Remove unselected years
   For $i = Number(GUICtrlRead ($Combo_Year_Begin))To Number(GUICtrlRead ($Combo_Year_End))
	  Local $sTemp = $sConfigData
	  $sTemp = StringRegExpReplace ($sTemp, '(<year>)(?!' & $i & ')(.+)(<year>)(.+)\r\n', '')
	  $sTemp = StringRegExpReplace ($sTemp, '\r\n(<year>)(?!' & $i & ')(.+)(<year>)(.+)', '')
	  $sTemp = StringRegExpReplace ($sTemp, '(<year>)(?!' & $i & ')(.+)(<year>)(.+)', '')
	  $sMakeList &= $sTemp & @CRLF
   Next

   ;Remove redundant @CRLF
   $sMakeList = StringRegExpReplace ($sMakeList, '[\r\n]+', @CRLF)
   $sMakeList = StringRegExpReplace ($sMakeList, '[\r\n]+$', '')

   ;Get Make List
   $sMakeList = StringRegExpReplace ($sMakeList, '(<year>)(.+)(--- <make>)', '')
   $sMakeList = StringRegExpReplace ($sMakeList, '(<make>)(.+)', '')

   ;String => Array
   Local $aMakeList = StringSplit ($sMakeList, @CRLF, $STR_ENTIRESPLIT + $STR_NOCOUNT)

   ;Remove dupplicates
   $aMakeList = _ArrayUnique ($aMakeList, 0, 0, 0, $ARRAYUNIQUE_NOCOUNT)
   ;Sort from less to greater
   _ArraySort ($aMakeList)


   ;Array => String
   Local $sMakes = _ArrayToString ($aMakeList, '|')

   ;---------------------------------
   ;Decide which line will be top line
   Local $sPreMake = GUICtrlRead ($Combo_Make)
   ;Load makes to combo box $Combo_Make
   GUICtrlSetData($Combo_Make, "")
   GUICtrlSetData($Combo_Make, $sMakes)

   Local $iTopIndex = _ArraySearch ($aMakeList, $sPreMake)

   If $iTopIndex <> -1 Then
	  _GUICtrlComboBox_SetCurSel($Combo_Make, $iTopIndex)
   Else
	  _GUICtrlComboBox_SetCurSel($Combo_Make, 0)
   EndIf
   ;-------------------
   MakeSelected ()
   GUICtrlSetData ($Commu_Ctrl, 'Press <Write Config> to write a new YMME config file' & @CRLF & 'Press <Get VIN> after selecting YMME to begin getting VINs')
EndFunc


Func MakeSelected ()
   GUICtrlSetData ($Commu_Ctrl, 'Loading information ...' & @CRLF & 'Please wait!')
   ;Get data from config
   Local $sModelList = ""
   ;Remove unselected years
   For $i = Number(GUICtrlRead ($Combo_Year_Begin))To Number(GUICtrlRead ($Combo_Year_End))
	  Local $sTemp = $sConfigData
	  $sTemp = StringRegExpReplace ($sTemp, '(<year>)(?!' & $i & ')(.+)(<year>)(.+)\r\n', '')
	  $sTemp = StringRegExpReplace ($sTemp, '\r\n(<year>)(?!' & $i & ')(.+)(<year>)(.+)', '')
	  $sTemp = StringRegExpReplace ($sTemp, '(<year>)(?!' & $i & ')(.+)(<year>)(.+)', '')
	  $sModelList &= $sTemp & @CRLF
   Next

   ;Remove redundant @CRLF
   $sModelList = StringRegExpReplace ($sModelList, '[\r\n]+', @CRLF)
   $sModelList = StringRegExpReplace ($sModelList, '[\r\n]+$', '')

   ;Remove unselected makes
   Local $sSelectedMake = GUICtrlRead ($Combo_Make)
   $sModelList = StringRegExpReplace ($sModelList, '((<year>)(.+)(--- <make>))(?!' & $sSelectedMake & ')(.+)\r\n', '')
   $sModelList = StringRegExpReplace ($sModelList, '\r\n((<year>)(.+)(--- <make>))(?!' & $sSelectedMake & ')(.+)', '')
   $sModelList = StringRegExpReplace ($sModelList, '((<year>)(.+)(--- <make>))(?!' & $sSelectedMake & ')(.+)', '')

   ;Get models
   $sModelList = StringRegExpReplace ($sModelList, '(.+)( --- <model>)', '')
   $sModelList = StringRegExpReplace ($sModelList, '(<model>)', '')

   ;String => Array
   Local $aModelList = StringSplit ($sModelList, @CRLF, $STR_ENTIRESPLIT + $STR_NOCOUNT)
   ;Remove dupplicates
   $aModelList = _ArrayUnique ($aModelList, 0, 0, 0, $ARRAYUNIQUE_NOCOUNT)
   ;Sort from less to greater
   _ArraySort ($aModelList)
   ;Array => String
   Local $sModels = '<All Models>|' & _ArrayToString ($aModelList, '|')

   ;---------------------------------
   ;Load makes to combo box $Combo_Make
   GUICtrlSetData($Combo_Model, "")
   GUICtrlSetData($Combo_Model, $sModels)
   _GUICtrlComboBox_SetCurSel($Combo_Model, 0)

   ;-------------------
   GUICtrlSetData ($Commu_Ctrl, 'Press <Write Config> to write a new YMME config file' & @CRLF & 'Press <Get VIN> after selecting YMME to begin getting VINs')

EndFunc


;====================================================================================================================
;                  FUNCTION DISCRIPTION:
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func GetInfoContent ($sLink)
   Local $html = GetHtmlSourceUsingHttpRequest ($sLink)
   ;Start mark and end mark used to mark 2 ends of the neccessary content

   Local $sStartMark = '<div aria-label='
   Local $sEndMark = '</div>'
   ;If the mark exists
   If StringInStr ($html, $sStartMark, 0, 1, 1) <> 0 Then
	  ;Get the content
	  Local $iStart = StringInStr ($html, $sStartMark, 0, 1, 1)
	  Local $iEnd = StringInStr ($html, $sEndMark, 0, 1, $iStart)
	  Local $sInfoContent = StringMid ($html, $iStart, $iEnd - $iStart)

	  ;Convert string to array
	  Local $aInfoContent = StringSplit ($sInfoContent, @LF, $STR_ENTIRESPLIT)

	  Local $aRefinedInfo [0]
	  ;Get neccessary string in each line
	  For $i = 1 To $aInfoContent[0]
		 If StringInStr ($aInfoContent[$i], '<a class="btn btn-default"') <> 0 Then
			ReDim $aRefinedInfo [UBound($aRefinedInfo)+1]
			$aRefinedInfo [UBound($aRefinedInfo)-1] = GetItemStringByMark ($aInfoContent[$i], '">', '</a>')
		 EndIf
	  Next
   Else
	  Local $aRefinedInfo [1] = ('No infomation')
   EndIf
   ;Return an array of the information
   Return $aRefinedInfo
EndFunc

;====================================================================================================================
;                  FUNCTION DISCRIPTION:
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func GetHtmlSourceByInet ($sLink)
   ;Get html source
   Local $html = ''
   For $i = 1 to 10
	  Local $iError = 0
	  ConsoleWrite ('Begin to get source    :' & $sLink & @CRLF)
	  $html = InetRead($sLink, $INET_FORCERELOAD )
	  $iError += @error
	  $html = BinaryToString($html)
	  $iError += @error
	  Sleep (300)
	  If $iError = 0 Then
		 ConsoleWrite ('Complete getting source' & @CRLF)
		 ExitLoop
	  Else
		 If $i = 10 Then ConsoleWrite ('Error: ' & $iError & ': Can not get source' & @CRLF)
	  EndIf
   Next
   ;Refined html source
   $html = StringReplace ($html, '><', '>' & @CRLF & '<')
   _ClipBoard_SetData ($html)
   Exit
   Return $html
EndFunc


;====================================================================================================================
;                  FUNCTION DISCRIPTION:
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func GetHtmlSourceUsingHttpRequest ($sLink)
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
			GUICtrlSetData ($Commu_Ctrl, 'Object error! Relink ...')
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
			ConsoleWrite ('   => Object failed/ Relink')
			GUICtrlSetData ($Commu_Ctrl, 'The request action with HTTP Object has failed!' & @CRLF & 'Check the internet connection and wait for the tool to reconnect.')
			Sleep (3000)
		 Else
			If $oHTTP.Status = 403 Then
			   ConsoleWrite ('   => Access denied')
			   GUICtrlSetData ($Commu_Ctrl, 'Error 403: Access denied' & @CRLF & 'The owner of this website has banned your IP address!' & @CRLF & 'You must use fake IP to continue using the tool!')
			   Local $sText = ('=====================================================================================' & @CRLF & _
						 'Error 403: Access denied'    & @CRLF  & @CRLF & _
						 '          The owner of this website has banned your IP address!'   & @CRLF & _
						 '          Current Link: ' & $sLink)
			   Write_Error (@CRLF & @CRLF & $sText)
			   Sleep (30000)
			ElseIf $oHTTP.Status = 200 Then
			   ExitLoop
			Else ;Other status
			   Local $sText = ('=====================================================================================' & @CRLF & _
						 'Error staus: ' &  $oHTTP.Status   & @CRLF  & @CRLF & _
						 '          An unexpected error has occured'   & @CRLF & _
						 '          Current Link: ' & $sLink)
			   Write_Error (@CRLF & @CRLF & $sText)
			EndIf
		 EndIf
	  WEnd

	  ConsoleWrite (' => Status ' & $oHTTP.Status & ' => Get response')
	  $html = $oHTTP.Responsetext
	  ;Release object
	  ConsoleWrite (' => Replease COM')
	  $oHTTP = 0

   ;Check content
   If StringInStr ($html, 'The page you were looking for was not found') <> 0 Then
	  $html = ''
   EndIf
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




Func Autoit_Exit ()
   Exit
EndFunc