#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include <Array.au3>
#include "_IE.au3"

HotKeySet ("{End}", "AutoIT_Exit")



Global $iTryAttach = 1
Global $iVisible = 1
Global $iWait = 1
Global $iTakeFocus = 1
Global $sFilePath = @ScriptDir

;Create an IE object with a link
Global $sLink = 'https://www.knfiltros.com/search/part_search.aspx'
Global $oIE = _IECreate ($sLink, $iTryAttach, $iVisible, $iWait, $iTakeFocus)

;Get object Year then get the content inside
ConsoleWrite ('Begin to get year' & @CRLF)
Local $oSearchYear = _IEGetObjById ($oIE, 'AppSearchControlHorizontal_ddl_year')
Local $aYears = GetContent ('AppSearchControlHorizontal_ddl_year', 'No check')


;Loop year
For $iYear = 2 To UBound ($aYears) - 3
   ;Take focus on the combobox object Year
   _IEAction ($oSearchYear, 'focus')
   ;Press down button to select option
   PressDown (1)

   ;Get object Make then get the content inside
   ConsoleWrite ('Begin to get make' & @CRLF)
   Local $oSearchMake = _IEGetObjById ($oIE, 'AppSearchControlHorizontal_ddl_make')
   Local $aMakes = GetContent ('AppSearchControlHorizontal_ddl_make')

   ;Find make Ford position
   For $iMake = 1 To UBound ($aMakes) - 2
	  If $aMakes[$iMake] = 'Ford' Then
		 Local $iFordPosition = $iMake
		 ExitLoop
	  EndIf
   Next

   ;Take focus on combobox Make and press down until see Make Ford
   _IEAction ($oSearchMake, 'focus')
   PressDown ($iFordPosition)
   Sleep (1000)


   ;Get object Model then get the content inside
   ConsoleWrite ('Begin to get model' & @CRLF)
   Local $oSearchModel = _IEGetObjById ($oIE, 'AppSearchControlHorizontal_ddl_model')
   Local $aModels = GetContent ('AppSearchControlHorizontal_ddl_model')



   ;Loop model
   For $iModel = 1 To UBound ($aModels) - 2

	  ;Take focus on combobox Model
	  _IEAction ($oSearchModel, 'focus')
	  PressDown (1)

		 ;Get object Engine then get the content inside
		 ConsoleWrite ('Begin to get engine' & @CRLF)
		 Local $oSearchEngine = _IEGetObjById ($oIE, 'AppSearchControlHorizontal_ddl_enginesize')
		 Local $aEngines = GetContent ('AppSearchControlHorizontal_ddl_enginesize')

		 ;Write YMME to a file
		 For $iEngine = 1 To UBound ($aEngines) - 2
			Local $sFileName = 'knfiltros_YMME'
			Local $sTxt = $aYears[$iYear] & @TAB & $aMakes[$iMake] & @TAB & $aModels[$iModel] & @TAB & $aEngines[$iEngine] & @CRLF
			WriteTxtFile ($sFileName, $sTxt, 'append')
			Sleep (200)
		 Next
   Next
Next





;#FUNCTION# ====================================================================================================================
;Author ........:
;Modified ......:
;===============================================================================================================================
Func GetContent ($sID, $sCheck = 'Check')
   If $sCheck = 'Check' Then
	  WaitChange ()
   EndIf

   $oObj = _IEGetObjById ($oIE, $sID)

   ;Get tag option objects
   Local $oTags = _IETagNameGetCollection($oObj, 'option')
   Local $sText = ''
   ;Get text from tag option
   For $oTag in $oTags
	  $sText &= $oTag.innertext & @CRLF
   Next
   ;Convert text to array
   Local $aArray = StringSplit ($sText, @CRLF, $STR_ENTIRESPLIT + $STR_NOCOUNT )
   Return $aArray
EndFunc

;#FUNCTION# ====================================================================================================================
;Author ........:
;Modified ......:
;===============================================================================================================================
Func PressDown ($numb)
   For $i=1 To $numb
	  ControlSend ('K&N Búsqueda', '', 'Internet Explorer_Server1', '{DOWN}')
	  Sleep (100)
   Next
EndFunc

;#FUNCTION# ====================================================================================================================
;Author ........:
;Modified ......:
;===============================================================================================================================
Func TakeFocus (Byref $oObj)
   While StringInStr  (_IEPropertyGet ($oObj, "outerhtml"), 'disabled') = 0
   WEnd
   While StringInStr  (_IEPropertyGet ($oObj, "outerhtml"), 'disabled') <> 0
   WEnd
   _IEAction($oObj, 'focus')
EndFunc

;#FUNCTION# ====================================================================================================================
;Author ........:
;Modified ......:
;===============================================================================================================================
Func WriteTxtFile ($sFileName, $sTxt, $sMode = "append")
   If $sMode = "overwrite" Then
	  Local $hFileOpen = FileOpen ($sFilePath & "\" & $sFileName & ".txt",$FO_OVERWRITE)
   Else
	  Local $hFileOpen = FileOpen ($sFilePath & "\" & $sFileName & ".txt",$FO_APPEND)
   EndIf
   FileWrite($hFileOpen, $sTxt)
   FileClose($hFileOpen)
EndFunc


;#FUNCTION# ====================================================================================================================
;Author ........:
;Modified ......:
;===============================================================================================================================
Func AutoIT_Exit()
   Exit
EndFunc




;#FUNCTION# ====================================================================================================================
;Author ........:
;Modified ......:
;===============================================================================================================================
Func WaitChange ()
	  Local $oCheckValue = _IEGetObjById ($oIE, 'ScriptManager1')
	  Static Local $sCheckValue = $oCheckValue.value
	  While 1
		 Sleep (100)
		 If $oCheckValue.value <> $sCheckValue Then
			$sCheckValue = $oCheckValue.value
			ExitLoop
		 EndIf
	  WEnd
	  Sleep (1000)
	  ControlSend ('K&N Búsqueda', '', 'Internet Explorer_Server1', '{ESC}')
EndFunc



