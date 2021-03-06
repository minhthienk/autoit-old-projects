#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

;Script Start - Add your code below here
;~ #include <IE.au3>
#include <Array.au3>
#include "_IE.au3"

HotKeySet ("{ESC}", "AutoIT_Exit")


Global $iTryAttach = 1
Global $iVisible = 1
Global $iWait = 1
Global $iTakeFocus = 1
Global $sFilePath = @ScriptDir

;Create an IE object with a link
Local $sLink = 'https://www.ebay.com/itm/New-13-Ford-C-Max-Engine-Control-Module-ECM-ECU-VIN-C-7th-digit-thru-VD-OEM/252261909257?hash=item3abbfb4309:g:RL4AAOSwa-dWoh2O'
Local $oIE = _IECreate ($sLink, $iTryAttach, $iVisible, $iWait, $iTakeFocus)

;Get object Year then get the content inside
ConsoleWrite ('Begin to get year' & @CRLF)
Local $oSearchYear = _IEGetObjById ($oIE, 'Year')
Local $aYears = GetContent ($oSearchYear)

;Loop year
For $iYear = 1 To UBound ($aYears) - 2
   ;Take focus on the combobox object Year
   _IEAction ($oSearchYear, 'focus')

   ;Press down button to select option
   PressDown (2)

   ;Get object Make then get the content inside
   ConsoleWrite ('Begin to get make' & @CRLF)
   Local $oSearchMake = _IEGetObjById ($oIE, 'Make')
   Local $aMakes = GetContent ($oSearchMake)

   ;Find make Ford position
   For $iMake = 1 To UBound ($aMakes) - 2
	  If $aMakes[$iMake] = 'Subaru' Then
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
   Local $oSearchModel = _IEGetObjById ($oIE, 'Model')
   Local $aModels = GetContent ($oSearchModel)

   ;Loop model
   For $iModel = 1 To UBound ($aModels) - 2

	  ;Take focus on combobox Model
	  _IEAction ($oSearchModel, 'focus')
	  PressDown (1)

	  ;Get object Trim then get the content inside
	  ConsoleWrite ('Begin to get trim' & @CRLF)
	  Local $oSearchTrim = _IEGetObjById ($oIE, 'Trim')
	  Local $aTrims = GetContent ($oSearchTrim)

	  ;Loop Trim
	  For $iTrim = 1 To UBound ($aTrims) - 2

		 ;Take focus on combobox Trim
		 _IEAction ($oSearchTrim, 'focus')
		 PressDown (1)

		 ;Get object Engine then get the content inside
		 ConsoleWrite ('Begin to get engine' & @CRLF)
		 Local $oSearchEngine = _IEGetObjById ($oIE, 'Engine')
		 Local $aEngines = GetContent ($oSearchEngine)

		 ;Write YMME to a file
		 For $iEngine = 1 To UBound ($aEngines) - 2
			Local $sFileName = 'Ebay_YMME'
			Local $sTxt = $aYears[$iYear] & @TAB & $aMakes[$iMake] & @TAB & $aModels[$iModel] & @TAB & $aTrims[$iTrim] & @TAB & $aEngines[$iEngine] & @CRLF
			WriteTxtFile ($sFileName, $sTxt, 'append')
			Sleep (200)
		 Next
	  Next
   Next
Next





;#FUNCTION# ====================================================================================================================
;Author ........:
;Modified ......:
;===============================================================================================================================
Func GetContent ($oObj)
   While StringInStr  (_IEPropertyGet ($oObj, "outerhtml"), 'disabled') <> 0
   WEnd
   Sleep (200)
   Local $oTags = _IETagNameGetCollection($oObj, 'option')
   Local $sText = ''
   For $oTag in $oTags
	  $sText &= $oTag.innertext & @CRLF
   Next

   Local $aArray = StringSplit ($sText, @CRLF, $STR_ENTIRESPLIT + $STR_NOCOUNT )
   Return $aArray
EndFunc

;#FUNCTION# ====================================================================================================================
;Author ........:
;Modified ......:
;===============================================================================================================================
Func PressDown ($numb)
   For $i=1 To $numb
	  ControlSend ('New 13 Ford C-Max Engine Control Module ECM ECU VIN C 7th digit thru VD OEM', '', 'Internet Explorer_Server1', '{DOWN}')
	  Sleep (100)
   Next
EndFunc

;#FUNCTION# ====================================================================================================================
;Author ........:
;Modified ......:
;===============================================================================================================================
Func TakeFocus ($oObj)
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
