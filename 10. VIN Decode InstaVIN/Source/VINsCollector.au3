#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         Thien Nguyen

 Script Function: Getting VINs and create database for INNOVA VIN Decode

 ***********
 Change Log:
 v01.00.01
   - Created demo version

 v01.00.02
   - Removed sleep when finding YMME which has been collected
   - Resized communication screen
   - Updated "Get Info" function, the last version can not get enough info from all pages (only get from page 1)
   - Updated "Get VINs" function: added full VIN next to VIN8

 v01.00.03
   - Updated COMErrorHandler: added Error Position marked by current link
   - Fixed bug display missing items because of the lack of "@CRLF" when delecting unslected years
   - Updated GetHtmlSourceUsingHttpRequest to display http error status when Ip is banned and write error log

#ce ----------------------------------------------------------------------------

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

;#include "COMErrorHandler.au3"
#include "Library.au3"

; Set Hotkey for the program
;~ HotKeySet("{ESC}", "Autoit_Exit")

Global $bGetVINs_Flag = False
Global $bGetInfo_Flag = False
Global $sFilePath = @ScriptDir
Global $sConfigData = LoadConfig ('YMME_Config')
Global $sDataFileName = 'VINs_Collector_Database'


Func Set_GetVINs_Flag ()
   $bGetVINs_Flag = True
EndFunc

Func Set_GetInfo_Flag ()
   $bGetInfo_Flag = True
EndFunc



#Region ### START GUI section ### Form=
   Opt('GUIOnEventMode', 1)
   $Form1 = GUICreate('VINs Collector v01.00.03 Aug0818', 329, 225, -1, -1)
   GUISetOnEvent($GUI_EVENT_CLOSE, 'Autoit_Exit')
   GUISetBkColor(0xFFFFFF)
   ;-------------------------------------------
   ;Create Labels
   GUICtrlCreateLabel('From', 20, 23, 50, 17)
   GUICtrlCreateLabel('To', 170, 23, 50, 17)
   GUICtrlCreateLabel('Make', 20, 53, 50, 17)
   GUICtrlCreateLabel('Model', 20, 83, 50, 17)
   ;-------------------------------------------
   ;Select Beginning Year
   $Combo_Year_Begin = GUICtrlCreateCombo('', 60, 20, 100, 20)
   GUICtrlSetOnEvent($Combo_Year_Begin, 'YearBeginSelected')
   ;-------------------------------------------
   ;Select End Year
   $Combo_Year_End = GUICtrlCreateCombo('', 190, 20, 100, 20)
   GUICtrlSetOnEvent($Combo_Year_End, 'YearEndSelected')
   ;-------------------------------------------
   ;Select Make
   $Combo_Make = GUICtrlCreateCombo('', 60, 50, 230, 20)
   GUICtrlSetOnEvent($Combo_Make, 'MakeSelected')
   ;-------------------------------------------
   ;Select Model
   $Combo_Model = GUICtrlCreateCombo('', 60, 80, 230, 20)
   GUICtrlSetOnEvent($Combo_Model, 'ModelSelected')

   ;-------------------------------------------
   ;Create Begin button
   $Button_Begin = GUICtrlCreateButton('Get VINs', 210, 110, 80, 25)
   GUICtrlSetOnEvent($Button_Begin, 'Set_GetVINs_Flag')

   ;-------------------------------------------
   ;Create Get Info button
   $Button_Info = GUICtrlCreateButton('Write Config', 60, 110, 80, 25)
   GUICtrlSetOnEvent($Button_Info, 'Set_GetInfo_Flag')


   ;-------------------------------------------
   ;CREATE GUI NOTIFICATION PLACE
	  $Commu_Ctrl = GUICtrlCreateLabel('', 20, 150, 289, 50)
	  $CopyRight = GUICtrlCreateLabel('Created by Thien Nguyen', 100, 210, 309, 50)
   ;-------------------------------------------
   ;SHOW GUI
#EndRegion ### END Koda GUI section ###


LoadYearBegin ()
GUICtrlSetData ($Commu_Ctrl, 'Press <Write Config> to write a new YMME config file' & @CRLF & 'Press <Get VIN> after selecting YMME to begin getting VINs')
GUISetState(@SW_SHOW)

While 1
   If $bGetVINs_Flag = True Then InstaVIN()
   ;If $bGetInfo_Flag = True Then InstaVIN ()
WEnd



Func InstaVIN()

   $sListVIN = LoadFileNames ('Select VIN File')

   Local $aVIN = StringSplit ($sListVIN, @CRLF, $STR_ENTIRESPLIT  +  $STR_NOCOUNT)

   ;create IE object with the link
   Local $oIE = _IECreate ('about:blank',0,1,1,0)

   ;----------------------------
   ;Decode vin here
   Local $aVIN8[0]
   Local $aVIN_RemoveDup[0]

   Local $sTxt = ''
   Local $sPreviousDecoded = ''
   For $l = 0 To UBound ($aVIN) - 1
		 GUICtrlSetData ($Commu_Ctrl, 'Decoding ' & $aVIN[$l] & @CRLF &'Previous decoded: ' & $sPreviousDecoded)

		 ;Get VIN Decode Infor
		 Local $sLinkVINDecode = 'https://www.instavin.com/order?VIN=' & $aVIN[$l]

		 _IENavigate ($oIE, $sLinkVINDecode, 1)

		 While 1
			Local $html = _IEPropertyGet ($oIE, 'outerhtml')
			If StringInStr($html, 'Please wait...') = 0 Then ExitLoop
		 WEnd

		 ;_ClipBoard_SetData($html)


		 Local $sYearMakeModel = GetItemStringByMark ($html, '<h1 class="OrderFreeVinCheckResults_Header">', '</h1>')


		 Local $sTrim = GetItemStringByMark ($html, 'Trim Level:', '</div>')


		 Local $sEngine = GetItemStringByMark ($html, 'Engine:', '</div>')


		 $sTxt &= $sYearMakeModel & @TAB & $sTrim & @TAB & ' ' & $sEngine & @TAB & StringLeft($aVIN[$l],8)  & @TAB & $aVIN[$l] & @CRLF
		 $sPreviousDecoded = $sYearMakeModel & '  ' & $sTrim & '  ' & $sEngine

   Next

		 ;----------------------------
		 ;WRITE YMME CONFIG FILE
		 Local $sFileName = $sDataFileName
		 WriteTxtFile ($sFileName, $sTxt, "append")

   $bGetVINs_Flag = False
EndFunc



Func LoadFileNames ($sTitle)
   ;Open YMME config file and get data

   Local $FilePath = FileOpenDialog ($sTitle, @ScriptDir, '(*.txt)')

   Local $hFileOpen = FileOpen ($FilePath,$FO_READ )
   Local $sFileRead = FileRead($hFileOpen)
   FileClose($hFileOpen)
   ;Remove redundant @CRLF
   $sFileRead = StringRegExpReplace ($sFileRead, '[\r\n]+', @CRLF)
   $sFileRead = StringRegExpReplace ($sFileRead, '[\r\n]+$', '')
   Return $sFileRead
EndFunc




