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

 v01.00.04
   - ArrayUnique Case sensetive

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

#include "COMErrorHandler.au3"
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
   If $bGetVINs_Flag = True Then GetVINs()
   If $bGetInfo_Flag = True Then GetInfo ()
WEnd


;====================================================================================================================
;                  FUNCTION DISCRIPTION: GET INFOR OF YMME OF DECODE THIS
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func GetVINs()
   GUICtrlSetData ($Commu_Ctrl, 'Please wait ...')

   ;Load Decoded VIN file and get data
   Local $hFileOpen = FileOpen ($sFilePath & "\" & $sDataFileName & ".txt",$FO_READ )
   Local $sCollectedYMME = FileRead($hFileOpen)
   FileClose($hFileOpen)
   ;Remove redundant data, just keep Year  Make  Model
   $sCollectedYMME = StringRegExpReplace ($sCollectedYMME, '([^\t]+)(\t)( )(.+)', '')
   $sCollectedYMME = StringRegExpReplace ($sCollectedYMME, '(\t)', '  ')
   ;String ==> Array
   Local $aCollectedYMME = StringSplit ($sCollectedYMME, @CRLF, $STR_ENTIRESPLIT + $STR_NOCOUNT)
   ;Remove dupplicates
   $aCollectedYMME = _ArrayUnique ($aCollectedYMME, 0, 0, 1, $ARRAYUNIQUE_NOCOUNT)


   ;Get Makes
   $sMake = GUICtrlRead ($Combo_Make)
   ;Get Models
   If GUICtrlRead ($Combo_Model) = '<All Models>' Then
	  Local $sModelList = _GUICtrlComboBox_GetList ($Combo_Model)
	  Local $aModelList = StringSplit ($sModelList, '|', $STR_ENTIRESPLIT + $STR_NOCOUNT)
   Else
	  Local $aModelList [2] = ['Nothinghere', GUICtrlRead ($Combo_Model)]
   EndIf
   ;Loop Range year
   For $i = Number(GUICtrlRead ($Combo_Year_Begin))To Number(GUICtrlRead ($Combo_Year_End))
	  ;Loop Model
	  For $j = 1 To UBound ($aModelList) - 1


		 ;Compare if the YMME has already collected => pass
		 If _ArraySearch ($aCollectedYMME, $i & '  ' & $sMake & '  ' & $aModelList[$j] & '  ') <> -1 Then
			GUICtrlSetData ($Commu_Ctrl, $i & '  ' & $sMake & '  ' & $aModelList[$j] & '  has already been collected before' & @CRLF & 'Please check!')
			MsgBox ($MB_TOPMOST, 'Message from VIN Collector', $i & '  ' & $sMake & '  ' & $aModelList[$j] & '  has already been collected before' & @CRLF & 'Please check!', 3)
			ContinueLoop
		 EndIf

		 ;Form links to get VINs
		 Local $sLinkVIN = 'https://www.decodethis.com/archives/' & $i & '/' & $sMake & '/' & $aModelList[$j]
		 Local $aVIN [0]
		 Local $iLastPageNum = GetLastPageNum ($sLinkVIN)

		 For $k = 1 To $iLastPageNum
			GUICtrlSetData ($Commu_Ctrl, 'Getting VINs (page ' &  $k & '/' & $iLastPageNum & ')' & @CRLF &  $i & '/' & $sMake & '/' & $aModelList[$j])
			_ArrayConcatenate ($aVIN, GetInfoContent ($sLinkVIN & '?page=' & $k))
		 Next
		 ;Decode VIN
		 DecodeVIN_1YearModel($aVIN, $aModelList[$j])
	  Next

   Next
   GUICtrlSetData ($Commu_Ctrl, 'Completed collecting VINs')
   MsgBox ($MB_TOPMOST, "Message from VIN Collector", 'Completed collecting VINs' & @CRLF & 'Please check!')


   $bGetVINs_Flag = False
EndFunc


Func DecodeVIN_1YearModel($aVIN, $sRealModel)
   ;----------------------------
   ;Decode vin here
   Local $aVIN8[0]
   Local $aVIN_RemoveDup[0]

   Local $sTxt = ''
   Local $sPreviousDecoded = ''
   For $l = 0 To UBound ($aVIN) - 1
	  If _ArraySearch ($aVIN8, StringLeft($aVIN[$l], 8)) = -1 Then
		 ;Save VIN8 to compare
		 ReDim $aVIN8[UBound($aVIN8)+1]
		 $aVIN8[UBound($aVIN8)-1] = StringLeft($aVIN[$l], 8)


		 GUICtrlSetData ($Commu_Ctrl, 'Decoding ' & $aVIN[$l] & @CRLF &'Previous decoded: ' & $sPreviousDecoded)

		 ;Get VIN Decode Infor
		 Local $sLinkVINDecode = 'https://www.decodethis.com/vin/' & $aVIN[$l]

		 Local $html = GetHtmlSourceUsingHttpRequest ($sLinkVINDecode)

		 Local $sYear = GetItemStringByMark ($html, '"name">Model Year', '</span>')
		 $sYear = StringRegExpReplace ($sYear, '(<)(.+)(>)', '')
		 $sYear = StringRegExpReplace ($sYear, '[\r\n]', '')
		 $sYear = StringReplace ($sYear, ' ', '')


		 Local $sMake = GetItemStringByMark ($html, '"name">Make', '</span>')
		 $sMake = StringRegExpReplace ($sMake, '(<)(.+)(>)', '')
		 $sMake = StringRegExpReplace ($sMake, '[\r\n]', '')
		 $sMake = StringRegExpReplace ($sMake, '\s\s+', ' ')
		 If StringLeft($sMake, 1) = ' ' Then $sMake = StringRight($sMake, StringLen($sMake) - 1)
		 If StringRight($sMake, 1) = ' ' Then $sMake = StringLeft($sMake, StringLen($sMake) - 1)


		 Local $sModel = GetItemStringByMark ($html, '"name">Series', '</span>')
		 $sModel = StringRegExpReplace ($sModel, '(<)(.+)(>)', '')
		 $sModel = StringRegExpReplace ($sModel, '[\r\n]', '')
		 $sModel = StringRegExpReplace ($sModel, '\s\s+', ' ')
		 If StringLeft($sModel, 1) = ' ' Then $sModel = StringRight($sModel, StringLen($sModel) - 1)
		 If StringRight($sModel, 1) = ' ' Then $sModel = StringLeft($sModel, StringLen($sModel) - 1)
		 If $sModel <> $sRealModel Then $sModel = $sModel & ' '


		 Local $sTrim = GetItemStringByMark ($html, '"name">Trim Level', '</span>')
		 $sTrim = StringRegExpReplace ($sTrim, '(<)(.+)(>)', '')
		 $sTrim = StringRegExpReplace ($sTrim, '[\r\n]', '')
		 $sTrim = StringRegExpReplace ($sTrim, '\s\s+', ' ')
		 If StringLeft($sTrim, 1) = ' ' Then $sTrim = StringRight($sTrim, StringLen($sTrim) - 1)
		 If StringRight($sTrim, 1) = ' ' Then $sTrim = StringLeft($sTrim, StringLen($sTrim) - 1)


		 Local $sEngine = GetItemStringByMark ($html, '"name">Engine Type', '</span>')
		 $sEngine = StringRegExpReplace ($sEngine, '(<)(.+)(>)', '')
		 $sEngine = StringRegExpReplace ($sEngine, '[\r\n]', '')
		 $sEngine = StringRegExpReplace ($sEngine, '\s\s+', ' ')
		 If StringLeft($sEngine, 1) = ' ' Then $sEngine = StringRight($sEngine, StringLen($sEngine) - 1)
		 If StringRight($sEngine, 1) = ' ' Then $sEngine = StringLeft($sEngine, StringLen($sEngine) - 1)


		 Local $sTrans = GetItemStringByMark ($html, '"name">Transmission', '</span>')
		 $sTrans = StringRegExpReplace ($sTrans, '(<)(.+)(>)', '')
		 $sTrans = StringRegExpReplace ($sTrans, '[\r\n]', '')
		 $sTrans = StringRegExpReplace ($sTrans, '\s\s+', ' ')
		 If StringLeft($sTrans, 1) = ' ' Then $sTrans = StringRight($sTrans, StringLen($sTrans) - 1)
		 If StringRight($sTrans, 1) = ' ' Then $sTrans = StringLeft($sTrans, StringLen($sTrans) - 1)


		 $sTxt &= $sYear & @TAB & $sMake & @TAB & $sModel & @TAB & $sTrim & @TAB & ' ' & $sEngine & @TAB & $sTrans & @TAB & StringLeft($aVIN[$l],8)  & @TAB & $aVIN[$l] & @CRLF
		 $sPreviousDecoded = $sYear & '  ' & $sMake & '  ' & $sModel & '  ' & $sTrim & '  ' & $sEngine & '  ' & $sTrans
	  EndIf
   Next

		 ;----------------------------
		 ;WRITE YMME CONFIG FILE
		 Local $sFileName = $sDataFileName
		 WriteTxtFile ($sFileName, $sTxt, "append")
EndFunc






;====================================================================================================================
;                  FUNCTION DISCRIPTION: GET INFOR OF YMME OF DECODE THIS
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func GetInfo()
   GUICtrlSetData ($Commu_Ctrl, 'Please wait ...')
   Local $aYMME[0]
   ;-------------------------------
   ;Link containing years
   Local $GetYearLink = 'https://www.decodethis.com/archives'
   ;Get years infor
   GUICtrlSetData ($Commu_Ctrl, 'Getting  ' & 'YEARS' & '...')
   Local $aYear = GetInfoContent ($GetYearLink)

   ;-------------------------------
   ;LOOP YEAR
   For $i = 0 To UBound($aYear)-1
	  ;Only get data from 1996
	  If $aYear[$i] = '2017' Then ExitLoop
	  ;Form link containing makes of a year
	  Local $sGetMakeLink = $GetYearLink & '/'& $aYear[$i]
	  ;Get makes infor
	  Local $aMake[0]
	  Local $iLastPageNum = GetLastPageNum ($sGetMakeLink)
	  For $z = 1 To $iLastPageNum
		 GUICtrlSetData ($Commu_Ctrl, 'Getting ' & $aYear[$i] & ' MAKES ' & '(page ' & $z & '/' & $iLastPageNum & ') ...')
		 _ArrayConcatenate ($aMake, GetInfoContent ($sGetMakeLink & '?page=' & $z))
	  Next
	  ;-------------------------------
	  ;LOOP MAKE
	  For $j = 0 To UBound($aMake)-1
		 ;Pass non-word line
		 If StringRegExpReplace ($aMake[$j], '[\r\n]', '') = '' Then ContinueLoop
		 ;Form link containing models of a year make
		 Local $sGetModelLink = $sGetMakeLink & '/' & $aMake[$j]
		 ;Get models infor
		 Local $aModel[0]
		 Local $iLastPageNum = GetLastPageNum ($sGetModelLink)
		 For $z = 1 To $iLastPageNum
			GUICtrlSetData ($Commu_Ctrl, 'Getting ' & $aYear[$i] & ' ' & $aMake[$j] & ' MODELS ' & '(page ' & $z & '/' & $iLastPageNum & ') ...')
			_ArrayConcatenate ($aModel, GetInfoContent ($sGetModelLink & '?page=' & $z))
		 Next

		 ;-------------------------------
		 ;SAVE YMME TO AN ARRAY
		 For $k = 0 To UBound($aModel)-1
			;Pass non-word line
			If StringRegExpReplace ($aModel[$k], '[\r\n]', '') = '' Then ContinueLoop
			ReDim $aYMME[UBound($aYMME)+1]
			$aYMME[UBound($aYMME)-1] = '<year>'&$aYear[$i] & '<year> --- <make>' & $aMake[$j]  & '<make> --- <model>' & $aModel[$k] & '<model>'
		 Next
	  Next
   Next
   ;----------------------------
   ;WRITE YMME CONFIG FILE
   GUICtrlSetData ($Commu_Ctrl, 'Writing YMME config file ...')
   Local $sFileName = 'YMME_Config'
   Local $sTxt = _ArrayToString ($aYMME, @CRLF)

   WriteTxtFile ($sFileName, $sTxt, "overwrite")
   GUICtrlSetData ($Commu_Ctrl, 'Completed writing YMME config file')
   LoadYearBegin ()
   MsgBox ($MB_TOPMOST, "Message from VIN Collector", 'Completed writing YMME config file' & @CRLF & 'Please check!')
   $bGetInfo_Flag = False
EndFunc


