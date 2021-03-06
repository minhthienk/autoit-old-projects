#cs ----------------------------------------------------------------------------


#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include <IE.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <Clipboard.au3>
#include <Array.au3>
#include <Excel.au3>
#include <Timers.au3>
#include <MsgBoxConstants.au3>
;~ #include "COMErrorHandler.au3"


HotKeySet ('{Pause}', AutoIT_Exit)


Global $bWeb_Attach = 0
Global $bWeb_Visible = 1
Global $bWeb_Wait = 1
Global $bWeb_TakeFocus = 0
Global $bImage_Download = 0

Global $sLink_SelectYMME = 'http://repair.alldata.com/alldata/vehicle/selection.action'
Global $sLink_YMME = ''
;~ Global $sLink_SelectYMME = 'http://repair.alldata.com/alldata/article/display.action?componentId= 800&iTypeId=376&nonStandardId=1934336&vehicleId=55746&windowName=mainADOnlineWindow'

Global $iSubscription_Num = 2



Global $sFilePath = @ScriptDir
Global $vWorksheet = 'Data'
Global $sMake = 'Chrysler Truck'
Global $sFileName = $sMake & '.xlsx'



Main ()



;====================================================================================================================
;                  FUNCTION DISCRIPTION: MAIN FUNCTION
;                  RETURN			   :
;====================================================================================================================
Func Main ()
   ConsoleWrite ('Main' & @CRLF)
   Close_All_IE()

   ;Create file to save result ---------------------------------------------
   ;Create application object
   Local $oExcel = _Excel_Open()
   ;Open an existing workbook and return its object identifier.
   Local $sWorkbook = $sFilePath & '\' & $sFileName
   Local $oWorkbook = _Excel_BookAttach ($sWorkbook)
   If $oWorkbook = 0 Then Local $oWorkbook = _Excel_BookOpen($oExcel, $sWorkbook)
   ;Row count var
   Local $iRow = 2
   ;Get the closest empty row
   While _Excel_RangeRead ($oWorkbook, $vWorksheet, 'A' & $iRow, 1, True) <> ''
	  $iRow += 1
   WEnd

   ;---------------------------------------------------------------------------
   ;Gán trang web cho biến object
   Local $oIE = IECreate_Check_Error($sLink_SelectYMME, $bWeb_Attach, $bWeb_Visible, $bWeb_Wait, $bWeb_TakeFocus)
   ;--------------------
   ;ĐOẠN CODE SỬ DỤNG FUNCTION "Check_Login_Alldata" ĐỂ KIỂM TRA ĐĂNG NHẬP
   If Check_Login_Alldata ($oIE) = "Not yet loged in before, this function has helped log in" Then
	  ;Reload trang DTC
	  IENavigate_Check_Error ($oIE, $sLink_SelectYMME)
   EndIf
   ;--------------------
   ;ĐOẠN CODE SỬ DỤNG FUNCTION "Check_Subscription_Alldata" ĐỂ KIỂM TRA SUBSCIPTION
   Check_Subscription_Alldata ($oIE, $sLink_SelectYMME)
   ;----------------------------------------
   ;SELECT VEHICLE
   Local $bNextVehicleFlag = False
   Local $bFirstVehicleFlag = False
   Local $sPreviousVehicle = _Excel_RangeRead ($oWorkbook, $vWorksheet, 'A' & $iRow - 1, 1, True) & ' - ' & _
							 _Excel_RangeRead ($oWorkbook, $vWorksheet, 'B' & $iRow - 1, 1, True) & ' - ' & _
							 _Excel_RangeRead ($oWorkbook, $vWorksheet, 'C' & $iRow - 1, 1, True) & ' - ' & _
						     _Excel_RangeRead ($oWorkbook, $vWorksheet, 'D' & $iRow - 1, 1, True)
   For $iYearCount = 2000 To 1996 Step -1
	  $aSelectData = YearVehicleSelectData($oIE, $iYearCount, $sMake)
	  For $iYMMECount = 0 To UBound($aSelectData, $UBOUND_ROWS) - 1
		 Local $sCurrentVehicle = $aSelectData[$iYMMECount][1] & ' - ' & _
								  $aSelectData[$iYMMECount][2] & ' - ' & _
								  $aSelectData[$iYMMECount][3] & ' - ' & _
								  $aSelectData[$iYMMECount][4]

		 If _Excel_RangeRead ($oWorkbook, $vWorksheet, 'A2', 1, True) = '' Then $bFirstVehicleFlag = True


		 ;Check to only process for the next vehicle
		 If $bNextVehicleFlag = False And $bFirstVehicleFlag = False Then
			If $sCurrentVehicle = $sPreviousVehicle Then
			   $bNextVehicleFlag = True
			   ContinueLoop
			Else
			   ContinueLoop
			EndIf
		 EndIf

;~ 		 ;Check to only process for a specific vehicle|Option
;~ 		 If $sCurrentVehicle <> '' Then
;~ 			ContinueLoop
;~ 		 EndIf



		 ;Show the Selected YMME
		 ;Select YMME, the last parameter is Last or Not Last
		 SelectVehicle ($oIE, 'year', $aSelectData[$iYMMECount][5], $aSelectData[$iYMMECount][9])
		 SelectVehicle ($oIE, 'make', $aSelectData[$iYMMECount][6], $aSelectData[$iYMMECount][10])
		 SelectVehicle ($oIE, 'model', $aSelectData[$iYMMECount][7], $aSelectData[$iYMMECount][11])
		 SelectVehicle ($oIE, 'engine', $aSelectData[$iYMMECount][8], $aSelectData[$iYMMECount][12])

		 $oForm = _IEGetObjById ($oIE, 'vehicleSelectionForm')
		 _IEFormSubmit($oForm)


		 $sLink_YMME = _IEPropertyGet($oIE, 'locationurl')
		 ;--------------------
		 ;Get Procedure link
		 Local $aProcedureLink = GetSearchLink ($oIE, 'check')
		 If $aProcedureLink[0] = 'No links' Then
			IENavigate_Check_Error ($oIE, $sLink_SelectYMME)
			Check_Subscription_Alldata ($oIE, $sLink_SelectYMME)
			ContinueLoop
		 EndIf

		 ;Process link by link
		 For $sLink In $aProcedureLink
			IENavigate_Check_Error ($oIE, $sLink)
			Check_Subscription_Alldata ($oIE, $sLink) ;//////////
			;Get data
			Local $aData = RetrieveData($oIE)
			;Write Procedure Name
			_Excel_RangeWrite ($oWorkbook, $vWorksheet, $aSelectData[$iYMMECount][1], 'A' & $iRow, True, True)
			;Write Procedure Name
			_Excel_RangeWrite ($oWorkbook, $vWorksheet, $aSelectData[$iYMMECount][2], 'B' & $iRow, True, True)
			;Write Procedure Name
			_Excel_RangeWrite ($oWorkbook, $vWorksheet, $aSelectData[$iYMMECount][3], 'C' & $iRow, True, True)
			;Write Procedure Name
			_Excel_RangeWrite ($oWorkbook, $vWorksheet, $aSelectData[$iYMMECount][4], 'D' & $iRow, True, True)
			;Write Procedure Name
			_Excel_RangeWrite ($oWorkbook, $vWorksheet, $aData[0], 'E' & $iRow, True, True)
			;Write Procedure
			_Excel_RangeWrite ($oWorkbook, $vWorksheet, $aData[1], 'F' & $iRow, True, True)
			;Increase Row number
			$iRow += 1
		 Next
		 ;Back To Select Vehicle Page
		 IENavigate_Check_Error ($oIE, $sLink_SelectYMME)
		 Check_Subscription_Alldata ($oIE, $sLink_SelectYMME)
	  Next
   Next

   Return $oIE
EndFunc





;====================================================================================================================
;                  FUNCTION DISCRIPTION: MAIN FUNCTION
;                  RETURN			   :
;====================================================================================================================
Func YearVehicleSelectData(Byref $oIE, $sYear, $sMake)
   ConsoleWrite ('YearVehicleSelectData' & @CRLF)
   Local $aSelectData [0][13]
   Local $bLastYear = False
   Local $bLastMake = False
   Local $bLastModel = False
   Local $bLastEngine = False
   SelectVehicle ($oIE, 'year', $sYear)
   ;--------------------------------------------------------------------
   ;Get make list
   Local $aMakeList = GetSelectData ($oIE, 'make')
   ;Process Make by Make
   For $iMakeCount = 0 To UBound($aMakeList, $UBOUND_ROWS) - 1
	  ;Only filter some necessary Makes
	  If $aMakeList[$iMakeCount][1] <> $sMake Then ContinueLoop
	  ;Check if processing the last item in the list
	  If $iMakeCount <> UBound($aMakeList, $UBOUND_ROWS) - 1 Then
		 SelectVehicle ($oIE, 'make', $aMakeList[$iMakeCount][0])
	  Else
		 SelectVehicle ($oIE, 'make', $aMakeList[$iMakeCount][0], 'Last')
		 Local $bLastMake = True
	  EndIf
	  ;--------------------------------------------------------------------
	  ;Get model list
	  Local $aModelList = GetSelectData ($oIE, 'model')
	  ;Process Model by Model
	  For $iModelCount = 0 To UBound($aModelList, $UBOUND_ROWS) - 1
		 ;Check if processing the last item in the list
		 If $iModelCount <> UBound($aModelList, $UBOUND_ROWS) - 1 Then
			SelectVehicle ($oIE, 'model', $aModelList[$iModelCount][0])
		 Else
			SelectVehicle ($oIE, 'model', $aModelList[$iModelCount][0], 'Last')
			$bLastModel = True
		 EndIf
		 ;--------------------------------------------------------------------
		 ;Get engine list
		 Local $aEngineList = GetSelectData ($oIE, 'engine')

		 ;Process Model by Model
		 For $iEngineCount = 0 To UBound($aEngineList, $UBOUND_ROWS) - 1
			;Check if processing the last item in the list
			If $iEngineCount = UBound($aEngineList, $UBOUND_ROWS) - 1 Then $bLastEngine = True

			ReDim $aSelectData[UBound($aSelectData, $UBOUND_ROWS ) + 1][13]

			;YMME
			$aSelectData[UBound($aSelectData, $UBOUND_ROWS) - 1][1] = $sYear
			$aSelectData[UBound($aSelectData, $UBOUND_ROWS) - 1][2] = $aMakeList[$iMakeCount][1]
			$aSelectData[UBound($aSelectData, $UBOUND_ROWS) - 1][3] = $aModelList[$iModelCount][1]
			$aSelectData[UBound($aSelectData, $UBOUND_ROWS) - 1][4] = $aEngineList[$iEngineCount][1]
			;Value
			$aSelectData[UBound($aSelectData, $UBOUND_ROWS) - 1][5] = $sYear
			$aSelectData[UBound($aSelectData, $UBOUND_ROWS) - 1][6] = $aMakeList[$iMakeCount][0]
			$aSelectData[UBound($aSelectData, $UBOUND_ROWS) - 1][7] = $aModelList[$iModelCount][0]
			$aSelectData[UBound($aSelectData, $UBOUND_ROWS) - 1][8] = $aEngineList[$iEngineCount][0]
			;Last item??
			If $bLastYear = True Then $aSelectData[UBound($aSelectData, $UBOUND_ROWS) - 1][9] = 'Last'
			If $bLastMake = True Then $aSelectData[UBound($aSelectData, $UBOUND_ROWS) - 1][10] = 'Last'
			If $bLastModel = True Then $aSelectData[UBound($aSelectData, $UBOUND_ROWS) - 1][11] = 'Last'
			If $bLastEngine = True Then $aSelectData[UBound($aSelectData, $UBOUND_ROWS) - 1][12] = 'Last'
		 Next
		 $bLastEngine = False
	  Next
	  $bLastModel = False
   Next
   Return $aSelectData
EndFunc
























;====================================================================================================================
;                  FUNCTION DISCRIPTION: MAIN FUNCTION
;                  RETURN			   :
;====================================================================================================================
Func GetSelectData (Byref $oIE, $sType = 'year')
   ConsoleWrite ('GetSelectData' & @CRLF)
   ;Get html source
   Local $sHtml = _IEPropertyGet($oIE, 'outerhtml')

   $sHtml = StringReplace($sHtml, '><', '>' & @CRLF & '<')


   ;Get the part of source which contains the list of option
   Local $sNecessaryHtml = GetItemStringByMark ($shtml, '<select name="' & $sType & '"', '</select>')
   ;Refine the data, remove redundant characters
   $sNecessaryHtml = StringRegExpReplace($sNecessaryHtml, '^.+', '')
   $sNecessaryHtml = StringRegExpReplace($sNecessaryHtml, '[\t\r\n]+', '')
   $sNecessaryHtml = StringReplace($sNecessaryHtml, '><', '>' & @CRLF & '<')
   ;Split string of option into an array
   Local $aSelectRawData = StringSplit ($sNecessaryHtml, @CRLF,  $STR_ENTIRESPLIT + $STR_NOCOUNT)
   ;The array to contain elements of an selection
   Local $aSelectUsableData[0][2]
   ;Browse the array, to get elements
   For $vElement In $aSelectRawData
	  ReDim $aSelectUsableData[UBound($aSelectUsableData, $UBOUND_ROWS ) + 1][2]
	  $aSelectUsableData[UBound($aSelectUsableData, $UBOUND_ROWS) - 1][0] = GetItemStringByMark ($aSelectRawData[UBound($aSelectUsableData, $UBOUND_ROWS) - 1], 'value="', '"')
	  $aSelectUsableData[UBound($aSelectUsableData, $UBOUND_ROWS) - 1][1] = GetItemStringByMark ($aSelectRawData[UBound($aSelectUsableData, $UBOUND_ROWS) - 1], '>', '<')
   Next
   Return $aSelectUsableData
EndFunc



;====================================================================================================================
;                  FUNCTION DISCRIPTION: CREATE IE OBJECT AND CHECK ERROR
;                  RETURN			   :
;====================================================================================================================
Func SelectVehicle (Byref $oIE, $sType = 'year', $sID = '', $sLast = 'Not Last')
   ConsoleWrite ('SelectVehicle' & @CRLF)
   ;Select Object
   Sleep (100)
   Local $oObject = _IEGetObjByName ($oIE, $sType)




   If $sType = 'model' Then
	  ;Get model list
	  Local $aModelList = GetSelectData ($oIE, 'model')
	  If UBound($aModelList) = 1 Then
		 ControlSend('ALLDATA Repair - Vehicle Selection - Windows Internet Explorer', '', '[CLASS:Internet Explorer_Server; INSTANCE:1]', '{Tab}')
		 ControlSend('ALLDATA Repair - Vehicle Selection - Internet Explorer', '', '[CLASS:Internet Explorer_Server; INSTANCE:1]', '{Tab}')
		 ControlSend('ALLDATA Repair - Vehicle Selection - Windows Internet Explorer', '', '[CLASS:Internet Explorer_Server; INSTANCE:1]', '{Tab}')
		 ControlSend('ALLDATA Repair - Vehicle Selection - Internet Explorer', '', '[CLASS:Internet Explorer_Server; INSTANCE:1]', '{Tab}')
		 ControlSend('ALLDATA Repair - Vehicle Selection - Windows Internet Explorer', '', '[CLASS:Internet Explorer_Server; INSTANCE:1]', '{Down}')
		 ControlSend('ALLDATA Repair - Vehicle Selection - Internet Explorer', '', '[CLASS:Internet Explorer_Server; INSTANCE:1]', '{Down}')
		 Sleep (100)
		 _IELoadWait($oIE)
		 Return
	  EndIf
   EndIf


   _IEFormElementSetValue ($oObject, $sID)
   _IEAction($oObject, 'focus')

   If $sType <> 'engine' Then
	  If $sLast <> 'Last' Then
		 ControlSend('ALLDATA Repair - Vehicle Selection - Windows Internet Explorer', '', '[CLASS:Internet Explorer_Server; INSTANCE:1]', '{Down}')
		 ControlSend('ALLDATA Repair - Vehicle Selection - Internet Explorer', '', '[CLASS:Internet Explorer_Server; INSTANCE:1]', '{Down}')
		 Sleep (100)
;~ 		 Local $hStarttime = _Timer_Init()
		 _IELoadWait($oIE)
;~ 		 ConsoleWrite('time to wait for showing new data ' & _Timer_Diff($hStarttime) & @CRLF)

		 ControlSend('ALLDATA Repair - Vehicle Selection - Windows Internet Explorer', '', '[CLASS:Internet Explorer_Server; INSTANCE:1]', '{Up}')
		 ControlSend('ALLDATA Repair - Vehicle Selection - Internet Explorer', '', '[CLASS:Internet Explorer_Server; INSTANCE:1]', '{Up}')
		 Sleep (100)
;~ 		 Local $hStarttime = _Timer_Init()
		 _IELoadWait($oIE)
;~ 		 ConsoleWrite('time to wait for showing new data ' & _Timer_Diff($hStarttime) & @CRLF)
	  Else
		 ControlSend('ALLDATA Repair - Vehicle Selection - Windows Internet Explorer', '', '[CLASS:Internet Explorer_Server; INSTANCE:1]', '{Up}')
		 ControlSend('ALLDATA Repair - Vehicle Selection - Internet Explorer', '', '[CLASS:Internet Explorer_Server; INSTANCE:1]', '{Up}')
		 Sleep (100)
;~ 		 Local $hStarttime = _Timer_Init()
		 _IELoadWait($oIE)
;~ 		 ConsoleWrite('time to wait for showing new data ' & _Timer_Diff($hStarttime) & @CRLF)

		 ControlSend('ALLDATA Repair - Vehicle Selection - Windows Internet Explorer', '', '[CLASS:Internet Explorer_Server; INSTANCE:1]', '{Down}')
		 ControlSend('ALLDATA Repair - Vehicle Selection - Internet Explorer', '', '[CLASS:Internet Explorer_Server; INSTANCE:1]', '{Down}')
		 Sleep (100)
;~ 		 Local $hStarttime = _Timer_Init()
		 _IELoadWait($oIE)
;~ 		 ConsoleWrite('time to wait for showing new data ' & _Timer_Diff($hStarttime) & @CRLF)
	  EndIf
   EndIf
EndFunc







;====================================================================================================================
;                  FUNCTION DISCRIPTION: MAIN FUNCTION
;                  RETURN			   :
;====================================================================================================================
Func RetrieveData(ByRef $oIE)
   ConsoleWrite ('RetrieveData' & @CRLF)
   Local $sProcedure = _IEPropertyGet($oIE, 'innertext')
   $sProcedure = GetItemStringByMark ($sProcedure, 'Image(s) Only', 'var classElements', 2, 1)


   ;Remove redundant spaces
   $sProcedure = StringRegExpReplace ($sProcedure, ' +', ' ')
   $sProcedure = StringRegExpReplace ($sProcedure, ' +$', ' ')
   $sProcedure = StringReplace ($sProcedure, ' ' & @CRLF, @CRLF)

   ;Remove Zoom and Print
   $sProcedure = StringReplace ($sProcedure, ' Zoom and Print Options', '')

   ;Remove redundant @CRLF
   $sProcedure = StringRegExpReplace ($sProcedure, '[\r\n]+', @CRLF)
   $sProcedure = StringRegExpReplace ($sProcedure, '[\r\n]+$', '')
   $sProcedure = StringRegExpReplace ($sProcedure, '^[\r\n]+', '')

   ;Add ' - '
   $sProcedure = StringReplace ($sProcedure, @CRLF &  ' ', @CRLF & '- ')

   ;Replace 2 ' - ' by 1
   $sProcedure = StringReplace ($sProcedure, @CRLF &  '- - ', @CRLF & '+ ')

   ;Replace '- ' + number to number
   $sProcedure = StringRegExpReplace ($sProcedure, '- (?=\d+.)', '')

   ;Get title of page
   $sTitle = _IEPropertyGet($oIE, 'title')

   Local $aData[2] = [$sTitle, $sProcedure]

   Return $aData
EndFunc





;====================================================================================================================
;                  FUNCTION DISCRIPTION:
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func GetItemStringByMark ($sString, $sStartMark, $sEndMark, $iOccurrenceStart = 1, $iOccurrenceEnd = 1)
   ConsoleWrite ('GetItemStringByMark' & @CRLF)
   If StringInStr ($sString, $sStartMark, 0, 1, 1) <> 0 Then
	  Local $iStart = StringInStr ($sString, $sStartMark, 0, $iOccurrenceStart, 1) + StringLen ($sStartMark)
	  Local $iEnd = StringInStr ($sString, $sEndMark, 0, $iOccurrenceEnd, $iStart)
	  Local $sItemString = StringMid ($sString, $iStart, $iEnd - $iStart)
   Else
	  Local $sItemString = ""
   EndIf
   Return $sItemString
EndFunc




;====================================================================================================================
;                  FUNCTION DESCRIPTION: OPEN DTC LINK FROM CONFIG FILE
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func GetSearchLink (Byref $oIE, $sSearch_String)
   ConsoleWrite ('GetSearchLink' & @CRLF)
   Sleep (1000)
   ;-------------------------------------
   ;Check link có ra page not found hay không, nếu có thì navigate vô trang YMME để có ô search
   Local $sHTML_Innertext = _IEPropertyGet ($oIE, "innertext")
   If StringInStr ($sHTML_Innertext, "Page not found") <> 0 Or StringInStr ($sHTML_Innertext, "DOCTYPE html PUBLIC") <> 0 Or StringInStr ($sHTML_Innertext, "The page you requested can not be displayed") <> 0  Then
	  IENavigate_Check_Error ($oIE, $sLink_YMME)
	  Check_Subscription_Alldata ($oIE, $sLink_YMME)
   EndIf
   ;-------------------------------------
   Do
	  ;Lấy object form search
	  Local $oForm = _IEFormGetObjByName($oIE, "simpleSearch")
	  ;Lấy object Search box
	  Local $oSearchBox = _IEFormElementGetObjByName($oForm, "searchQuery")
	  ;Set search string
	  _IEFormElementSetValue($oSearchBox, $sSearch_String)
	  ;Submit form, no wait for page load to complete
	  _IEFormSubmit($oForm, 0)
	  ;Wait for the page load to complete
	  _IELoadWait($oIE)
	  ;------------------------------------
	  ;ĐOẠN CODE SỬ DỤNG FUNCTION "Check_Subscription_Alldata" ĐỂ KIỂM TRA SUBSCIPTION
	  Check_Subscription_Alldata ($oIE, $sLink_YMME)
	  If _IEPropertyGet ($oIE, "title") = "ALLDATA Repair - Search Results" Then ExitLoop
   Until 0

   ;--------------------
   ;Get object of Service and Repair
	  Local $oServiceRepair = _IEGetObjById ($oIE, 'cat10_records');
	  If @error = 0 Then
		 Local $oLinks = _IETagNameGetCollection($oServiceRepair, 'a')
		 Local $aNecessaryLinks [0]
		 For $oLink In $oLinks
			If StringInStr ($oLink.innertext, 'Procedure') Then
			   ;------------------------------------
			   Local $sTemp = _IEPropertyGet ($oLink,"outerhtml")
   ;~ 			MsgBox (0, '', $sTemp)
			   ;------------------------------------
			   ;ĐOẠN CODE LẤY ID TẠO LINK
			   Local $aIDs [4]
			   For $i = 1 To 4 Step 1
				  $aIDs [$i-1] =  StringMid ($sTemp, Stringinstr ($sTemp, ",", 0, $i) + 1, Stringinstr ($sTemp, ",", 0, $i + 1) - Stringinstr ($sTemp, ",", 0, $i) - 1)
			   Next

			   $sLink = "http://repair.alldata.com/alldata/article/display.action?componentId=" & $aIDs [0] & "&iTypeId=" & $aIDs [1] & "&nonStandardId=" & $aIDs [2] & "&vehicleId=" & $aIDs [3] & "&windowName=mainADOnlineWindow"

			   ReDim $aNecessaryLinks[UBound($aNecessaryLinks)+1]
			   $aNecessaryLinks[UBound($aNecessaryLinks)-1] = $sLink
			EndIf
		 Next
		 ;-----------------------------------------
		 ;Nếu không có cái link nào tron đây
		 If UBound($aNecessaryLinks) = 0 Then
			ReDim $aNecessaryLinks[1]
			$aNecessaryLinks[0] = 'No links'
		 EndIf
	  Else
			Local $aNecessaryLinks [1]
			$aNecessaryLinks [0] = 'No links'
	  EndIf
   Return $aNecessaryLinks
EndFunc





;====================================================================================================================
;                  FUNCTION DISCRIPTION: Select Vehicle
;                  RETURN			   :
;====================================================================================================================
Func IECreate_Check_Error ($sLink, $bAttach, $bVisible, $bWait, $bTakeFocus)
   ConsoleWrite ('IECreate_Check_Error' & @CRLF)
   Do
	  ;MsgBox (0, "", "Open: " & $sLink)
	  __IELockSetForegroundWindow($LSFW_LOCK)
;~ 	  $WinHandle = WinGetHandle ('')
	  Local $oIE = _IECreate($sLink, $bAttach, $bVisible, $bWait, $bTakeFocus)
;~ 	  WinActivate ($WinHandle)
	  If @error <> 0 Then Sleep(1000)
	  ;MsgBox (0, "", @error)
   Until @error = 0
   Sleep (2000)
   Return $oIE
EndFunc


;====================================================================================================================
;                  FUNCTION DISCRIPTION: NAVIGATE IE OBJECT AND CHECK ERROR
;                  RETURN			   : A STRING OF PROCEDURE NAME
;====================================================================================================================
Func IENavigate_Check_Error (ByRef $oIE, $sLink)
   ConsoleWrite ('IENavigate_Check_Error' & @CRLF)
   Local $icount = 0
   Local $bFlag = True
   Do
	  ;~ $icount = $icount + 1
	  __IELockSetForegroundWindow($LSFW_LOCK)
;~ 	  $WinHandle = WinGetHandle ('')
	  _IENavigate ($oIE, $sLink, 1)
;~ 	  WinActivate ($WinHandle)
	  Sleep (500)

	  If @error = 0 Then
		 $bFlag = True
	  Else
		 $bFlag = False
	  EndIf
   ;ConsoleWrite ("Navigate count: " & $icount & " --- Error code: " & @error & @CRLF)
   Until $bFlag = True
   Sleep(1000)
EndFunc



;====================================================================================================================
;                  FUNCTION DISCRIPTION: CHECK WHETHER USER HAS LOGGED IN ALLDATA
;				   INPUT               : $oIE
;                  OUTPUT              : A STRING OF RESULT
;====================================================================================================================
Func Check_Login_Alldata (Byref $oIE)
   ConsoleWrite ('Check_Login_Alldata' & @CRLF)
   Local $sResult = "Already Loged in before"
   ;ĐOẠN CODE ĐỂ KIỂM TRA ĐĂNG NHẬP ALLDATA, NẾU CHƯA ĐĂNG NHẬP THÌ ĐĂNG NHẬP
   ;Collect tất cả text trong tag <body> và lưu vào biến $sTxt_Body
   Local $sLink_Login = "https://repair.alldata.com/alldata/secure/login.action"
   Local $oBodys = _IETagNameGetCollection($oIE, "body")
   Local $sTxt_Body = ""
   For $oBody In $oBodys
	   $sTxt_Body &= $oBody.innertext & @CRLF
   Next
   ;Kiểm tra nếu trong $sTxt_Body có  "HTTP Status 404 - No result" hoặc "Please Log In"
   If StringInStr ($sTxt_Body, "HTTP Status 404 - No result", 0, 1) <> 0 Or StringInStr ($sTxt_Body, "Please Log In", 0, 1) <> 0 Then
	  ;Mở trang login
	  IENavigate_Check_Error ($oIE, $sLink_Login)
	  ;Lấy object form login
	  Local $oForm = _IEFormGetObjByName($oIE, "customer_login_center")
	  ;Lấy object LoginName
	  Local $oLoginName = _IEFormElementGetObjByName($oForm, "j_username")
	  ;Set LoginName
	  _IEFormElementSetValue($oLoginName, "innovard")
	  ;Lấy object Password
	  Local $oPassword = _IEFormElementGetObjByName($oForm, "j_password")
	  ;Set password
	  _IEFormElementSetValue($oPassword, "Inn0v@VN123")
	  ;Submit form, no wait for page load to complete
	  _IEFormSubmit($oForm, 0)
	  ;Wait for the page load to complete
	  _IELoadWait($oIE)
	  $sResult = "Not yet loged in before, this function has helped log in"
   EndIf
   Return $sResult
EndFunc





;====================================================================================================================
;                  FUNCTION DISCRIPTION: CHECK WHETHER USER HAS CLICKED SUBSRIPTION
;				   INPUT               : $oIE, $sLink, $iSubscription_Num
;                  OUTPUT              : AN IE OJECT OF THE DTC OR PROCEDURE LINK
;====================================================================================================================
Func Check_Subscription_Alldata (Byref $oIE, $sLink)
   ConsoleWrite ('Check_Subscription_Alldata' & @CRLF)
   Sleep (1000)
   Local $sResult = ""
   Do
	  ;ĐOẠN CODE ĐỂ KIỂM TRA SUBSCRIPTION ALLDATA, NẾU CHƯA CÓ THÌ CLICK SUBSCRIPTION
	  ;Collect tất cả text trong tag <body> và lưu vào biến $sTxt_Body
	  Local $sTxt_Title = _IEPropertyGet ($oIE, "title")
	  ;Collect tất cả text trong tag <span> và lưu vào biến $sTxt_Span
	  Local $oSpans = _IETagNameGetCollection($oIE, "body")
	  Local $sTxt_Span = ""
	  For $oSpan In $oSpans
		 $sTxt_Span &= $oSpan.innertext & @CRLF
	  Next
	  ;Kiểm tra nếu trong $sTxt_Title có  "Subscription License Management"
	  If $sTxt_Title = "ALLDATA Repair - Subscription License Management" Then
		 ;Kiểm tra nếu trong $sTxt_Span có "Your License"
		 If StringInStr ($sTxt_Span, "Your License", 0, 1) = 0 Then
			;Lấy Object của nut OK
			Local $oOK_Button = _IEGetObjById ($oIE, "ok_button")
			;Nhấn nút OK
			_IEAction ($oOK_Button, "click" )
			;Chờ cho action done
			Sleep (1000)
			;Collect tất cả text trong tag <Img> để tìm nút nhấn đỏ, xanh. Sau đó lưu object vào mảng $aImg_Object
			Local $oImgs = _IEImgGetCollection($oIE)
			Local $iCount = 0
			Local $aImg_Object [500]
			For $oImg In $oImgs
				$aImg_Object [$iCount] = $oImg
				$iCount = $iCount + 1
			Next

			;Click vào nút đỏ release subscription
			_IEAction ($aImg_Object [$iSubscription_Num*2], "click")
			Sleep (1000)
			;Click vào nút xanh enter subscription
			_IEAction ($aImg_Object [$iSubscription_Num*2 - 1], "click")
			Sleep (1000)
			;Wait for the page load to complete
			_IELoadWait($oIE)
			$sResult = "Not yet had subsription, this function has helped clicked subscription"
		 Else
			$sResult = "Already had subscription  - But need reload"
		 EndIf
	  Else
		 $sResult = "Already had subscription  - No need reload"
	  EndIf
	  ;--------------------
	  If $sResult = "Already had subscription  - But need reload" Then
		 ;Reload trang DTC
		 IENavigate_Check_Error ($oIE, $sLink)
	  EndIf
   Until $sResult = "Already had subscription  - No need reload"
EndFunc


Func AutoIT_Exit()
   Exit
EndFunc



Func Close_All_IE()
   $Proc = "iexplore.exe"
   While ProcessExists($Proc)
      ProcessClose($Proc)
   Wend
EndFunc




