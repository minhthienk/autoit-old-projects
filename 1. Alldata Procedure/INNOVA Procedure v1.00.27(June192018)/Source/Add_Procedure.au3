#include-once

#include <MsgBoxConstants.au3>
#include <FileConstants.au3>
#include <InetConstants.au3>
#include <Clipboard.au3>
#include <IE.au3>
#include <Excel.au3>
#include <WinAPIFiles.au3>
#include <Array.au3>

#include "General_Library.au3"
#include "Create_JAVASCRIPT_Procedure.au3"
#include "Create_NORMAL_Procedure.au3"
#include "GUI.au3"




;====================================================================================================================
;                  FUNCTION DISCRIPTION: FIND PROCEDURE
;====================================================================================================================

Func Find_Procedure ()
   Local $sDTC_Path = GUICtrlRead ($Input_DTC_Path)
   Local $hFileOpen = FileOpen($sDTC_Path, $FO_READ)
   Local $sDTC_HTML = FileRead ($hFileOpen)
   FileClose($hFileOpen)
   Local $aProcedure [0]
   Local $aTemp = StringSplit ($sDTC_HTML, @CRLF, $STR_ENTIRESPLIT)
   For $i = 1 to $aTemp[0]
	  If StringInStr ($aTemp[$i], "PROCEDURE_") <> 0 Then
		 ReDim $aProcedure [UBound ($aProcedure) + 1]
		 Local $sTemp = ""
		 $sTemp = $aTemp[$i]
		 $sTemp = StringMid ($sTemp, StringInStr ($sTemp, "PROCEDURE_", 0, -1), StringInStr ($sTemp, ".html", 0, -1) - StringInStr ($sTemp, "PROCEDURE_", 0, -1))
		 $aProcedure [UBound ($aProcedure) - 1] = $sTemp
	  EndIf
   Next

   ;Add procedures to combobox
   GUICtrlSetData($Combo_Procedure, "")
   GUICtrlSetData($Combo_Procedure, "(Select Procedure)|" & _ArrayToString($aProcedure, "|"))
   _GUICtrlComboBox_SetCurSel($Combo_Procedure, 0)

   MsgBox (0, "Message", "Found " & UBound ($aProcedure) & " PROCEDURE(S)" & @CRLF & @CRLF & "Please select one!" & @CRLF & "then input PROCEDURE LINK you want to replace.")
   $bAdd_Allow_Flag = True
   $bFind_Flag = False
EndFunc





;====================================================================================================================
;                  FUNCTION DISCRIPTION: ADD PROCEDURE
;====================================================================================================================
;http://repair.alldata.com/alldata/navigation/treedisplay.action?vehicleId=47132&componentId=621&iTypeId=383&nonStandardId=5221757&fromJs=true&openUrl=
Func Add_Procedure ()
   ;------------------------------------------------------
   ;LẤY CÁC GIÁ TRỊ ĐẦU VÀO
   Local $sDTC_Path = GUICtrlRead ($Input_DTC_Path)
   Local $sSelected_Procedure = GUICtrlRead ($Combo_Procedure)
   Local $sLink = GUICtrlRead ($Input_Procedure_Link)
   ;------------------------------------------------------
   ;LẤY YMME PATH ĐỂ SỬ DỤNG CHO VIỆC TẠO PROCEDURE
   Local $sFilePath_YMME = StringLeft ($sDTC_Path, StringInStr ($sDTC_Path, "\", 0, -3) - 1)
   Local $sYMME = StringMid ($sDTC_Path, StringInStr ($sDTC_Path, "\", 0, -4) + 1, StringInStr ($sDTC_Path, "\", 0, -3) - StringInStr ($sDTC_Path, "\", 0, -4) -1)
   ;------------------------------------------------------
   ;MỞ TRANG PROCEDURE
   Local $oIE = IECreate_Check_Error($sLink, $bWeb_Attach, $bWeb_Visible, $bWeb_Wait, $bWeb_TakeFocus)
   ;------------------------------------------------------
   ;CHECK SUBSCRIPTION
   Check_Subscription_Alldata ($oIE, $sLink)
   ;------------------------------------------------------
   ;BIẾN TITLE ĐỂ CHECK DẠNG PROCEDURE: PROCEDURE DẠNG ĐƠN HAY DẠNG PARTS
   Local $sTitle = _IEPropertyGet ($oIE, "title")
   ;------------------------------------------------------
   ;CHECK DẠNG PROCEDURE
   If StringInStr ($sTitle, "ALLDATA Repair - Vehicle Information") = 0 Then
	  ;------------------------------------------------------
	  ;CHECK NẾU PROCEDURE ĐÃ ĐƯỢC TẠO CHƯA NẾU
	  If Check_Log_File ($sYMME, "Log File Procedure Successful.txt", $sLink) = "Not Exist" Then
		 Local $sProcedure_Path = Create_NORMAL_Procedure ($oIE, $sFilePath_YMME, $sLink)
	  Else
		 Local $sTxt_File_Name = Create_Procedure_Name ($sLink)
		 Local $sProcedure_Path = "../../PROCEDURE/" & $sTxt_File_Name & "/" & $sTxt_File_Name & ".html"
	  EndIf
   Else
	  $sLink = Create_JAVA_link ($sLink)
	  Local $sProcedure_Path = Create_JAVASCRIPT_Procedure ($oIE, $sFilePath_YMME, $sLink, "", $sYMME)
   EndIf


   ;------------------------------------------------------
   ;THAY THẾ FILE HTML CŨ BẰNG MỚI
   ;Mở file lấy DTC html source
   Local $hFileOpen = FileOpen($sDTC_Path, $FO_READ)
   Local $sDTC_HTML = FileRead ($hFileOpen)
   FileClose($hFileOpen)
   ;Thay thế link procedure cũ bằng mới
   Local $sOld_Path = "../../PROCEDURE/" & $sSelected_Procedure & "/" & $sSelected_Procedure & ".html"
   $sDTC_HTML = StringReplace ($sDTC_HTML, $sOld_Path, $sProcedure_Path)
   ;Tạo lại file DTC html
   Local $sFilePath_Title = StringLeft ($sDTC_Path, StringInStr ($sDTC_Path, "\", 0, -1) - 1)
   Local $sTxt_Title = StringMid ($sDTC_Path, StringInStr ($sDTC_Path, "\", 0, -1) + 1, StringInStr ($sDTC_Path, ".html", 0, -1) - StringInStr ($sDTC_Path, "\", 0, -1)-1)
   Create_HTML  ($sFilePath_Title, $sTxt_Title, $sTxt_Title, $sDTC_HTML)
   Notification ("Done! Please Check!", "Normal")
   Return $oIE
EndFunc







;====================================================================================================================
;                  FUNCTION DISCRIPTION: CREATE JAVA LINK
;                  RETURN			   : A
;====================================================================================================================
Func Create_JAVA_link ($sProcedure_Link)
   ;ĐOẠN CODE LẤY CÁC ID TRONG URL
   ;Gắn thêm dấu và vào cuối string để đánh dấu
   $sProcedure_Link = $sProcedure_Link & "&"
   ;Lấy $sComponentID
   Local $iComponentID_Pos = StringInStr ($sProcedure_Link, "componentId=", 0, 1) + StringLen ("componentId=")
   Local $iID_End_Pos = StringInStr ($sProcedure_Link, "&", 0, 1,  $iComponentID_Pos)
   Local $sComponentID = StringMid ($sProcedure_Link, $iComponentID_Pos, $iID_End_Pos - $iComponentID_Pos)
   ;Lấy $sITypeId
   Local $iITypeId_Pos = StringInStr ($sProcedure_Link, "iTypeId=", 0, 1) + StringLen ("iTypeId=")
   Local $iID_End_Pos = StringInStr ($sProcedure_Link, "&", 0, 1,  $iITypeId_Pos)
   Local $sITypeId = StringMid ($sProcedure_Link, $iITypeId_Pos, $iID_End_Pos - $iITypeId_Pos)
   ;Lấy $sNonStandardId
   Local $iNonStandardId_Pos = StringInStr ($sProcedure_Link, "nonStandardId=", 0, 1) + StringLen ("nonStandardId=")
   Local $iID_End_Pos = StringInStr ($sProcedure_Link, "&", 0, 1,  $iNonStandardId_Pos)
   Local $sNonStandardId = StringMid ($sProcedure_Link, $iNonStandardId_Pos, $iID_End_Pos - $iNonStandardId_Pos)
   ;Lấy $sVehicleId
   Local $iVehicleId_Pos = StringInStr ($sProcedure_Link, "vehicleId=", 0, 1) + StringLen ("vehicleId=")
   Local $iID_End_Pos = StringInStr ($sProcedure_Link, "&", 0, 1,  $iVehicleId_Pos)
   Local $sVehicleId = StringMid ($sProcedure_Link, $iVehicleId_Pos,  $iID_End_Pos - $iVehicleId_Pos)
   ;Tạo PROCEDURE name
   Local $sTemp = ""
   $sTemp = "javascript:navigateOnTree('/alldata/','" & $sComponentID & "', '" & $sITypeId &"', '" & $sNonStandardId & "', '" & $sVehicleId & "');"
   Return $sTemp
EndFunc







