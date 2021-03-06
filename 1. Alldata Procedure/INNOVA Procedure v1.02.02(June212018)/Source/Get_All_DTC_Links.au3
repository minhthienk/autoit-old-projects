
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




Func Scan_DTCs ()
   ;Gán trang web cho biến object
   Local $oIE = IECreate_Check_Error($sLink_YMME, $bWeb_Attach, $bWeb_Visible, $bWeb_Wait, $bWeb_TakeFocus)
   Sleep (1000)
   ;------------------------------------
   ;ĐOẠN CODE SỬ DỤNG FUNCTION "Check_Login_Alldata" ĐỂ KIỂM TRA ĐĂNG NHẬP
   If Check_Login_Alldata ($oIE) = "Not yet loged in before, this function has helped log in" Then
	  ;Reload trang DTC
	  IENavigate_Check_Error ($oIE, $sLink_YMME)
   EndIf
   ;------------------------------------
   ;ĐOẠN CODE SỬ DỤNG FUNCTION "Check_Subscription_Alldata" ĐỂ KIỂM TRA SUBSCIPTION
   Check_Subscription_Alldata ($oIE, $sLink_YMME)
   ;------------------------------------
   ;ĐOẠN CODE LẤY TEXT TRONG TAG <TITLE> ĐỂ KIỂM TRA XEM LINK ĐÓ CÓ PHẢI LINK YMME KHÔNG
   Local $sTxt_Title = _IEPropertyGet ($oIE, "title")
   $sTxt_Title = Standardize_String ($sTxt_Title)
   If StringInStr ($sTxt_Title, "ALLDATA Repair - Vehicle Information - ") <> 0 Then
	  ;------------------------------------
	  ;ÐOẠN CODE LẤY YMME ĐẶT TÊN CHO FOLDER
	  $sYMME = _IEPropertyGet ($oIE, "title")
	  $sYMME = StringRight ($sYMME, StringLen ($sYMME) - StringInStr ($sYMME, "-", 0, 2) - 1)
	  $sYMME = Standardize_File_Name ($sYMME)
	  Local $sFilePath_YMME        = @ScriptDir & "\INNOVA"      &"\" & $sYMME
	  ;------------------------------------
	  ;OPEN CONFIG FILE
	  Notification ("Opening config file for: " & @CRLF & $sYMME,"Normal")
	  Sleep (1000)
	  Local $hFileOpen = FileOpen($sFilePath_YMME & "\" & "Scan DTC Config" & ".txt", $FO_READ)
	  Local $sConfig = FileRead ($hFileOpen)
	  FileClose($hFileOpen)
	  ;------------------------------------
	  ;CHECK IF THE CONFIG FILE EXISTS
	  If $sConfig = "" Then
		 Notification ("Found NO CONFIG FILE" & @CRLF & "Please WRITE CONFIG file first!","Normal")
	  Else
		 ;GET SYSTEM STRINGS
		 Local $sPos_Temp = StringInStr ($sConfig, "<<<-- SYSTEM STRINGS -->>>") + StringLen ("<<<-- SYSTEM STRINGS -->>>" & @CRLF)
		 Local $sCount_Temp = StringInStr ($sConfig, "<<<-- PART STRINGS -->>>") - StringInStr ($sConfig, "<<<-- SYSTEM STRINGS -->>>") - StringLen ("<<<-- SYSTEM STRINGS -->>>" & @CRLF & @CRLF & @CRLF)
		 Local $sConfig_SystemStrings = StringMid ($sConfig, $sPos_Temp, $sCount_Temp)
		 Local $aConfig_SystemStrings = StringSplit ($sConfig_SystemStrings, @CRLF,  $STR_ENTIRESPLIT +  $STR_NOCOUNT)
		 ;GET PART STRINGS
		 Local $sPos_Temp = StringInStr ($sConfig, "<<<-- PART STRINGS -->>>") + StringLen ("<<<-- PART STRINGS -->>>" & @CRLF)
		 Local $sCount_Temp = StringInStr ($sConfig, "<<<-- SEARCH STRINGS -->>>") - StringInStr ($sConfig, "<<<-- PART STRINGS -->>>") - StringLen ("<<<-- PART STRINGS -->>>" & @CRLF & @CRLF & @CRLF)
		 Local $sConfig_PartStrings = StringMid ($sConfig, $sPos_Temp, $sCount_Temp)
		 Local $aConfig_PartStrings = StringSplit ($sConfig_PartStrings, @CRLF,  $STR_ENTIRESPLIT +  $STR_NOCOUNT)
		 ;GET SEARCH STRINGS
		 Local $sPos_Temp = StringInStr ($sConfig, "<<<-- SEARCH STRINGS -->>>") + StringLen ("<<<-- SEARCH STRINGS -->>>" & @CRLF)
		 Local $sCount_Temp = StringInStr ($sConfig, "Last Saved") - StringInStr ($sConfig, "<<<-- SEARCH STRINGS -->>>") - StringLen ("<<<-- SEARCH STRINGS -->>>" & @CRLF & @CRLF)
		 Local $sConfig_SearchStrings = StringMid ($sConfig, $sPos_Temp, $sCount_Temp)
		 Local $aConfig_SearchStrings = StringSplit ($sConfig_SearchStrings, @CRLF,  $STR_ENTIRESPLIT +  $STR_NOCOUNT)

		 ;SUB STRINGS = SYSTEM STRINGS + PART STRINGS
		 Local $aConfig_SubStrings = $aConfig_SystemStrings
		 For $i = 0 to UBound ($aConfig_PartStrings) - 1
			_ArrayAdd ($aConfig_SubStrings, $aConfig_PartStrings[$i])
		 Next


		 ;------------------------------------
		 ;Put search strings from config file into a 2-dimension array
		 Local $aSearch_Strings [5000][50]
		 Local $D1 = 0, $D2 = 0, $D1_Max = 1, $D2_Max = 1
		 For $vElement In $aConfig_SearchStrings
			$D1 = StringMid ($vElement, StringInStr ($vElement, "[", 0 ,1) + 1, StringInStr ($vElement, "]", 0 ,1) - StringInStr ($vElement, "[", 0 ,1) - 1)
			$D2 = StringMid ($vElement, StringInStr ($vElement, "[", 0 ,2) + 1, StringInStr ($vElement, "]", 0 ,2) - StringInStr ($vElement, "[", 0 ,2) - 1)
			$sSearch_String = StringRight ($vElement, StringLen ($vElement) - StringInStr ($vElement, "]", 0 ,2))
			If Number ($D1) > $D1_Max Then $D1_Max = Number ($D1)
			If Number ($D2) > $D2_Max Then $D2_Max = Number ($D2)
			$aSearch_Strings [$D1][$D2] = $sSearch_String
		 Next
		 ;------------------------------------
		 Local $sLast_Save = StringRight ($sConfig, StringLen ($sConfig) - StringInStr ($sConfig, " ", 0, - 1))
		 For $i = $sLast_Save To $D1_Max
			;NẾU KHÔNG THẤY STRING THÌ BỎ QUA
			If $aSearch_Strings [$i][0] = "" Then ContinueLoop
			;Khai báo các biến dùng để xác định số lượng parts trong DTC
			Local $iDTC_Part_Nums = 0
			;Vòng lặp để xác định số lượng parts trong DTC
			For $j = 0 to $D2_Max
			   If $aSearch_Strings [$i][$j] <> "" Then $iDTC_Part_Nums += 1
			Next
			;------------------------------------
			;ĐOẠN CODE TẠO DTC, CHÈN LINK CHO DTC NHIỀU PARTS
			Local $sInsert_Path = ""
			For $j = $iDTC_Part_Nums - 1 To 0 Step -1
			   Local $sResult_Open = Open_DTC_Link ($oIE, $aSearch_Strings [$i][$j])
			   If $sResult_Open = "No Link Found" Then
				  ExitLoop 2
			   EndIf
			   ;------------------------------------
			   ;CHECK IF THE LINK ALREADY EXISTED OR NOT, IF LINK EXISTS BUT IT IS PART LINK => STILL DO IT
			   If Check_Log_File ($sYMME, "Log File DTC Successful.txt", $sLink_DTC) = "Not Exist" Or $j <> 0 Then


				  Local $sSub_String = Get_SubString ($aSearch_Strings [$i][$j])
				  Local $sSub_Name = ""
				  For $vElement In $aConfig_SubStrings
					 Local $sConfig_Alter_String = StringLeft ($vElement, StringInStr ($vElement, " *** ") - 1)
					 Local $sConfig_Sub_String = StringRight ($vElement, StringLen ($vElement) - StringLen ($sConfig_Alter_String & " *** "))
					 If $sSub_String = $sConfig_Sub_String Then
						$sSub_Name = $sConfig_Alter_String
						ExitLoop
					 EndIf
				  Next



				  ;Lấy link DTC đầu tiên làm Main Link
				  ;Kiểm tra DTC giống nhau thì lấy tên Alter phía sau
				  If $j = 0 Then
					 If $i = 0 Then
						If StringLeft($aSearch_Strings [$i][0],5) = StringLeft($aSearch_Strings [$i + 1][0],5) Then
						   $sSub_Name = " (" & Standardize_File_Name ($sSub_Name) & ")"
						Else
						   $sSub_Name = ""
						EndIf
					 Elseif $i = $D1_Max Then
						If StringLeft($aSearch_Strings [$i][0],5) = StringLeft($aSearch_Strings [$i - 1][0],5) Then
						   $sSub_Name = " (" & Standardize_File_Name ($sSub_Name) & ")"
						Else
						   $sSub_Name = ""
						EndIf
					 Else
						If StringLeft($aSearch_Strings [$i][0],5) = StringLeft($aSearch_Strings [$i + 1][0],5) Or StringLeft($aSearch_Strings [$i][0],5) = StringLeft($aSearch_Strings [$i - 1][0],5) Then
						   $sSub_Name = " (" & Standardize_File_Name ($sSub_Name) & ")"
						Else
						   $sSub_Name = ""
						EndIf
					 EndIf
					 DTC_Procedure_Alldata ($oIE, $sInsert_Path, $sSub_Name, "Main")
				  ;Link part
				  Else
					 $sInsert_Path = DTC_Procedure_Alldata ($oIE, $sInsert_Path, "_" & $sSub_Name, "Not Main")
				  EndIf
			   Else ;Exist
				  Notification ("Found a DTC has been GENERATED BEFORE" & @CRLF & "Please CHECK!", "Normal")
			   EndIf
			Next
			Save_Current_Work ($sFilePath_YMME, $i)
		 Next
		 If $sResult_Open <> "No Link Found" Then
			Notification ("DONE" & @CRLF & "Please CHECK!", "Normal")
			MsgBox ($MB_TOPMOST, "Message", "DONE" & @CRLF & "Please CHECK!")
		 EndIf
	  EndIf
   ;If the link is not YMME link
   Else
	  Notification ("The link is not Vehicle Link" & @CRLF & "Please ENTER A VEHICLE LINK!", "Normal")
   EndIf
   Return $oIE
EndFunc



;====================================================================================================================
;                  FUNCTION DESCRIPTION: OPEN DTC LINK FROM CONFIG FILE
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func Open_DTC_Link (Byref $oIE, $sSearch_String)
   Local $sResult = ""
   ;------------------------------------------
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
	  ;Set search string (ADDED "Testing and Inspection" BEFORE THE SEARCH STRING TO MAKE SURE THE LINK IS RIGHT
	  _IEFormElementSetValue($oSearchBox, "Testing and Inspection >> " & $sSearch_String)
	  ;Submit form, no wait for page load to complete
	  _IEFormSubmit($oForm, 0)
	  ;Wait for the page load to complete
	  _IELoadWait($oIE)
	  ;------------------------------------
	  ;ĐOẠN CODE SỬ DỤNG FUNCTION "Check_Subscription_Alldata" ĐỂ KIỂM TRA SUBSCIPTION
	  Check_Subscription_Alldata ($oIE, $sLink_YMME)

	  ;NẾU SEARCH KHÔNG RA LINK THÌ BỎ "Testing and Inspection >> "
	  Local $oLink = _IEGetObjById ($oIE, "category_link")
	  If @error = $_IEStatus_NoMatch Then
		 IENavigate_Check_Error ($oIE, $sLink_YMME)
		 Check_Subscription_Alldata ($oIE, $sLink_YMME)
		 ;Lấy object form search
		 Local $oForm = _IEFormGetObjByName($oIE, "simpleSearch")
		 ;Lấy object Search box
		 Local $oSearchBox = _IEFormElementGetObjByName($oForm, "searchQuery")
		 ;Set search string (ADDED "Testing and Inspection" BEFORE THE SEARCH STRING TO MAKE SURE THE LINK IS RIGHT
		 _IEFormElementSetValue($oSearchBox, $sSearch_String)
		 ;Submit form, no wait for page load to complete
		 _IEFormSubmit($oForm, 0)
		 ;Wait for the page load to complete
		 _IELoadWait($oIE)
		 ;------------------------------------
		 ;ĐOẠN CODE SỬ DỤNG FUNCTION "Check_Subscription_Alldata" ĐỂ KIỂM TRA SUBSCIPTION
		 Check_Subscription_Alldata ($oIE, $sLink_YMME)
	  EndIf

	  If _IEPropertyGet ($oIE, "title") = "ALLDATA Repair - Search Results" Then
		 ;------------------------------------

		 Local $oDIVs = _IETagNameGetCollection ($oIE, "div")
		 Local $iCompare_Result = 0
		 Local $oLink
		 Local $bFlag_None = False
		 For $oDIV In $oDIVs
			If $oDIV.id = "category_link" Then
			   If Compare_Strings ($sSearch_String, $oDIV.innertext) > $iCompare_Result Then
				  $iCompare_Result  = Compare_Strings ($sSearch_String, $oDIV.innertext)
				  $oLink = $oDIV
				  $bFlag_None = True
			   EndIf
			EndIf
		 Next
		 ;------------------------------------
		 ;ĐOẠN CODE CHECK NẾU KHÔNG CÓ KẾT QUẢ SEARCH THÌ THÔNG BÁO USER CHƯA CHỌN ĐÚNG XE
		 If $bFlag_None = False Then
			Notification ("Vehicle selected in IE does not match the config file" & @CRLF & "Please select the correct vehicle and REBEGIN!","Normal")
			$sResult = "No Link Found"
			ExitLoop
		 Else
			$sResult = "Link Found"
		 EndIf

		 ;------------------------------------
		 Local $sTemp = _IEPropertyGet ($oLink,"innerhtml")
		 ;------------------------------------
		 ;ĐOẠN CODE LẤY ID TẠO LINK
		 Local $aIDs [4]
		 For $i = 1 To 4 Step 1
			$aIDs [$i-1] =  StringMid ($sTemp, Stringinstr ($sTemp, ",", 0, $i) + 1, Stringinstr ($sTemp, ",", 0, $i + 1) - Stringinstr ($sTemp, ",", 0, $i) - 1)
		 Next
		 $sLink_DTC = "http://repair.alldata.com/alldata/article/display.action?componentId=" & $aIDs [0] & "&iTypeId=" & $aIDs [1] & "&nonStandardId=" & $aIDs [2] & "&vehicleId=" & $aIDs [3] & "&windowName=mainADOnlineWindow"
		 IENavigate_Check_Error ($oIE, $sLink_DTC)
		 ;------------------------------------
		 ;ĐOẠN CODE SỬ DỤNG FUNCTION "Check_Subscription_Alldata" ĐỂ KIỂM TRA SUBSCIPTION
		 Check_Subscription_Alldata ($oIE, $sLink_DTC)
		 ExitLoop
	  EndIf
   Until 0
   Return $sResult
EndFunc



;====================================================================================================================
;                  FUNCTION DESCRIPTION: COMPARE THE SIMILARITY OF 2 STRINGS
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func Compare_Strings ($sString1, $sString2)
   ;Chuẩn hóa strings
   $sString1 = Standardize_String ($sString1)
   $sString2 = Standardize_String ($sString2)
   ;Chuyển strings thành array
   Local $aString1 = StringSplit ($sString1, "")
   Local $aString2 = StringSplit ($sString2, "")
   ;Lấy mảng dài hơn làm Max Bound
   If $aString1 [0] <= $aString2 [0] Then
	  Local $iMax_Bound = $aString1 [0]
   Else
	  Local $iMax_Bound = $aString2 [0]
   EndIf

   ;Xem độ giống của string từ dưới lên trên
   Local $iSame_Count = 0
   For $i = 0 To $iMax_Bound - 1
	  If $aString1 [$aString1[0] - $i] = $aString2 [$aString2[0] - $i] Then $iSame_Count += 1
   Next
   ;Trả lại số lần giống
   Return $iSame_Count
EndFunc




;====================================================================================================================
;                  FUNCTION DESCRIPTION: WRITE CONFIG FILE FOR SCANNING FUNCTION
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func Save_Current_Work ($sFilePath_YMME, $iPosition)
	  Local $hFileOpen = FileOpen($sFilePath_YMME & "\" & "Scan DTC Config" & ".txt", $FO_READ)
	  Local $sConfig = FileRead ($hFileOpen)
	  FileClose($hFileOpen)
	  ;Delete the last position
	  $sConfig = StringLeft ($sConfig, StringInStr ($sConfig, " ", 0, - 1))
	  ;Write new position
	  $sConfig = $sConfig & $iPosition
	  ;WRITE CONFIG FILE
	  Write_Log_File ($sFilePath_YMME, "Scan DTC Config", $sConfig, "overwrite")
EndFunc


;====================================================================================================================
;                  FUNCTION DESCRIPTION: WRITE CONFIG FILE FOR SCANNING FUNCTION
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func Write_Config ()
   ;Gán trang web cho biến object
   Local $oIE = IECreate_Check_Error($sLink_YMME, $bWeb_Attach, $bWeb_Visible, $bWeb_Wait, $bWeb_TakeFocus)
   Sleep (1000)
   _IEAction($oIE, "refresh")
   ;------------------------------------
   ;ĐOẠN CODE SỬ DỤNG FUNCTION "Check_Subscription_Alldata" ĐỂ KIỂM TRA SUBSCIPTION
   Check_Subscription_Alldata ($oIE, $sLink_YMME)
   ;------------------------------------
   ;ĐOẠN CODE LẤY TEXT TRONG TAG <TITLE> ĐỂ KIỂM TRA XEM LINK ĐÓ CÓ PHẢI LINK YMME KHÔNG
   Local $sTxt_Title = _IEPropertyGet ($oIE, "title")
   $sTxt_Title = Standardize_String ($sTxt_Title)
   If StringInStr ($sTxt_Title, "ALLDATA Repair - Vehicle Information - ") <> 0 Then
	  ;------------------------------------
	  ;ÐOẠN CODE LẤY YMME ĐẶT TÊN CHO FOLDER
	  $sYMME = _IEPropertyGet ($oIE, "title")
	  $sYMME = StringRight ($sYMME, StringLen ($sYMME) - StringInStr ($sYMME, "-", 0, 2) - 1)
	  $sYMME = Standardize_File_Name ($sYMME)
	  ;Tạo các thư mục cần thiết
	  Local $sFilePath_Alldata_DTC = @ScriptDir & "\INNOVA"
	  If FileExists ($sFilePath_Alldata_DTC) = 0 Then	DirCreate($sFilePath_Alldata_DTC)
	  Local $sFilePath_YMME        = @ScriptDir & "\INNOVA"      &"\"&$sYMME
	  If FileExists ($sFilePath_YMME) = 0 Then	DirCreate($sFilePath_YMME)
	  ;------------------------------------
	  Notification ("Expanding The Tree View ..." & @CRLF & "It's gonna take a whille!","Normal")
	  ;-----------------------------------
	  ;CLICK TREE VIEW LEVEL 1
	  $i1 = 1
	  Local $aSearch_Strings [1]
	  Local $aLevel [1]
	  Local $iOrder = 0
		 ;Click "vehicle"
		 Local $aChildren_Item_1 = Click_Tree_View_By_Text ($oIE, "Vehicle")

		 For $i2 = 1 To $aChildren_Item_1[0]
			If $aChildren_Item_1[$i2] = "A L L Diagnostic Trouble Codes ( DTC )" Or $aChildren_Item_1[$i2] = " A L L Diagnostic Trouble Codes ( DTC )"Then

			   ;Click "A L L Diagnostic Trouble Codes ( DTC )"
			   $aChildren_Item_2 = Click_Tree_View_By_Text ($oIE, $aChildren_Item_1[$i2])

			   ;-----------------------------------
			   For $i3 = 1 To $aChildren_Item_2 [0]
				  If $aChildren_Item_2[$i3] = "Information for A L L Diagnostic Trouble Codes ( DTC )" Or $aChildren_Item_2[$i3] = " Information for A L L Diagnostic Trouble Codes ( DTC )" Then
					 ;CLick "Information for A L L Diagnostic Trouble Codes ( DTC )"
					 $aChildren_Item_3 = Click_Tree_View_By_Text ($oIE, $aChildren_Item_2[$i3])
					 ;-----------------------------------
					 For $i4 = 1 To $aChildren_Item_3 [0]
						If $aChildren_Item_3[$i4] = "Testing and Inspection" Or $aChildren_Item_3[$i4] = " Testing and Inspection" Then
						   ;Click "Testing and Inspection"
						   $aChildren_Item_4 = Click_Tree_View_By_Text ($oIE, $aChildren_Item_3[$i4])
						   ;-----------------------------------
						   For $i5 = 1 To $aChildren_Item_4 [0]
							  If $aChildren_Item_4[$i5] = "P Code Charts" Or $aChildren_Item_4[$i5] = " P Code Charts" Then
								 ;Click "P Code Charts"
								 $aChildren_Item_5 = Click_Tree_View_By_Text ($oIE, $aChildren_Item_4[$i5])
								 ;-----------------------------------
								 For $i6 = 1 To $aChildren_Item_5 [0]
;~ 		 If  $aChildren_Item_5 [$i6] = "P0AA6" Or $aChildren_Item_5 [$i6] = "P0AA4" Or $aChildren_Item_5 [$i6] = "P0AA7" Or $aChildren_Item_5 [$i6] = "P0ACD" Or $aChildren_Item_5 [$i6] = "P0AC8" Or $aChildren_Item_5 [$i6] = "P3190" Or $aChildren_Item_5 [$i6] = "P3004" Or $aChildren_Item_5 [$i6] = "P3193" Then
									;Click Pxxx
									$aChildren_Item_6 = Click_Tree_View_By_Text ($oIE, $aChildren_Item_5[$i6])
									;Put data into array

									$aSearch_Strings [$iOrder] = $aChildren_Item_5[$i6]
									$aLevel [$iOrder] = $i6
									$sReplace_String_1 = $aSearch_Strings [$iOrder]

									If $aChildren_Item_6 [0] = 0 Then
									   $iOrder += 1
									   ReDim $aSearch_Strings[UBound($aSearch_Strings) + 1]
									   ReDim $aLevel[UBound($aLevel) + 1]
									EndIf
									;----------------------------------
									For $i7 = 1 To $aChildren_Item_6 [0]
									   ;Click System 1
									   $aChildren_Item_7 = Click_Tree_View_By_Text ($oIE, $aChildren_Item_6[$i7])
									   ;Put data into array
									   $aSearch_Strings [$iOrder] = $sReplace_String_1 & " >> " & $aChildren_Item_6[$i7]
									   $aLevel [$iOrder] = $i6 & "\" & $i7
									   $sReplace_String_2 = $aSearch_Strings [$iOrder]
									   If $aChildren_Item_7 [0] = 0 Then
										  $iOrder += 1
										  ReDim $aSearch_Strings[UBound($aSearch_Strings) + 1]
										  ReDim $aLevel[UBound($aLevel) + 1]
									   EndIf
									   ;-----------------------------------
									   For $i8 = 1 To $aChildren_Item_7 [0]
										  ;Click System 2
										  $aChildren_Item_8 = Click_Tree_View_By_Text ($oIE, $aChildren_Item_7[$i8])
										  ;Put data into array
										  $aSearch_Strings [$iOrder] = $sReplace_String_2 & " >> " & $aChildren_Item_7[$i8]
										  $aLevel [$iOrder] = $i6 & "\" & $i7 & "\" & $i8
										  $sReplace_String_3 = $aSearch_Strings [$iOrder]
										  If $aChildren_Item_8 [0] = 0 Then
											 $iOrder += 1
											 ReDim $aSearch_Strings[UBound($aSearch_Strings) + 1]
											 ReDim $aLevel[UBound($aLevel) + 1]
										  EndIf
										  ;-----------------------------------
										  For $i9 = 1 To $aChildren_Item_8 [0]
											 ;Click System 3
											 $aChildren_Item_9 = Click_Tree_View_By_Text ($oIE, $aChildren_Item_8[$i9])
											 ;Put data into array
											 $aSearch_Strings [$iOrder] = $sReplace_String_3 & " >> " & $aChildren_Item_8[$i9]
											 $aLevel [$iOrder] = $i6 & "\" & $i7 & "\" & $i8 & "\" & $i9
											 $sReplace_String_4 = $aSearch_Strings [$iOrder]
											 If $aChildren_Item_9 [0] = 0 Then
												$iOrder += 1
												ReDim $aSearch_Strings[UBound($aSearch_Strings) + 1]
												ReDim $aLevel[UBound($aLevel) + 1]
											 EndIf
											 ;-----------------------------------
											 For $i10 = 1 To $aChildren_Item_9 [0]
												;Click System Part
												$aChildren_Item_10 = Click_Tree_View_By_Text ($oIE, $aChildren_Item_9[$i10])
												;Put data into array
												$aSearch_Strings [$iOrder] = $sReplace_String_4 & " >> " & $aChildren_Item_9[$i10]
												$aLevel [$iOrder] = $i6 & "\" & $i7 & "\" & $i8 & "\" & $i9 & "\" & $i10
												$sReplace_String_5 = $aSearch_Strings [$iOrder]
												If $aChildren_Item_10 [0] = 0 Then
												   $iOrder += 1
												   ReDim $aSearch_Strings[UBound($aSearch_Strings) + 1]
												   ReDim $aLevel[UBound($aLevel) + 1]
												EndIf
												;-----------------------------------
												For $i11 = 1 To $aChildren_Item_10 [0]
												   ;Click Reserve
												   $aChildren_Item_10 = Click_Tree_View_By_Text ($oIE, $aChildren_Item_10[$i11])
												   ;Put data into array
												   $aSearch_Strings [$iOrder] = $sReplace_String_5 & " >> " & $aChildren_Item_10[$i11]
												   $aLevel [$iOrder] = $i6 & "\" & $i7 & "\" & $i8 & "\" & $i9 & "\" & $i10 & "\" & $i11
												   $sReplace_String_6 = $aSearch_Strings [$iOrder]
												   If $aChildren_Item_11 [0] = 0 Then
													  $iOrder += 1
													  ReDim $aSearch_Strings[UBound($aSearch_Strings) + 1]
													  ReDim $aLevel[UBound($aLevel) + 1]
												   EndIf
												Next
											 Next
										  Next
									   Next
									Next

;~ 		 EndIf  ;--------------

								 Next
							  EndIf
						   Next
						EndIf
					 Next
				  EndIf
			   Next
			EndIf
		 Next
	  ;-----------------------------------------
	  ;DELETE EMPTY STRINGS AT THE END OF THE ARRAY
	  If $aSearch_Strings[UBound ($aSearch_Strings) - 1] = "" Then _ArrayDelete ($aSearch_Strings, UBound ($aSearch_Strings) - 1)


	  ;-----------------------------------------
	  ;GET SUBSTRINGS ARRAY
	  Local $aSub_Strings [0]
	  For $i = 0 To UBound ($aSearch_Strings) - 1
		 Local $sSub_String = Get_SubString ($aSearch_Strings[$i])
		 If $sSub_String <> "" And _ArraySearch ($aSub_Strings, $sSub_String) = -1 Then
			ReDim $aSub_Strings [UBound ($aSub_Strings) + 1]
			$aSub_Strings [UBound ($aSub_Strings) - 1] = $sSub_String
		 EndIf
		 _ArraySort ($aSub_Strings)
	  Next
	  ;-----------------------------------------
	  MsgBox ($MB_TOPMOST, "Message", "The Procedure Generator needs your help!!!!")
	  Sleep (200)
	  WinActivate ($sVersion)



	  ;GET PARTSTRINGS ARRAY
	  Local $aPart_Strings = User_Input_Part_Strings (_ArrayToString($aSub_Strings, @CRLF))
	  ;-----------------------------------------
	  ;GET NOT PARTSTRINGS ARRAY
	  Local $aSystem_Strings = Array_Minus ($aSub_Strings, $aPart_Strings)

	  ;-----------------------------------------
	  ;GET PART STRINGS AND NOT PART STRINGS
	  Local $sPart_Strings = _ArrayToString ($aPart_Strings, @CRLF)
	  Local $sSystem_Strings = _ArrayToString ($aSystem_Strings, @CRLF)
	  ;--------------------------
	  Notification ("Writing Config file for: " & @CRLF & $sYMME, "Normal")
	  Sleep (1000)
	  ;WRITE CONFIG FILE
	  Write_Log_File ($sFilePath_YMME, "Scan DTC Config",  "This config file is to save System Strings, Part Strings and Search Strings" & @CRLF & "for the tool to get all DTC links of the vehicle: " & $sYMME & @CRLF & "Model Year Link: " & $sLink_YMME, "overwrite")

	  Write_Log_File ($sFilePath_YMME, "Scan DTC Config",  @CRLF & @CRLF & "<<<-- SYSTEM STRINGS -->>>", "append")
	  For $iSys = 1 To UBound($aSystem_Strings)
		 Write_Log_File ($sFilePath_YMME, "Scan DTC Config", @CRLF & "" & $iSys & " *** " & $aSystem_Strings [$iSys - 1], "append")
	  Next


	  Write_Log_File ($sFilePath_YMME, "Scan DTC Config",  @CRLF & @CRLF & "<<<-- PART STRINGS -->>>", "append")
	  For $iPart = 1 To UBound($aPart_Strings)
		 Write_Log_File ($sFilePath_YMME, "Scan DTC Config", @CRLF & "" & $iPart + UBound($aSystem_Strings) & " *** " & $aPart_Strings [$iPart - 1], "append")
	  Next


	  Write_Log_File ($sFilePath_YMME, "Scan DTC Config",  @CRLF & @CRLF & "<<<-- SEARCH STRINGS -->>>", "append")

	  Local $bPart_Flag = False
	  Local $i = -1
	  Local $j = 0
	  For $iOrder = 0 To UBound ($aSearch_Strings) - 1
		 ;Nếu không có part
		 If _ArraySearch ($aPart_Strings, Get_SubString ($aSearch_Strings[$iOrder])) = -1 Then
			$j = 0
			$i += 1
			Write_Log_File ($sFilePath_YMME, "Scan DTC Config",  @CRLF & "[" & $i & "]" & "[" & $j & "]" & $aSearch_Strings[$iOrder], "append")
			$bPart_Flag = False
		 ;Nếu có part
		 Else
			;Nếu Không có DTC chứa part phía trước nó
			If $bPart_Flag = False Then
			   $j = 0
			   $i += 1
			   Write_Log_File ($sFilePath_YMME, "Scan DTC Config",  @CRLF & "[" & $i & "]" & "[" & $j & "]" & $aSearch_Strings[$iOrder], "append")
			   $bPart_Flag = True
			;Nếu có DTC chứa part phía trước nó
			Else
			   ;Nếu DTC có stt PART STRING lớn hơn phía trước là 1
			   If Check_Search_String_Type ($aLevel[$iOrder], $aLevel[$iOrder - 1]) = "Not first part" Then
				  $j += 1
				  $i = $i
				  Write_Log_File ($sFilePath_YMME, "Scan DTC Config",  @CRLF & "[" & $i & "]" & "[" & $j & "]" & $aSearch_Strings[$iOrder], "append")
				  $bPart_Flag = True
			   Else
				  $j = 0
				  $i += 1
				  Write_Log_File ($sFilePath_YMME, "Scan DTC Config",  @CRLF & "[" & $i & "]" & "[" & $j & "]" & $aSearch_Strings[$iOrder], "append")
				  $bPart_Flag = True
			   EndIf
			EndIf
		 EndIf
	  Next

	  Write_Log_File ($sFilePath_YMME, "Scan DTC Config",  @CRLF & "Last Saved: New", "append")
	  ;-------------------------
	  ;Complete writing
	  Notification ("Completed writing config file for """ & $sYMME & """" & @CRLF & "Please CHECK!", "Normal")

   Else
	  Notification ("The link is not Vehicle Link" & @CRLF & "Please ENTER A VEHICLE LINK!", "Normal")
   EndIf

EndFunc

Func Check_Search_String_Type ($sCur_Level, $sPre_Level)
   $sResult = ""
   StringReplace ($sCur_Level,"\","/")
   Local $iRep_Num_1 = @extended
   StringReplace ($sPre_Level,"\","/")
   Local $iRep_Num_2 = @extended
   $sCur_Level = "\" & $sCur_Level & "\"
   $sPre_Level = "\" & $sPre_Level & "\"
   Local $iCount = 0
   If $iRep_Num_1 = $iRep_Num_2 Then
	  For $ii = 1 To @extended + 1
		 $iMinus =  Number (StringMid ($sCur_Level, StringInStr ($sCur_Level, "\", 0, $ii) + 1, StringInStr ($sCur_Level, "\", 0, $ii + 1) - StringInStr ($sCur_Level, "\", 0, $ii) - 1)) _
			      - Number (StringMid ($sPre_Level, StringInStr ($sPre_Level, "\", 0, $ii) + 1, StringInStr ($sPre_Level, "\", 0, $ii + 1) - StringInStr ($sPre_Level, "\", 0, $ii) - 1))
		 If $iMinus = 1  Then
			$iCount += 1
		 ElseIf $iMinus = 0 Then
			$iCount += 0
		 Else
			$iCount = 999
		 EndIf
	  Next
   Else
	  $sResult = "First part"
   EndIf

   If $iCount = 1 Then
	  $sResult = "Not first part"
   Else
	  $sResult = "First part"
   EndIf

   Return $sResult

EndFunc




;====================================================================================================================
;                  FUNCTION DESCRIPTION: USER INPUTS PART STRINGS INTO COMMNUNICATION SCREEN
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func User_Input_Part_Strings ($sSub_Strings)
   Notification_Clear()
   ;ASK USER TO HELP ON PARTS STRINGS
   _GUICtrlEdit_SetReadOnly ($Commu_Ctrl, False)
   Notification ("SUB STRINGS:" & @CRLF & "----------------------------------------" & @CRLF & $sSub_Strings & @CRLF & "----------------------------------------" _
			   & @CRLF & "NOTE: Please help split PARTS STRINGS" & @CRLF & "out of the strings above then paste them below" & @CRLF & "and Press ENTER then type EXECUTE " & @CRLF & "----------------------------------------", "Normal")

   While _GUICtrlEdit_GetLine ($Commu_Ctrl, _GUICtrlEdit_GetLineCount ($Commu_Ctrl) - 2) <> "EXECUTE"
	  Sleep (100)
   WEnd
   _GUICtrlEdit_SetReadOnly ($Commu_Ctrl, True)

   Local $sNoti = GUICtrlRead ($Commu_Ctrl)
   Local $iStart_Pos = StringInStr ($sNoti, "----------------------------------------", 0, -1) + StringLen ("----------------------------------------")
   Local $iEnd_Pos = StringInStr ($sNoti, "EXECUTE", 0, -1) + StringLen ("EXECUTE")
   Local $sTemp = StringMid ($sNoti, $iStart_Pos, $iEnd_Pos - $iStart_Pos)

   Local $aTemp = StringSplit ($sTemp, @CRLF, $STR_ENTIRESPLIT)
   Local $aPart_Strings [0]
   For $i = 1 To $aTemp[0]
	  If StringInStr ($sSub_Strings, $aTemp[$i]) <> 0 Then
		 ReDim $aPart_Strings[UBound ($aPart_Strings) + 1]
		 $aPart_Strings[UBound ($aPart_Strings) - 1] = $aTemp[$i]
	  EndIf
   Next
   Notification ("----------------------------------------" & @CRLF & "EXECUTING ..." & @CRLF & "Please wait for the App to create config file!", "Normal")
   Sleep (2000)
   Return $aPart_Strings
EndFunc



;====================================================================================================================
;                  FUNCTION DESCRIPTION: GET SYSTEMS AND PARTS STRINGS EXISTING
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func Array_Minus ($aArray_A, $aArray_B)
   For $vElement In $aArray_B
	  Local $Pos = _ArraySearch ($aArray_A, $vElement)
	  _ArrayDelete ($aArray_A, $Pos)
   Next
   _ArraySort ($aArray_A)
   $aArray_A = _ArrayUnique ($aArray_A, 0, 0, 0, $ARRAYUNIQUE_NOCOUNT)
   Return $aArray_A
EndFunc







;====================================================================================================================
;                  FUNCTION DESCRIPTION: GET SUBSTRING
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func Get_SubString ($sSearch_String)
   If StringInStr($sSearch_String, ">> ", 0, 1) <> 0 Then
	  Local $iStart_Pos = StringInStr($sSearch_String, ">> ", 0, 1) + StringLen(">> ")
	  Local $sSub_String = StringRight ($sSearch_String, StringLen($sSearch_String) - $iStart_Pos + 1)
   Else
	  $sSub_String = ""
   EndIf
   Return $sSub_String
EndFunc


;====================================================================================================================
;                  FUNCTION DESCRIPTION: GET SUBSTRING
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func Get_DTC_String ($sSearch_String)
   Local $iEnd_Pos = StringInStr($sSearch_String, ">> ", 0, 1)
   Local $sDTC_String = StringLeft ($sSearch_String, $iEnd_Pos - 2)
   Return $sDTC_String
EndFunc




;====================================================================================================================
;                  FUNCTION DESCRIPTION: STANDARDIZE LIST
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func Standardize_List ($sList)
   ;Thêm xuống dòng để dễ chuẩn hóa
   $sList = @CRLF & $sList & @CRLF


   ;Replace {space} enter  bằng 1 enter
   Do
	  $sList = StringReplace ( $sList, " " & @CRLF, @CRLF)
   Until @extended = 0

   ;Replace {space} enter  bằng 1 enter
   Do
	  $sList = StringReplace ( $sList, @CRLF & " ", @CRLF)
   Until @extended = 0

   ;Replace 2 enter  bằng 1 enter
   Do
	  $sList = StringReplace ( $sList, @CRLF & @CRLF, @CRLF)
   Until @extended = 0
   If StringLeft ($sList, 2) = @CRLF Then $sList = StringRight ($sList, StringLen($sList) - 2)
   If StringRight ($sList, 2) = @CRLF Then $sList = StringLeft ($sList, StringLen($sList) - 2)
   Return $sList
EndFunc




;====================================================================================================================
;                  FUNCTION DESCRIPTION: CLICK TREE VEIEW BY TEXT
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func Click_Tree_View_By_Text (Byref $oIE, $sTxt)
;~ Local $hTimer = TimerInit() ;
	  ;-------------------------------------------
	  ;GET ID
	  ;Get ID Click by text
	  Local $sID_Click = Get_CLickID_By_Text ($oIE, $sTxt)
;~ ConsoleWrite($sTxt & " ---- Time Difference 0: " & TimerDiff($hTimer) & @CRLF)
	  ;Get ID children
	  Local $sID_Children = StringReplace ($sID_Click, "ygtvt", "ygtvc")
	  ;Get ID Status
	  Local $sID_Status = StringReplace ($sID_Click, "ygtvt", "ygtvtableel")

	  ;-------------------------------------------
	  Sleep (100)
	  Local $sStatus = _IEGetObjById ($oIE, $sID_Status).GetAttribute("class")
	  If StringInStr ($sStatus, "expanded") = 0 Then
		 ;CLICK
		 ;Get Object Click
		 Local $oClick = _IEGetObjById ($oIE, $sID_Click)
		 ;Click object
		 If _IEGetObjById ($oIE, $sID_Children).innertext = "" Then _IEAction($oClick, "click")
	  EndIf

	  Sleep (100)
	  Local $sStatus = _IEGetObjById ($oIE, $sID_Status).GetAttribute("class")
	  If StringInStr ($sStatus, "expanded")  <> 0 Then
		 ;-------------------------------------------
		 ;WAIT
		 ;Wait for the text to appear
		 Do ;Loop until get the SubOject
			Sleep (200)
		 Until _IEGetObjById ($oIE, $sID_Children).innertext <> ""
		 ;--------------------------------------------
		 ;Get Text From Children
		 Local $sChildren_Item = Standardize_List (_IEGetObjById ($oIE, $sID_Children).innertext)
		 Local $aChildren_Item = StringSplit ($sChildren_Item, @CRLF, $STR_ENTIRESPLIT)
	  Else
		 Local $aChildren_Item = [0]
	  EndIf

   Return $aChildren_Item
EndFunc






;====================================================================================================================
;                  FUNCTION DESCRIPTION: GET ID OF AN ELEMENT IN TREE VIEW BY TEXT
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func Get_CLickID_By_Text (ByRef $oIE, $sTxt)
   ;----------------------------------------------------
   Local $oTags = _IETagNameGetCollection($oIE, "td")
   Local $Array [0]
   Local $i = 0, $iMark = 0

   While StringLeft ($sTxt, 1) = " "
	  $sTxt = StringRight ($sTxt, StringLen ($sTxt) - 1)
   WEnd

   While StringRight ($sTxt, 1) = " "
	  $sTxt = StringLeft ($sTxt, StringLen ($sTxt) - 1)
   WEnd

   ;Loop until found the string
;~ Local $hTimer = TimerInit() ;
   Do
	  For $oTag In $oTags
		 ReDim $Array [$i + 1]
		 $Array [$i] = $oTag
		 If $oTag.innertext = $sTxt  Then
			$iMark = $i
		 EndIf
		 $i += 1
	  Next
   Until $iMark <> 0

;~ ConsoleWrite($sTxt & " ---- Time Difference Get Click ID: " & TimerDiff($hTimer) & @CRLF)

   Local $oTemp = $Array [$iMark]
   Local $sID_Content = $oTemp.id
   Local $sID_CLick = StringReplace ($sID_Content, "ygtvcontentel", "ygtvt")
   Return $sID_CLick
EndFunc


