#cs
Local $sPart_Strings = "Part 1" & @CRLF & "Part 2" & @CRLF & "Part 3"
Local $sSystem_Strings = "SFI System" & @CRLF & "Engine" & @CRLF & "Trans"
Local $aDTC_And_Inside = [19, "P013A", "P0123", "SFI System", "Engine", "Trans", "P1225",  "Part 1", "Part 2", _
			              "P013C", "SFI System", "Part 1", "Part 2", "Trans" , "Part 1", "Part 2", "Part 3", "P0014", "SFI System", "P0000"]
Local $aDTC            = [6, "P013A", "P0123", "P1225", "P013C", "P0014", "P0000"]


   Local $aSearch_Strings [1000][20] = Get_Search_Strings ($aDTC, $aDTC_And_Inside, $sPart_Strings, $sSystem_Strings)

   For $i = 0 To 500
	  For $j = 0 to 10
		 If $aSearch_Strings [$i][$j] <> "" Then Write_Log_File ($sFilePath_YMME, "Scan DTC Config",  @CRLF & $aSearch_Strings [$i][$j], "append")

	  Next
   Next
Exit
#CE

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

;http://repair.alldata.com/alldata/navigation/treedisplay.action?vehicleId=54276&componentId=1&iTypeId=0&nonStandardId=0&fromJs=true&openUrl=#ygtvlabelel1

Func Scan_DTCs ()

   ;Gán trang web cho biến object
   Local $oIE = IECreate_Check_Error($sLink_YMME, $bWeb_Attach, $bWeb_Visible, $bWeb_Wait, $bWeb_TakeFocus)
   Sleep (1000)

   _IEAction($oIE, "refresh")
   ;------------------------------------------------------------------------------------------------------------------
   ;ĐOẠN CODE SỬ DỤNG FUNCTION "Check_Login_Alldata" ĐỂ KIỂM TRA ĐĂNG NHẬP
   If Check_Login_Alldata ($oIE) = "Not yet loged in before, this function has helped log in" Then
	  ;Reload trang DTC
	  IENavigate_Check_Error ($oIE, $sLink_YMME)
   EndIf
   ;------------------------------------------------------------------------------------------------------------------
   ;ĐOẠN CODE SỬ DỤNG FUNCTION "Check_Subscription_Alldata" ĐỂ KIỂM TRA SUBSCIPTION
   Check_Subscription_Alldata ($oIE, $sLink_YMME)
   ;------------------------------------------------------------------------------------------------------------------
   ;ĐOẠN CODE LẤY TEXT TRONG TAG <TITLE> ĐỂ KIỂM TRA XEM LINK ĐÓ CÓ PHẢI LINK YMME KHÔNG
   Local $sTxt_Title = _IEPropertyGet ($oIE, "title")
   $sTxt_Title = Standardize_String ($sTxt_Title)
   If StringInStr ($sTxt_Title, "ALLDATA Repair - Vehicle Information - ") <> 0 Then
	  ;------------------------------------
	  ;ÐOẠN CODE LẤY YMME ĐẶT TÊN CHO FOLDER
	  $sYMME = _IEPropertyGet ($oIE, "title")
	  $sYMME = StringRight ($sYMME, StringLen ($sYMME) - StringInStr ($sYMME, "-", 0, 2) - 1)
	  $sYMME = Standardize_File_Name ($sYMME)
	  Local $sFilePath_YMME        = @ScriptDir & "\INNOVA Prepair Procedures"      &"\" & $sYMME
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
		 $sConfig_SearStrings = StringRight ($sConfig, StringLen($sConfig) - StringInStr ($sConfig, "<<<-- SEARCH STRINGS -->>>") - StringLen ("<<<-- SEARCH STRINGS -->>>" & @CRLF))
		 $aConfig_SearStrings = StringSplit ($sConfig_SearStrings, @CRLF,  $STR_ENTIRESPLIT)
;		 $aSearch_Strings [0][20]
;		 Local $D1, $D2
;		 For $vElement In $aConfig_SearStrings
;			$D1 = StringMid ($vElement,

		 For $i = 1 to $aConfig_SearStrings[0]
			$sTemp = $aConfig_SearStrings[$i]
			$sTemp = StringRight ($sTemp, StringLen ($sTemp) - StringInStr ($sTemp, "]", 0 , 2) - 1)
			Open_DTC_Link ($oIE, $sTemp)
			;------------------------------------------------------------------------------------------------------------------
			;CHECK IF THE LINK ALREADY EXISTED OR NOT
			If Check_Log_File ($sYMME, "Log File DTC Successful.txt", $sLink_DTC) = "Not Exist" Then
			   DTC_Procedure_Alldata ($oIE)
			   Notification ("DONE" & @CRLF & "Please CHECK!", "Normal")
			Else ;Exist
			   Notification ("Found a DTC has been GENERATED BEFORE" & @CRLF & "Please CHECK!", "Normal")
			EndIf
		 Next
	  EndIf
   ;If the link is not YMME link
   Else
	  Notification ("The link is not Vehicle Link" & @CRLF & "Please ENTER A VEHICLE LINK!", "Normal")
   EndIf
   Return $oIE
EndFunc



Func Open_DTC_Link (Byref $oIE, $sSearch_String)
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
	  ;------------------------------------------------------------------------------------------------------------------
	  ;ĐOẠN CODE SỬ DỤNG FUNCTION "Check_Subscription_Alldata" ĐỂ KIỂM TRA SUBSCIPTION
	  Check_Subscription_Alldata ($oIE, $sLink_YMME)
	  If _IEPropertyGet ($oIE, "title") = "ALLDATA Repair - Search Results" Then
		 ;------------------------------------------------------------------------------------------------------------------
		 ;GET LINK OBJECT THEN CLICK
		 Local $oLink = _IEGetObjById ($oIE, "category_link")
		 Local $sTemp = _IEPropertyGet ($oLink,"innerhtml")
		 ;------------------------------------------------------------------------------------------------------------------
		 ;ĐOẠN CODE LẤY ID TẠO LINK
		 Local $aIDs [4]
		 For $i = 1 To 4 Step 1
			$aIDs [$i-1] =  StringMid ($sTemp, Stringinstr ($sTemp, ",", 0, $i) + 1, Stringinstr ($sTemp, ",", 0, $i + 1) - Stringinstr ($sTemp, ",", 0, $i) - 1)
		 Next
		 $sLink_DTC = "http://repair.alldata.com/alldata/article/display.action?componentId=" & $aIDs [0] & "&iTypeId=" & $aIDs [1] & "&nonStandardId=" & $aIDs [2] & "&vehicleId=" & $aIDs [3] & "&windowName=mainADOnlineWindow"
		 IENavigate_Check_Error ($oIE, $sLink_DTC)
		 ;------------------------------------------------------------------------------------------------------------------
		 ;ĐOẠN CODE SỬ DỤNG FUNCTION "Check_Subscription_Alldata" ĐỂ KIỂM TRA SUBSCIPTION
		 Check_Subscription_Alldata ($oIE, $sLink_DTC)
		 ExitLoop
	  EndIf
   Until 0
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
   ;------------------------------------------------------------------------------------------------------------------
   ;ĐOẠN CODE SỬ DỤNG FUNCTION "Check_Login_Alldata" ĐỂ KIỂM TRA ĐĂNG NHẬP
   If Check_Login_Alldata ($oIE) = "Not yet loged in before, this function has helped log in" Then
	  ;Reload trang DTC
	  IENavigate_Check_Error ($oIE, $sLink_YMME)
   EndIf
   ;------------------------------------------------------------------------------------------------------------------
   ;ĐOẠN CODE SỬ DỤNG FUNCTION "Check_Subscription_Alldata" ĐỂ KIỂM TRA SUBSCIPTION
   Check_Subscription_Alldata ($oIE, $sLink_YMME)
   ;------------------------------------------------------------------------------------------------------------------
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
	  Local $sFilePath_Alldata_DTC = @ScriptDir & "\INNOVA Prepair Procedures"
	  If FileExists ($sFilePath_Alldata_DTC) = 0 Then	DirCreate($sFilePath_Alldata_DTC)
	  Local $sFilePath_YMME        = @ScriptDir & "\INNOVA Prepair Procedures"      &"\"&$sYMME
	  If FileExists ($sFilePath_YMME) = 0 Then	DirCreate($sFilePath_YMME)
	  ;------------------------------------
	  Notification ("Expanding The Tree View ...","Normal")
	  Local $sID_Sub = ""
	  ;-----------------------------------
	  ;CLICK TREE VIEW LEVEL 1
	  $sID_Sub = Click_Tree_View_By_Text ($oIE, $sID_Sub, "A L L Diagnostic Trouble Codes ( DTC )", "First")
	  ;-----------------------------------
	  ;CLICK TREE VIEW LEVEL 2
	  $sID_Sub = Click_Tree_View_By_Text ($oIE, $sID_Sub, "Information for A L L", "Not First")
	  ;-----------------------------------
	  ;CLICK TREE VIEW LEVEL 3
	  $sID_Sub = Click_Tree_View_By_Text ($oIE, $sID_Sub, "Testing and Inspection", "Not First")
	  ;-----------------------------------
	  ;CLICK TREE VIEW LEVEL P CODES
	  $sID_Sub_PCode = Click_Tree_View_By_Text ($oIE, $sID_Sub, "P Code", "Not First")
	  ;-----------------------------------
	  ;CLICK ALL ELEMENT IN THE SUB OBJECT
	  Do ;Loop until get the SubOject
		 ;Get Object by ID
		 $oIE_SubObject = _IEGetObjById ($oIE, $sID_Sub_PCode)
		 Sleep (200)
	  Until @error = 0
	  ;Lấy DTC lưu vào mảng
	  Local $aDTC = Innertext2Array ($oIE_SubObject)
	  ;Array to save elements have already expanded
	  Local $Array_All [5000]
	  Local $j = 0
	  Do ;Loop until expand all elements
		 Local $oTags = _IETagNameGetCollection($oIE_SubObject, "td")
		 Local $txt = ""
		 Local $i = 0
		 ;Collect ID then click
		 For $oTag In $oTags
			If StringInStr ($oTag.id, "ygtvt") <> 0 Then
			   If _ArraySearch ($Array_All, $oTag.id) = -1 Then
				  ;Click object
				  _IEAction($oTag, "click")
				  $Array_All [$j] = $oTag.id
				  $txt &= $oTag.id & @CRLF
				  $j += 1
				  $i += 1
			   EndIf
			EndIf
		 Next
	  Until $i = 0
	  ;Lấy DTC, Systems và Parts luu vào mảng
	  Local $aDTC_And_Inside = Innertext2Array ($oIE_SubObject)
	  ;--------------------------
	  ;GET SYSTEMS AND PARTS STRINGS
	  Local $aSystems_And_Parts = Array_Minus ($aDTC_And_Inside, $aDTC)
	  Local $sSystems_And_Parts = _ArrayToString ($aSystems_And_Parts, @CRLF)
	  ;--------------------------
	  ;GET SYSTEMS STRINGS AND PARTS STRINGS IN 2 DIFF VAR
	  $i = 0
	  $j = 0
	  Local $aPart_Strings [0]
	  For $i = 0 To UBound($aSystems_And_Parts) - 1 Step 1
		 If StringInStr ($aSystems_And_Parts[$i], "Part ") <> 0 Then
			ReDim $aPart_Strings [UBound ($aPart_Strings) + 1]
			$aPart_Strings[$j] = $aSystems_And_Parts[$i]
			$j += 1
		 EndIf
	  Next
	  Local $aSystem_Strings = Array_Minus ($aSystems_And_Parts, $aPart_Strings)
	  Local $sPart_Strings = _ArrayToString ($aPart_Strings, @CRLF)
	  Local $sSystem_Strings = _ArrayToString ($aSystem_Strings, @CRLF)
	  ;--------------------------
	  Notification ("Writing Config file for: " & @CRLF & $sYMME,"Normal")
	  Sleep (1000)
	  ;WRITE CONFIG FILE
	  Write_Log_File ($sFilePath_YMME, "Scan DTC Config",  "This config file is to save System Strings, Part Strings and Search Strings" & @CRLF & "for the soft to get all DTC links of the vehicle", "overwrite")
	  Write_Log_File ($sFilePath_YMME, "Scan DTC Config",  @CRLF & @CRLF & "<<<-- SYSTEM STRINGS -->>>" & @CRLF & $sSystem_Strings, "append")
	  Write_Log_File ($sFilePath_YMME, "Scan DTC Config",  @CRLF & @CRLF & "<<<-- PART STRINGS -->>>" & @CRLF & $sPart_Strings, "append")
	  Write_Log_File ($sFilePath_YMME, "Scan DTC Config",  @CRLF & @CRLF & "<<<-- SEARCH STRINGS -->>>", "append")
	  Local $aSearch_Strings [1000][20] = Get_Search_Strings ($aDTC, $aDTC_And_Inside, $sPart_Strings, $sSystem_Strings)
	  For $i = 0 To 500
		 For $j = 0 to 10
			If $aSearch_Strings [$i][$j] <> "" Then
			   Write_Log_File ($sFilePath_YMME, "Scan DTC Config",  @CRLF & "[" & $i & "]" & "[" & $j & "]" & $aSearch_Strings [$i][$j], "append")
			EndIf
		 Next
	  Next
	  ;-------------------------
	  ;Complete writing
	  Notification ("Completed writing config file for """ & $sYMME & """" & @CRLF & "Please CHECK!", "Normal")
   Else
	  Notification ("The link is not Vehicle Link" & @CRLF & "Please ENTER A VEHICLE LINK!", "Normal")
   EndIf
   Return $oIE
EndFunc




;====================================================================================================================
;                  FUNCTION DESCRIPTION: GET THE SEARCH STRING ARRAY
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func Get_Search_Strings ($aDTC, $aDTC_And_Inside, $sPart_Strings, $sSystem_Strings)
   ;Mảng lưu search string
   Local $aSearch_Strings [1000][20]
   ;Các biến i, j dùng cho vòng lặp, D1, D2 để xác định dimension cho Search_String array
   Local $i = 0, $j = 0, $D1 = 0, $D2 = 0
   ;Các biến để xác định vị trí trên và dưới một đoạn text trong mảng $aDTC_And_Inside
   Local $sLowPos = 0
   Local $sHighPos = 0
   ;Gắn 2 giá trị End cuối mảng End
   ReDim $aDTC[UBound($aDTC) + 1]
   $aDTC [UBound($aDTC)-1] = "P0000"
   ReDim $aDTC_And_Inside[UBound($aDTC_And_Inside) + 1]
   $aDTC_And_Inside [UBound($aDTC_And_Inside)-1] = "P0000"
   ;Xét mảng $aDTC
   For $i = 0 To UBound($aDTC) - 2 Step 1
	  ;Lấy vị trí DTC dưới và trên
	  $sLowPos = _ArraySearch ($aDTC_And_Inside, $aDTC [$i], $sLowPos)
	  $sHighPos = _ArraySearch ($aDTC_And_Inside, $aDTC [$i+1], $sHighPos)
	  ;Xem thử các phần tử của mảng trong đoạn vị trí dưới đến trên có chứa Part hay system không
	  Local $bSystem_Flag = False
	  Local $bPart_Flag = False
	  For $j = $sLowPos To $sHighPos Step 1
		 If StringInStr ($sSystem_Strings, $aDTC_And_Inside [$j]) <> 0 Then $bSystem_Flag = True
		 If StringInStr ($sPart_Strings, $aDTC_And_Inside [$j]) <> 0 Then $bPart_Flag = True
	  Next
	  ;Xét 4 trường hợp của đoạn text trả về
	  Switch $bSystem_Flag & " " & $bPart_Flag
		 Case "False False"
			$aSearch_Strings [$D1][$D2] = $aDTC_And_Inside [$sLowPos]
			   $D1 += 1
			   $D2 = 0
		 Case "True False"
			For $j = $sLowPos + 1 To $sHighPos - 1 Step 1
			   $aSearch_Strings [$D1][$D2] = $aDTC_And_Inside [$sLowPos] & " >> " & $aDTC_And_Inside [$j]
				  $D1 += 1
				  $D2 = 0
			Next
		 Case "False True"
			For $j = $sLowPos + 1 To $sHighPos - 1 Step 1
			   $aSearch_Strings [$D1][$D2] = $aDTC_And_Inside [$sLowPos] & " >> " & $aDTC_And_Inside [$j]
				  $D1 += 0
				  $D2 += 1
			Next
			$D1 += 1
			$D2 = 0
		 Case "True True"
			$D1 -= 1
			For $j = $sLowPos + 1 To $sHighPos - 1 Step 1
			   If StringInStr ($sSystem_Strings, $aDTC_And_Inside [$j]) <> 0 Then
				  Local $sTemp = $aDTC_And_Inside [$sLowPos] & " >> " & $aDTC_And_Inside [$j]
					 $D1 += 1
					 $D2 = 0
			   EndIf
			   If StringInStr ($sPart_Strings, $aDTC_And_Inside [$j]) <> 0 Then
				  $aSearch_Strings [$D1][$D2] = $sTemp & " >> " & $aDTC_And_Inside [$j]
				  $D2 +=1
			   EndIf
			Next
			$D1 += 1
			$D2 = 0
	  EndSwitch
   Next
   Return $aSearch_Strings
EndFunc


;====================================================================================================================
;                  FUNCTION DESCRIPTION: CLICK TREE VEIEW BY TEXT
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func Click_Tree_View_By_Text (Byref $oIE, $sID_Sub, $sTxt, $sMode)
   If $sMode = "First" Then
	  ;Wait for the text to appear
	  Do ;Loop until the string appear
		 Sleep (500)
		 Local $sHTML = _IEPropertyGet ($oIE, "outertext")
	  Until StringInStr ($sHTML, "A L L Diagnostic Trouble Codes ( DTC )") <> 0
	  Sleep (1000)
	  ;Get object
	  $oIE_SubObject = $oIE
	  ;Get ID by text
	  $sID_Click = GetIDByText ($oIE, "A L L")
	  ;Get ID from the previous $sID_Click
	  $sID_Sub = StringReplace ($sID_Click, "vt", "vc")
	  ;Get Object by ID
	  $oTreeView = _IEGetObjById ($oIE, $sID_Click)
	  ;Click object
	  _IEAction($oTreeView, "click")
   Else
	  ;Wait for the text to appear
	  Do ;Loop until get the SubOject
		 ;Get Object by ID
		 $oIE_SubObject = _IEGetObjById ($oIE, $sID_Sub)
		 Sleep (200)
	  Until @error = 0
	  ;Get ID by text
	  $sID_Click = GetIDByText ($oIE_SubObject, $sTxt)
	  $sID_Sub = StringReplace ($sID_Click, "vt", "vc")
	  ;Get Object by ID
	  $oTreeView = _IEGetObjById ($oIE, $sID_Click)
	  ;Click object
	  _IEAction($oTreeView, "click")
   EndIf
   Return $sID_Sub
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
;                  FUNCTION DESCRIPTION: GET TEXT IN THE SUB OBJECT AND PUT INTO AN ARRAY
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func Innertext2Array ($oIE_SubObject)
   $sText = _IEPropertyGet ($oIE_SubObject, "innertext")
   ;Gắn 1 cái xuống dòng phía cuối để xử lý string thừa
   $sText &= @CRLF
   ;-----------------------------------
   ;REMOVE REDUNDANT LINES
   Do ;Loop until replace all "Space + CRLF" = CRLF
	  $sText = StringReplace ($sText, " " & @CRLF, @CRLF)
   Until StringInStr ($sText, " " & @CRLF) = 0
   Do ;Loop until replace all "CRLF + Space" = CRLF
	  $sText = StringReplace ($sText,@CRLF &  " ", @CRLF)
   Until StringInStr ($sText, " " & @CRLF) = 0
   Do ;Loop until replace all 2xCRLF = 1xCRLF
	  $sText = StringReplace ($sText, @CRLF & @CRLF, @CRLF)
   Until StringInStr ($sText, @CRLF & @CRLF) = 0
   If StringLeft ($sText, 2) = @CRLF Then $sText = StringRight ($sText, StringLen ($sText) - 2)
   If StringRight ($sText, 2) = @CRLF Then $sText = StringLeft ($sText, StringLen ($sText) - 2)
   ;-----------------------------------
   ;STORE THE STRING INTO AN ARRAY
   Local $aText = StringSplit($sText, @CRLF, $STR_ENTIRESPLIT )
   _ArrayDelete ($aText, 0)
   Return $aText
EndFunc


;====================================================================================================================
;                  FUNCTION DESCRIPTION: GET ID OF AN ELEMENT IN TREE VIEW BY TEXT
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func GetIDByText ($oIE, $sTxt)
   ;----------------------------------------------------
   Local $oTags = _IETagNameGetCollection($oIE, "td")
   Local $Array [5000]
   Local $i = 0, $iMark = 0
   ;Loop until found the string
   Do
	  For $oTag In $oTags
		 $Array [$i] = $oTag
		 If StringInStr ($oTag.innertext, $sTxt) <> 0 Then
			$iMark = $i
		 EndIf
		 $i += 1
	  Next
   Until $iMark <> 0

   $iMark = $iMark - 1
   $oTemp = $Array [$iMark]
   Return $oTemp.id
EndFunc




#cs
   ;ASK USER TO HELP ON PARTS STRINGS
   _GUICtrlEdit_SetReadOnly ($Commu_Ctrl, False)
   Notification ("NOTE: Please help delete all the strings below" & @CRLF & "which are not PARTS STRINGS" & @CRLF & "then Type ""EXECUTE"" and Press ""ENTER"""  & @CRLF & "<<<----------->>>" & @CRLF & $sSystems_And_Parts & @CRLF & "<<<----------->>>", "Normal")
   While _GUICtrlEdit_GetLine ($Commu_Ctrl, _GUICtrlEdit_GetLineCount ($Commu_Ctrl) - 2) <> "EXECUTE"
	  Sleep (100)
   WEnd
   _GUICtrlEdit_SetReadOnly ($Commu_Ctrl, True)
   Notification ("You have typed EXECUTE" & @CRLF & "Please wait for the App to get config file", "Normal")

   Local $sTemp = _GUICtrlEdit_GetText ($Commu_Ctrl)
   Local $sPart_Strings = StringMid ($sTemp, StringInStr ($sTemp, ">>>", 0 , -2) + 5, StringInStr ($sTemp, "<<<", 0 , -1) - StringInStr ($sTemp, ">>>", 0 , -2) - 7)

   Local $aPart_Strings = StringSplit ($sPart_Strings, @CRLF,  $STR_ENTIRESPLIT)

   Local $aSystem_Strings = Array_Minus ($aSystems_And_Parts, $aPart_Strings)
   Local $sSystem_Strings = _ArrayToString ($aSystem_Strings, @CRLF)
#ce