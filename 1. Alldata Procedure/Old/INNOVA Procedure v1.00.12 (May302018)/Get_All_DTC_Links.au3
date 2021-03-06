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

HotKeySet ("{ESC}", "Autoit_Exit")

$sLink_YMME = "http://repair.alldata.com/alldata/navigation/treedisplay.action?vehicleId=54276&componentId=1&iTypeId=0&nonStandardId=0&fromJs=true&openUrl=#ygtvlabelel1"


Func Get_All_DTC_Links ()

   Local $oIE = _IECreate ($sLink_YMME)
   Sleep (1000)
   _IEAction($oIE, "refresh")
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
   Local $aDTC [1000]
   $aDTC = Innertext2Array ($oIE_SubObject)
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
   Local $aDTC_And_Inside [1000]
   $aDTC_And_Inside = Innertext2Array ($oIE_SubObject)
   ;--------------------------
   ;GET SYSTEMS AND PARTS STRINGS
   Local $aSystems_And_Parts [1000]
   $aSystems_And_Parts  = Array_Minus ($aDTC_And_Inside, $aDTC)
   $sSystems_And_Parts = _ArrayToString ($aSystems_And_Parts, @CRLF)
   ;--------------------------
   $i = 0
   $j = 0
   Local $aPart_Strings [1000]
   For $i = 0 To UBound($aSystems_And_Parts) - 1 Step 1
	  If StringInStr ($aSystems_And_Parts[$i], "Part ") <> 0 Then
		 $aPart_Strings[$j] = $aSystems_And_Parts[$i]
		 $j += 1
	  EndIf
	  $i += 1
   Next
   Local $aSystem_Strings = Array_Minus ($aSystems_And_Parts, $aPart_Strings)
   Local $sPart_Strings = _ArrayToString ($aPart_Strings, @CRLF)
   Local $sSystem_Strings = _ArrayToString ($aSystem_Strings, @CRLF)




   Write_Log_File ($sFilePath_YMME, "Scan DTC Config",  "This config file is to save System Strings, Part Strings and Search Strings", "overwrite")
   Write_Log_File ($sFilePath_YMME, "Scan DTC Config",  @CRLF & @CRLF & "<<<-- SYSTEM STRINGS -->>>" & @CRLF & $sSystem_Strings, "append")
   Write_Log_File ($sFilePath_YMME, "Scan DTC Config",  @CRLF & @CRLF & "<<<-- PART STRINGS -->>>" & @CRLF & $sPart_Strings, "append")
   Write_Log_File ($sFilePath_YMME, "Scan DTC Config",  @CRLF & @CRLF & "<<<-- SEARCH STRINGS -->>>" & @CRLF, "append")
   Local $aSearch_Strings [1000][20] = Get_Search_Strings ($aDTC, $aDTC_And_Inside, $sPart_Strings, $sSystem_Strings)
   For $i = 0 To 500
	  For $j = 0 to 10
		 If $aSearch_Strings [$i][$j] <> "" Then
			;MsgBox (0, "", $aSearch_Strings [$i][$j])
			Write_Log_File ($sFilePath_YMME, "Scan DTC Config",  @CRLF & $aSearch_Strings [$i][$j], "append")
		 EndIf
	  Next
   Next
   Exit
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
   ;-----------------------------------
   ;REMOVE REDUNDANT LINES
   Do ;Loop until replace all "Space + CRLF" = CRLF
	  $sText = StringReplace ($sText, " " & @CRLF, @CRLF)
   Until StringInStr ($sText, " " & @CRLF) = 0
   Do ;Loop until replace all 2xCRLF = 1xCRLF
	  $sText = StringReplace ($sText, @CRLF & @CRLF, @CRLF)
   Until StringInStr ($sText, @CRLF & @CRLF) = 0
   If StringLeft ($sText, 2) = @CRLF Then $sText = StringRight ($sText, StringLen ($sText) - 2)
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