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

Func Autoit_Exit ()
   Exit
EndFunc



;Local $sPart_Strings = "Part 1" & @CRLF & "Part 2" & @CRLF & "Part 3"
;Local $sSystem_Strings = "SFI System" & @CRLF & "Engine" & @CRLF & "Trans"
;Local $aDTC_And_Inside = [19, "P013A", "P0123", "SFI System", "Engine", "Trans", "P1225",  "Part 1", "Part 2", _
;			              "P013C", "SFI System", "Part 1", "Part 2", "Trans" , "Part 1", "Part 2", "Part 3", "P0014", "SFI System", "P0000"]
;Local $aDTC            = [6, "P013A", "P0123", "P1225", "P013C", "P0014", "P0000"]


;Local $sLink = "http://repair.alldata.com/alldata/navigation/treedisplay.action?vehicleId=47006&componentId=0&iTypeId=0&nonStandardId=0"
Local $sLink = "http://repair.alldata.com/alldata/navigation/treedisplay.action?vehicleId=54276&componentId=1&iTypeId=0&nonStandardId=0&fromJs=true&openUrl=#ygtvlabelel1"
Local $oIE = _IECreate ($sLink)
Sleep (1000)
_IEAction($oIE, "refresh")



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
Local $aDTC_And_Inside = Innertext2Array ($oIE_SubObject)






Local $aSystems_And_Parts
$aSystems_And_Parts  = Get_Systems_And_Parts ($aDTC, $aDTC_And_Inside)
$sSystems_And_Parts = _ArrayToString ($aSystems_And_Parts, @CRLF)

MsgBox (0, "", $sSystems_And_Parts)
Exit





Local $sSearch_Strings = Get_Search_Strings ($aDTC, $aDTC_And_Inside, $sPart_Strings, $sSystem_Strings)


For $i = 0 To 10
   For $j = 0 to 10
	  If $sSearch_Strings [$i][$j] <> "" Then MsgBox (0, "", "Search String: " & $i & " " & $j & "  " & $sSearch_Strings [$i][$j])
   Next
Next

Exit



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
Func Get_Systems_And_Parts ($aDTC, $aDTC_And_Inside)
   Local $aSystems_And_Parts = $aDTC_And_Inside
   For $vElement In $aDTC
	  Local $Pos = _ArraySearch ($aSystems_And_Parts, $vElement)
	  _ArrayDelete ($aSystems_And_Parts, $Pos)
   Next
   _ArraySort ($aSystems_And_Parts)
   $aSystems_And_Parts = _ArrayUnique ($aSystems_And_Parts, 0, 0, 0, $ARRAYUNIQUE_NOCOUNT)
   Return $aSystems_And_Parts
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
;                  FUNCTION DESCRIPTION: GET THE SEARCH STRING ARRAY
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func Get_Search_Strings ($aDTC, $aDTC_And_Inside, $sPart_Strings, $sSystem_Strings)
   ;Mảng lưu search string
   Local $sSearch_Strings [100][100]
   ;Các biến i, j dùng cho vòng lặp, D1, D2 để xác định dimension cho Search_String array
   Local $i = 0, $j = 0, $D1 = 0, $D2 = 0
   ;Các biến để xác định vị trí trên và dưới một đoạn text trong mảng $aDTC_And_Inside
   Local $sLowPos
   Local $sHighPos
   ;Xét mảng $aDTC
   For $i = 1 To $aDTC [0] - 1 Step 1
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
			$sSearch_Strings [$D1][$D2] = $aDTC_And_Inside [$sLowPos]
			   $D1 += 1
			   $D2 = 0
		 Case "True False"
			For $j = $sLowPos + 1 To $sHighPos - 1 Step 1
			   $sSearch_Strings [$D1][$D2] = $aDTC_And_Inside [$sLowPos] & " >> " & $aDTC_And_Inside [$j]
				  $D1 += 1
				  $D2 = 0
			Next
		 Case "False True"
			For $j = $sLowPos + 1 To $sHighPos - 1 Step 1
			   $sSearch_Strings [$D1][$D2] = $aDTC_And_Inside [$sLowPos] & " >> " & $aDTC_And_Inside [$j]
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
				  $sSearch_Strings [$D1][$D2] = $sTemp & " >> " & $aDTC_And_Inside [$j]
				  $D2 +=1
			   EndIf
			Next
			$D1 += 1
			$D2 = 0
	  EndSwitch
   Next
   Return $sSearch_Strings
EndFunc











;====================================================================================================================
;                  FUNCTION DESCRIPTION: GET ID OF AN ELEMENT IN TREE VIEW BY TEXT
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func GetIDByText ($oIE, $sTxt)
   ;----------------------------------------------------
   Local $oTags = _IETagNameGetCollection($oIE, "td")
   Local $Array [1000]
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
Exit