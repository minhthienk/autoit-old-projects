
#include <MsgBoxConstants.au3>
#include <FileConstants.au3>
#include <InetConstants.au3>

#include <Clipboard.au3>
#include <IE.au3>
#include <Excel.au3>
#include <WinAPIFiles.au3>

#include "General_Library.au3"
#include "Create_JAVASCRIPT_Procedure.au3"
#include "Create_NORMAL_Procedure.au3"

;====================================================================================================================
;                  FUNCTION DISCRIPTION: MAIN FUNCTION
;====================================================================================================================
Func Main_function_DTC ()
   ;Gán trang web cho biến object
   Local $oIE = IECreate_Check_Error($sLink_DTC, $bWeb_Attach, $bWeb_Visible, $bWeb_Wait, $bWeb_TakeFocus)
   ;------------------------------------------------------------------------------------------------------------------
   ;ĐOẠN CODE SỬ DỤNG FUNCTION "Check_Login_Alldata" ĐỂ KIỂM TRA ĐĂNG NHẬP
   If Check_Login_Alldata ($oIE) = "Not yet loged in before, this function has helped log in" Then
	  ;Reload trang DTC
	  IENavigate_Check_Error ($oIE, $sLink_DTC)
   EndIf
   ;------------------------------------------------------------------------------------------------------------------
   ;ĐOẠN CODE SỬ DỤNG FUNCTION "Check_Subscription_Alldata" ĐỂ KIỂM TRA SUBSCIPTION
   Check_Subscription_Alldata ($oIE, $sLink_DTC)
   ;------------------------------------------------------------------------------------------------------------------
   ;ĐOẠN CODE LẤY TEXT TRONG TAG <TITLE> ĐỂ KIỂM TRA XEM LINK ĐÓ CÓ PHẢI LINK DTC HAY KHÔNG
   Local $sTxt_Title = _IEPropertyGet ($oIE, "title")
   $sTxt_Title = Standardize_String ($sTxt_Title)
   If StringInStr ($sTxt_Title, "A L L Diagnostic Trouble Codes ( DTC ) |Testing and Inspection") <> 0 And StringInStr($sTxt_Title, "Code Charts:") <> 0 Then
	  ;------------------------------------------------------------------------------------------------------------------
	  ;CHECK IF THE LINK ALREADY EXISTED OR NOT
	  Local $sYMME = Get_YMME ($oIE)
	  If Check_Log_File ($sYMME, "Log File DTC Successful.txt", $sLink_DTC) = "Not Exist" Then
		 DTC_Procedure_Alldata ($oIE)
		 Notification ("DONE" & @CRLF & "Please CHECK!", "Normal")
	  Else ;Exist
		 Notification ("Found a DTC has been GENERATED BEFORE" & @CRLF & "Please CHECK!", "Normal")
	  EndIf
   Else
	  Notification ("The link is not DTC Link" & @CRLF & "Please ENTER A DTC LINK!", "Normal")
   EndIf
   Return $oIE
EndFunc




;====================================================================================================================
;                  FUNCTION DISCRIPTION: CHECK IF A LINK ALREADY EXISTED IN A LOG FILE
;====================================================================================================================
Func Check_Log_File ($sYMME, $sFileName,$sLink)
   Local $sResult = ""
   Local $sFilePath_YMME = @ScriptDir & "\INNOVA Prepair Procedures" & "\"& $sYMME
   Local $sFilePath_DTCLog = $sFilePath_YMME & "\" & $sFileName
   If FileExists ($sFilePath_DTCLog) Then
	  Local $hFileOpen = FileOpen($sFilePath_DTCLog, $FO_READ)
	  Local $sFileContent = FileRead ($hFileOpen)
	  FileClose($hFileOpen)
	  If StringInStr ($sFileContent, $sLink) <> 0 Then
		 $sResult = "Exist"
	  Else
		 $sResult = "Not Exist"
	  EndIf
   Else
	  $sResult = "Not Exist"
   EndIf
   Return $sResult
EndFunc





;====================================================================================================================
;                  FUNCTION DISCRIPTION: CREATE DTC PROCEDURE FROM ALLDATA
;====================================================================================================================
Func DTC_Procedure_Alldata ($oIE)
   ;Chỉnh sửa text trong Procedure của All data cho phù hợp với Innova
   $sTxt_Body = Modify_Body_HTML ($oIE)
   ;------------------------------------------------------------------------------------------------------------------
   ;ĐOẠN CODE LẤY TEXT TRONG TAG <TITLE> ĐỂ ĐẶT TÊN CHO FOLDER, TITLE HTML VÀ FILE HTML
   Local $sTxt_Title_Alldata = _IEPropertyGet ($oIE, "title")
   Local $sTxt_Title = $sTxt_Title_Alldata
   $sTxt_Title = Standardize_String ($sTxt_Title)
   ;Code lấy DTC Code trong $sTxt_Title  để làm title cho file html và folder name cho DTC Procedure
   $sTxt_Title = StringMid ($sTxt_Title, StringInStr($sTxt_Title, "Code Charts: ") + StringLen("Code Charts: "), 5)
   ;------------------------------------------------------------------------------------------------------------------
   ;ĐOẠN CODE LẤY TEXT ĐỂ LẤY TÊN XE ĐẶT TÊN CHO FOLDER
   Local $sYMME = Get_YMME ($oIE)
   $sYMME = Standardize_File_Name ($sYMME)
   Notification ("Begin to generate DTC: " & $sTxt_Title & @CRLF & " of " & $sYMME, "Normal")
   Sleep ("1000")
   ;------------------------------------------------------------------------------------------------------------------
   ;ĐOẠN CODE TẠO THƯ MỤC VÀ TẢI HÌNH ẢNH VỀ THƯ MỤC ĐÓ
   ;Tạo các thư mục cần thiết
   Local $sFilePath_Alldata_DTC = @ScriptDir & "\INNOVA Prepair Procedures"
   If FileExists ($sFilePath_Alldata_DTC) = 0 Then	DirCreate($sFilePath_Alldata_DTC)
   Local $sFilePath_YMME        = @ScriptDir & "\INNOVA Prepair Procedures"      &"\"&$sYMME
   If FileExists ($sFilePath_YMME) = 0 Then	DirCreate($sFilePath_YMME)
   Local $sFilePath_DTC         = @ScriptDir & "\INNOVA Prepair Procedures"      &"\"&$sYMME      &"\DTC"
   If FileExists ($sFilePath_DTC) = 0 Then DirCreate($sFilePath_DTC)
   Local $sFilePath_Title       = @ScriptDir & "\INNOVA Prepair Procedures"      &"\"&$sYMME      &"\DTC"       &"\"& $sTxt_Title
   If FileExists ($sFilePath_Title) = 0 Then	DirCreate($sFilePath_Title)
   ;Tải hình ảnh của procedure vào thư mục
   Download_Procedure_Image_Collection ($sFilePath_Title, $oIE)
   ;------------------------------------------------------------------------------------------------------------------
   ;ĐOẠN CODE INSERT HÌNH ẢNH CHO PROCEDURE
   $sTxt_Body = Insert_Images_HTML ($oIE, $sTxt_Body)
   Notification ("Downloaded all images for DTC: " & $sTxt_Title & @CRLF & "of " & $sYMME & @CRLF & @CRLF & "Waiting for the next process ...", "Normal")
   ;------------------------------------------------------------------------------------------------------------------
   ;ĐOẠN CODE LẤY TEXT VÀ LINK PROCEDURE TRONG TAG <A>
   Local $oAs = _IETagNameGetCollection($oIE, "a")
   Local $aHyperlink_Innertexts [1000]
   Local $aHyperlink_Links [1000]
   Local $iCount_Hyperlink = 0
   For $oA In $oAs
	  If StringInStr ($oA.innertext, "See:", 0, 1) <> 0 Then
		 $aHyperlink_Innertexts [$iCount_Hyperlink] = $oA.innertext
		 $aHyperlink_Links [$iCount_Hyperlink] = $oA.href
		 $iCount_Hyperlink = $iCount_Hyperlink + 1
	  EndIf
   Next


   ;------------------------------------------------------------------------------------------------------------------
   ;ĐOẠN CODE CREATE PROCEDURE VÀ TRẢ VỀ MỘT STRING MANG ĐƯỜNG DẪN TỚI FILE HTML CỦA PROCEDURE
   $iCount_Hyperlink = 0
   While $aHyperlink_Links [$iCount_Hyperlink] <> ""
	  ;------------------------------------------------------------------------------------------------------------------
	  ;CHECK IF THE LINK ALREADY EXISTED OR NOT
	  If Check_Log_File ($sYMME, "Log File Procedure Successful.txt", $aHyperlink_Links [$iCount_Hyperlink]) = "Not Exist" Then
		 ;Xem thử trong đường link có text "repair.alldata.com" hay không, để tránh trường hợp xuất hiện "javascript"
		 If StringInStr ($aHyperlink_Links [$iCount_Hyperlink], "repair.alldata.com", 0, 1) <> 0 Then
			Local $sProcedure_Path = Create_NORMAL_Procedure ($oIE, $sFilePath_YMME, $aHyperlink_Links [$iCount_Hyperlink])
			Local $sHTML_Procedure_Hyperlink = "<a href=""" & $sProcedure_Path & """ target=""_blank"">" & "(More info)" & "</a>"
			$sTxt_Body = StringReplace ($sTxt_Body, Standardize_String ($aHyperlink_Innertexts [$iCount_Hyperlink]),$sHTML_Procedure_Hyperlink, 1, 0)
		 Else
			;ĐOẠN CODE XỬ LÝ LINK PROCEDURE CHỨA JAVASCRIPT
			;Lấy info để xử lý lỗi procedure sau này
			Local $sInfo = $sYMME & @CRLF & $sTxt_Title & @CRLF & "Procedure number: " & ($iCount_Hyperlink + 1) & @CRLF & $aHyperlink_Links [$iCount_Hyperlink]
			Local $sProcedure_Path = Create_JAVASCRIPT_Procedure ($oIE, $sFilePath_YMME, $aHyperlink_Links [$iCount_Hyperlink], $sInfo)
			Local $sHTML_Procedure_Hyperlink = "<a href=""" & $sProcedure_Path & """ target=""_blank"">" & "(More info)" & "</a>"
			$sTxt_Body = StringReplace ($sTxt_Body, Standardize_String ($aHyperlink_Innertexts [$iCount_Hyperlink]),$sHTML_Procedure_Hyperlink, 1, 0)
		 EndIf
	  Else
		 Local $sProcedure_Link = $aHyperlink_Links [$iCount_Hyperlink]
		 Local $sTxt_File_Name = Create_Procedure_Name ($sProcedure_Link)
		 Local $sProcedure_Path = "../../PROCEDURE/" & $sTxt_File_Name & "/" & $sTxt_File_Name & ".html"
		 Local $sHTML_Procedure_Hyperlink = "<a href=""" & $sProcedure_Path & """ target=""_blank"">" & "(More info)" & "</a>"
		 $sTxt_Body = StringReplace ($sTxt_Body, Standardize_String ($aHyperlink_Innertexts [$iCount_Hyperlink]),$sHTML_Procedure_Hyperlink, 1, 0)
		 Notification ("Found A PROCEDURE has already been generated" & @CRLF & "Waiting for the next process ...", "Normal")
	  EndIf
	  $iCount_Hyperlink = $iCount_Hyperlink + 1
   WEnd

   ;------------------------------------------------------------------------------------------------------------------
   Create_HTML  ($sFilePath_Title, $sTxt_Title, $sTxt_Title, $sTxt_Body)
   Local $sLog_Txt = $sTxt_Title_Alldata & @CRLF & $sLink_DTC
   Write_Log_File ($sFilePath_YMME, "Log File DTC Successful",  @CRLF & @CRLF & @CRLF & $sLog_Txt, "append")
EndFunc



Func Get_YMME (Byref $oIE)
   ;------------------------------------------------------------------------------------------------------------------
   ;ĐOẠN CODE LẤY TEXT TRONG TAG SPAN ĐỂ LẤY TÊN XE ĐẶT TÊN CHO FOLDER
   Local $oSpans = _IETagNameGetCollection($oIE, "span")
   Local $aSpans [1000]
   Local $iCount = 0
   Local $iMark = 0
   For $oSpan In $oSpans
	  $aSpans [$iCount] = $oSpan.innertext
	  If $aSpans [$iCount] = "Save Article " Then $iMark = $iCount
	  $iCount = $iCount + 1
   Next
   Local $sYMME = $aSpans [$iMark - 1]
   $sYMME = Standardize_File_Name ($sYMME)
   Return $sYMME
EndFunc


