#cs ----------------------------------------------------------------------------
NOTE:
#ce ----------------------------------------------------------------------------


#include-once

#include <MsgBoxConstants.au3>
#include <FileConstants.au3>
#include <InetConstants.au3>

#include <Clipboard.au3>
#include <IE.au3>
#include <Excel.au3>
#include <WinAPIFiles.au3>


;====================================================================================================================
;                  FUNCTION DISCRIPTION: CREATE OTHER PROCEDURE IN DTC FROM ALLDATA
;				   RETURN              : A STRING OF PROCEDURE PATH
;====================================================================================================================
Func Create_NORMAL_Procedure ($oIE, $sFilePath_YMME, $sProcedure_Link)
   _IENavigate ($oIE, $sProcedure_Link)
   ;ĐOẠN CODE SỬ DỤNG FUNCTION "Check_Subscription_Alldata" ĐỂ KIỂM TRA SUBSCIPTION
   Check_Subscription_Alldata ($oIE)
   ;------------------------------------------------------------------------------------------------------------------
   ;ĐOẠN CODE LẤY TEXT TRONG TAG <TITLE> ĐỂ LÀM TITLE CHO HTML PROCEDURE VÀ KIỂM TRA XEM LINK ĐÓ CÓ PHẢI LINK DTC HAY KHÔNG
   Local $sTxt_Title_Alldata = _IEPropertyGet ($oIE, "title")
   Local $sTxt_Title = $sTxt_Title_Alldata
   $sTxt_Title = Standardize_String ($sTxt_Title)
   ;------------------------------------------------------------------------------------------------------------------
   ;ĐOẠN CODE KIỂM TRA NẾU LÀ LINK DTC THÌ TRẢ VỀ PATH HTML CỦA DTC, NẾU LÀ LINK PROCEDURE THÌ LÀM PROCEDURE VÀ TRẢ VỀ PATH CỦA PROCEDURE
   ;Code lấy DTC Code trong $sTxt_Title  để làm đường dẫn
   If StringInStr ($sTxt_Title, "A L L Diagnostic Trouble Codes ( DTC ) |Testing and Inspection") <> 0 And StringInStr($sTxt_Title, "Code Charts:") <> 0 Then
	  ;Code lấy DTC Code trong $sTxt_Title  để làm đường dẫn
	  $sTxt_Title = StringMid ($sTxt_Title, StringInStr($sTxt_Title, "Code Charts: ") + StringLen("Code Charts: "), 5)
	  Local $sProcedure_Path = "../" & $sTxt_Title & "/" & $sTxt_Title & ".html"
   Else
	  ;------------------------------------------------------------------------------------------------------------------
	  ;Chỉnh sửa text trong Procedure của All data cho phù hợp với Innova
	  $sTxt_Body = Modify_Body_HTML ($oIE)
	  ;------------------------------------------------------------------------------------------------------------------
	  ;Code lấy tên procedure trong $sTxt_Title để làm title cho file html và folder name cho Procedure
	  ;Mẫu: Computers and Control Systems |Testing and Inspection, Reading and Clearing Diagnostic Trouble Codes: DTC Check / Clear
	  $sTxt_Title = StringRight ($sTxt_Title, StringLen ($sTxt_Title) - StringInStr($sTxt_Title, ": ") - 1)
	  ;Chuẩn tên theo window
	  $sTxt_Title = Standardize_File_Name ($sTxt_Title)
	  ;Thêm string "Procedure: " phía trước
	  $sTxt_Title = "Procedure: " & $sTxt_Title
	  ;------------------------------------------------------------------------------------------------------------------
	  ;ĐOẠN CODE LẤY CÁC ID TRONG URL ĐỂ ĐẶT TÊN CHO PROCEDURE
	  Local $sTxt_File_Name = Create_Procedure_Name ($oIE)
	  Notification ("Begin to generate: " & $sTxt_File_Name)
	  Sleep (1000)
	  ;------------------------------------------------------------------------------------------------------------------
	  ;ĐOẠN CODE TẠO THƯ MỤC VÀ TẢI HÌNH ẢNH VỀ THƯ MỤC ĐÓ
	  ;Tạo các thư mục cần thiết
	  Local $sFilePath_PROCEDURE  = $sFilePath_YMME      &"\PROCEDURE"
	  If FileExists ($sFilePath_PROCEDURE) = 0 Then DirCreate($sFilePath_PROCEDURE)
	  Local $sFilePath_Title      = $sFilePath_YMME      &"\PROCEDURE"       &"\"& $sTxt_File_Name
	  If FileExists ($sFilePath_Title) = 0 Then	DirCreate($sFilePath_Title)
	  ;Tải hình ảnh của procedure vào thư mục
	  Download_Procedure_Image_Collection ($sFilePath_Title, $oIE)
	  ;------------------------------------------------------------------------------------------------------------------
	  ;ĐOẠN CODE INSERT HÌNH ẢNH CHO PROCEDURE
	  $sTxt_Body = Insert_Images_HTML ($oIE, $sTxt_Body)
	  ;------------------------------------------------------------------------------------------------------------------
	  ;ĐOẠN CODE LẤY TEXT TRONG TAG <A> ĐỂ REMOVE TẤT CẢ "SEE:............"
	  Local $oAs = _IETagNameGetCollection($oIE, "a")
	  Local $aHyperlink_Innertexts [1000]
	  Local $iCount_Hyperlink = 0
	  Local $sTemp = ""
	  For $oA In $oAs
		 If StringInStr ($oA.innertext, "See:", 0, 1) <> 0 Then
			$sTemp = Standardize_String ($oA.innertext)
			$sTxt_Body = StringReplace ($sTxt_Body, $sTemp,"", 1, 0)
		 EndIf
	  Next
	  ;------------------------------------------------------------------------------------------------------------------
	  Create_HTML  ($sFilePath_Title, $sTxt_File_Name,  $sTxt_Title, $sTxt_Body)
	  Local $sProcedure_Path = "../../PROCEDURE/" & $sTxt_File_Name & "/" & $sTxt_File_Name & ".html"
	  ;------------------------------------------------------------------------------------------------------------------
	  ;ĐOẠN CODE WRITE LOG FILE PROCEDURE
	  Local $sLog_Txt = "File name: " & $sTxt_File_Name & @CRLF & $sTxt_Title_Alldata & @CRLF & $sProcedure_Link
	  Write_Log_File ($sFilePath_YMME,"Log File Procedure Successful", @CRLF & @CRLF & @CRLF & $sLog_Txt, "append")
   EndIf
   ;------------------------------------------------------------------------------------------------------------------
   ;Trả về một string đường dẫn của PROCEDURE
   Return $sProcedure_Path
EndFunc






