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
Func Create_NORMAL_Procedure ($sFilePath_YMME, $sProcedure_Link)
   Local $oIE_Procedure = _IECreate ($sProcedure_Link,  $bWeb_Attach, $bWeb_Visible, $bWeb_Wait, $bWeb_TakeFocus)
   ;ĐOẠN CODE SỬ DỤNG FUNCTION "Check_Subscription_Alldata" ĐỂ KIỂM TRA SUBSCIPTION
   $oIE_Procedure = Check_Subscription_Alldata ($oIE_Procedure, $sProcedure_Link, $iSubscription_Num)
   ;------------------------------------------------------------------------------------------------------------------
   ;ĐOẠN CODE LẤY TEXT TRONG TAG <TITLE> ĐỂ LÀM TITLE CHO HTML PROCEDURE VÀ KIỂM TRA XEM LINK ĐÓ CÓ PHẢI LINK DTC HAY KHÔNG
   Local $sTxt_Title_Alldata = _IEPropertyGet ($oIE_Procedure, "title")
   Local $sTxt_Title = $sTxt_Title_Alldata
   $sTxt_Title = Standardize_String ($sTxt_Title)
   ;------------------------------------------------------------------------------------------------------------------
   ;ĐOẠN CODE KIỂM TRA NẾU LÀ LINK DTC THÌ TRẢ VỀ PATH HTML CỦA DTC, NẾU LÀ LINK PROCEDURE THÌ LÀM PROCEDURE VÀ TRẢ VỀ PATH CỦA PROCEDURE
   If StringInStr ($sTxt_Title, "A L L Diagnostic Trouble Codes ( DTC ) |Testing and Inspection") <> 0 And StringInStr($sTxt_Title, "Code Charts:") <> 0 Then
	  ;Nếu trong $sTxt_Title có chứa các string nhận dạng DTC
	  $sTxt_Title = StringMid ($sTxt_Title, StringInStr($sTxt_Title, "Code Charts: ") + StringLen("Code Charts: "), 5)
	  Local $sProcedure_Path = "../" & $sTxt_Title & "/" & $sTxt_Title & ".html"
	  ;Code lấy DTC Code trong $sTxt_Title  để làm đường dẫn
   Else
	  ;------------------------------------------------------------------------------------------------------------------
	  ;Chỉnh sửa text trong Procedure của All data cho phù hợp với Innova
	  $sTxt_Body = Modify_Body_HTML ($oIE_Procedure)
	  ;------------------------------------------------------------------------------------------------------------------
	  ;Code lấy tên procedure trong $sTxt_Title để làm title cho file html và folder name cho Procedure
	  ;Mẫu: Computers and Control Systems |Testing and Inspection, Reading and Clearing Diagnostic Trouble Codes: DTC Check / Clear
	  $sTxt_Title = StringRight ($sTxt_Title, StringLen ($sTxt_Title) - StringInStr($sTxt_Title, ": ") - 1)
	  ;Chuẩn tên theo window
	  $sTxt_Title = Standardize_File_Name ($sTxt_Title)
	  ;Thêm string "Procedure: " phía trước
	  $sTxt_Title = "Procedure: " & $sTxt_Title
	  ;------------------------------------------------------------------------------------------------------------------
	  ;ĐOẠN CODE LẤY DÙNG $sTxt_Title ĐỂ TẠO HTML VÀ FOLDER FILE NAME VÀ LƯU VÀO $sTxt_File_Name
	  ;Thay thế ": " phía sau chữ Procedure thanh khoảng trắng
	  Local $sTxt_File_Name = StringReplace ($sTxt_Title, ": ", " ")
	  ;Chuyển tất cả thành chữ hoa
	  $sTxt_File_Name = StringUpper ($sTxt_File_Name)
	  ;Chuyển khoảng trắng thành gạch dưới
	  $sTxt_File_Name = StringReplace ($sTxt_File_Name, " ", "_")
	  ;------------------------------------------------------------------------------------------------------------------
	  ;ĐOẠN CODE LẤY CÁC ID TRONG URL ĐỂ ĐẶT TÊN CHO PROCEDURE
	  ;Lấy URL của procedure để lấy các ID đặt tên cho procedure, tránh trùng tên procedure
	  $sIE_Procedure_URL = _IEPropertyGet ($oIE_Procedure, "locationurl")
	  ;Lấy vị trí các ID
	  Local $iComponentID_Pos = StringInStr ($sIE_Procedure_URL, "componentId=", 0, 1) + StringLen ("componentId=")
	  Local $iITypeId_Pos = StringInStr ($sIE_Procedure_URL, "&iTypeId=", 0, 1) + StringLen ("&iTypeId=")
	  Local $iNonStandardId_Pos = StringInStr ($sIE_Procedure_URL, "&nonStandardId=", 0, 1) + StringLen ("&nonStandardId=")
	  Local $iVehicleId_Pos = StringInStr ($sIE_Procedure_URL, "&vehicleId=", 0, 1) + StringLen ("&vehicleId=")
	  ;Lấy string các ID
	  Local $sComponentID = StringMid ($sIE_Procedure_URL, $iComponentID_Pos, $iITypeId_Pos - $iComponentID_Pos - StringLen ("&iTypeId="))
	  Local $sITypeId = StringMid ($sIE_Procedure_URL, $iITypeId_Pos, $iNonStandardId_Pos - $iITypeId_Pos - StringLen ("&nonStandardId="))
	  Local $sNonStandardId = StringMid ($sIE_Procedure_URL, $iNonStandardId_Pos, $iVehicleId_Pos - $iNonStandardId_Pos - StringLen ("&vehicleId="))
	  $sTxt_File_Name = $sTxt_File_Name & "_" & $sComponentID & "_"& $sITypeId & "_"& $sNonStandardId
	  ;------------------------------------------------------------------------------------------------------------------
	  ;ĐOẠN CODE TẠO THƯ MỤC VÀ TẢI HÌNH ẢNH VỀ THƯ MỤC ĐÓ
	  ;Tạo các thư mục cần thiết
	  Local $sFilePath_PROCEDURE  = $sFilePath_YMME      &"\PROCEDURE"
	  If FileExists ($sFilePath_PROCEDURE) = 0 Then DirCreate($sFilePath_PROCEDURE)
	  Local $sFilePath_Title      = $sFilePath_YMME      &"\PROCEDURE"       &"\"& $sTxt_File_Name
	  If FileExists ($sFilePath_Title) = 0 Then	DirCreate($sFilePath_Title)
	  ;Tải hình ảnh của procedure vào thư mục
	  Download_Procedure_Image_Collection ($sFilePath_Title, $oIE_Procedure)
	  ;------------------------------------------------------------------------------------------------------------------
	  ;ĐOẠN CODE INSERT HÌNH ẢNH CHO PROCEDURE
	  $sTxt_Body = Insert_Images_HTML ($oIE_Procedure, $sTxt_Body)
	  ;------------------------------------------------------------------------------------------------------------------
	  ;ĐOẠN CODE LẤY TEXT TRONG TAG <A> ĐỂ REMOVE TẤT CẢ "SEE:............"
	  Local $oAs = _IETagNameGetCollection($oIE_Procedure, "a")
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
	  Local $sLog_Txt = $sTxt_Title_Alldata & @CRLF & _IEPropertyGet ($oIE_Procedure, "locationurl")
	  Write_Log_File ($sFilePath_YMME,"Log File Procedure", $sLog_Txt, "append")
	  _IEQuit($oIE_Procedure)
   EndIf
   ;------------------------------------------------------------------------------------------------------------------
   ;Trả về một string đường dẫn của PROCEDURE
   Return $sProcedure_Path
EndFunc






