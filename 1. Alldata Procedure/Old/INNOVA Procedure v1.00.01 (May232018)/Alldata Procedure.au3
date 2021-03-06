#cs ----------------------------------------------------------------------------
NOTE:
Làm file Log lưu lại hình bị lỗi khi tải (Nếu tải lâu hơn bao nhiêu giây thì phải vào function check mạng, note lại tốc độ mạng => tải lại)

Lưu ý trường hợp trong link procedure xuất hiện javascript

;_ClipBoard_SetData ($sYMME,$CF_TEXT)
;Exit

http://repair.alldata.com/alldata/article/display.action?componentId=3926&iTypeId=423&nonStandardId=3951433&vehicleId=54277&windowName=mainADOnlineWindow
http://repair.alldata.com/alldata/article/display.action?componentId=3926&iTypeId=423&nonStandardId=3956429&vehicleId=54277&windowName=mainADOnlineWindow
http://repair.alldata.com/alldata/article/display.action?componentId=3926&iTypeId=423&nonStandardId=3952079&vehicleId=53841&windowName=mainADOnlineWindow

Link chứa procedure có javascript
http://repair.alldata.com/alldata/article/display.action?componentId=3926&iTypeId=423&nonStandardId=5364910&vehicleId=52950&windowName=mainADOnlineWindow

Link chứa DTC có Part:
http://repair.alldata.com/alldata/article/display.action?componentId=3926&iTypeId=423&nonStandardId=5349152&vehicleId=52950&windowName=mainADOnlineWindow

Func Write_Log_File_Error ($sTxt)
	  Local $hFileOpen = FileOpen ("C:\Users\K\Desktop\Alldata DTC" & "\" & "Log File Error" & ".txt",$FO_APPEND)
	  FileWrite($hFileOpen, $sTxt & @CRLF & @CRLF)
EndFunc
#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include <MsgBoxConstants.au3>
#include <FileConstants.au3>
#include <InetConstants.au3>

#include <Clipboard.au3>
#include <IE.au3>
#include <Excel.au3>
#include <WinAPIFiles.au3>
#include "General_Alldata_Procedure_Library.au3"
#include "Create_JAVASCRIPT_Procedure.au3"
#include "Create_NORMAL_Procedure.au3"

Global $bWeb_Attach = 0
Global $bWeb_Visible = 0
Global $bWeb_Wait = 1
Global $bWeb_TakeFocus = 0

Global $bImage_Download = 0
;===================================================================================================================
;                  GENERAL CODE
;===================================================================================================================

; Set Hotkey for the program
HotKeySet("{ESC}", "Autoit_Exit")
;HotKeySet("^z", "Main_function")

Main_function ()

while (1)
WEnd


Func test_func ()
EndFunc


Func Autoit_Exit ()
   Exit
EndFunc



;====================================================================================================================
;                  FUNCTION DISCRIPTION: MAIN FUNCTION
;====================================================================================================================
Func Main_function ()
   ;ĐOẠN CODE THÔNG BÁO CHƯƠNG TRÌNH BẮT ĐẦU VÀ YÊU CẦU USER NHẬP DỮ LIỆU CẦN THIẾT
   ;Msg bắt đầu chương trình
   Sleep (500)
   MsgBox (0, "Alldata DTC Procedure", "This is the software to get DTC Procedure from Alldata" & @CRLF & "Please press OK to continue" & @CRLF & "While the software is processing, press ESC to Exit")
   Sleep (500)
   Local $sLink = "http://repair.alldata.com/alldata/article/display.action?componentId=3926&iTypeId=423&nonStandardId=5349152&vehicleId=52950&windowName=mainADOnlineWindow"
   Global $iSubscription_Num = 2

   ;Local $sLink = User_Input_Data ()
   ;------------------------------------------------------------------------------------------------------------------
   ;Gán trang web cho biến object
   Local $oIE = _IECreate($sLink, $bWeb_Attach, $bWeb_Visible, $bWeb_Wait, $bWeb_TakeFocus)
   ;------------------------------------------------------------------------------------------------------------------
   ;ĐOẠN CODE SỬ DỤNG FUNCTION "Check_Login_Alldata" ĐỂ KIỂM TRA ĐĂNG NHẬP
   If Check_Login_Alldata ($oIE) = "Not yet loged in before, this function has helped log in" Then
	  ;Reload trang DTC
	  $oIE = _IECreate($sLink,  $bWeb_Attach, $bWeb_Visible, $bWeb_Wait, $bWeb_TakeFocus)
   EndIf
   ;------------------------------------------------------------------------------------------------------------------
   ;ĐOẠN CODE SỬ DỤNG FUNCTION "Check_Subscription_Alldata" ĐỂ KIỂM TRA SUBSCIPTION
   $oIE = Check_Subscription_Alldata ($oIE, $sLink, $iSubscription_Num)
   ;------------------------------------------------------------------------------------------------------------------
   DTC_Procedure_Alldata ($oIE)
   Sleep (500)
   MsgBox (0, "Alldata DTC Procedure", "The process is done. Please Check!")
   Exit
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
   ;ĐOẠN CODE LẤY TEXT TRONG TAG SPAN ĐỂ LẤY TÊN XE ĐẶT TÊN CHO FOLDER
   ;Collect tất cả text trong tag <title> và lưu vào biến $sTxt_Title_Alldata, sau đó gán $sTxt_Title = $sTxt_Title_Alldata
   ;Việc này giúp giữ $sTxt_Title_Alldata để phần tích nếu cần sau này
   Local $oSpans = _IETagNameGetCollection($oIE, "span")
   Local $aSpans [1000]
   Local $iCount = 0
   Local $iMark = 0
   For $oSpan In $oSpans
	  $aSpans [$iCount] = $oSpan.innertext
	  If $aSpans [$iCount] = "Save Article " Then $iMark = $iCount
	  $iCount = $iCount + 1
   Next
   Local $sYMME = $aSpans [$iMark-1]
   $sYMME = Standardize_File_Name ($sYMME)
   ;------------------------------------------------------------------------------------------------------------------
   ;ĐOẠN CODE TẠO THƯ MỤC VÀ TẢI HÌNH ẢNH VỀ THƯ MỤC ĐÓ
   ;Tạo các thư mục cần thiết
   Local $sFilePath_Alldata_DTC = @ScriptDir & "\Alldata DTC"
   If FileExists ($sFilePath_Alldata_DTC) = 0 Then	DirCreate($sFilePath_Alldata_DTC)
   Local $sFilePath_YMME        = @ScriptDir & "\Alldata DTC"      &"\"&$sYMME
   If FileExists ($sFilePath_YMME) = 0 Then	DirCreate($sFilePath_YMME)
   Local $sFilePath_DTC         = @ScriptDir & "\Alldata DTC"      &"\"&$sYMME      &"\DTC"
   If FileExists ($sFilePath_DTC) = 0 Then DirCreate($sFilePath_DTC)
   Local $sFilePath_Title       = @ScriptDir & "\Alldata DTC"      &"\"&$sYMME      &"\DTC"       &"\"& $sTxt_Title
   If FileExists ($sFilePath_Title) = 0 Then	DirCreate($sFilePath_Title)
   ;Tải hình ảnh của procedure vào thư mục
   Download_Procedure_Image_Collection ($sFilePath_Title, $oIE)
   ;------------------------------------------------------------------------------------------------------------------
   ;ĐOẠN CODE INSERT HÌNH ẢNH CHO PROCEDURE
   $sTxt_Body = Insert_Images_HTML ($oIE, $sTxt_Body)
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
   ;_ClipBoard_SetData ($sText_Test,$CF_TEXT)





   ;------------------------------------------------------------------------------------------------------------------
   ;ĐOẠN CODE CREATE PROCEDURE VÀ TRẢ VỀ MỘT STRING MANG ĐƯỜNG DẪN TỚI FILE HTML CỦA PROCEDURE
   $iCount_Hyperlink = 0
   While $aHyperlink_Links [$iCount_Hyperlink] <> ""
	  ;Xem thử trong đường link có text "repair.alldata.com" hay không, để tránh trường hợp xuất hiện "javascript"
	  If StringInStr ($aHyperlink_Links [$iCount_Hyperlink], "repair.alldata.com", 0, 1) <> 0 Then
		 Local $sProcedure_Path = Create_NORMAL_Procedure ($sFilePath_YMME, $aHyperlink_Links [$iCount_Hyperlink])
		 Local $sHTML_Procedure_Hyperlink = "<a href=""" & $sProcedure_Path & """ target=""_blank"">" & "(More info)" & "</a>"
		 $sTxt_Body = StringReplace ($sTxt_Body, $aHyperlink_Innertexts [$iCount_Hyperlink],$sHTML_Procedure_Hyperlink, 1, 0)
	  Else
		 ;ĐOẠN CODE XỬ LÝ LINK PROCEDURE CHỨA JAVASCRIPT
		 Create_JAVASCRIPT_Procedure ($sFilePath_YMME, $aHyperlink_Links [$iCount_Hyperlink])
	  EndIf
	  $iCount_Hyperlink = $iCount_Hyperlink + 1
   WEnd
   ;------------------------------------------------------------------------------------------------------------------
   Create_HTML  ($sFilePath_Title, $sTxt_Title, $sTxt_Title, $sTxt_Body)
   Local $sLog_Txt = $sTxt_Title_Alldata & @CRLF & _IEPropertyGet ($oIE, "locationurl")
   Write_Log_File ($sFilePath_YMME, "Log File DTC",  $sLog_Txt, "append")

   _IEQuit($oIE)
EndFunc

