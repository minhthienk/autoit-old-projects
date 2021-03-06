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
;                  FUNCTION DISCRIPTION: DISPLAY INPUT BOXES FOR USER TO ENTER INITIAL DATA
;                  RETURN			   : LINK TO DTC PROCEDURE
;====================================================================================================================
Func User_Input_Data ()
   ;ĐOẠN CODE THÔNG BÁO CHƯƠNG TRÌNH BẮT ĐẦU VÀ YÊU CẦU USER NHẬP DỮ LIỆU CẦN THIẾT
   ;Msg bắt đầu chương trình
   Sleep (500)
   MsgBox (0, "Alldata DTC Procedure", "This is the software to get DTC Procedure from Alldata" & @CRLF & "Please press OK to continue" & @CRLF & "While the software is processing, press ESC to Exit")
   Sleep (500)
   ;Yêu cầu user nhập Subcription License của Alldata
   Global $iSubscription_Num = InputBox("Alldata DTC Procedure", "Enter INNOVA Subscription License Number:" & @CRLF & "Type 1, 2, 3, 4 or 5")
   While $iSubscription_Num <> "0" And $iSubscription_Num <> "1" And $iSubscription_Num <> "2" And $iSubscription_Num <> "3" And $iSubscription_Num <> "4" And $iSubscription_Num <> "5"
	  $iSubscription_Num = InputBox("Alldata DTC Procedure", "The number is not correct" & @CRLF & "Enter INNOVA Subscription License Number again:" & @CRLF & "(Type 1, 2, 3, 4 or 5)")
   WEnd
   ;Yêu cầu user nhập link alldata DTC
   Local $sLink = InputBox("Alldata DTC Procedure", "Enter the DTC link:")
   While StringInStr ($sLink, "repair.alldata.com", 0, 1) = 0
	  $sLink = InputBox("Alldata DTC Procedure", "This is not repair alldata Link" & @CRLF & "Enter an Alldata DTC link:")
   WEnd
   ;Thông báo chuognw trình  sẽ bắt đầu thực hiện
   MsgBox (0, "Alldata DTC Procedure", "The software is going to process." & @CRLF & "Please wait until the message ""Done"" popping up" & @CRLF & "While the software is processing, press ESC to Exit")
   Return $sLink
EndFunc




;====================================================================================================================
;                  FUNCTION DISCRIPTION: CHECK WHETHER USER HAS LOGGED IN ALLDATA
;				   INPUT               : $oIE
;                  OUTPUT              : A STRING OF RESULT
;====================================================================================================================
Func Check_Login_Alldata (Byref $oIE)
   Notification ("Checking Login ...", "Normal")
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
   Notification ("Checking Lisence ...", "Normal")
   Local $sResult
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
			Local $sTxt_Img = ""
			Local $oImgs = _IEImgGetCollection($oIE)
			Local $iCount
			Local $aImg_Object [20]
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
			$sResult = "Not yet had subsription, this function has helped clicked subscription"
		 Else
			$sResult = "Already had subscription  - But need reload"
		 EndIf
	  Else
		 $sResult = "Already had subscription  - No need reload"
	  EndIf
	  ;------------------------------------------------------------------------------------------------------------------
	  If $sResult <> "Already had subscription  - No need reload" Then
		 ;Reload trang DTC
		 IENavigate_Check_Error ($oIE, $sLink)
	  EndIf
   Until $sResult = "Already had subscription  - No need reload"
EndFunc



;====================================================================================================================
;                  FUNCTION DISCRIPTION: STANDARDIZE FILE NAME AS WINDOW SPECIFIED
;                  INPUT               : $sFile_Name
;                  OUTPUT              : $sFile_Name WITH INAPPOPRIATE CHARACTERS REMOVED
;====================================================================================================================
Func Standardize_File_Name ($sFile_Name)
   ;Remove ký tự được biệt
   $sFile_Name = StringReplace ($sFile_Name, "/", " ")
   $sFile_Name = StringReplace ($sFile_Name, "\", " ")
   $sFile_Name = StringReplace ($sFile_Name, ":", " ")
   $sFile_Name = StringReplace ($sFile_Name, "*", " ")
   $sFile_Name = StringReplace ($sFile_Name, "?", " ")
   $sFile_Name = StringReplace ($sFile_Name, """", " ")
   $sFile_Name = StringReplace ($sFile_Name, "<", " ")
   $sFile_Name = StringReplace ($sFile_Name, ">", " ")
   $sFile_Name = StringReplace ($sFile_Name, "|", " ")
   $sFile_Name = StringReplace ($sFile_Name, ",", " ")
   $sFile_Name = StringReplace ($sFile_Name, "-", " ")
   ;Chuyển 2 khoảng trắng thành 1 khoảng trắng
   While StringInStr ($sFile_Name, "  ", 0, 1) <> 0
		$sFile_Name = StringReplace ($sFile_Name, "  ", " ")
   WEnd
   Return $sFile_Name
EndFunc


;====================================================================================================================
;                  FUNCTION DISCRIPTION: STANDARDIZE STRING
;                  INPUT               : $sString
;                  OUTPUT              : $sString AFTER STANDARDIZED
;====================================================================================================================
Func Standardize_String ($sString)
   ;Replace 2 khoảng trắng bằng 1 khoảng trắng
   Do
   $sString = StringReplace ( $sString, "  ", " ")
   Until @extended = 0
   ;Cắt các khoảng trắng thừa bên trái
   While StringLeft ($sString, 1) = " "
   $sString = StringTrimLeft ($sString, 1)
   WEnd
   ;Cắt các khoảng trắng thừa bên phải
   While StringRight ($sString, 1) = " "
   $sString = StringTrimRight ($sString, 1)
   WEnd
   Do
	  $sString = StringReplace ( $sString, " :", ":")
   Until @extended = 0
   Return $sString
EndFunc



;====================================================================================================================
;                  FUNCTION DISCRIPTION: INSERT IMAGES TO HTML BODY
;                  INPUT               : $oIE, $sTxt_Body
;                  OUTPUT              : $sTxt_Body WITH HTML CODE FOR INSERTING PROCEDURE IMAGES IN IT
;====================================================================================================================
Func Insert_Images_HTML ($oIE, $sTxt_Body)
   ;ĐOẠN CODE INSERT HÌNH ẢNH CHO PROCEDURE
   ;Lấy thông tin của tất cả hình ảnh trong procedure;tên Image name mẫu 11_lx14ls460A-A203088E09.png&width=356&height=190
   Local $aProcedure_Image_Name [1000] = Get_Procedure_Image_Info_Collection ($oIE)
   Local $sTemp
   Local $iCount = 0
   While StringInStr ($sTxt_Body, "Zoom and Print Options", 0, 1) <> 0
	  ;Chỉnh sửa thông tin hình ảnh có được thành code html. ;Code mẫu: <img src="11_lx14ls460A-A203088E09.jpg" width="356" height="190"/>
	  $sTemp = """" & $aProcedure_Image_Name [$iCount]
	  $sTemp = StringReplace($sTemp,".png",".jpg")
	  $sTemp = StringReplace($sTemp,"&width=",""" width=""")
	  $sTemp = StringReplace($sTemp,"&height=",""" height=""")
	  $sTemp = $sTemp & """"
	  $sTemp = "<br>" & @CRLF & "<img src=" & $sTemp & " border = ""1""" & "/><br>"
	  ;Thay thế các đoạn text "Zoom and Print Options" bằng link hình ảnh
	  $sTxt_Body = StringReplace ($sTxt_Body, "Zoom and Print Options", $sTemp, 1)
	  $iCount = $iCount + 1
   WEnd
   Return $sTxt_Body
EndFunc

;====================================================================================================================
;                  FUNCTION DISCRIPTION: MODIFY TEXT IN ALLDATA PROCEDURE
;                                        INTO INNOVA PROCEDURE
;====================================================================================================================
Func Modify_Body_HTML ($oIE)
   Sleep (300)
   ;ĐOẠN CODE LẤY TEXT TRONG TAG <TR> ĐỂ LÀM PROCEDURE
   ;Collect tất cả text trong tag <tr> và lưu vào biến $sTxt_Body_Alldata, sau đó gán $sTxt_Body = $sTxt_Body_Alldata
   ;Việc này giúp giữ $sTxt_Body_Alldata để phần tích nếu cần sau này
   Local $oTrs = _IETagNameGetCollection($oIE, "tr")
   Local $sTxt_Body_Alldata = ""
   For $oTr In $oTrs
	   $sTxt_Body_Alldata &= $oTr.innertext & @CRLF
	Next
   Local $sTxt_Body = $sTxt_Body_Alldata
   ;------------------------------------------------------------------------------------------------------------------
   ;ĐOẠN CODE CHỈNH SỬA TEXT TRONG BODY CHO PHÙ HỢP VỚI INNOVA HTML
   ;Cắt tất cả text phía trước dòng chữ "Text Only  Image(s) Only"
   $sTxt_Body = StringRight ($sTxt_Body, StringLen ($sTxt_Body) - StringInStr($sTxt_Body, "Text Only  Image(s) Only", 0, 1, 1) - StringLen ("Text Only  Image(s) Only"))
   ;Cắt tất cả khoảng trắng và xuống dòng đầu trang
   While StringLeft ($sTxt_Body,1) = " " Or StringLeft ($sTxt_Body,1) = @CR Or StringLeft ($sTxt_Body,1) = @LF
		 $sTxt_Body = StringRight ($sTxt_Body, StringLen ($sTxt_Body) - 1)
   WEnd
   ;Xuống 1 dòng đầu trang
   $sTxt_Body = @CRLF & $sTxt_Body
   ;Cắt tất cả text phía sau dòng chữ "var classElements"
   $sTxt_Body = StringLeft ($sTxt_Body, StringInStr($sTxt_Body, "var classElements", 0, 1, 1) - 1)
   ;Replace xuống dòng nhiều lần bằng tối đa 2 lần xuống dòng trong $sTxt_Body
   Do
	  $sTxt_Body = StringReplace ( $sTxt_Body, @CRLF&@CRLF&@CRLF, @CRLF&@CRLF)
   Until @extended = 0
   ;Replace ký tự khoảng trắng thừa bằng ký tự trống HTML, chỉ replace bộ 3 khoảng trắng
   Do
	  $sTxt_Body = StringReplace ( $sTxt_Body, "   ", "&nbsp;&nbsp;&nbsp;")
   Until @extended = 0
   ;Replace 2 khoảng trắng bằng 1 khoảng trắng
   Do
	  $sTxt_Body = StringReplace ( $sTxt_Body, "  ", " ")
   Until @extended = 0
   Do
	  $sTxt_Body = StringReplace ( $sTxt_Body, " :", ":")
   Until @extended = 0
   ;Thay thế ký tự xuống dòng bằng tag <br>
   $sTxt_Body = StringReplace ( $sTxt_Body, @CRLF, "<br>"&@CRLF)
   ;------------------------------------------------------------------------------------------------------------------
   ;ĐOẠN CODE THAY THẾ CÁC STRING OEM TRONG ALLDATA THÀNH STRING INNOVA
   $sTxt_Body = StringReplace ( $sTxt_Body, "Techstream", "Scan Tool")
   ;------------------------------------------------------------------------------------------------------------------
   ;ĐOẠN CODE IN ĐẬM CÁC STRING CẦN THIẾT TRONG BODY
   ;Collect tất cả text trong tag <b> (text in đậm) và lưu vào biến $sTxt_Bold và mảng $aData_Bold
   Local $oBolds = _IETagNameGetCollection($oIE, "b")
   Local $sTxt_Bold = ""
   Local $aData_Bold [1000]   ;Mảng dùng để lưu các text BOLD
   Local $iCount			  ;Biến	dùng để đếm thứ tự cho mảng
   Local $iMax_Data_Bold	  ;Biến	dùng để lưu giá trị lớn nhất của mảng
   Local $sTemp
   For $oBold In $oBolds
	  $sTemp = $oBold.innertext
	  ;Collect tất cả text trong tag <a> để xem trong text bold có chứa text của link hay ko
	  Local $oAs = _IETagNameGetCollection($oBold, "a")
	  For $oA In $oAs
		 $sTemp = StringReplace ($sTemp, $oA.innertext, " " & $oA.innertext & " ")
	  Next
	  ;Replace " :" = ":"
	  Do
		 $sTemp = StringReplace ( $sTemp, " :", ":")
	  Until @extended = 0
	  ;Cắt các khoảng trắng thừa bên trái
	  While StringLeft ($sTemp, 1) = " "
		 $sTemp = StringTrimLeft ($sTemp, 1)
	  WEnd
	  ;Cắt các khoảng trắng thừa bên phải
	  While StringRight ($sTemp, 1) = " "
		 $sTemp = StringTrimRight ($sTemp, 1)
	  WEnd
	  ;Replace 2 khoảng trắng bằng 1 khoảng trắng
	  Do
		 $sTemp = StringReplace ( $sTemp, "  ", " ")
	  Until @extended = 0
	  $sTxt_Bold &= $sTemp & @CRLF
	  $aData_Bold [$icount] = $sTemp
	  $icount = $icount + 1
   Next

   $iMax_Data_Bold = $icount - 1
   For $i=0 To $iMax_Data_Bold Step 1
	  ;Dấu hiệu thay thế: (Xuống dòng) + Bold Text & <br>
	  $sTxt_Body = StringReplace ($sTxt_Body, @CRLF & $aData_Bold[$i] & "<br>", @CRLF&"<b>" & $aData_Bold[$i] & "</b><br>" , 1)
   Next
   Return $sTxt_Body
EndFunc



;====================================================================================================================
;                  FUNCTION DISCRIPTION: CREATE HTML FILE FROM
;				   INPUT               : $sFilePath, $sTxt_Title,$HTML_body
;                  OUTPUT              : AN HTML FILE IN $sFilePath
;====================================================================================================================
Func Create_HTML ($sFilePath, $sTxt_Name, $sTxt_Title, $HTML_body)
   Local $sHTML = ""
   $sHTML &= "<HTML>" & @CRLF
   $sHTML &= "<HEAD>" & @CRLF
   $sHTML &= "<TITLE>"&$sTxt_Title&"</TITLE>" & @CRLF
   $sHTML &= "</HEAD>" & @CRLF
   $sHTML &= "<BODY>"
   $sHTML &= "<OL>"
   $sHTML &= $HTML_body & @CRLF
   $sHTML &= "</OL>" & @CRLF
   $sHTML &= "</BODY>" & @CRLF
   $sHTML &= "</HTML>"
   Local $hFileOpen = FileOpen ($sFilePath & "\" & $sTxt_Name & ".html",$FO_OVERWRITE)
   FileWrite($hFileOpen, $sHTML)
   FileClose($hFileOpen)
EndFunc



;====================================================================================================================
;                  FUNCTION DISCRIPTION: CREATE LOG FILE OF DTC OR PROCEDURE
;				   INPUT               : $sFilePath, $sWhichLogFile,  $sTxt, $sMode
;                  OUTPUT              : AN LOG FILE IN $sFilePath
;====================================================================================================================
Func Write_Log_File ($sFilePath, $sWhichLogFile,  $sTxt, $sMode)
   If $sWhichLogFile = "Log File DTC Successful" Then
	  If $sMode = "overwrite" Then
		 Local $hFileOpen = FileOpen ($sFilePath & "\" & "Log File DTC Successful" & ".txt",$FO_OVERWRITE)
	  Else
		 Local $hFileOpen = FileOpen ($sFilePath & "\" & "Log File DTC Successful" & ".txt",$FO_APPEND)
	  EndIf
	  FileWrite($hFileOpen, $sTxt)
	  FileClose($hFileOpen)


   ElseIf $sWhichLogFile = "Log File Procedure Successful" Then
	  If $sMode = "overwrite" Then
		 Local $hFileOpen = FileOpen ($sFilePath & "\" & "Log File Procedure Successful" & ".txt",$FO_OVERWRITE)
	  Else
		 Local $hFileOpen = FileOpen ($sFilePath & "\" & "Log File Procedure Successful" & ".txt",$FO_APPEND)
	  EndIf
	  FileWrite($hFileOpen, $sTxt)
	  FileClose($hFileOpen)


   ElseIf $sWhichLogFile = "Log File Procedure Failed" Then
	  If $sMode = "overwrite" Then
		 Local $hFileOpen = FileOpen ($sFilePath & "\" & "Log File Procedure Failed" & ".txt",$FO_OVERWRITE)
	  Else
		 Local $hFileOpen = FileOpen ($sFilePath & "\" & "Log File Procedure Failed" & ".txt",$FO_APPEND)
	  EndIf
	  FileWrite($hFileOpen, $sTxt)
	  FileClose($hFileOpen)


   Else ;$sWhichLogFile = "Scan DTC Config"
	  If $sMode = "overwrite" Then
		 Local $hFileOpen = FileOpen ($sFilePath & "\" & "Scan DTC Config" & ".txt",$FO_OVERWRITE)
	  Else
		 Local $hFileOpen = FileOpen ($sFilePath & "\" & "Scan DTC Config" & ".txt",$FO_APPEND)
	  EndIf
	  FileWrite($hFileOpen, $sTxt)
	  FileClose($hFileOpen)
   EndIf
EndFunc



;====================================================================================================================
;                  FUNCTION DISCRIPTION: DOWNLOAD ALL IMAGES IN A DTC OR PROCEDURE LINK
;				   INPUT               : $FilePath, $oIE
;                  OUTPUT              : ALL IMAGES STRORED IN $FilePath
;====================================================================================================================
Func Download_Procedure_Image_Collection ($FilePath, $oIE)
   ;Lấy tất cả link hình ảnh và lưu vào mảng $aImage_Links
   Local $aImage_Links [1000] = Get_Image_Link_Collection ($oIE)
   ;Mảng dùng để lưu links hình ảnh cần cho procedure
   Local $aProcedure_Image_Links [1000]
   ;Biến tạm để xử lý string
   Local $sTemp
   ;Doạn Code Lấy link hình ảnh cần cho procedure với độ phân giải cao vào lưu vào $aProcedure_Image_Links
   Local $i = 0
   Local $j = 0
   While $aImage_Links [$i] <> ""
	  If StringInStr ($aImage_Links [$i], "&", 0, 1) <> 0 Then
		 $sTemp = $aImage_Links [$i]
		 ;Cắt phần text thừa bên trái trong link cũ
		 $sTemp = StringRight ($sTemp, StringLen ($sTemp) - StringLen ("http://repair.alldata.com/alldata/imagesWLinks?file=//"))
		 ;Cắt phần text thừa bên phải trong link cũ
		 $sTemp = StringLeft ($sTemp, StringInStr ($sTemp, "&", 0, 1) -1)
		 ;Tạo link mới độ phân giải cao
		 $sTemp = "http://repair.alldata.com/alldata/images?t_file=/" & $sTemp
		 $aProcedure_Image_Links [$j] = $sTemp
		 $j = $j + 1
	  EndIf
	  $i = $i + 1
   WEnd
   ;Tải tất cả hình ảnh cần cho procedure với độ phân giải cao
   $i = 0
   While $aProcedure_Image_Links [$i] <> ""
	  Image_Download_Convert2JPG($FilePath, $aProcedure_Image_Links [$i])
	  $i = $i + 1
   WEnd
EndFunc




;====================================================================================================================
;                  FUNCTION DISCRIPTION: DOWNLOAD AN IMAGE A THE LINK AND CONVERT TO JPG
;				   INPUT               : $sFilePath, $sLink
;                  OUTPUT              : AN JPG IMAGE
;====================================================================================================================
Func Image_Download_Convert2JPG($sFilePath, $sLink)
   ;Function này dùng để tải về hình ảnh và lưu vào vị trí FilePath
   ;Cắt phần tên hình ảnh của link để làm tên lưu cho hình ảnh
   Local $iTrim_Pos = StringInStr ($sLink, "/", 0, -1)
   Local $ImageName
   $ImageName = StringRight ($sLink, StringLen($sLink) - $iTrim_Pos)
   $ImageName = StringLeft($ImageName, StringLen ($ImageName) - 4) & ".jpg"
   If ($bImage_Download = 1) Then
	  ;Download the file in the background with the selected option of 'force a reload from the remote site.'
	  Local $hDownload = InetGet($sLink, $sFilePath &"\"& $ImageName, $INET_FORCERELOAD, $INET_DOWNLOADBACKGROUND)
	  ;Wait for the download to complete by monitoring when the 2nd index value of InetGetInfo returns True.
	  Local $bNoti_Replace_Previous_Flag = False
	  Do
		 Local $sDownloadRead = InetGetInfo($hDownload, $INET_DOWNLOADREAD)
		 Local $sDownloadSize = InetGetInfo($hDownload, $INET_DOWNLOADSIZE)
		 If $sDownloadSize = 0 Then $sDownloadSize = "Unknown"
		 If $bNoti_Replace_Previous_Flag = False Then
			Notification ("Downloading:  " & @CRLF & $ImageName & @CRLF & $sDownloadRead & "/" & $sDownloadSize & " (bytes)", "Normal")
			$bNoti_Replace_Previous_Flag = True
		 Else
			Notification ("Downloading:  " & @CRLF & $ImageName & @CRLF & $sDownloadRead & "/" & $sDownloadSize & " (bytes)", "Replace Previous")
		 EndIf
		 Sleep(200)
	  Until InetGetInfo($hDownload, $INET_DOWNLOADCOMPLETE)
	  Sleep (500)
	  InetClose($hDownload)
	  Sleep (500)
	  Notification ("Downloaded:  " & @CRLF & $ImageName, "Replace Previous")
   EndIf
EndFunc




;====================================================================================================================
;                  FUNCTION DISCRIPTION: GET ALL IMAGES INFO IN A DTC OR PROCEDURE LINK
;				   INPUT               : $oIE
;                  OUTPUT              : AN ARRAY OF ALL IMAGES INFO
;====================================================================================================================
Func Get_Procedure_Image_Info_Collection ($oIE)
   ;Lấy tất cả link hình ảnh và lưu vào mảng $aImage_Links
   Local $aImage_Links [1000] = Get_Image_Link_Collection ($oIE)
   ;Mảng dùng để lưu tên hình ảnh cần cho procedure
   Local $aProcedure_Image_Names [1000]
   ;Biến tạm để xử lý string
   Local $sTemp
   ;Doạn Code Lấy tên hình ảnh cần cho procedure lưu vào $aProcedure_Image_Names
   Local $i = 0
   Local $j = 0
   While $aImage_Links [$i] <> ""
	  If StringInStr ($aImage_Links [$i], "&", 0, 1) <> 0 Then
		 $sTemp = $aImage_Links [$i]

		 ;Cắt phần text thừa bên phải trong link
		 $sTemp = StringLeft ($sTemp, StringInStr ($sTemp, "&", 0, 3) -1)

		 ;Cắt phần text thừa bên trái trong link
		 $sTemp = StringRight ($sTemp, StringLen ($sTemp) - StringInStr ($sTemp, "/", 0, -1))

		 ;Lấy tên .jpg
		 $aProcedure_Image_Names [$j] = $sTemp
		 $j = $j + 1
	  EndIf
	  $i = $i + 1
   WEnd
   Return $aProcedure_Image_Names
EndFunc



;====================================================================================================================
;                  FUNCTION DISCRIPTION: DOWNLOAD ALL IMAGES IN A NORMAL WEBSITE (NOT ALLDATA)
;				   INPUT               : $FilePath, $oIE
;                  OUTPUT              : ALL IMAGES STORED IN $FilePath
;====================================================================================================================
Func Download_Image_Collection ($FilePath, $oIE)
   ;Function này dùng để tải tất cả hình ảnh có trong một trang web và lưu vào vị trí $FilePath
   Local $aImage_Links [1000] = Get_Image_Link_Collection ($oIE)
   Local $i = 0
   While $aImage_Links [$i] <> ""
	  Image_Download($FilePath,$aImage_Links [$i])
	  $i = $i + 1
   WEnd
EndFunc



;====================================================================================================================
;                  FUNCTION DISCRIPTION: GET ALL IMAGES LINKS IN A NORMAL WEBSITE (NOTE ALLDATA)
;				   INPUT               : oIE
;                  OUTPUT              : AN ARRAY OF ALL IMAGES LINKS
;====================================================================================================================
Func Get_Image_Link_Collection ($oIE)
   ;Function này dùng để lưu tất cả link hình ảnh có trong 1 trang web
   ;Giá trị trả về là một mảng gồm tất cả link hình ảnh
   Local $oImgs = _IEImgGetCollection($oIE)
   Local $iCount			  ;Biến	dùng để đếm thứ tự cho mảng
   Local $aImage_Links [1000]
   For $oImg  In $oImgs
	  $aImage_Links [$iCount] = $oImg.src
	  $icount = $icount + 1
   Next
   Return $aImage_Links
EndFunc



;====================================================================================================================
;                  FUNCTION DISCRIPTION: DOWNLOAD AN IMAGE A THE LINK
;				   INPUT               : $sFilePath, $sLink
;                  OUTPUT              : AN JPG IMAGE
;====================================================================================================================
Func Image_Download($sFilePath, $sLink)
   ;Function này dùng để tải về hình ảnh và lưu vào vị trí FilePath
   ;Cắt phần tên hình ảnh của link để làm tên lưu cho hình ảnh
   Local $iTrim_Pos = StringInStr ($sLink, "/", 0, -1)
   Local $ImageName
   $ImageName = StringRight ($Link, StringLen($sLink) - $iTrim_Pos)
   ; Download the file in the background with the selected option of 'force a reload from the remote site.'
   Local $hDownload = InetGet($sLink, $sFilePath &"\"& $ImageName, $INET_FORCERELOAD, $INET_DOWNLOADBACKGROUND)
   ; Wait for the download to complete by monitoring when the 2nd index value of InetGetInfo returns True.
   Do
	  Sleep(250)
   Until InetGetInfo($hDownload, $INET_DOWNLOADCOMPLETE)
   ; Download the file by waiting for it to complete. The option of 'get the file from the local cache' has been selected.
   ;InetGet($Link, $sFilePath &"\"& $ImageName, 1, 0)
EndFunc



;====================================================================================================================
;                  FUNCTION DISCRIPTION: CREATE PROCEDURE NAME
;                  RETURN			   : A STRING OF PROCEDURE NAME
;====================================================================================================================
Func Create_Procedure_Name ($sProcedure_Link)
   ;ĐOẠN CODE LẤY CÁC ID TRONG URL ĐỂ ĐẶT TÊN CHO PROCEDURE
   Local $sIE_Procedure_URL = $sProcedure_Link
   ;Gắn thêm dấu và vào cuối string để đánh dấu
   $sIE_Procedure_URL = $sIE_Procedure_URL & "&"
   ;Lấy $sComponentID
   Local $iComponentID_Pos = StringInStr ($sIE_Procedure_URL, "componentId=", 0, 1) + StringLen ("componentId=")
   Local $iID_End_Pos = StringInStr ($sIE_Procedure_URL, "&", 0, 1,  $iComponentID_Pos)
   Local $sComponentID = StringMid ($sIE_Procedure_URL, $iComponentID_Pos, $iID_End_Pos - $iComponentID_Pos)
   ;Lấy $sITypeId
   Local $iITypeId_Pos = StringInStr ($sIE_Procedure_URL, "iTypeId=", 0, 1) + StringLen ("iTypeId=")
   Local $iID_End_Pos = StringInStr ($sIE_Procedure_URL, "&", 0, 1,  $iITypeId_Pos)
   Local $sITypeId = StringMid ($sIE_Procedure_URL, $iITypeId_Pos, $iID_End_Pos - $iITypeId_Pos)
   ;Lấy $sNonStandardId
   Local $iNonStandardId_Pos = StringInStr ($sIE_Procedure_URL, "nonStandardId=", 0, 1) + StringLen ("nonStandardId=")
   Local $iID_End_Pos = StringInStr ($sIE_Procedure_URL, "&", 0, 1,  $iNonStandardId_Pos)
   Local $sNonStandardId = StringMid ($sIE_Procedure_URL, $iNonStandardId_Pos, $iID_End_Pos - $iNonStandardId_Pos)
   ;Lấy $sVehicleId
   Local $iVehicleId_Pos = StringInStr ($sIE_Procedure_URL, "vehicleId=", 0, 1) + StringLen ("vehicleId=")
   Local $iID_End_Pos = StringInStr ($sIE_Procedure_URL, "&", 0, 1,  $iVehicleId_Pos)
   Local $sVehicleId = StringMid ($sIE_Procedure_URL, $iVehicleId_Pos,  $iID_End_Pos - $iVehicleId_Pos)
   ;Tạo PROCEDURE name
   Local $sTxt_File_Name = "PROCEDURE"
   $sTxt_File_Name = $sTxt_File_Name & "_" & $sComponentID & "_"& $sITypeId & "_"& $sNonStandardId & "_"& $sVehicleId
   Return $sTxt_File_Name
EndFunc




;====================================================================================================================
;                  FUNCTION DISCRIPTION: CREATE IE OBJECT AND CHECK ERROR
;                  RETURN			   :
;====================================================================================================================
Func IECreate_Check_Error ($sLink, $bAttach, $bVisible, $bWait, $bTakeFocus)
   Do
	  ;MsgBox (0, "", "Open: " & $sLink)
	  Local $oIE = _IECreate($sLink, $bAttach, $bVisible, $bWait, $bTakeFocus)
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
   Do
ConsoleWrite ("Begin navigate" & @CRLF)
	  __IELockSetForegroundWindow($LSFW_LOCK)
	  _IENavigate ($oIE, $sLink, 1)
	  __IELockSetForegroundWindow($LSFW_UNLOCK)
ConsoleWrite ("finsih loading"& @CRLF)
	  If @error <> 0 Then
		   MsgBox(0, "Error", "There was a problem opening webpage!  " & @error)
	  EndIf
   Until @error = 0
ConsoleWrite ("Check IE Object: " & IsObj ($oIE)& @CRLF)
EndFunc

