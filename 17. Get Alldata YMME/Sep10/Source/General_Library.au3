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
;                  FUNCTION DISCRIPTION: CREATE HTML FILE FROM
;				   INPUT               : $sFilePath, $sTxt_Title,$HTML_body
;                  OUTPUT              : AN HTML FILE IN $sFilePath
;====================================================================================================================
Func Create_HTML ($sFilePath, $sTxt_Name, $sTxt_Title, $HTML_body)
   Local $bOL_Flag = False
   If StringInStr ($HTML_body, "<OL>") <> 0 Then  $bOL_Flag = True
   Local $sHTML = ""
   If $bOL_Flag = False Then
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
   Else
	  $sHTML &= $HTML_body
   EndIf
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









