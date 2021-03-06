#include <MsgBoxConstants.au3>
#include <FileConstants.au3>
#include <InetConstants.au3>

#include <Clipboard.au3>
#include <IE.au3>
#include <Excel.au3>
#include <WinAPIFiles.au3>
#include <ImageSearch.au3>


; Set Hotkey for the program
HotKeySet("{ESC}", "Autoit_Exit")
Erwin ()
while (1)
WEnd

Func Autoit_Exit ()
   Exit
EndFunc

Func Initial_Values (ByRef $iCount_Year, Byref $iCount_Model)
   $hFileOpen = FileOpen(@ScriptDir & "\Final Values File.txt", $FO_READ)
   $iCount_Year = FileReadLine($hFileOpen, 2)
   $iCount_Model = FileReadLine($hFileOpen, 6)
   Local $sFilePath = FileReadLine($hFileOpen, 9)
   FileClose($hFileOpen)

   Local $bFlag = 1;
   If $iCount_Year = "" Then $iCount_Year = 0
   If $iCount_Model = "" Then
	  $iCount_Model = 0
	  $bFlag = 0
   EndIf
   If $iCount_Model <> "" Then DirRemove ($sFilePath, $DIR_REMOVE )
   Return $bFlag
EndFunc


Func Erwin ()
   MsgBox (0,"","Begin")
   Local $sTemp
   ;----------------------------------------------------------
   ;MỞ TRANG VÀ LẤY OBJECT
   Local $oIE = _IEAttach ("erWin Online")
   WinActivate ("erWin Online | Audi of America | erWin Online - Internet Explorer")
   ;----------------------------------------------------------
   ;KHAI BÁO CÁC BIẾN ĐẾM CHO VÒNG LẶP
   Local $iCount_Year, $iCount_Model, $iCount_Cate1, $iCount_Cate2
   ;----------------------------------------------------------
   ;LẤY GIÁ TRỊ BẮT ĐẦU
   Local $bFlag = Initial_Values ($iCount_Year, $iCount_Model)
   ;----------------------------------------------------------
   ;FOLDER CHUNG
   Local $sFilePath_erWin_Download = @ScriptDir & "\" & "erWin Download"
   If FileExists ($sFilePath_erWin_Download) = 0 Then DirCreate($sFilePath_erWin_Download)
   ;----------------------------------------------------------
   ;ĐOẠN CODE CHỌN YEAR
   ;Lấy các option year
   Local $oSelect_Year = _IEGetObjById ($oIE, "f_modelYear")
   Local $aYear_Options = Get_Option ($oSelect_Year)
   ;Vòng lặp chọn year
   While $aYear_Options [1][$iCount_Year] <> ""
	  ;Tao folder Year
	  $sTemp = Standardize_File_Name ($aYear_Options [2][$iCount_Year])
	  Local $sFilePath_Year = $sFilePath_erWin_Download & "\" & $sTemp
	  If FileExists ($sFilePath_Year) = 0 Then DirCreate($sFilePath_Year)
	  ;Chọn year
	  $oSelect_Year = _IEGetObjById ($oIE, "f_modelYear")
	  _IEFormElementOptionSelect ($oSelect_Year, $aYear_Options [1][$iCount_Year])
	  _IELoadWait ($oIE)
	  ;Write Log file
	  Write_Log_File (@ScriptDir,  @CRLF & @CRLF & $aYear_Options [2][$iCount_Year], "append")
	  ;----------------------------------------------------------
	  ;ĐOẠN CODE CHỌN MODEL
	  ;Lấy các option model
	  Local $oSelect_Model = _IEGetObjById ($oIE, "f_cartypeId")
	  Local $aModel_Options = Get_Option ($oSelect_Model)
	  ;Vòng lặp chọn model
	  If $bFlag = 0 Then $iCount_Model = 0
	  $bFlag = 0
	  While $aModel_Options [1][$iCount_Model] <> ""
		 ;Tao folder Model
		 $sTemp = Standardize_File_Name ($aModel_Options [2][$iCount_Model])
		 Local $sFilePath_Model = $sFilePath_Year & "\" & $sTemp
		 If FileExists ($sFilePath_Model) = 0 Then DirCreate($sFilePath_Model)
		 ;Chọn model
		 $oSelect_Model = _IEGetObjById ($oIE, "f_cartypeId")
		 _IEFormElementOptionSelect ($oSelect_Model, $aModel_Options [1][$iCount_Model])
		 _IELoadWait ($oIE)
		 ;Write Log file
		 Write_Log_File (@ScriptDir,  "     " & $aModel_Options [2][$iCount_Model], "append")
		 ;Write final values file to use as initial values in case running the script again
		 Write_Final_Values_File (@ScriptDir,  "$iCount_Year: " & @CRLF & $iCount_Year & @CRLF &  $aYear_Options [2][$iCount_Year] & @CRLF & @CRLF & "$iCount_Model:" & @CRLF & $iCount_Model & @CRLF & $aModel_Options [2][$iCount_Model] & @CRLF & @CRLF & $sFilePath_Model, "overwrite")
		 ;----------------------------------------------------------
		 ;ĐOẠN CODE CHỌN CATEGORY 1
		 ;Lấy các option category 1
		 Local $oSelect_Cate1 = _IEGetObjByName ($oIE, "mainTopicCode")
		 Local $aCate1_Options = Get_Option_Cate1 ($oSelect_Cate1)
		 ;Vòng lặp chọn category 1
		 $iCount_Cate1 = 0
		 While $aCate1_Options [1][$iCount_Cate1] <> ""
			$sTemp = Standardize_File_Name ($aCate1_Options [2][$iCount_Cate1])
			;Tao folder Cate1
			Local $sFilePath_Cate1 = $sFilePath_Model & "\" & $sTemp
			If FileExists ($sFilePath_Cate1) = 0 Then DirCreate($sFilePath_Cate1)
			;Chọn category 1
			$oSelect_Cate1 = _IEGetObjByName ($oIE, "mainTopicCode")
			_IEFormElementOptionSelect ($oSelect_Cate1, $aCate1_Options [1][$iCount_Cate1])
			_IELoadWait ($oIE)
			;Write Log file
			Write_Log_File (@ScriptDir,  "          " & $aCate1_Options [2][$iCount_Cate1], "append")
			;----------------------------------------------------------
			;ĐOẠN CODE CHỌN CATEGORY 2
			;Lấy các option category 2
			Local $oSelect_Cate2 = _IEGetObjByName ($oIE, "topicCode")
			If @error <> 7 Then ;Match
			   Local $aCate2_Options = Get_Option ($oSelect_Cate2)
			   ;Vòng lặp chọn category 2
			   $iCount_Cate2 = 0
			   While $aCate2_Options [1][$iCount_Cate2] <> ""
			   ;Tao folder Cate2
			   $sTemp = Standardize_File_Name ($aCate2_Options [2][$iCount_Cate2])
			   Local $sFilePath_Cate2 = $sFilePath_Cate1 & "\" & $sTemp
			   If FileExists ($sFilePath_Cate2) = 0 Then DirCreate($sFilePath_Cate2)
				  ;Chọn category 2
				  $oSelect_Cate2 = _IEGetObjByName ($oIE, "topicCode")
				  _IEFormElementOptionSelect ($oSelect_Cate2, $aCate2_Options [1][$iCount_Cate2])
				  _IELoadWait ($oIE)
				  ;Write Log file
				  Write_Log_File (@ScriptDir,  "               " & $aCate2_Options [2][$iCount_Cate2], "append")
				  ;Download files
				  Local $sDownloaded_Links = Download_Documents ($sFilePath_Cate2, $oIE)
				  ;Write Log file
				  Write_Log_File (@ScriptDir,  $sDownloaded_Links, "append")
				  $iCount_Cate2 = $iCount_Cate2 + 1
			   WEnd
			Else
			   ;Download files
			   Local $sDownloaded_Links = Download_Documents ($sFilePath_Cate1, $oIE)
			   ;Write Log file
			   Write_Log_File (@ScriptDir,  $sDownloaded_Links, "append")
			EndIf
			$iCount_Cate1 = $iCount_Cate1 + 1
		 WEnd
		 $iCount_Model = $iCount_Model + 1
	  WEnd
	  $iCount_Year = $iCount_Year + 1
   WEnd

EndFunc



;====================================================================================================================
Func Standardize_File_Name ($sFile_Name)
   ;Remove ký tự được biệt
   $sFile_Name = StringReplace ($sFile_Name, "/", " - ")
   $sFile_Name = StringReplace ($sFile_Name, "\", " - ")
   $sFile_Name = StringReplace ($sFile_Name, ":", " ")
   $sFile_Name = StringReplace ($sFile_Name, "*", " ")
   $sFile_Name = StringReplace ($sFile_Name, "?", " ")
   $sFile_Name = StringReplace ($sFile_Name, """", " ")
   $sFile_Name = StringReplace ($sFile_Name, "<", "--")
   $sFile_Name = StringReplace ($sFile_Name, ">", "--")
   $sFile_Name = StringReplace ($sFile_Name, "|", " - ")
   ;Chuyển 2 khoảng trắng thành 1 khoảng trắng
   While StringInStr ($sFile_Name, "  ", 0, 1) <> 0
		$sFile_Name = StringReplace ($sFile_Name, "  ", " ")
   WEnd
   Return $sFile_Name
EndFunc


;====================================================================================================================
;                  FUNCTION DISCRIPTION: DOWNLOAD DOCUMENTS FROM ARWIN BY MOUSE AND KEYBOARD
;				   INPUT               : $sFilePath, $sLink
;                  OUTPUT              :
;====================================================================================================================
Func Download_Control ($sFilePath, $sLink)
   MouseMove (1,1,5)
   Local $x, $y
   _ClipBoard_SetData ($sLink)
   Local $iTimeout = 30
   While $iTimeout > 20
	  $iTimeout = 0
	  Sleep (500)
	  WinActivate ("erWin Online | Audi of America | erWin Online - Internet Explorer")
	  Send ("^t")
	  Sleep (1000)
	  Send ("^v")
	  Sleep (1000)
	  Send ("{ENTER}")
	  Do
		 Local $Search = _ImageSearch (@ScriptDir & "\Save_As_Button.bmp", 1, $x, $y, 0)
		 Sleep (1000)
		 $iTimeout = $iTimeout + 1
	  Until $Search = 1 Or $iTimeout > 20
	  If $iTimeout > 20 Then
		 WinActivate ("New tab - Internet Explorer")
		 Sleep (500)
		 Send ("^w")
		 Sleep (1000)
	  EndIf
   WEnd



   Local $Search = _ImageSearch (@ScriptDir & "\Save_As_Button.bmp", 1, $x, $y, 0)
   If $Search = 1 Then
	  MouseClick ("left", $x, $y, 1, 10)
	  Sleep (500)
	  Send ("{DOWN}")
	  Sleep (500)
	  Send ("{ENTER}")
	  Sleep (500)
	  ControlClick ("Save As", "", "ToolbarWindow325", "left" ,1 , 5)
	  Sleep (500)
	  Send ($sFilePath)
	  Sleep (3000)
	  Send ("{ENTER}")
	  Sleep (500)
	  ControlClick ("Save As", "", "Button2", "left")
	  Sleep (500)
	  WinActivate ("New tab - Internet Explorer")
	  Sleep (500)
	  Send ("^w")
	  Sleep (500)
   EndIf
EndFunc




;====================================================================================================================
;                  FUNCTION DISCRIPTION: DOWNLOAD DOCUMENTS FROM ARWIN
;				   INPUT               : $oIE
;                  OUTPUT              :
;====================================================================================================================
Func Download_Documents ($sFilePath, $oIE)
   ;ĐOẠN CODE LẤY TEXT VÀ LINK PROCEDURE TRONG TAG <A>
   Local $oAs = _IETagNameGetCollection($oIE, "a")
   Local $sDownloaded_Links = ""
   For $oA In $oAs
	  If $oA.title = "Download document" Then
		 Download_Control ($sFilePath, $oA.href)
		 Sleep (200)
		 $sDownloaded_Links = $sDownloaded_Links & "                    " & $oA.href
	  EndIf
   Next
   If $sDownloaded_Links = "" Then $sDownloaded_Links = "                     No download links here !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
   Return $sDownloaded_Links
EndFunc



;====================================================================================================================
;                  FUNCTION DISCRIPTION: GET OPTION FROM A SELECT FORM
;				   INPUT               : $oSelect
;                  OUTPUT              :
;====================================================================================================================
Func Get_Option ($oSelect)
   ;------------------------------------------------------------------------------------------------------------------
   ;ĐOẠN CODE LẤY TEXT VÀ LINK PROCEDURE PARTS TRONG TAG <A>
   Local $oOptions = _IETagNameGetCollection($oSelect, "option")
   Local $Text = ""
   Local $aOptions [3][1000]
   Local $iCount = 0
   For $oOption In $oOptions
	  If $oOption.value <> "" Then
		 $aOptions [1][$iCount] = $oOption.value
		 $aOptions [2][$iCount] = $oOption.innertext
		 $iCount = $iCount + 1
	  EndIf
   Next
   Return $aOptions
EndFunc



Func Get_Option_Cate1 ($oSelect)
   ;------------------------------------------------------------------------------------------------------------------
   ;ĐOẠN CODE LẤY TEXT VÀ LINK PROCEDURE PARTS TRONG TAG <A>
   Local $oOptions = _IETagNameGetCollection($oSelect, "option")
   Local $Text = ""
   Local $aOptions [3][1000]
   Local $iCount = 0
   For $oOption In $oOptions
	  If $oOption.innertext = "Repair Manual" Then
		 $aOptions [1][$iCount] = $oOption.value
		 $aOptions [2][$iCount] = $oOption.innertext
		 $iCount = $iCount + 1
	  EndIf
   Next
   Return $aOptions
EndFunc




;====================================================================================================================
;                  FUNCTION DISCRIPTION: CREATE LOG FILE
;				   INPUT               : $sFilePath,  $sTxt, $sMode
;                  OUTPUT              : AN LOG FILE IN $sFilePath
;====================================================================================================================
Func Write_Log_File ($sFilePath,  $sTxt, $sMode)
   If $sMode = "overwrite" Then
	  Local $hFileOpen = FileOpen ($sFilePath & "\" & "Log File" & ".txt",$FO_OVERWRITE)
   Else
	  Local $hFileOpen = FileOpen ($sFilePath & "\" & "Log File" & ".txt",$FO_APPEND)
   EndIf
   FileWrite($hFileOpen, $sTxt & @CRLF)
EndFunc


;====================================================================================================================
;                  FUNCTION DISCRIPTION: CREATE LOG FILE
;				   INPUT               : $sFilePath,  $sTxt, $sMode
;                  OUTPUT              : AN LOG FILE IN $sFilePath
;====================================================================================================================
Func Write_Final_Values_File ($sFilePath,  $sTxt, $sMode)
   If $sMode = "overwrite" Then
	  Local $hFileOpen = FileOpen ($sFilePath & "\" & "Final Values File" & ".txt",$FO_OVERWRITE)
   Else
	  Local $hFileOpen = FileOpen ($sFilePath & "\" & "Final Values File" & ".txt",$FO_APPEND)
   EndIf
   FileWrite($hFileOpen, $sTxt & @CRLF)
EndFunc



