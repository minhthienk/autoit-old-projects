#include-once

#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <GuiEdit.au3>
#include <GuiComboBox.au3>
#include <Clipboard.au3>



#include "General_Library.au3"
#include "Create_JAVASCRIPT_Procedure.au3"
#include "Create_NORMAL_Procedure.au3"
#include "Create_DTC.au3"
#include "Get_All_DTC_Links.au3"
#include "Add_Procedure.au3"
#include "Permission.au3"
#include "GUI.au3"

;====================================================================================================================
;                  FUNCTION DISCRIPTION: SUBMIT ENTERED KEY THEN CHECK AND INFORM USER
;				   RETURN              :
;====================================================================================================================
Func KeySubmit ()
   Notification ("Please wait for the server to check your key ...", "Normal")
   ;-------------------------------------
   ;LẤY KEY ĐƯỢC NHẬP VÀO
   $sKey = GUICtrlRead ($Input_Key)
   ;-------------------------------------
   ;CHECK KẾT QUẢ CỦA KEY TRONG BỘ NHỚ VÀ TRÊN SERVER
   $sKey_Result = Key_Check ($sKey)
   ;-------------------------------------
   ;CHECK CÁC TRƯỜNG HỢP CỦA KEY

   ;Key không hợp lệ
   If $sKey_Result = "Invalid"  Then
	  Notification ("The KEY you entered is INVALID" & @CRLF & "Please enter a VALID KEY!", "Normal")
	  MsgBox (0, "Message", "The KEY you entered is INVALID" & @CRLF & "Please enter a VALID KEY!")
   Else
	  GUICtrlSetData ($Key_Warning, @CRLF & "Bạn có " & Round ($iRemaining_Hours,3) & " giờ sử dụng" & @CRLF & "từ " & Get_Time ())
	  Notification ("You have " & Round ($iRemaining_Hours,3) & " hour(s) left", "Normal")
	  MsgBox (0, "", "You have " & Round ($iRemaining_Hours,3) & " hour(s) left")
	  $iUse_Hours = Number ($sKey_Result)
	  $bTotal_Allow_Flag = True
   EndIf
   ;-------------------------------------
   ;TRẢ LẠI GIÁ TRỊ CỦA FLAG ĐỂ NGẮT NÚT BẤM
   $bKeySubmit_Flag = False
EndFunc





;====================================================================================================================
;                  FUNCTION DISCRIPTION: CHECK VALID KEY OR NOT
;				   RETURN              :
;====================================================================================================================
Func Key_Check ($sKey)
   Local $sResult = "Invalid"
   Local $aKeys = Keys_Storage ()
   For $vElement In $aKeys
	  If $sKey = $vElement Then
		 $sResult = Key_Check_On_Server ($sKey)
		 ExitLoop
	  Else
		 $sResult = "Invalid"
	  EndIf
   Next
   Sleep (500)
   Return $sResult
EndFunc



;====================================================================================================================
;                  FUNCTION DISCRIPTION: CHECK VALID KEY OR NOT
;				   RETURN              :
;====================================================================================================================
Func Key_Check_On_Server ($sKey)
   Local $iMax_Time = 12
   Local $sResult = "Key not on server"
   Local $sLink_Expired_Keys = "https://docs.google.com/forms/d/e/1FAIpQLSdCsPkt2Add3RP5FU4ZVdYuVjUpdWZC8ODC1QcboGN3vqGpkw/viewanalytics"
   ;-----------------------------
   ;OPEN KEYS LINK AND GET CONTENT
   Local $oIE = IECreate_Check_Error($sLink_Expired_Keys, 0, 0, 1, 0)
   Local $sContent = _IEPropertyGet ($oIE, "innertext")
   If StringInStr ($sContent, "Prepair Procedure Keys") <> 0  Then
	  ;-----------------------------
	  ;PUT KEY AND DATE INTO AN ARRAY $aKeys_On_Server
	  Local $aContent = StringSplit ($sContent, @CRLf, $STR_ENTIRESPLIT)
	  Local $aKeys_On_Server[0]
	  Local $i = 0
	  For $i = 1 to $aContent[0]
		 If StringInStr ($aContent[$i], ">>>") <> 0 Then
			ReDim $aKeys_On_Server[UBound($aKeys_On_Server) + 1]
			$aKeys_On_Server[UBound($aKeys_On_Server) - 1] = $aContent[$i]
		 EndIf
		 $i += 1
	  Next
	  ;----------------------------
	  ;COMPARE USER ENTERED KEY AND KEY ON SERVER
	  $iRemaining_Hours = $iMax_Time
	  For $vElement In $aKeys_On_Server
		 If StringInStr ($vElement, $sKey) <> 0  Then
			$sResult = "Key on server"
			$iRemaining_Hours = $iMax_Time - (Time2Hour(Get_Time ()) - Time2Hour ($vElement))
			If $iRemaining_Hours <= 0 Then
			   $sResult = "Invalid"
			   $iRemaining_Hours = 0
			   ExitLoop
			EndIf
		 EndIf
	  Next
	  ;----------------------------
	  ;UPDATE EXPIRED KEY TO SERVER
	  If $iRemaining_Hours = $iMax_Time Then Update_Expired_Key ($sKey)
   Else
	  $sResult = "Invalid"
	  MsgBox (0, "", "")
	  $iRemaining_Hours = 0
   EndIf
   Return $sResult
EndFunc






Func Time2Hour ($sTime)
   ;-------------------------------------------
   ;LẤY GIÁ TRỊ THỜI GIAN TRONG STRING
   Local $sYear = StringRight ($sTime, 4)
   Local $sMonth = StringMid ($sTime, StringInStr ($sTime, "/", 0, -1) - 2, 2)
   Local $sDate = StringMid ($sTime, StringInStr ($sTime, "/", 0, -2) - 2, 2)
   Local $sSecond = StringMid ($sTime, StringInStr ($sTime, ":", 0, -1) + 1, 2)
   Local $sMinute = StringMid ($sTime, StringInStr ($sTime, ":", 0, -1) - 2, 2)
   Local $sHour = StringMid ($sTime, StringInStr ($sTime, ":", 0, -2) - 2, 2)
   ;-------------------------------------------
   ;CHUYỂN ĐỔI CÁC ĐƠN VỊ THỜI GIAN VỀ GIỜ
   ;Đổi năm thành giờ
   Local $iYear2Hour = 0
	  If Mod($sYear, 4) <> 0 Then
		 $iYear2Hour = Number ($sYear - 2018)*365*24
	  Else
		 $iYear2Hour = Number ($sYear - 2018)*366*24
	  EndIf
   ;Đổi tháng thành giờ
   Local $iMonth2Hour = 0
	  If Mod($sYear, 4) <> 0 Then
		 Switch $sMonth
			Case "01"
			   $iMonth2Hour = 0
			Case "02"
			   $iMonth2Hour = 31*24*1
			Case "03"
			   $iMonth2Hour = 31*24*1 + 28*24*1
			Case "04"
			   $iMonth2Hour = 31*24*2 + 28*24*1
			Case "05"
			   $iMonth2Hour = 31*24*2 + 28*24*1 + 30*24*1
			Case "06"
			   $iMonth2Hour = 31*24*3 + 28*24*1 + 30*24*1
			Case "07"
			   $iMonth2Hour = 31*24*3 + 28*24*1 + 30*24*2
			Case "08"
			   $iMonth2Hour = 31*24*4 + 28*24*1 + 30*24*2
			Case "09"
			   $iMonth2Hour = 31*24*5 + 28*24*1 + 30*24*2
			Case "10"
			   $iMonth2Hour = 31*24*5 + 28*24*1 + 30*24*3
			Case "11"
			   $iMonth2Hour = 31*24*6 + 28*24*1 + 30*24*3
			Case "12"
			   $iMonth2Hour = 31*24*6 + 28*24*1 + 30*24*4
		 EndSwitch
	  Else
		 Switch $sMonth
			Case "01"
			   $iMonth2Hour = 0
			Case "02"
			   $iMonth2Hour = 31*24*1
			Case "03"
			   $iMonth2Hour = 31*24*1 + 29*24*1
			Case "04"
			   $iMonth2Hour = 31*24*2 + 29*24*1
			Case "05"
			   $iMonth2Hour = 31*24*2 + 29*24*1 + 30*24*1
			Case "06"
			   $iMonth2Hour = 31*24*3 + 29*24*1 + 30*24*1
			Case "07"
			   $iMonth2Hour = 31*24*3 + 29*24*1 + 30*24*2
			Case "08"
			   $iMonth2Hour = 31*24*4 + 29*24*1 + 30*24*2
			Case "09"
			   $iMonth2Hour = 31*24*5 + 29*24*1 + 30*24*2
			Case "10"
			   $iMonth2Hour = 31*24*5 + 29*24*1 + 30*24*3
			Case "11"
			   $iMonth2Hour = 31*24*6 + 29*24*1 + 30*24*3
			Case "12"
			   $iMonth2Hour = 31*24*6 + 29*24*1 + 30*24*4
		 EndSwitch
	  EndIf
   ;Đổi ngày thành giờ
   Local $iDate2Hour = Number ($sDate-1)*24
   ;Đổi giờ thành giờ
   Local $iHour2Hour = Number ($sHour)*1
   ;Đổi phút thành giờ
   Local $iMinute2Hour = Number ($sMinute)/60
   Local $iTime2Hour = $iYear2Hour + $iMonth2Hour +  $iDate2Hour + $iHour2Hour + $iMinute2Hour
   Return $iTime2Hour
EndFunc


;====================================================================================================================
;                  FUNCTION DISCRIPTION: UPDATE EXPIRED KEY
;				   RETURN              :
;====================================================================================================================
Func Update_Expired_Key ($sKey)
   Local $sLink_Update_Used_Key = "https://docs.google.com/forms/d/e/1FAIpQLSdCsPkt2Add3RP5FU4ZVdYuVjUpdWZC8ODC1QcboGN3vqGpkw/viewform"
   Local $oIE = IECreate_Check_Error($sLink_Update_Used_Key, 0, 0, 1, 0)
   ;Lấy object form
   Local $oForm = _IEFormGetObjByName($oIE, "mG61Hd")
   ;Lấy object input
   Local $oLoginName = _IEFormElementGetObjByName($oForm, "entry.57893756")
   ;Set input
   _IEFormElementSetValue($oLoginName, ">>> " & $sKey & " <<< --- " & Get_Time ())
   ;Submit form, no wait for page load to complete
   _IEFormSubmit($oForm, 0)
   ;Wait for the page load to complete
   _IELoadWait($oIE)
   _IEQuit ($oIE)
   Sleep (500)
EndFunc



;====================================================================================================================
;                  FUNCTION DISCRIPTION: GET TIME FROM INTERNET
;				   RETURN              :
;====================================================================================================================
Func Get_Time ()
   Local $sLink_Time = "http://ngaygio24.com/xem-gio-viet-nam.html"
   Local $oIE = IECreate_Check_Error($sLink_Time, 0, 0, 1, 0)
   Local $oTimeID = _IEGetObjById ($oIE, "boxClock")
   Local $sTime = _IEPropertyGet ($oTimeID, "innertext")
   ;-------------------------------------
   ;CHỈNH SỬA STRING CHO HỢP LÝ
   $sTime = StringReplace ( $sTime, @CR, "")
   $sTime = StringReplace ( $sTime, @LF, "")
   $sTime = Standardize_String ($sTime)
   ;-------------------------------------
   _IEQuit ($oIE)
   Return $sTime
EndFunc



;====================================================================================================================
;                  FUNCTION DISCRIPTION: KEYS
;				   RETURN              :
;====================================================================================================================
Func Keys_Storage ()
   Local $aKeys = [ _
   "jiwYrvxBXeW0J8V", _
   "6puSL7rcudsebFz", _
   "gklaO0JWDzG8UrU", _
   "r0csmMSpolEYzj0", _
   "RFGRirAsXSkopuL", _
   "pRQUg1iQfhMbRPU", _
   "DL2Nee61KEPLfpo", _
   "ybLsJ2xbjDJFslV", _
   "tkFqUuaiTSPvH98", _
   "63WSPpEuYIc7cnl", _
   "b63WXfjhVoxHfzl", _
   "NfjJSRdJ5RGxcJh", _
   "r87KqVZETuhbOk9", _
   "FmKSMRtZorksNXg", _
   "QIrQS23bFAWuChS", _
   "Wo9eE1SRTOad9ZF", _
   "dVteI3oQT966Daa", _
   "GCXWfAfRu6Y0oxr", _
   "CuJkckEJwFahmw1", _
   "tpBXlbcmcDVQaNA", _
   "SnrBTU5yOv5el16", _
   "Bq0AijxsnTC5foO", _
   "NXOrwxW4ZLYoG9r", _
   "4ApU0fiszXn5hhJ", _
   "9c3uVhTEm9Yu0J8", _
   "fQDbSS1Df0zSCq6", _
   "Hyl23ZmHItzmDJp", _
   "NWpcc0box7tQ9FD", _
   "lcOwt4re1auYUy1", _
   "6CmkWhsBavwbWHl", _
   "PzSEF2cBJOONPCn", _
   "pItzj3w5yyVbBxi", _
   "FbeHUBCIBQHMUna", _
   "xi2up8O2HskkfcR", _
   "teZ7sPBXbcNPoK0", _
   "FBmORWnT44JJrSw", _
   "4pj8GKNajVE8yE7", _
   "t4TnCLw8vzPYckW", _
   "FTCmSK7AO6igjPq", _
   "uDjzbaSeLXBLx0f", _
   "obBUJSYv11aHJgW", _
   "iKmq48rPqDdUP17", _
   "2jNT4sOG4pwt6Ra", _
   "Y4SGwbpSurfUPGr", _
   "Xc8nA2E5KFhvODX", _
   "6Z5MemFwPC3lQqV", _
   "ZKDiQH8xN9KgbOq", _
   "JaJFbSTYAswGNn1", _
   "VfWfQKgbYmc0JSr", _
   "iNXROBU8aSqOLxU", _
   "6C0IqLMQ8Qdb08e", _
   "ObIVcrL2V6iU8CK", _
   "GnzS91wzVgfX8PY", _
   "HKWFqeJjPDv9nUH", _
   "UIvdaz56PXhcc9O", _
   "fmWxv1e1Te1nfBt", _
   "jHvqIv7t8gtgb8R", _
   "cZaA6eyVrH4brDC", _
   "tn23w2QMdveLwil", _
   "oiuCJKTQka1BFd1", _
   "c5GXSYkhGZmyTBk", _
   "Itq7PeJbF6NhT7J", _
   "lsCzMvzBLt6u5eN", _
   "Bm9rEljECznsoEe", _
   "oBKfcYu8eIFBA4x", _
   "2pUYBC0V296rwoS", _
   "nyqeBEeLVEuTdRc", _
   "fGvWGGEHjW64EVb", _
   "zov7qvp8kVy9vbQ", _
   "0GJ6XsubzA3g5qb", _
   "1Rp8tGQIfmSBNZV", _
   "rjhXEUkLX9qy9wX", _
   "4E1AjmWdVWeyy69", _
   "8lR0wqfFrxifTy6", _
   "8qbHg7zX9aWjphq", _
   "IRtNajbKUgsOY4c", _
   "qiIJa6otX9gpaUu", _
   "u9RK2yak8FEcHkY", _
   "tlqimrA9HHeeQKk", _
   "9Kh80VNEPOMwZeA", _
   "cdYUWEO80VzFpco", _
   "3syFfhKN9xyEqyJ", _
   "hrJyxwHFRXN5Gc4", _
   "yBniPZxQ18uK3SN", _
   "QyluBFFVU0kG83s", _
   "biZb3jUqP7G7miR", _
   "dyzdB9Lf0Ahms4u", _
   "Yym6eTR0x24nS28", _
   "dadQma0oEX2LMG9", _
   "I5nABVF4ZYwVOZ6", _
   "QEDJgwUF4jOZzBu", _
   "ut3jwgaGzttP23z", _
   "lPINVCKZb1rdg17", _
   "mv46IQEuo87ud2j", _
   "RnYNLUlqshth1Ai", _
   "gsaMstQ3PLenzwN", _
   "BAft7nFybiBy79h", _
   "jzYWzCk2FYo39CY", _
   "Ijvf6bJQ3juLceL", _
   "0Yqi17Zwn9HZfkJ", _
   "d1UsmfkmJ956NTG", _
   "VSwf2W3j7l0OLCV", _
   "eMrpbGKXHPYfljb", _
   "VTvXPVO1Mf89UcU", _
   "C43VhJcE5iclFLq", _
   "eq3erSiYZ8Dgzdj", _
   "XGnyMj0d2cnX8kl", _
   "X2OrxktshU518xi", _
   "T6vqYmO82o8n7IM", _
   "Gkc5uBT0qgVuKJD", _
   "IuRbezG7b8XDQhm", _
   "fct1bw7xDzNVU4v", _
   "JmEVv5cTUSzYE7l", _
   "pwWBNmKElIN6rxg", _
   "jJH8BsF1YWpisVV", _
   "h3XEHEufJ7guQ8C", _
   "6LwoYj4nv49RcJU", _
   "cJNhmobuZl5qMCN", _
   "vqyPvzDXBM3ArLK", _
   "fR9UUE5np75OWak", _
   "u1kKdtJjxOHTrhw", _
   "fBJ0TPb8gRoYpNp", _
   "IStCiNHZO6NV1Yo", _
   "cDyhWE0fYdVPp3s", _
   "KQ4EBdmzwoy7hBD", _
   "J7xFtfp0UPPu2Ll", _
   "EweUB577gLlV0CR", _
   "ZNk0ziyB42tUpm2", _
   "5X8DR7iW5S8gwvd", _
   "net5Ficy8br1Hm6", _
   "BfCx8lUqVV9cicK", _
   "lvSwKQTEE34f6Nw", _
   "ETEz3tSvMVVZleg", _
   "EEL8spY4pB345bC", _
   "131SXX0hygffVqR", _
   "YmUUa6dGC3cSfhQ", _
   "yrv3HUWA0tbVryC", _
   "QFU2Lt7vBnJGYZM", _
   "phOiBOGKrIiOiEL", _
   "WnzS82xUe1BpGFJ", _
   "wlnkK1JGIAnvdYB", _
   "LXO9soXQ3NqKiWF", _
   "t50TiT7VCncdWZJ", _
   "sdFQW3Bh3qeELQQ", _
   "LMo4fuypjmVpH3B", _
   "v2tTwnh9xwYbwfo", _
   "FyNx8XlJUgEZ8z7", _
   "kQivnSovwdRtDaE", _
   "iEybqsJZ2KhRjR9", _
   "s6Djsr0M93kcGLG", _
   "LgTAxCLafxsWqar", _
   "dGROi1CzEcqvpVx", _
   "3xgd3H1DpVVArjp", _
   "LhzeFvl6fjISc9s", _
   "zswZzPY1isf9ETl", _
   "5964dCTN42Pgkyb", _
   "aIHCj9HPv6GvTxl", _
   "1UT7MdGPVeoJbfo", _
   "lgBlLK5aMAD0d7t", _
   "4kYa6n8bZZFKimS", _
   "Nr2Mta6mIZRWSvl", _
   "GolVWsqnoV54HJw", _
   "C69YkWSTOuCJKsy", _
   "ycZxD1lAvix9P3K", _
   "A4aOmLxqcEcHdrC", _
   "qJ4Oza1Jyizpedo", _
   "KaWSJVMGkdbBvq4", _
   "CBAgsokZ0oBpdU0", _
   "NtktjetwfMgNxFi", _
   "9XZrPiOdFGuKqIu", _
   "zDTlagswsCLVD08", _
   "lmIOKLyKOuNr7ur", _
   "FPOSL674TumwrSX", _
   "ADz2FjAwqLXCfz2", _
   "RrAYsBHRKBq1P9R", _
   "554ZGZsqd7IEBgY", _
   "Ic1myFACnBrJAkm", _
   "K4MtUVoWeD4LpTe", _
   "ukec2pZXmCjEMiY", _
   "U4afH2tb5vr0BFf", _
   "6ytWZ5VB7uL1uFK", _
   "w1j3HkbmOXOON1H", _
   "EOIdyi8WRUsPgvM", _
   "XflTIrtz8rHyJzC", _
   "SRrgXXHChY6pjhO", _
   "2UJ4zOXaYAESrLU", _
   "GyQlNdHVtdTuJac", _
   "lU8BapEm4sl3KoT", _
   "epMBjL4aSrpd7ve", _
   "4EeOBCmGj3ovfVX", _
   "5ANn5oVqnB1CgUY", _
   "ZMVXNpXfqSfM8vq", _
   "5JRaBiKTbIFJRSJ", _
   "ArgnjNwzA6Wl8YQ", _
   "hef4EUyKK5kdM0r", _
   "BmoOl1Vs5prwX37", _
   "hwSYbKeBNsUaTEz", _
   "KXgPj42QT67a0V9", _
   "FelZ9NPb4XXNaPE", _
   "c3jWN8zvK67OKxd", _
   "YnQTsgqCCYXRzAj", _
   "ZxMwONsjhRGDX8x", _
   "Ug2E03xV2PkUakz", _
   "AfnasprKjUxGheL", _
   "NBwbNODc1K8jp8D", _
   "aK3tHRQwNkOdxhG", _
   "jnaAPrxk4wwkiKy", _
   "WlkuUkOMMqozDWF", _
   "Mb9eNrWjAhQqKjb", _
   "BcpXA66tvH0A7gE", _
   "TTOnlXnQOuK6D3Y", _
   "ME0750sDzYyDLRp", _
   "7yJWy40DKcWgq6A", _
   "p2w6gbxApFzoBKy", _
   "HJwEavt4D2c5H0F", _
   "8mKMbNQlzofsvIH", _
   "Gm8PDDlOhbeIvSX", _
   "8tFB8KnYeASoMLY", _
   "QvR10CdtjOREz9f", _
   "atXCT643Rd55ukg", _
   "qajtsfSna24rvcO", _
   "K9MOeApHDxndYmF", _
   "p2gE02n6VsXeB9Q", _
   "rRlDfaDGaKRqgae", _
   "lkgVow8HZF4ykUL", _
   "O1TG1ZKAPk8M239", _
   "ITvC6pLQcCvKbMX", _
   "fuRx44iu7dDPLlT", _
   "PG52nKYsWqx2LEF", _
   "9NPYgVsmE34nxIl", _
   "cnivRoHR9Rjyj5D", _
   "fdAGZO6abToTQ7P", _
   "0vXFw9XugJnMy6E", _
   "byZtaIUXLF3neGZ", _
   "Bc81jqSkSRqGh6f", _
   "KDwplORgyWsTatK", _
   "HDK4TiGHYxqGLks", _
   "JngVLJ1ayPakRiV", _
   "dGhe40mj4rHqOQy", _
   "DzlQHarBGzd9iWw", _
   "xTdAn0jm7WtjhPZ", _
   "0BU438r923RVDqV", _
   "gLjM06IdE06JNvk", _
   "S19aWEC87AOV6Ub", _
   "3qye3vzdPrs35dm", _
   "rC6gD7LL5nO56nk", _
   "FGIzh50xJDphLi8", _
   "3aky4hxh7Itv4v6", _
   "LC1onGM7hHOsRPv", _
   "VEvvWaSvSVLENfv", _
   "jIW5UcjBHVjbgjn", _
   "ocPZBoXFvNQ5C6O", _
   "f8uQ2R8QOqdTOqt", _
   "AVCuFzZY4BXdHaM", _
   "WqTQOcxuQuJ8Rvj", _
   "F4Salq7XICZ7GIB", _
   "HkbEawL7nnWgeUG", _
   "uRnE64H7HF1N7LO", _
   "Bk8ack6TCSTIUy5", _
   "pwvOY96jU8JtC8C", _
   "2cyWBTAPdg0Adzo", _
   "KxsGEcojfs8kwan", _
   "6tOm3GwcLcog9Sp", _
   "DMXN6gCr8lA9cfj", _
   "ZzMox1KazELn0MG", _
   "UQlLQnpYgQMy2a4", _
   "s46JMNhBHwp4QtB", _
   "k2qZykoWwHIApeH", _
   "TeX3kmD4geAT9Yf", _
   "c12atQZoonJLOgi", _
   "bZolL4olhGqYKPw", _
   "k62G5d72mgCFdMY", _
   "VOZPuuzrrHrM95A", _
   "KrMRKEnK6Y8V3rZ", _
   "hzBe1SJnKqbIIw0", _
   "Ow4OyEB5gUPsCgX", _
   "3bsr8diL2x5eUtz", _
   "neqIfeXjiAjqWeL", _
   "2EeFevkr21phVl0", _
   "2RB2k2EomvEOuBf", _
   "iVMQbzTTFbQ3vmq", _
   "qU9yeKDoHS1uI8f", _
   "Oj4HZKRFY1pX1vh", _
   "eU1A8n7VbnpEKBf", _
   "kXUc91pc7KUeWo0", _
   "AwagqS96Lj26jsh", _
   "AOepJOYvBa1H44W", _
   "umfH21hmyBcGRXr", _
   "ho0QNkW8Jec2UM7", _
   "nXSa7C8vxOMIoJb", _
   "aJ4zLwxome0ehq5", _
   "zyJwiMtkdaB20Gx", _
   "0l19gwEaLG91mpp", _
   "2hay2VeeGrzFLXy", _
   "nziTJdygyo7RbiF", _
   "rJ8203kK8BFB9lm", _
   "lBxn53lSt1nlZ9F", _
   "HE9ItKPKOY1J3sk", _
   "Txg9YNyKfkFYDQx", _
   "4xNqgM84onlNZOl", _
   "O2vIG3245Ob7BXr", _
   "Hv6tW9CnDIkJwR2", _
   "wKLRHgJ49X9iQLF", _
   "0OWE7q5SmZeRS6O", _
   "ZDui1OT0OpJWQzu", _
   "u6wxoPJkYrQf4Xt", _
   "Q17Wh9SRWOAZqpx", _
   "BqYKN5ZypvHMVM6", _
   "ErefvEehK0lc1v4", _
   "CauHWSNeEJq4aQS", _
   "2XIdt9V98jLM9Hw", _
   "n9yoVZwBhh3a1E6", _
   "c3j70oRzsFbTSKj", _
   "cVU4Xd7YG5JuAq9", _
   "vDIyUaxCWuetFrA", _
   "c9tYr1gRgRPk4dF", _
   "y56kTFmuZf66VoD", _
   "FqZ0VvBZ6HUiMqZ", _
   "GQJ2Hc8rQUy8PAG", _
   "HrpjXydpY7gYfda", _
   "WfcfIP7QcVoYAu0", _
   "iC6Pvvus7TibVxR", _
   "NPdVbP4n5rcNbDs", _
   "vfUaK5LRICvDIJz", _
   "ZXwwqMUxLdeDFTZ", _
   "ZJ8cLC6BmcxX5jV", _
   "rt440JewOiIZ5XY", _
   "PPaA38DdrdUKxc0", _
   "FepbXZXEDqQXKzp", _
   "tLqS3RSJUmlm6OM", _
   "7ToBdH9lQxbcga0", _
   "OXTdV9FkFx4NvoK", _
   "YsqrAfWumXGa2K6", _
   "M9b1gJ4fLuFTdBf", _
   "3uXqdhR8wCDNZIc", _
   "ORD7DmKagjryAfU", _
   "4MOFggOLm8MsZ9r", _
   "coTFwjjlK6X3qrd", _
   "sFfHD6RcRzwYjMh", _
   "EQO0JYPdVdcpeEj", _
   "DbO8kSzU659sUFr", _
   "Xc57yrg7Vj5iRRv", _
   "tA2SZVhRcSexqtJ", _
   "xpmElhIUiIqudCx", _
   "pU5eLDQXIMevkHC", _
   "0Qia2rx7m8qNxSg", _
   "JFCFE1wy0u5gBl3", _
   "mOw4HYiPU7XchkQ", _
   "OzuqMIIwXn88PkR", _
   "sq1n4KsCwQdaG1R", _
   "QgKXaXImaul6pNr", _
   "ay7BnoXF2PPzPPW", _
   "VDRJi7IzNOmLNIT", _
   "akopBxBFYg5IXQv", _
   "sjs0kWAS6d2Hb06", _
   "HdXFDW17MstOBL8", _
   "RvIUC45HgsFoP9o", _
   "QaVpFQVtp8kAwmH", _
   "luBxpEhOIoA62IC", _
   "fqd7WX4U8wcSfSm", _
   "bh59bo1Dbypsc6e", _
   "w1MmhJBXOlFaofo", _
   "0UzOqlpdXpUDRF3", _
   "SbtIyVU6NAZdK3J", _
   "1zBuES2KWmpr4Zv", _
   "Brs1iI1MRy5ymd6", _
   "12tE3ZnYquzBHwf", _
   "BKJ2IgJMYXKaReD", _
   "A0V34nFC4vORHzW", _
   "Mi8DnoD3wGNGDHi", _
   "ghy3yZ3JejSGJrY", _
   "Cs0MYhEOKjrc9wM", _
   "owwcwM4ExYDpKJV", _
   "qGl5iOSt2VPR736", _
   "SkOwuzEsi9AVEnT", _
   "1CKv5tz5pI16f3q", _
   "7IY4GHNaLoPja49", _
   "Ondy5NIPwevMa9F", _
   "WKw0TR73ip6ItBI", _
   "5UhHlh9lCXe9TOj", _
   "2oSLyXcbS8T7o8r", _
   "RgxW3hIp7VvZCyt", _
   "qceO3iArSiAi5XP", _
   "KZtITCGvDKhhZeR", _
   "MFEAS0KUM2zGFoE", _
   "Blwslyg8IBQ679D", _
   "JSqoBtvQZFiw8lU", _
   "bHQMxogGCrxqIhC", _
   "yXqphigAzscwxkI", _
   "KZkehEDRzVrg8il", _
   "7CmtLmcNZxq415l", _
   "tuqylzsXGTvRxLC", _
   "zwPrCFfZxl0AUHT", _
   "MS2WRoqxdQmqhkV", _
   "yPT9LLC4AF1mQ59", _
   "OGbO4ZXO6SOZHtR", _
   "peKnHn6tsdVZE4y", _
   "BmkKU3WIB9g90PD", _
   "exWAWPkyHYdH25l", _
   "VwfLMNMFvnRD22C", _
   "kjEPeYx8onD3kfk", _
   "8YyGAw6xh56IvzX", _
   "DRKQggZAWme7bYl", _
   "VSkZaniZAmktIbb", _
   "8dK97nKRpz31zGj", _
   "zJtkWvWq0nM08fN", _
   "B2H8t3UzFcRwYdi", _
   "fIKgAiKnq3ohmjJ", _
   "B1ZMwnjN6JcGg5B", _
   "wE0YXHdoIQaX3lq", _
   "Cd3z1ybcOltMoOh", _
   "ctkxGu1eEErixUW", _
   "4rCGHax7kGy1RnM", _
   "YW69NBY1XVDZVdE", _
   "FwAw5z8T9ooTyi0", _
   "HmwpCD9qwSYCaxB", _
   "rEnn3hIcHC49gv3", _
   "PYxr2YU9PfK8cQa", _
   "bRgBHmOME8UDV0x", _
   "r4xbu5u4DG3GGh5", _
   "rgFOmKS0J0D2shI", _
   "83O7nwWvOyoJhfn", _
   "sR6HLC22tLd5WRb", _
   "MxkeTSYcrcqgtsJ", _
   "4wBgbtJqjeqbXeb", _
   "BGUwAK1NDAdR67m", _
   "8nFut6yIVqzC7CW", _
   "WB5ggktX6BnBw9b", _
   "rwoLnrQ6BPw2PN0", _
   "tR9x1TpIpDpELw3", _
   "Sa6wIorBT457c94", _
   "1tkHcGaq38THdSS", _
   "P8dxLtyVPJ6sj8K", _
   "QB7aXaFYCULxiE3", _
   "Ak6oGA9YQks50QM", _
   "s17RfGk3IVKMFRf", _
   "AJ019UbnCSebn4j", _
   "krc9VkSBozIflP8", _
   "UUS8TzMB3ZokYLZ", _
   "dzEzlO6AG5blaxl", _
   "xWl5OImE1LWiumB", _
   "RCtR8wx2xmnjNwD", _
   "HXLlL4KM6bV0IBe", _
   "bjcxK2Y6YfllTks", _
   "zs0lMeLHEx2yH2D", _
   "2AmS4lpHT8NqPcT", _
   "U4vDhm39fI5VGdA", _
   "X9Y3zbcgiyXNSYv", _
   "50nSBaIxsmbNrf9", _
   "7zxfSKjfyoSFh7W", _
   "xiE5XJM2DR5ZNw0", _
   "x6hudlMnjFVzYBg", _
   "6bPefpqujMf93Yr", _
   "UpzgCGMv7FSbGCz", _
   "WynEZSi020v8HnT", _
   "oSCzG4Gdwn1Zpzo", _
   "PYNBdZ52bZXB8EX", _
   "8HwvBb541C3VR3K", _
   "8dLFYJ4Pj8xSTVy", _
   "fdXqUv72RsfNpu1", _
   "Xswe89UGgVxDYks", _
   "6ORDUnhS094xJ1f", _
   "Wrlg8bgEX8hITPR", _
   "Bpf9PywV9elBynw", _
   "brFXsAdceTWajrL", _
   "1MysrebgDX89qON", _
   "xixXjVAKw2P3SWB", _
   "dkBwggTlbVDhq18", _
   "2amp2Ozr10d0Hrz", _
   "ZAMEtlITxCeLdhG", _
   "AigjBFSQJ0Tpb9P", _
   "JW8f8h9I8vb5twM", _
   "ITooVeGBHgm5JQ1", _
   "FAl57aXcuEzRt9j", _
   "K8b1kJgxfJhZzOw", _
   "qAp72aW0Ekb4820", _
   "z4hX2PId5jo2Xh7", _
   "cI3NfoRvjSENg8J", _
   "C4i9KikmyaBI2GJ", _
   "A6vrngHMUhLnBxa", _
   "AgmAwcjn5QooPex", _
   "1qRbVSDSsFx73jI", _
   "5D5AnW2IquQhPFp", _
   "f9g1p1HAGKiasBD", _
   "UxBheoeRnFqQspG", _
   "Ku4BDcftAeICSNc", _
   "4D9OkMD0onIMoTs", _
   "gmlpHtxJ628MXL9", _
   "5Kq42nMfbLgEaRT", _
   "tebwpPzrQL7Whj5", _
   "0sKs1CloDAMkB5k", _
   "5UP4r6s5P1mBqVl", _
   "gHsoBfxW4p9Fbyw", _
   "i6TcWkAyM8IEGQB", _
   "ZImOVro7o0PMam3", _
   "u6GnQQyGxqhm3bQ", _
   "snCIqJDY9kVhR6Z", _
   "cnew7YNgg6yOdIZ", _
   "m5EKQBUk3Gmczeq", _
   "jBMzsP68eiEiGOK", _
   "x7Y6NPPdBzZeUCc", _
   "xBzGlnhGePoXs8S", _
   "lMTuoWY5fFjfStt", _
   "fa2wSSt7dTPOkn1", _
   "18pBTtghn0M1xfT", _
   "3H6KtsYLTOfv75V", _
   "mQjwsM0EMRXDSqQ", _
   "1CKf9085rt21He4", _
   "i1mOP5ZZ9hq3u9t", _
   "l9goDhDkbnIQbyN", _
   "tsrhEATcX1wLxah", _
   "lKWj6EN5ztZC6l4", _
   "s5bGroAHFyNvTVU", _
   "MOxNZVT7w14EkFh", _
   "bPq1TlykNZDCK19", _
   "QZWEVMg6mZtZNYj", _
   "Rilb3nBywLLI0kS", _
   "P8YCHhked20ULEE", _
   "tJy35MiZFcly5so", _
   "VCUswxdyYkPNjbY", _
   "IW9U6Eg6hwNwsTM", _
   "L8HtzTKl9P9pyT5", _
   "2oXDtgEFGU9topg", _
   "qiyS8c0YJ6bfKhd", _
   "cc9Zh4L7VDPNsHS", _
   "jTzAtNieWI24iSj", _
   "plrlhx1lau19RpU", _
   "0pDK14pRdchOfG3", _
   "xuHYMni7jZe56qH", _
   "O9h3baJ2GryIN1u", _
   "aexqA52m9Elrnuq", _
   "416PYegVt5OXkrM", _
   "GAnvmvM7vxukWfS", _
   "El9iPmuSEw9uLgZ", _
   "lTmO5xwEjXFFYr9", _
   "nyKo4OFZsJFfnme", _
   "abGEVDMTZduSMjR", _
   "fpSCZKtfmGsazmc", _
   "5S4kjGf8oEZBJg7", _
   "C7BbVkgkXH56K3S", _
   "ryM4qKljCo2TY9Y", _
   "zmOQaKzfRc7AWqz", _
   "OT5TCqh2EXTVOgG", _
   "BiQlpyFDViHx8GC", _
   "kRrdNijOn7DiBmj", _
   "GghsP8TgxmNeEDz", _
   "nNf3smvUZ548jDQ", _
   "OzG81fa4eake6ig", _
   "cnodhvPSEkwN6La", _
   "V3JZLELFfc24a07", _
   "oaJoZSH2Aakdy6n", _
   "EqTdPcBIOdOPAij", _
   "MgPfZvkR1z1Xu1p", _
   "2ArQmd2uGBOE84H", _
   "q5YvzQ1OBOXSLKA", _
   "3ii4SVSmNZfQvHw", _
   "L4KIqMQfJgYbkzy", _
   "MoUpIOneJAevHGs", _
   "F9oAYeoGgwbyfQQ", _
   "q8Tl08ViSbi8t9k", _
   "UcFGL5ny2HHlsy4", _
   "Kqd4y6rxGpUuRuL", _
   "rpihlOAqXmHopzT", _
   "ZJtPEiVINmmGHI0", _
   "5ltMBB80DqUCR4x", _
   "zR3gS8CUuLvFv7Z", _
   "DPfXntcwUvTwdc5", _
   "Azf1kdlC7a14lac", _
   "BHgti5oTBTURIWS", _
   "zP3TRYUScoCJy4l", _
   "0oa1kLIJeQvtJam", _
   "mFjQoScp0aWFGFu", _
   "gwH0eY89HJQLbOG", _
   "kzWw38R9T5BiwMz", _
   "tjG8LWKY9dops1L", _
   "ZMZeqXNBxgvshcl", _
   "RNfW1BtSi3oYf5c", _
   "LDYkLr1N5Ec7Njp", _
   "1HEd9uPxpzAIh2A", _
   "QyoDTgdhrPut41U", _
   "wdntwr9TJPvOpWm", _
   "yctjUXEa3hWtJbX", _
   "yxOVumftuS6moh4", _
   "mvHcze77qL24ubT", _
   "vP8z1OBlW6uCozO", _
   "rKJwOepe00r4jRQ", _
   "2j07CgIBpF5QnS1", _
   "qlqhFPUGbeHTtc4", _
   "nCaXl9q08VyeGLN", _
   "Q4LzdRAepJuKdRH", _
   "QlKO33RLa6RTgL0", _
   "HCW7bXot4SsxwuW", _
   "ItKA5tLp4NrCoA7", _
   "lpiq8Ue9MASoLlp", _
   "qi8n0maRkbp9CjB", _
   "o4Z9yuoBBNVpLy7", _
   "9IngtvPqIHF4UuT", _
   "wlmNpIITkJmtXlg", _
   "o57vBqJzAM0s48a", _
   "ukUCslE7fVZljBk", _
   "kLTodpHmJbX4nlA", _
   "w8Ux2GJ9ZHTW5dt", _
   "LuJv1AnGOXcuPrL", _
   "Yoie7p7MRrUPTSw", _
   "YAsUKXQHHdAzn07", _
   "Th9w1AXekdZzHBZ", _
   "9gwpceNBr0cqqOM", _
   "gmH7fWmfpBfXc12", _
   "wjqRUATkUDUt1xb", _
   "vBZ4gPRlkPaAZbj", _
   "8i1D1C0uFxapBcV", _
   "EcR0HAJTf3brMvo", _
   "6KTGhWicy0Fih1n", _
   "aLDkRXyTzr7B9HM", _
   "Kq4HJBrEu8W0Mwq", _
   "5OB8qJGCxegLxE2", _
   "iDGrqoun63Tu7BC", _
   "S8MP3ENc8LQcgDU", _
   "pBrd3jxXP9VhpsF", _
   "6580mhzrWd2nxnh", _
   "3hI1XZlUFhLrECI", _
   "QxcEFLMsUJNZGJr", _
   "5KxWsKM5MxOIpys", _
   "u4dJ8sPEjJ1ZZc5", _
   "sJH4kuh5tVZxpdT", _
   "y47iaOidLfO2o6p", _
   "YMMgCltLisGDg4V", _
   "ooexEFB18n0Cc2a", _
   "Dw5gUQfRWINnKSu", _
   "lDVh98FEVpDAMJ2", _
   "MbAIUiqGLS6nZSk", _
   "BvHSsD5jjfEW9Ra", _
   "YDDzhpYmMdzjJBj", _
   "x8FJmZrgjfyE6B7", _
   "LXew5vXrL4h4xml", _
   "aZ0OVMlDLvxhKW1", _
   "jKh7qNKfOYtGSat", _
   "geKobeKVTGVYwTN", _
   "FJZqbw7KM507KFZ", _
   "pqkuti4t0lOcFLA", _
   "5KSqtY48tn00bqN", _
   "FnQSFiDdS6SxH3O", _
   "gMMU5f3HKGDEmef", _
   "r2xcXpB7VdkhiDp", _
   "wS6tR2vcMFOra7s", _
   "RCGBjaaaCvquxvB", _
   "pAY0QZcF7g0pHoJ", _
   "3tg5hD2HSY5wLlG", _
   "oOtaGNazI8bmf9p", _
   "Q0qtQ4wygp9QjHU", _
   "hUJrOfNbKEOqN9X", _
   "ZAZOnHLdAr0gx39", _
   "i5dkWzQOMqLCAH7", _
   "Lx1i9srFKgZDpz1", _
   "OvSzuweDhJfHhjg", _
   "Tr8N1IDzpdz8u9l", _
   "0Ldq2Nh4Kb1KRMl", _
   "mwxT0n7wlaOaFlZ", _
   "oEGpar251rt1bsN", _
   "alKU6Suu9CrYQ4F", _
   "gMNSnQ1ihZKYajm", _
   "hnOE59jwk5rQZz3", _
   "RAuwQ7JHvNymNIj", _
   "OPYPmVQ3CLwAdyO", _
   "aAMtLVP4uOTU5vX", _
   "HqHnRPoWWpxGGXG", _
   "RLn2JqNwHvv45MN", _
   "FwghPhTDaHbCC63", _
   "kLWNTTX31FV7fHj", _
   "BArWABkNvOr39Vy", _
   "mSCVyeaDLpL76He", _
   "A8tR8Xs9imdif1C", _
   "tnRTCPpT5JOQbQU", _
   "QOcioJ1hRooSh95", _
   "VDge9sKXJacKAwf", _
   "1kI59lAHC1tGqJE", _
   "yWSciDXMDamaBqE", _
   "dNQtSLsOru64YI7", _
   "61Zm9bQH153apj5", _
   "TVR8o1Lh8ZnTxiU", _
   "kYtEfXIFOQzSNX6", _
   "AVbVYrEUHkHy9eq", _
   "vRvIxKbL3PevQO1", _
   "49j0P60n1LHXtqe", _
   "tjMYGTr7tmECKCp", _
   "oMapLbM7ZdDxTLs", _
   "5gNAUHyDsflVc2F", _
   "r4euh6AaGQQknTS", _
   "kKilGRlEIOWvEI0", _
   "AaGCd8jKC2bnpUo", _
   "RxyJt5WictJAOTr", _
   "auMUg9ur4qcsLAO", _
   "qLHuURO5AHBDIIj", _
   "DCFCP004PMV90OA", _
   "4S6O3MqP2gjep0i", _
   "Yq9NJ7TqzD1hLBv", _
   "EqYURl3v9z97xnl", _
   "yyJlVwWixqs2cBb", _
   "KPtDDWO0zRnBU6R", _
   "bqIVNAivKP75AA3", _
   "SMeJE88H8YccLGg", _
   "guU5v4p0u6l456Z", _
   "cZzaZw0Hkyxeief", _
   "vVEPHjga3z8zK2m", _
   "HBxNRBypUH677tq", _
   "Vr7rgWOsQ4VPH0c", _
   "8MgioAjg3pNTLMh", _
   "LCBa3O0wUVHGdkF", _
   "cSHkcbwEVt3z6vJ", _
   "GaedOOPzElZj7tS", _
   "P0TlknByYkbD72R", _
   "G3KLv66RNtPAOgv", _
   "eh2I9G13P80gDT9", _
   "hkWXbGco5kOLBfh", _
   "dKpO3G5q5Z0lRAt", _
   "Us1TUKDK4PtXdBy", _
   "XcxU6XeyfqUI914", _
   "3rZkU6fI9ptQDA0", _
   "2InRtKvHGSPMrHT", _
   "bbjkgk80kWjzxj8", _
   "uKbYS5ARUpN9NEg", _
   "nGXRkS8yvw5zwpm", _
   "fv5LRkXGIVgFE8B", _
   "BBR7S6dYBsgeWIT", _
   "EvviJ94t0X3h79F", _
   "cBDhmDWnwylbRyt", _
   "yr4kDPZIVWD5T0z", _
   "qVFHCSoLfi1DS8A", _
   "X5LmdCpOkVbEiKt", _
   "5vmBvdFdMFncCxB", _
   "1XqwRNEcoytCUST", _
   "NKBM6EqaCFn7MFA", _
   "MV0m96gJXcpskM5", _
   "Py9JRfcyJHo5g2v", _
   "hPfAAwyO8z2GWAf", _
   "8pjTTgyyEJzHsbO", _
   "zQGo8yGJWEMT7np", _
   "O1mqXykBtpRUbU4", _
   "TA1LX3ukGDh9iGl", _
   "ueAUClIoFrwvTnA", _
   "vLoeVo1qhajshtv", _
   "QA0ViAwfOqyZBmJ", _
   "WOoEtBbQVE6rddK", _
   "1f9O4Rm16XU8XQ4", _
   "oLnyMG1JgBUotOE", _
   "fTuoPerSilg8f4H", _
   "zTeaME0qN8orrTu", _
   "4H8tZdO8pTEIT65", _
   "Y8I6eCpXIRcighC", _
   "7cuKihetulV9xzP", _
   "T49NYJ3n1QS9ooC", _
   "wZXGYJSzrLyVXX7", _
   "9r0BcaUHxbcfoCH", _
   "kD9MEroRxk3I4Qu", _
   "dRc2PVxw3zcR9mP", _
   "dC6tJjzlf1Ilj0V", _
   "UcMCxho9QUGcn1a", _
   "UYPGC3kgQ4Qpjcw", _
   "idll4ZplLI4h41K", _
   "5aQxuUHBau4FbLo", _
   "owfgCySXjQYrZYp", _
   "Cawj7xk2E2dfBs8", _
   "uniyokgUj9sfJo4", _
   "2Ag9lBE1R9s7lrv", _
   "JMFGOPl37E3925Q", _
   "P65Tq7BVrElwK3f", _
   "y9oBohFBUt86Cgi", _
   "TrZ0to55VN3WgIu", _
   "rlxReYQILQMrUKE", _
   "MMfkQIG1IDi7Yl9", _
   "C1mhVLQHCQOJuTY", _
   "CFUw34eJ5srZVGM", _
   "nFeYRNUmfjstNYu", _
   "k6WM2m27uS7n8Li", _
   "yu3gpN77MchqrZr", _
   "QBDYyMEJ6YgA3Jc", _
   "Sz2MqoysGStZkn1", _
   "ZOaPOPqDwDAGC87", _
   "dLVxErUndOT62aP", _
   "r5jyiRS0Gq9ZqwF", _
   "gkSgJbKViJQ8PIL", _
   "CRxw1bMC9P3aTR9", _
   "DsOvESZKuz0G5Z3", _
   "iX88zN3rMS0i1xF", _
   "09rxEmOG8XfURLe", _
   "VB2LYzsZZIiYKj8", _
   "Wak1HXfWVGv5PPP", _
   "EBKeSCxGU6WpjAE", _
   "aoGiwqT1VFRQO2Q", _
   "i2IEirKTQB06052", _
   "Etmfna1xzOAk4PL", _
   "TqBxmhcSVsfBq9G", _
   "NxHwH8j3raWhB6Z", _
   "WFSOSp5zRHkMO9E", _
   "pjODTwoaDue4SYA", _
   "NY4jlFiVHY84OIN", _
   "GgCrVPDep8JYFeY", _
   "yHLg33NYmpPVFEh", _
   "GccFKo6vZaSRilS", _
   "fzWNA2Ug1JUD3zK", _
   "jWIz1aZUrlPcSMA", _
   "GCavYYl3jVfHJIU", _
   "FlM5ZnOm8E5Hiph", _
   "fSiYiOeOUkUO8BC", _
   "lPpDPpJMj3WxPIg", _
   "npQyTqSOw88Lrr1", _
   "u8bnrphmvO0l8aJ", _
   "OXEjJxnxSJDwhZH", _
   "Z7aLW6qXiw61utv", _
   "GW1Bmm7dCwPkwG8", _
   "yl3USZTC6r8RGaw", _
   "X4xIbjMP00lPgPk", _
   "RmM6AgXGKlZ1chQ", _
   "sD9Km1gUSMMhClw", _
   "tXW5ulqRHeoUdpW", _
   "eErHSn8g4iFW7IF", _
   "CB33ubsJP6r1b54", _
   "bSuofuK9IJGkkDM", _
   "QWX5rYAhevSQm1o", _
   "pmRC4fyJLf3FwY0", _
   "JYxzsUXckA6fj1s", _
   "9B8W36TSLhvlYyL", _
   "Z6GB9NuGq2nJBqS", _
   "lqAqwfjo2VynXY1", _
   "2ntdERm3fr01Wwc", _
   "PXfdeR2bEESoV3B", _
   "Wc1ZkNpLaPZoxX7", _
   "akVY42C6eo5OwnZ", _
   "YLTH75ipMahfj5p", _
   "bStltgmEDJqRAxm", _
   "T2WbZQHd87aag7f", _
   "TqTTQbRCcuMS9bd", _
   "ZxC564qFX5Upec0", _
   "Ja6IuVDRFtnKLkt", _
   "VODi7qDEGyHrzFo", _
   "B81j8tV00Ht76rd", _
   "lvuxHPRwMY4WG4k", _
   "TkE5qs22f6K0Z3E", _
   "hMpxnXc1qtZ7mxB", _
   "qiTh2m1ZYfTlwo1", _
   "NVGxGqpIIMlNzKE", _
   "VceDjdnOipwr5SI", _
   "SLFFk4ZsXSKSOZJ", _
   "HSveTNf4mripg01", _
   "DPbA14YkmeECIiI", _
   "ROI8dGJURcDmvhV", _
   "n7GHf6VoajeqU87", _
   "BvjxEcWTuI40zVv", _
   "G9acCWS0W2ohbNi", _
   "f5tkoDibS7HwaKo", _
   "HFO3ixvvtuTmy5j", _
   "dyHJlRXnCPBkTEy", _
   "FgvaL1ttSrK1OOU", _
   "z2hZZEEQcyaAH1o", _
   "f8GJajQVsQv6kRV", _
   "9pjEM4yHsB45tU6", _
   "BIBOv3zSR2EEx8M", _
   "HhSWGHmEM6kDy29", _
   "ZudN6LrHFyb2gXO", _
   "Mo50hT1hoscttXq", _
   "snfkFx54AvAwnr6", _
   "KlA95CH5XihDGm2", _
   "jIQBjWOTtIrCwC6", _
   "5gd6PboyIVQ8drA", _
   "aJDsbjELSRpLcow", _
   "QILKIauousxT63j", _
   "iedPKnH2GkgGdwr", _
   "E5sFwLadFcP3q1U", _
   "J3aiiaC9SymT4J0", _
   "CrxGIritnIDOkYn", _
   "t35sS41UXCjGsuX", _
   "am0SB2whA9HjF9n", _
   "rDSV6ZPdpBg7xGj", _
   "SXtnRn5D3gzjP8s", _
   "dE63aEiDuSA0HyJ", _
   "qeQaj2m1RqN8YUg", _
   "Al9kkG9eI52b4Y8", _
   "EP3sAIhiSgdnwRq", _
   "PFxW8LSW7Atdjpu", _
   "xovQFPm9SMEAMTW", _
   "p6b57ctsfHEv5Zx", _
   "51LAWTwQlVoGlRs", _
   "1IXKmgJnCrF40cK", _
   "zb23pxsL4IisVuI", _
   "xAMp3JqmwoTSp1f", _
   "vX0U8kluMkxf4qe", _
   "U4bstQG9Pu0RAwW", _
   "pjF6Epdhz6hCu6J", _
   "qSN3I98M5QGH0i0", _
   "b88dAcGnT90Kour", _
   "oCF2txW02K8o8pF", _
   "f5bN83dXXGnCFep", _
   "DsFH48Ppm8xME71", _
   "SaViwY1em8brDXc", _
   "riMVcGDxiekExCw", _
   "73JFoQO42TA2QVl", _
   "IH5LyXOE6UBw6Fr", _
   "Ui5RZWUmMudcXk1", _
   "fOcjM7wAnZwkphd", _
   "JBOWrQ49S2RxrSt", _
   "0Du1vBtPh8rVp7g", _
   "xFmDBvOImdcXUde", _
   "bU5Sb3ZSvTgnIK8", _
   "mwiyCyChyGgMyG5", _
   "CU8CJzTlwGKcpvI", _
   "VH1o5RnBvMQuEA4", _
   "NI2hivmN5Keb5lR", _
   "Xh1LJYTt2n8g59c", _
   "OpupweaXvwHptzp", _
   "1GHq1ubxfLqK4OG", _
   "qt5UHDhvnvR5n9U", _
   "hulkmNLIsK4PC3p", _
   "ySEg4PwidQe1teQ", _
   "hGge8gqVELy1zYe", _
   "r5ehyiGZmkX7UQe", _
   "haqc99ud8JHPzDh", _
   "7ujk1muylI66m2b", _
   "pFIjm26yGyBOW5K", _
   "lMboyFhfEP74qkU", _
   "ywmwIlR0nsL6gK3", _
   "P0oxFcOrgP0Df0h", _
   "qkte9BAOHSCpNo9", _
   "lbFJtMmz9piDsDj", _
   "Wsi48gVpETqspFU", _
   "g8wrQIi5FnZD9OD", _
   "5O6OwSsO5D8FO5n", _
   "gR07I6LsbjVEz4C", _
   "cxb7ZB6HrefcPXR", _
   "trDYQcLt40Ja9aK", _
   "7UpCBW4OW0MLz6B", _
   "XRujqIUmtvflj3N", _
   "umaZFhyvY5AzaKN", _
   "OmdkcM3jLe7J1A0", _
   "4YMZGLmhvLnu3Kv", _
   "5d2r1QyOhJj9Fqw", _
   "9nP1uDxXtbxiZ9k", _
   "xQdUjVay6o9uFCO", _
   "LEMr6pM4SMeXbDU", _
   "kIAjKraFII0iTNj", _
   "Us9sUgR279LuTIy", _
   "ktKJUKh5Dio1H98", _
   "DmfN3dGLzjgsAAS", _
   "9Xc35T0I2ar43Jm", _
   "ILgtOg2WykKU62n", _
   "C7fYbv8qTt38M8H", _
   "K9rndiw0WGOQY4N", _
   "TChidUp8P8NBqtC", _
   "jDwEtObfaAiGH7O", _
   "KJkWjCYT1RSEFdQ", _
   "0iFGHF36Nrmp07h", _
   "5VIt1fJKUKLFKLK", _
   "Xxqhv5Uj5xbRXxs", _
   "5Y4EWO7XCwTo67T", _
   "sCa5SbxER85BIBs", _
   "JRKuzSxSqts9Ro0", _
   "jec2Q0doP7RiOhp", _
   "0Y9HaTGNxRQBlka", _
   "ghwve25h2fcyeOx", _
   "NWUdhx7ZDufUDRO", _
   "ymyXefzmHYC2eU8", _
   "S0QzskYN3qDNQWA", _
   "V0z7Z901KkPenGe", _
   "EI26DtjL97kogL0", _
   "O2AcTnh4PjNxDtk", _
   "FLBjOkxlv6zZsAF", _
   "o0aBE20vMO04hEr", _
   "9Fc8j4uaBerGEO2", _
   "ypOMeTzRxbsvn8R", _
   "xmDjc12e7rPgmsB", _
   "2eRTINEBAB6qD95", _
   "lhtjR88WK2DMtzk", _
   "DxIXZqauz8XAWeb", _
   "CvZD5wdBuDOEJFD", _
   "FgZ3F7cQQdX729p", _
   "te0H25fZuqu4zzZ", _
   "0hnzts4vlyjiXmp", _
   "PG9ZLXVvdxGS9EN", _
   "LywgQdTOUhJMZuu", _
   "T4TYVTfhMYlfJEB", _
   "siBhcBmdFnIlbb0", _
   "5qvdaMppzaEwuFu", _
   "WgYAOtOndkjRbxH", _
   "TpWZ2rPBCSoTaRj", _
   "g2VMTkdWEXtCnXq", _
   "NpNWx1PwEigzpDc", _
   "RZPjPSjmoL9cNVz", _
   "I6vgrOrFguPDrLf", _
   "NvSUDmyK3xZai94", _
   "0wmOsLgbXQZrsdu", _
   "nnZarhNd11Hp9QU", _
   "24zcJWIRJFkuaYF", _
   "6cQa8LHu6BOcLa9", _
   "MjVysqJJvKkptr9", _
   "tbyVFaUNntwTjve", _
   "5KBskq8vfAzzrqO", _
   "SmHXnXF3eel3tR0", _
   "rlQ7ZujRaUy4TKE", _
   "oL83yYWZbnCATxc", _
   "sVwU5m60vAfkk2G", _
   "6SUK63Tvce6Zo7A" ]
   Return $aKeys
EndFunc





