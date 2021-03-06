#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>
#include <Array.au3>

#include "ImageSearch.au3"



$aArray = SearchFile()
$sMsg = ""
For $element in $aArray
   $sMsg = $sMsg & $element & @CRLF
Next

If $sMsg = "" Then
   MsgBox(0,"Thông báo", "Không tìm thấy hình nào trên thư mục chương trình." & @CRLF & "Vui lòng kiểm tra")
   Exit
Else
   MsgBox(0,"Thông báo", "Các hình tìm thấy:" & @CRLF & @CRLF & $sMsg)
   $sMsg = "Chương trình sẽ kiểm tra liên tục các hình này trên màn hình máy tính" & @CRLF & "Nếu không tìm thấy hình, chương trình sẽ phát đoạn âm thanh ""sound.mp3"" " & @CRLF & @CRLF & "NOTE: CẦN THẬN TRƯỜNG HỢP KHUNG HÌNH ĐANG KIỂM TRA CÓ CHỨA HÌNH ĐỘNG"
   MsgBox(0, "Thông báo", $sMsg)
   MsgBox(0, "Thông báo", "Nhấn ""OK"" để bắt đầu")
   Sleep(3000)
EndIf



While 1
   For $element in $aArray
	  Local $x = 0, $y = 0
	  Local $Search = _ImageSearch($element, 1 ,$x,$y,0)
	  If $Search = 1 then
		 Sleep(1000)
	  Else
		 MsgBox(0,"Thông báo", "Không còn tìm thấy các hình đã lưu trên màn hình" & @CRLF & "Âm thanh được phát sau 3 giây nữa!!!", 3)
		 SoundPlay("sound.mp3", 0)
		 MsgBox(0,"Thông báo", "Không còn tìm thấy các hình đã lưu cho trên màn hình" & @CRLF & "Nhấn ""OK"" để thoát" & @CRLF & @CRLF & "Script được viết bởi Thiện Tứ, cảm ơn đã sử dụng")
		 Exit
	  EndIf
   Next
WEnd





Func SearchFile()
    ; Assign a Local variable the search handle of all files in the current directory.
    Local $hSearch = FileFindFirstFile('*.bmp')

    ; Check if the search was successful, if not display a message and return False.
    If $hSearch = -1 Then
        MsgBox($MB_SYSTEMMODAL, "", "Error: No files/directories matched the search pattern.")
        Return False
    EndIf

    ; Assign a Local variable the empty string which will contain the files names found.
    Local $sFileName = "", $iResult = 0
	Local $aArray[0]

    While 1
        $sFileName = FileFindNextFile($hSearch)
        ; If there is no more file matching the search.
		 Local $bError = @error
		;
		_ArrayAdd($aArray, $sFileName)
        If $bError Then ExitLoop
    WEnd
    $aArray = _RemoveEmptyArrayElements ( $aArray )
    ; Close the search handle.
    Return $aArray
EndFunc


Func _RemoveEmptyArrayElements ( $_Array )
    Local $_Item
    For $_Element In $_Array
        If $_Element= '' Then
            _ArrayDelete ( $_Array, $_Item )
        Else
            $_Item+=1
        EndIf
    Next
    Return ( $_Array )
 EndFunc ;==> _RemoveEmptyArrayElements ( )

