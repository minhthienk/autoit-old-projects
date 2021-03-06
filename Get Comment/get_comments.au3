#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include <IE.au3>
#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>
#include <Date.au3>


While (1)

    $sUrl = InputBox('Link?', 'Input link to get comments' & @CRLF & 'Input "Exit" to exit')
    If ($sUrl = 'Exit') Then Exit


    $oIE = _IECreate ($sUrl, 1)

    $time = _NowTime(5)
    $time = StringReplace($time, ':', '-')

    $sFilePath = 'data ' & $time &'.txt'
    ; Open the file for writing (append to the end of a file) and store the handle to a variable.
    Local $hFileOpen = FileOpen($sFilePath, $FO_APPEND)
    If $hFileOpen = -1 Then
        MsgBox($MB_SYSTEMMODAL, "", "An error occurred whilst writing the temporary file.")
    EndIf


    $comment_tags = _IETagNameGetCollection($oIE, 'div') 
    $classvalue =  'review-item__text--'
    For $comment_tag In $comment_tags
        $class = $comment_tag.GetAttribute('class')
        If StringInStr(String($class), $classvalue) Then

            $title = GetItemStringByMark ($comment_tag.outerhtml, '<h2>', '</h2>')
            $title = '"' & StringReplace ($title, '"', '""') & '"' 

            $comment = getObjByClass($comment_tag, 'div', 'class', 'text__comment').innertext
            $comment = '"' & StringReplace ($comment, '"', '""') & '"' 

            $name = getObjByClass($comment_tag, 'p', 'class', 'review-item__displayName').innertext
            $name = '"' & StringReplace ($name, '"', '""') & '"' 

            $date = getObjByClass($comment_tag, 'p', 'class', 'review-item__createdDate').innertext
            $date = StringReplace ($date, '?', '')
            $date = '"' & StringReplace ($date, '"', '""') & '"' 


            $row = $title & @TAB & $comment & @TAB & $name & @TAB & $date & @CRLF

            ; Write data to the file using the handle returned by FileOpen.
            FileWrite($hFileOpen, $row)

            ; Close the handle returned by FileOpen.

            ConsoleWrite($row & @CRLF)
        EndIf
    Next
    FileClose($hFileOpen)
WEnd



Func getObjByClass(ByRef $oIE, $tag, $className, $classvalue)
    $tags = _IETagNameGetCollection($oIE, $tag)  
    For $tag In $tags
        $class = $tag.GetAttribute($className)
        If StringInStr(String($class), $classvalue) Then
            Return $tag
        EndIf
    Next
    Return False
EndFunc


Func GetItemStringByMark ($sString, $sStartMark, $sEndMark, $iOccurrenceStart = 1, $iOccurrenceEnd = 1)
   If StringInStr ($sString, $sStartMark, 0, 1, 1) <> 0 Then
      Local $iStart = StringInStr ($sString, $sStartMark, 0, $iOccurrenceStart, 1) + StringLen ($sStartMark)
      Local $iEnd = StringInStr ($sString, $sEndMark, 0, $iOccurrenceEnd, $iStart)
      Local $sItemString = StringMid ($sString, $iStart, $iEnd - $iStart)
   Else
      Local $sItemString = ""
   EndIf
   Return $sItemString
EndFunc


