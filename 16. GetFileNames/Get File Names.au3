#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>


SearchFile()

Func SearchFile()
   $sMasterFolderPath = FileSelectFolder ('Select Folder', @ScriptDir)


    ; Assign a Local variable the search handle of all files in the current directory.
    Local $hSearch = FileFindFirstFile($sMasterFolderPath & '\*')

    ; Check if the search was successful, if not display a message and return False.
    If $hSearch = -1 Then
        MsgBox($MB_SYSTEMMODAL, "", "Error: No files/directories matched the search pattern.")
        Return False
    EndIf

    ; Assign a Local variable the empty string which will contain the files names found.
    Local $sFileName = "", $iResult = 0

    While 1
        $sFileName = FileFindNextFile($hSearch)
        ; If there is no more file matching the search.
		 Local $bError = @error
		;
		WriteTxtFile ('StringList', $sFileName & @CRLF, "append")
        If $bError Then ExitLoop
    WEnd

    ; Close the search handle.
    FileClose($hSearch)

EndFunc   ;==>Example;====================================================================================================================
;                  FUNCTION DISCRIPTION: CLOSE ALL IE OBJECT
;                  INPUT			   :
; 				   OUTPUT			   :
;====================================================================================================================
Func WriteTxtFile ($sFileName, $sTxt, $sMode = "append")
   Local $sFilePath = @ScriptDir
   If $sMode = "overwrite" Then
	  Local $hFileOpen = FileOpen ($sFilePath & "\" & $sFileName & ".txt",$FO_OVERWRITE)
   Else
	  Local $hFileOpen = FileOpen ($sFilePath & "\" & $sFileName & ".txt",$FO_APPEND)
   EndIf
   FileWrite($hFileOpen, $sTxt)
   FileClose($hFileOpen)
EndFunc