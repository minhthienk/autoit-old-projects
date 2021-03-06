#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>
#include <array.au3>
#include <timers.au3>

GetFileNames()
Exit

Main()


Func Main ()
   Local $sFileNames = LoadFileNames ('Select FileNames File')
   Local $sWordList = LoadFileNames ('Select Wordlist File')



   Local $aFileNames = StringSplit ($sFileNames, @CRLF, $STR_ENTIRESPLIT  +  $STR_NOCOUNT)
   Local $aWordList = StringSplit ($sWordList, @CRLF, $STR_ENTIRESPLIT  +  $STR_NOCOUNT)





   For $iWordCount = 300 To UBound($aWordList) - 1

	  Local $sSearchWord = $aWordList[$iWordCount]
	  Local $iSentenceLenth = 20
	  Local $sChosenSentence  = ''
	  Local $iGap  = 999
	  Local $sTempSentence
	  Local $aTempArray[0][2]

	  ConsoleWrite('Procesing the word: ' & $sSearchWord & @CRLF)
	  If StringLen($sSearchWord) <= 1 Then ContinueLoop
	  ConsoleWrite('> 2 character' & @CRLF)


	  For $i=0 to UBound($aFileNames) - 1

		 If IsWordInSentence($aFileNames[$i], $sSearchWord) Then
			$sTempSentence = $aFileNames[$i]
			$sTempSentence = StringReplace($sTempSentence, '.mp3', '')

			;ReDim $aTempArray[UBound($aTempArray)+1][2]
			;$aTempArray[UBound($aTempArray)-1][0] = $sTempSentence
			;$aTempArray[UBound($aTempArray)-1][1] = StringLen($sTempSentence)


			If Abs($iSentenceLenth - StringLen($sTempSentence)) < $iGap Then
			   $iGap = Abs($iSentenceLenth - StringLen($sTempSentence))
			   $sChosenSentence = $sTempSentence
			EndIf
		 EndIf
	  Next

	  ;_ArrayDisplay($aTempArray)
	  ;MsgBox (0, '', $iGap & '  ' & $sChosenSentence)

	  Local $sLastFolder = Int ($iWordCount/100)*100 + 100


	  If $sChosenSentence <> '' Then
		 Local $sSourceFolderPath = 'E:\Vocabulary Audio\Original\Not classified'
		 Local $sDesFolderPath = 'E:\Vocabulary Audio\Classified\Level 1\' & $sLastFolder
		 FileCopy($sSourceFolderPath & '\' & $sChosenSentence & '.mp3', $sDesFolderPath & '\', $FC_OVERWRITE + $FC_CREATEPATH)
		 FileMove($sDesFolderPath & '\' & $sChosenSentence & '.mp3', $sDesFolderPath & '\' & ($iWordCount + 1) & '.[' & $sSearchWord & '] ' & $sChosenSentence & '.mp3', $FC_OVERWRITE + $FC_CREATEPATH)
	  EndIf

   Next


EndFunc
;
;
;
;
;
;
Func GetFileNames()
   Local $data[0][2]

   $sMasterFolderPath = 'E:\Vocabulary Audio\Original\Not classified'
   ;@DesktopDir;FileSelectFolder ('Select Folder', @ScriptDir)

   ; Assign a Local variable the search handle of all files in the current directory.
   Local $hSearch = FileFindFirstFile($sMasterFolderPath & '\*')

   ; Check if the search was successful, if not display a message and return False.
   If $hSearch = -1 Then
	  MsgBox($MB_SYSTEMMODAL, "", "Error: No files/directories matched the search pattern.")
	  Return False
   EndIf

   ; Assign a Local variable the empty string which will contain the files names found.
   Local $sFileName = "", $iResult = 0

   $count=0
   While 1
	  $sFileName = FileFindNextFile($hSearch)
	  ; If there is no more file matching the search.
	  If @Error Then ExitLoop
	  ;
	  ;ReDim $data[UBound($data)+1][2]
	  ;$data[UBound($data)-1][0] = $sFileName
	  ;$data[UBound($data)-1][1] = StringLen($sFileName)
	  $count = $count + 1
	  ConsoleWrite($count & @CRLF)
	  WriteTxtFile ('FileNames2', $sFileName & @CRLF, "append")
   WEnd

   ; Close the search handle.
   FileClose($hSearch)

   ;_ArraySort ($data, 0, 0, 0, 1)
   ;Return $data
EndFunc
;
;
;
;
;
;
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
;
;
;
;
;
;
Func LoadFileNames ($sTitle)
   ;Open YMME config file and get data

   Local $FilePath = FileOpenDialog ($sTitle, @ScriptDir, '(*.txt)')

   Local $hFileOpen = FileOpen ($FilePath,$FO_READ )
   Local $sFileRead = FileRead($hFileOpen)
   FileClose($hFileOpen)
   ;Remove redundant @CRLF
   $sFileRead = StringRegExpReplace ($sFileRead, '[\r\n]+', @CRLF)
   $sFileRead = StringRegExpReplace ($sFileRead, '[\r\n]+$', '')
   Return $sFileRead
EndFunc
;
;
;
;
;
;
Func IsWordInSentence($sSentence, $sWord)
   Local $bResult = False
   If StringInStr($sSentence, ' ' & $sWord & ' ') <> 0 Then     ;2 spaces at both ends
	  $bResult = True
   ElseIf StringInStr($sSentence, $sWord & ' ') = 1 Then        ;The word begins the sentence
	  $bResult = True
   ElseIf StringInStr($sSentence, ' ' & $sWord & '.') <> 0 Then ;Then word ends the sentence
	  $bResult = True
   EndIf


   Return $bResult
EndFunc
