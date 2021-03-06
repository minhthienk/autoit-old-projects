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
#include <IE.au3>

HotKeySet('{ESC}', _Exit)

Local $sListWord = LoadFile ('Select Word File')
Local $aWords = StringSplit ($sListWord, @CRLF, $STR_ENTIRESPLIT  +  $STR_NOCOUNT)

For $i = 0 To  UBound($aWords) - 1

   $sFile1 = $aWords[$i]
   $sFile2 = '#Blank mp3'
   $sTitleControl = 'Audacity'

   WinActivate (WinWait($sTitleControl))
   Send ('^+i')
   ControlSetText(WinWait('Select one or more audio files...'), '', 'Edit1', $sFile1 & '.mp3')
   ControlClick (WinWait('Select one or more audio files...'), '', 'Button2')


   Sleep (500)
   $sTitleControl = $sFile1

   WinActivate (WinWait($sTitleControl))
   Send ('^+i')
   ControlSetText(WinWait('Select one or more audio files...'), '', 'Edit1', $sFile2 & '.mp3')
   ControlClick (WinWait('Select one or more audio files...'), '', 'Button2')


   Sleep (500)
   $sTitleControl = $sFile1
   WinActivate (WinWait($sTitleControl))
   Send ('^+e')
   ControlSetText(WinWait('Export File'), '', 'Edit1', $sFile1 & '.mp3')
   ControlClick (WinWait('Export File'), '', 'Button2')


   Sleep (500)
   $sTitleControl = $sFile1
   WinActivate (WinWait($sTitleControl))
   Send ('^w')
   ControlClick (WinWait('Save changes?'), '', 'Button2')
   Sleep (1000)

Next

Func LoadFile ($sTitle)
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


Func _Exit ()
   Exit
EndFunc