#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         Thien Nguyen

 Script Function:
   Copy data from bonbanh.com

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include-once

#include <MsgBoxConstants.au3>
#include <InetConstants.au3>
#include <IE.au3 >
#include <Excel.au3>
#include <Clipboard.au3>



#Region ### START GUI section ### Form=
   Opt('GUIOnEventMode', 1)
   $Form1 = GUICreate('Get Mp3 Pronuncation Cambridge', 400, 200, -1, -1)
   GUISetBkColor(0xFFFFFF)
   ;-------------------------------------------
   ;Create Labels
   $Commu_Ctrl = GUICtrlCreateLabel('', 50, 160, 250, 60)


   GUICtrlCreateLabel('From word #', 50, 30, 250, 20)
   $StartNum = GUICtrlCreateInput('', 50, 50, 100, 20)


   GUICtrlCreateLabel('To word #', 250, 30, 250, 20)
   $EndNum = GUICtrlCreateInput('', 250, 50, 100, 20)


   ;Create Button
   $Button_OpenDialog = GUICtrlCreateButton("Save to ...", 50, 110, 100, 30)
   GUICtrlSetOnEvent($Button_OpenDialog, 'Set_OpenDialog_Flag')

   $Button_Download = GUICtrlCreateButton("Download", 250, 110, 100, 30)
   GUICtrlSetOnEvent($Button_Download, 'Set_Download_Flag')
#EndRegion ### END Koda GUI section ###


;SHOW GUI
GUISetState(@SW_SHOW)


;Set Hotkey for the program
HotKeySet("{ESC}", "Autoit_Exit")


Global $bOpenDialog_Flag = False
Global $bDownload_Flag = False
Global $sDownloadFolder = @ScriptDir

Func Set_OpenDialog_Flag ()
   $bOpenDialog_Flag = True
EndFunc

Func Set_Download_Flag ()
   $bDownload_Flag = True
EndFunc



While 1
   If $bOpenDialog_Flag = True Then OpenDialog()
   If $bDownload_Flag = True Then GetMp3 ()
WEnd




Func OpenDialog()
   $sDownloadFolder = FileSelectFolder ('Select Folder', $sDownloadFolder)
   $bOpenDialog_Flag = False
EndFunc


GUICtrlRead ( controlID [, advanced = 0] )



Func GetMp3 ()
   Close_All_IE()

   Local $sListWord = LoadFile ('Select Word File')
   Local $aWords = StringSplit ($sListWord, @CRLF, $STR_ENTIRESPLIT  +  $STR_NOCOUNT)

   ;_ArrayDisplay($aWords)


   ;Master Link
   Local $sMasterLink = 'https://dictionary.cambridge.org/pronunciation/english/'

   ;create IE object with the link
   Local $oIE = _IECreate ('about:blank',0,1,1,0)

   ;Start and end position of the loop
   Local $iStartNum = GUICtrlRead ($StartNum)
   Local $iEndNum = GUICtrlRead ($EndNum)

   If $iStartNum = '' Then $iStartNum = 1
   If $iEndNum = '' Then $iEndNum = UBound($aWords)

   For $i = $iStartNum - 1 To $iEndNum - 1
	  Local $sWord = $aWords[$i]

	  ;Only download word with 2 or more characters
	  If StringLen($sWord) < 2 Then ContinueLoop

	  ;Navigate to the link of word pronunciation
	  Local $sLink = $sMasterLink & $sWord
	  _IENavigate($oIE, $sLink)
	  Sleep (200)
	  ;Get links collection
	  Local $oDivs = _IETagNameGetCollection($oIE, 'div')

	  For $oDiv In $oDivs
		 ;Determine tage object containing data
		 If StringInStr($oDiv.innertext, 'How to pronounce ' & $sWord & ' ', 1) <> 0 And StringInStr($oDiv.innertext, @CRLF) = 0 And StringInStr($oDiv.innertext, 'American') <> 0 Then
			;Get innertext of the data
			Local $sWordType = $oDiv.innertext
			$sWordType = StringRegExpReplace($sWordType, '^.+pronounce ' & $sWord & ' ','')
			$sWordType = StringRegExpReplace($sWordType, ' in American.+','')
			Local $sMp3Name = $i + 1 & '. ' & $sWord & ' (' & $sWordType & ')' & '.mp3'
			;Get mp3 Link of dataa
			GUICtrlSetData($Commu_Ctrl, $sMp3Name)
			Local $oMetas = _IETagNameGetCollection($oDiv, 'meta')
			For $oMeta In $oMetas
			   If StringInStr($oMeta.GetAttribute('content'),'.mp3') Then
				  Download($oMeta.GetAttribute('content'), $sDownloadFolder & '\' & $sMp3Name)
			   EndIf
			Next
		 EndIf
	  Next
   Next


   _IEQuit ($oIE)
   GUICtrlSetData ($Commu_Ctrl, 'Done')
   $bDownload_Flag = False
EndFunc




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




Func Download($sUrl, $sFilePath)
   Local $bDone_Flag = False
   Local $sExceed_Time = 0

   While (1)
	  ; Download the file in the background with the selected option of 'force a reload from the remote site.'
	  Local $hDownload = InetGet($sUrl, $sFilePath, $INET_FORCERELOAD, $INET_DOWNLOADBACKGROUND)
	  $sExceed_Time = 0
	  ; Wait for the download to complete by monitoring when the 2nd index value of InetGetInfo returns True.
	  While (1)
		 Sleep(50)
		 $sExceed_Time = $sExceed_Time + 50
		 If InetGetInfo($hDownload, $INET_DOWNLOADCOMPLETE) Then
			$bDone_Flag = True
			ExitLoop
		 EndIf

		 If $sExceed_Time > 5000 Then ExitLoop
	  WEnd

	   ; Close the handle returned by InetGet.
	   InetClose($hDownload)

	   If $bDone_Flag = True Then ExitLoop
   WEnd
   Sleep(50)

EndFunc


Func Autoit_Exit ()
   Close_All_IE()
   Exit
EndFunc


Func Close_All_IE()
   $Proc = "iexplore.exe"
   While ProcessExists($Proc)
      ProcessClose($Proc)
   Wend
EndFunc






;Function: OPEN EXCEL ==========================================================
Func OpenExcel ()
   ;Set initial values for the parameters used to open Excel
   #Region Local Vars
	  Local $bVisible = True
	  Local $bDisplayAlerts = False
	  Local $bScreenUpdating = True
	  Local $bInteractive = True
	  Local $bForceNew = False
   #EndRegion
   ;
   Local $oExcel = _Excel_Open ($bVisible, $bDisplayAlerts, $bScreenUpdating, $bInteractive, $bForceNew)
   Return $oExcel
EndFunc


;Function: RANGE READ ===========================================================
Func RangeRead ($oWorkbook, $vWorksheet)
   ;Set initial values for the parameters used to open Excel
   #Region Local Vars
	  Local $vRange = Default
	  Local $iReturn = 1
			;1 - Value (default)
			;2 - Formula
			;3 - The displayed text
			;4 - Value2 (The only difference between Value and Value2 is that the Value2 property doesn’t use the Currency and Date data types)
	  Local $bForceFunc = False
   #EndRegion
   ;
   Local $aResult = _Excel_RangeRead ($oWorkbook, $vWorksheet, $vRange, $iReturn, $bForceFunc)
   Return $aResult
EndFunc