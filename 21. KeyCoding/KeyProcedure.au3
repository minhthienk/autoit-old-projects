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


   GUICtrlCreateLabel('Reserve', 50, 30, 250, 20)
   $StartNum = GUICtrlCreateInput('', 50, 50, 100, 20)


   GUICtrlCreateLabel('Reserve', 250, 30, 250, 20)
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
   If $bDownload_Flag = True Then GetData ()
WEnd




Func OpenDialog()
   $sDownloadFolder = FileSelectFolder ('Select Folder', $sDownloadFolder)
   $bOpenDialog_Flag = False
EndFunc


GUICtrlRead ( controlID [, advanced = 0] )



Func GetData ()
   Close_All_IE()
   Local $sLink = 'https://www.carandtruckremotes.com/searchresults.html?keywords=mazda&x=0&y=0#/?keywords=mazda&res_per_page=36&search_return=all&page='
   Local $oIE = _IECreate('about:blank',1,1,1,0)

   Local $sTxt = ""

   Local $sPrevious = ''
   Local $sCurrent = ''

   For $iPage = 1 To 15
	  _IENavigate($oIE, $sLink & $iPage)
	  While 1
		 Local $oTags = _IETagNameGetCollection($oIE, 'li')
		 For $oTag In $oTags
			If $oTag.getAttribute('class') = 'nxt-product-item' Then
			   Local $temp = ''
			   $temp = StringRegExpReplace($oTag.innertext, '[\r,\n]', '')
			   $temp = StringRegExpReplace($temp, '^ +', '')
			   $sCurrent &= $temp & @CRLF
			EndIf
		 Next

		 If $sCurrent <> $sPrevious Then
			$sTxt = $sTxt & $sCurrent
			$sPrevious = $sCurrent
			$sCurrent = ''
			ExitLoop
		 Else
			$sCurrent = ''
		 EndIf

	  WEnd

	  Sleep(500)

   Next


   _ClipBoard_SetData($sTxt)
   MsgBox (0,0,'done')

   Exit


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