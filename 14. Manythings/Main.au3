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

#include <Clipboard.au3>
#include <IE.au3 >
#include <Excel.au3>


;Set Hotkey for the program
HotKeySet("{PAUSE}", "Autoit_Exit")


Global $bOpenDialog_Flag = False
Global $bDownload_Flag = False

Global $sMasterFolderPath = @ScriptDir

Func Set_OpenDialog_Flag ()
   $bOpenDialog_Flag = True
EndFunc

Func Set_Download_Flag ()
   $bDownload_Flag = True
EndFunc


#Region ### START GUI section ### Form=
   Opt('GUIOnEventMode', 1)
   $Form1 = GUICreate('Get Mp3 <manythings.org>', 400, 200, -1, -1)
   GUISetBkColor(0xFFFFFF)
   ;-------------------------------------------
   ;Create Labels
   $Commu_Ctrl = GUICtrlCreateLabel('', 50, 160, 250, 60)


   ;Create edit
   ;GUICtrlCreateLabel('Link manythings.org', 152, 10, 250, 20)
   ;$UrlEdit = GUICtrlCreateInput('', 50, 30, 300, 20)

   GUICtrlCreateLabel('From Page', 50, 30, 250, 20)
   $PageEdit = GUICtrlCreateInput('', 50, 50, 100, 20)


   GUICtrlCreateLabel('To Page', 250, 30, 250, 20)
   $ToPageEdit = GUICtrlCreateInput('', 250, 50, 100, 20)

   ;Create Button
   $Button_OpenDialog = GUICtrlCreateButton("Save to ...", 50, 110, 100, 30)
   GUICtrlSetOnEvent($Button_OpenDialog, 'Set_OpenDialog_Flag')

   $Button_Download = GUICtrlCreateButton("Download", 250, 110, 100, 30)
   GUICtrlSetOnEvent($Button_Download, 'Set_Download_Flag')
#EndRegion ### END Koda GUI section ###


;~ ;Open excel then open workbook
;~ Local $oExcel = OpenExcel ()
;~ Local $sWorkbook = 'F:\Google Drive\#1. Work PC Sync\Autoit Scripts\Manythings\20k English Words.xlsx'
;~ Local $oWorkbook = _Excel_BookOpen($oExcel, $sWorkbook)

;~ ;Put an example range into an array
;~ Global $aWordRank = _Excel_RangeRead ($oWorkbook, 'Sheet1', 'B2:B2001')


;SHOW GUI
GUISetState(@SW_SHOW)


While 1
   If $bOpenDialog_Flag = True Then OpenDialog()
   If $bDownload_Flag = True Then GetFullMp3 ()
WEnd




Func OpenDialog()
   $sMasterFolderPath = FileSelectFolder ('Select Folder', @ScriptDir)
   $bOpenDialog_Flag = False
EndFunc



Func GetFullMp3 ()

   Local $sPageNum = Number(GUICtrlRead ($PageEdit))
   If $sPageNum = '' Then $sPageNum = 1

   Local $sPageNumEnd = Number(GUICtrlRead ($ToPageEdit))
   If $sPageNumEnd = '' Then $sPageNumEnd = 2171

   Local $sPageNumCount
   Local $oIE = _IECreate ('',0,0,1,0)



   For $sPageNumCount = $sPageNum  To $sPageNumEnd
	  Local $sUrl = 'http://www.manythings.org/audio/sentences/' & $sPageNumCount & '.html'
	  _IENavigate ($oIE, $sUrl, 1)
	  Sleep(1000)
	  Local $sTitle = _IEPropertyGet($oIE, 'title')
	  $sTitle = StringRegExpReplace($sTitle, '.+"(?=\w)','')
	  $sTitle = StringReplace($sTitle, '"','')

	  ;Create word folder
	  Local $sWordPath = $sMasterFolderPath & '\' & $sPageNumCount & '. ' & $sTitle
	  DirCreate ($sWordPath)

	  ;get data from the tag name <dt>
	  Local $oDts = _IETagNameGetCollection ($oIE, 'dt')
	  Local $iNumber = @extended
	  Local $iCount = 0

	  ;Loop object to get data
	  For $oDt In $oDts
		 $iCount = $iCount + 1
		 ;save the name
		 Local $sFileName = StringRegExpReplace($oDt.innertext, '^.+] +','')
		 $sFileName = StringRegExpReplace($sFileName, '[?!./\:*<>|]','')

		 ;get link mp3
		 Local $oLink = _IETagNameGetCollection ($oDt, 'a', 0)
		 ;save link
		 Local $sLink = $oLink.href

		 ;download mp3
		 GUICtrlSetData ($Commu_Ctrl, 'Page ' & $sPageNumCount & ' -- (' & $iCount & '/' & $iNumber  & ') ' & 'Downloading ' & $sFileName)
		 Download($oLink.href, $sWordPath & '\' & $sFileName & '.mp3')
	  Next
	  $bDownload_Flag = False
   Next
   _IEQuit ($oIE)
   GUICtrlSetData ($Commu_Ctrl, 'Done')
EndFunc


Func GetMp3 ()
   Local $sRank = GUICtrlRead ($RankEdit)
   Local $sWord = GUICtrlRead ($WordEdit)
   Local $sWordPath = $sMasterFolderPath & '\' & $sRank & '. ' & $sWord
   ;Create word folder
   DirCreate ($sWordPath)

   Local $sUrl = GUICtrlRead ($UrlEdit)
   Local $oIE = _IECreate ($sUrl,0,0,1,0)

   ;get data from the tag name <dt>
   Local $oDts = _IETagNameGetCollection ($oIE, 'dt')
   Local $iNumber = @extended
   Local $iCount = 0
   ;Loop object to get data
   For $oDt In $oDts
	  $iCount = $iCount + 1
	  ;save the name
	  Local $sFileName = StringRegExpReplace($oDt.innertext, '^.+] +','')
	  $sFileName = StringRegExpReplace($sFileName, '\W','_')
	  $sFileName = StringRegExpReplace($sFileName, '_+$','')
	  $sFileName = $sRank & '.' & '[' & $sWord & ']' & ' ' & $sFileName

	  ;get link mp3
	  Local $oLink = _IETagNameGetCollection ($oDt, 'a', 0)
	  ;save link
	  Local $sLink = $oLink.href


	  ;download mp3
	  GUICtrlSetData ($Commu_Ctrl, '(' & $iCount & '/' & $iNumber  & ') ' & 'Downloading ' & $sFileName)
	  Download($oLink.href, $sWordPath & '/' & $sFileName & '.mp3')
   Next
   _IEQuit ($oIE)
   $bDownload_Flag = False

   GUICtrlSetData ($Commu_Ctrl, 'Done')
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

		 If $sExceed_Time > 5000 Then Exit
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