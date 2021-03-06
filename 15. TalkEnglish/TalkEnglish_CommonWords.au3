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


;Set Hotkey for the program
HotKeySet("{ESC}", "Autoit_Exit")


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
   $Form1 = GUICreate('Get Mp3 <talkenglish.com>', 400, 200, -1, -1)
   GUISetBkColor(0xFFFFFF)
   ;-------------------------------------------
   ;Create Labels
   $Commu_Ctrl = GUICtrlCreateLabel('', 50, 160, 250, 60)


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
   ;The link contain words
   Local $sMasterLink = 'https://www.talkenglish.com/vocabulary/top-2000-vocabulary.aspx'

   ;create IE object with the link
   Local $oIE = _IECreate ($sMasterLink,0,0,1,0)

   ;Get links collection
   Local $oLinks = _IELinkGetCollection($oIE)
   Local $aLinks [1]
   ;convert object to an array containing links

   For $oLink In $oLinks
	  If StringInStr($oLink.href, 'how-to-use') <> 0 And StringInStr($oLink.href, 'aspx') = 0 Then
		 ReDim $aLinks[UBound($aLinks)+1]
		 $aLinks[UBound($aLinks)-1] = $oLink.href
	  EndIf
   Next

;~    _ArrayDisplay($aLinks)
   Local $sPageNum = Number(GUICtrlRead ($PageEdit))
   If $sPageNum = '' Then $sPageNum = 1

   Local $sPageNumEnd = Number(GUICtrlRead ($ToPageEdit))
   If $sPageNumEnd = '' Then $sPageNumEnd =  UBound($aLinks) - 1

;~    MsgBox (0, '', 'Page ' & $sPageNum & ' To ' & $sPageNumEnd, 10)

   For $sPageNumCount = $sPageNum  To $sPageNumEnd
	  Local $sUrl = $aLinks[$sPageNumCount]

	  _IENavigate ($oIE, $sUrl, 1)
	  Sleep(300)


	  Local $sTitle = _IEPropertyGet($oIE, 'title')
	  $sTitle = StringReplace($sTitle, 'How to Use ','')
	  $sTitle = StringReplace($sTitle, ' with Example Sentences','')

	  ;Create word folder
	  Local $sWordPath = $sMasterFolderPath & '\' & $sPageNumCount & '. ' & $sTitle
	  DirCreate ($sWordPath)

	  ;get data from the tag name <a>
	  Sleep(100)
	  Local $oAs = _IETagNameGetCollection ($oIE, 'a')
	  Local $iNumber = @extended
	  Local $iCount = 0
	  ;Loop object to get data

	  For $oA In $oAs

		 Local $sTempNum = StringRegExpReplace($oA.href, '^.+\D(?=\d+\.mp3)','')
		 $sTempNum = StringRegExpReplace($sTempNum, '.mp3)','')

		 If Number($sTempNum) >= 3 Then
			$iCount = $iCount + 1
			;save the name
			Local $sFileName = StringRegExpReplace($oA.innertext, '[."!?/|\<>]','')

			;download mp3
			GUICtrlSetData ($Commu_Ctrl, 'Page ' & $sPageNumCount & ' -- (' & $iCount & '/' & $iNumber  & ') ' & 'Downloading ' & $sFileName)
			Download($oA.href, $sWordPath & '\' & $sFileName & '.mp3')
		 EndIf
	  Next
   Next

   _IEQuit ($oIE)
   GUICtrlSetData ($Commu_Ctrl, 'Done')
   $bDownload_Flag = False
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