#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>
#include <Array.au3>
#include <IE.au3>
#include <GUIConstantsEx.au3>
#include <InetConstants.au3>
#include <Clipboard.au3>



Global $sMasterPath = @ScriptDir
Global $bGetHTML_Flag = False

Func Set_GetHtml_Flag ()
   $bGetHTML_Flag = True
EndFunc



#Region ### START GUI section ### Form=
   Opt('GUIOnEventMode', 1)
   $Form1 = GUICreate('Collector', 500, 200, -1, -1)
   GUISetOnEvent($GUI_EVENT_CLOSE, 'Autoit_Exit')
   GUISetBkColor(0xFFFFFF)
   ;-------------------------------------------
   ;Create Get Info button
   $Button_Info = GUICtrlCreateButton('Get Files', 200, 30, 100, 50)
   GUICtrlSetOnEvent($Button_Info, 'Set_GetHtml_Flag')
   ;-------------------------------------------
   ;CREATE GUI NOTIFICATION PLACE
   $Commu_Ctrl = GUICtrlCreateLabel('', 50, 110, 400, 50)
   $CopyRight = GUICtrlCreateLabel('Created by Thien Nguyen', 190, 180, 150, 50)
   ;-------------------------------------------
		 ;-------------------------------------------
		 ;Create input
		 ;$sDelay = GUICtrlCreateInput("", 32, 60, 265, 21)
		 ;GUICtrlCreateLabel("Input Vehicle Link ", 120, 40, 114, 17)



   ;SHOW GUI
   GUISetState(@SW_SHOW)
   GUICtrlSetData ($Commu_Ctrl, 'Press the button to begin to get the documents')
#EndRegion ### END Koda GUI section ###



While 1
   If $bGetHTML_Flag = True Then GetHTML ()
WEnd



Func GetHTML()
   Local $oIE = _IEAttach ('Honda Service Express')

   ;WinActivate('HONDA Service Express - Windows Internet Explorer')

   Local $oFrameContent = _IEFrameGetObjByName ($oIE, 'content')
   Local $oLinks = _IELinkGetCollection($oFrameContent)
   Local $iNumLinks = @extended
   Local $aLinks [0][3]

   Local $oFrameNavigation = _IEFrameGetObjByName ($oIE, 'navigation')
   Local $oFrameNaviTop = _IEFrameGetObjByName ($oFrameNavigation, 'navigation_top')
   Local $oModel = _IEGetObjByName($oFrameNaviTop, 'model')
   Local $oYear = _IEGetObjByName($oFrameNaviTop, 'year')

   Local $sModel = _IEFormElementGetValue ($oModel)
   Local $sYear = _IEFormElementGetValue ($oYear)


   Local $sTempTxt
   For $oLink In $oLinks
	  If StringInStr ($oLink.href, 'DTC Advanced Diagnostics') <> 0 And $oLink.innertext <> '' Then
		 ReDim $aLinks [UBound($aLinks)+1][3]
		 $sTempTxt = StringRegExpReplace($oLink.href, 'javascript.+pubs/', 'pubs/')
		 $sTempTxt = 'https://techinfo.honda.com/rjanisis/' & StringReplace($sTempTxt, ''')', '') & '.html'
		 $aLinks [UBound($aLinks)-1][0] = $sTempTxt
		 $sTempTxt = StringRegExpReplace($oLink.innertext, 'DTC Advanced Diagnostics: ','')
		 $sTempTxt = StringRegExpReplace($sTempTxt,  '[\/:*?"<>|]', '-')
		 $aLinks [UBound($aLinks)-1][1] = $sTempTxt
		 $aLinks [UBound($aLinks)-1][2] = StringLen($sTempTxt)
	  EndIf
   Next

   Local $oIEContent = _IECreate ('',0,1,1,0)

   ;Get lastlink to determine the last process
   Local $sLastLink = LogFile ($sMasterPath, 'Read')
   ;This flag is used to allow the loop to only process the undone links
   Local $sActivate_Flag = False
   For $i=0 To UBound($aLinks) - 1
	  ;Check last link
	  If StringInStr($sLastLink, 'techinfo') <> 0 Then
		 If $aLinks[$i][0] = $sLastLink Then
			Local $sActivate_Flag = True
			ContinueLoop
		 EndIf
		 ;If not activate the flag => bypass the link
		 If $sActivate_Flag = False Then ContinueLoop
	  EndIf



	  ;Navigate to the site needed to get html
	  _IENavigate($oIEContent, $aLinks[$i][0])


	  If StringInStr (_IEDocReadHTML ($oIEContent), 'FRAMESET') <> 0 Then
		 GUICtrlSetData ($Commu_Ctrl, 'Getting: ' & $aLinks[$i][1] & '.html')

		 ;Get frame => get html => write html
		 Local $oFrameDTCContent = _IEFrameGetObjByName ($oIEContent, 'content')
		 Local $shtml = _IEDocReadHTML ($oFrameDTCContent)
		 Create_HTML ($shtml, $sMasterPath & '\' & '[' & $sYear & ' ' & $sModel & '] ' & $aLinks[$i][1] & '.html')

		 ;Write log file
		 LogFile ($sMasterPath, 'Write', '[' & $sYear & ' ' & $sModel & '] ' & 'Successfully created html file from:' & @CRLF & $aLinks[$i][0])

		 ;Delay to be like a human
		 Sleep (3000)
	  Else
		 GUICtrlSetData ($Commu_Ctrl, 'Getting: ' & $aLinks[$i][1] & '.pdf')
		 ;Use manually click and button to control the save funtion

		 ;Wait for the IE content object and get handle
		 Local $hIEContent = WinWait('https://techinfo.honda.com/rjanisis/pubs')
		 ;Check if the pdf frame appears
		 While 1
			If ControlCommand($hIEContent, '', 'AfxWnd100su4', 'IsEnabled', '') Then ExitLoop
			Sleep(500) ; Important or you eat all the CPU which will not help the defrag!
		 WEnd
		 Sleep (1000)
		 ;Send CTRL + P to open print window
		 ControlSend($hIEContent, '', 'AfxWnd100su4', '^p')
		 ;Press OK
		 Sleep(500)
		 ControlSend(WinWait('Print'), '', 'Button44', '{Enter}')
		 ;Type new name
		 Sleep(500)
		 ControlSetText(WinWait('Print to PDF Document - Foxit Reader PDF Printer'), '', 'Edit1', '[' & $sYear & ' ' & $sModel & '] ' & $aLinks[$i][1] & '.pdf')

		 ;Press Save
		 Sleep(500)
		 ControlSend(WinWait('Print to PDF Document - Foxit Reader PDF Printer'), '', 'Button2', '{Enter}')


		 ;Write log file
		 LogFile ($sMasterPath, 'Write', '[' & $sYear & ' ' & $sModel & '] ' & 'Successfully created pdf file from:' & @CRLF & $aLinks[$i][0])


		 ;Delay to be like a human
		 Sleep (1000)
	  EndIf

   Next
   ;Quit IE Content
   _IEQuit($oIEContent)

   GUICtrlSetData ($Commu_Ctrl, 'Done')
   $bGetHTML_Flag = False

   SoundPlay (@ScriptDir & '\#beep-01a.wav', 1)
   Sleep (500)
   SoundPlay (@ScriptDir & '\#beep-01a.wav', 1)
   Sleep (500)
   SoundPlay (@ScriptDir & '\#beep-01a.wav', 1)
   Sleep (500)

EndFunc




;====================================================================================================================
;                  FUNCTION DISCRIPTION: DOWNLOAD AN IMAGE A THE LINK
;				   INPUT               : $sFilePath, $sLink
;                  OUTPUT              : AN JPG IMAGE
;====================================================================================================================
Func FileDownload($sLink, $sFileName)
   Local $sFilePath = @ScriptDir
   ; Download the file in the background with the selected option of 'force a reload from the remote site.'
;   Local $hDownload = InetGet($sLink, $sFilePath &"\"& $sFileName, $INET_FORCERELOAD, $INET_DOWNLOADBACKGROUND)
   ; Wait for the download to complete by monitoring when the 2nd index value of InetGetInfo returns True.
;   Do
;	  Sleep(250)
;   Until InetGetInfo($hDownload, $INET_DOWNLOADCOMPLETE)

   ; Download the file by waiting for it to complete. The option of 'get the file from the local cache' has been selected.
   InetGet($sLink, $sFilePath &"\"& $sFileName, $INET_LOCALCACHE , $INET_DOWNLOADWAIT)
EndFunc


;====================================================================================================================
;                  FUNCTION DISCRIPTION:
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func GetItemStringByMark ($sString, $sStartMark, $sEndMark)
   If StringInStr ($sString, $sStartMark, 0, 1, 1) <> 0 Then
	  Local $iStart = StringInStr ($sString, $sStartMark, 0, 1, 1) + StringLen ($sStartMark)
	  Local $iEnd = StringInStr ($sString, $sEndMark, 0, 1, $iStart)
	  Local $sItemString = StringMid ($sString, $iStart, $iEnd - $iStart)
   Else
	  Local $sItemString = ""
   EndIf
   Return $sItemString
EndFunc




;====================================================================================================================
;                  FUNCTION DISCRIPTION: CREATE HTML FILE FROM
;				   INPUT               : $sFilePath, $sTxt_Title,$HTML_body
;                  OUTPUT              : AN HTML FILE IN $sFilePath
;====================================================================================================================
Func Create_HTML ($html, $sFilePath)
   Local $hFileOpen = FileOpen ($sFilePath,$FO_OVERWRITE)
   FileWrite($hFileOpen, $html)
   FileClose($hFileOpen)
EndFunc



;====================================================================================================================
;                  FUNCTION DISCRIPTION: CREATE LOG FILE OF DTC OR PROCEDURE
;				   INPUT               : $sFilePath, $sWhichLogFile,  $sTxt, $sMode
;                  OUTPUT              : AN LOG FILE IN $sFilePath
;====================================================================================================================
Func LogFile ($sFilePath, $sMode, $sTxt = '')
   If $sMode = 'Read' Then
	  Local $hFileOpen = FileOpen ($sFilePath & '\#LogFile.txt', $FO_READ )
	  Local $sFileRead = FileReadLine($hFileOpen, 2)
	  FileClose($hFileOpen)
	  Return $sFileRead
   Else
	  Local $hFileOpen = FileOpen ($sFilePath & '\#LogFile.txt',$FO_OVERWRITE)
	  FileWrite($hFileOpen, $sTxt)
	  FileClose($hFileOpen)
	  Return 1
   EndIf
EndFunc




;====================================================================================================================
;                  FUNCTION DISCRIPTION: LOAD FILE CONTENT
;                  INPUT			   :
; 				   OUTPUT			   :
;====================================================================================================================
Func ReadLog ($sFileName)
   ;Open YMME config file and get data
   Local $hFileOpen = FileOpen ($sFilePath & "\" & $sFileName & ".txt",$FO_READ )
   Local $sFileRead = FileRead($hFileOpen)
   FileClose($hFileOpen)
   ;Remove redundant @CRLF
   $sFileRead = StringRegExpReplace ($sFileRead, '[\r\n]+', @CRLF)
   $sFileRead = StringRegExpReplace ($sFileRead, '[\r\n]+$', '')
   Return $sFileRead
EndFunc



Func Autoit_Exit ()
   Exit
EndFunc




#cs
Local $bAttach = 1
Local $bVisible = 1
Local $bWait = 1
Local $bTakeFocus = 1


Local $aModel =  ["ACCORD", "ACCORD HYBRID", "ACCORD PLUG-IN", _
				  "CIVIC", "CIVIC 3-DOOR", "CIVIC HYBRID", _
				  "CLARITY FUEL CELL", _
				  "CLARITY PLUG-IN", "CR-V", "CR-Z", _
				  "CROSSTOUR", "CRX", "DEL SOL", "ELEMENT", _
				  "FIT", "HR-V", "INSIGHT", _
				  "LEGEND", "ODYSSEY", "PASSPORT", _
				  "PILOT", "PRELUDE", "RIDGELINE", "S2000"]

Local $oIE = _IEAttach ('Honda Service Express')
Local $oIEContent = _IECreate ('')
WinActivate('HONDA Service Express - Windows Internet Explorer')

Local $oFrameHeader = _IEFrameGetObjByName ($oIE, 'header')
Local $oFrameNavigation = _IEFrameGetObjByName ($oIE, 'navigation')
Local $oFrameNaviTop = _IEFrameGetObjByName ($oFrameNavigation, 'navigation_top')
Local $oFrameNaviTree = _IEFrameGetObjByName ($oFrameNavigation, 'navigation_tree')
Local $oFrameFooter = _IEFrameGetObjByName ($oIE, 'footer')
Local $oFrameContent = _IEFrameGetObjByName ($oIE, 'content')


Local $oLinks = _IELinkGetCollection($oFrameContent)
Local $iNumLinks = @extended
Local $aLinks [0]

Local $sTempTxt
For $oLink In $oLinks
   If StringInStr ($oLink.href, 'DTC Advanced Diagnostics') <> 0 Then
	  ReDim $aLinks [UBound($aLinks)+1]
	  $sTempTxt = StringRegExpReplace($oLink.href, 'javascript.+pubs/', 'pubs/')
	  $sTempTxt = 'https://techinfo.honda.com/rjanisis/' & StringReplace($sTempTxt, ''')', '') & '.html'
	  $aLinks [UBound($aLinks)-1] = $sTempTxt
   EndIf
Next

$aLinks = _ArrayUnique($aLinks)
_IENavigate($oIEContent, $aLinks[10])
Local $oFrameDTCContent = _IEFrameGetObjByName ($oIEContent, 'content')
Local $shtml = _IEDocReadHTML ($oFrameDTCContent)

Create_HTML ($shtml)

Exit

Local $oModel = _IEGetObjByName($oFrameNaviTop, 'model')
Local $oYear = _IEGetObjByName($oFrameNaviTop, 'year')
Local $oSubject = _IEGetObjByName($oFrameNaviTop, 'subject')

For $vElement In $aModel
   _IEFormElementOptionSelect ($oModel, $vElement)
   For $j = 2019 To 1996 Step -1
	  _IEFormElementOptionSelect ($oYear, $j)
	  _IEFormElementOptionSelect ($oSubject, 'D')

	  ;Check if the master folder appears => get object
	  While 1
		 Local $oMasterFolderButton = _IEGetObjByName($oFrameNaviTree, 'joustEntry0')
		 If Not @error Then ExitLoop
		 Sleep (1000)
	  WEnd
	  Sleep(2000) ;Make sure when the object disappears and appears again
	  _IEAction($oMasterFolderButton, 'click')

	  ;Check if the ECM/PCM appears => get object
	  While 1
		 Local $oPCM = _IEGetObjByName($oFrameNaviTree, 'joustEntry12')
		 If Not @error Then ExitLoop
		 Sleep (1000)
	  WEnd
	  Sleep(2000)
	  _IEAction($oPCM, 'click')

	  Exit
   Next
Next

#ce

