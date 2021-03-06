#cs
	NOTE: Before using this script, need to open IE with url "https://techinfo.toyota.com/"
	login  => select TIS => select tab RM
#ce




#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>
#include <Array.au3>
#include <IE.au3>
#include <GUIConstantsEx.au3>
#include <InetConstants.au3>
#include <Clipboard.au3>
#include <StringConstants.au3>

HotKeySet('{ESC}', 'Autoit_Exit')

Global $sMasterPath = @ScriptDir
Global $bGetHTML_Flag = False

Func Set_GetHtml_Flag ()
	$bGetHTML_Flag = True
EndFunc

GUIInit()
;
While 1
	If $bGetHTML_Flag = True Then SelectVehicle()
WEnd

;====================================================================================================
;This function is Exit AutoIT
;====================================================================================================
Func Autoit_Exit ()
	Exit
EndFunc


;====================================================================================================
;This function is to create GUI
;====================================================================================================
Func GUIInit()
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
	;Create input
	;$sDelay = GUICtrlCreateInput("", 32, 60, 265, 21)
	;GUICtrlCreateLabel("Input Vehicle Link ", 120, 40, 114, 17)

	;SHOW GUI
	GUISetState(@SW_SHOW)
	GUICtrlSetData ($Commu_Ctrl, 'Press the button to begin to get the documents')
EndFunc


;====================================================================================================
;This function is to remove redundant enter characters in a string just keep 1 enter each line
;Also remove beginning and ending enters
;====================================================================================================
Func _StringRemoveRedundantEnter ($sString)
	;process the string
	$sString = StringRegExpReplace($sString, '[\r\n]+', @CRLF)	;replace 2 enters by 1 enter
	$sString = StringRegExpReplace($sString, '[\r\n]+$', '')	;remove beginning enters
	$sString = StringRegExpReplace($sString, '^[\r\n]+', '')	;remove ending enters
	;return processed string
	Return $sString
EndFunc


;====================================================================================================
;This function is to get options in a form object and put in to an array
;====================================================================================================
Func _IEGetOptions (Byref $oIE, Byref $oObject)
	Local $shtml = _IEPropertyGet ($oObject, 'outerhtml')		;get html of object
	_ClipBoard_SetData($shtml)

	exit

	$shtml = StringReplace($shtml, '><', '>' & @CRLF & '<')		;replace '><' by enter in between to have multi-line string
	$shtml = StringRegExpReplace($shtml, '<.+".+">', '')		;remove redundant text to get options text only
	$shtml = StringRegExpReplace($shtml, '<.+>', '')			;remove redundant text to get options text only
	$shtml = StringRegExpReplace($shtml, '<.+>', '')			;remove redundant text to get options text only
	Local $sAllOptions = _StringRemoveRedundantEnter ($shtml)	;remove redundant enters
	Local $aOption = StringSplit($sAllOptions, @CRLF,  $STR_ENTIRESPLIT + $STR_NOCOUNT)	;convert string options to an arrray
	;return the array of options
	Return $aOption
EndFunc



;====================================================================================================
;create html file from input string
;filepath needs to include the html name with extension ".html"
;====================================================================================================
Func CreateHTML ($html, $sFilePath)
	Local $hFileOpen = FileOpen ($sFilePath,$FO_OVERWRITE)
	FileWrite($hFileOpen, $html)
	FileClose($hFileOpen)
EndFunc


;====================================================================================================
;This function will run first when clicking the button on GUI
;====================================================================================================
Func SelectVehicle()
	;create IE object with link
	Local $oIE = _IECreate('https://www.innova.com/en-US/Dlc',1)
	;
	;get make object then set value to it, always set Ford to the object
	Local $oMake = _IEGetObjById($oIE, 'DlcMake')
	_IEFormElementSetValue($oMake,'Ford')	;set valuae to make
	Sleep(2000)								;wait for the selection to be applied
	;
	;year loop
	For $i = 2017 To 1996 Step -1
		;get year object and set value
		Local $oYear = _IEGetObjById($oIE, 'DlcYear')
		_IEFormElementSetValue($oYear, 2017)	;set valuae to year
		Sleep(2000)		
		;
		;get model
		Local $oModel = _IEGetObjById($oIE, 'DlcModel')
		Local $aModelOption = _IEGetOptions($oIE, $oModel)


	Next








	;Select each model of the model array
	For $i = 1 To UBound($aModelOption) - 1					;skip the first option "All"
		_IEFormElementSetValue($oModel, $aModelOption[$i])	;set value to model
		Sleep(2000)											;wait for the selection to be applied
		;
		;get year object then get all options of the object and put in an array
		Local $oYear = _IEGetObjByName($oIE, 'repairformwlw-select_key:{actionForm.year}')
		Local $aYearOption = _IEGetOptions($oIE, $oYear)
		;
		;select each model of the model array
		For $j = 1 To UBound($aYearOption) - 1					;skip the first option "All"
			Local $oYear = _IEGetObjByName($oIE, 'repairformwlw-select_key:{actionForm.year}')
			_IEFormElementSetValue($oYear, $aYearOption[$j])	;set value to model
			Sleep(2000)											;wait for the selection to be applied
			;
			;get button object then click the button 	
			Local $oSearchButton = _IEGetObjById($oIE, 'searchButton')	;get the search button object	
			_IEAction($oSearchButton, 'click')							;click the button		
			_IELoadWait($oIE)											;wait for IE to load completely
			;
			;get all links of the current IE site
			Local $oLinks = _IELinkGetCollection ($oIE)
			;
			;check each link to find to Repair Manual Link
			For $oLink In $oLinks
				If (StringInStr($oLink.href, '/RM') <> 0) Then	;if the current link has "/RM" in it => this link contains Repair Manual
					Local $sRM_Link = $oLink.href				;get the RM link
					ExitLoop									;exit loop after get the link 
				EndIf
			Next
			;
			;tranfer the link to function GetRepairManual
			ConsoleWrite($sRM_Link & @CRLF & @CRLF)
			GetRepairManual($sRM_Link)
			;
			;exit loop if the selected year is 1996
			If ($aYearOption[$j] = 1996) Then ExitLoop
		Next
	Next
EndFunc



;====================================================================================================
;check to see if the RM site has done loading document
;====================================================================================================
Func _IECheckLoadDone ($oObject, $sStringToCheck)
	;wait for the string to disappear
	While 1
		If (StringInStr(_IEPropertyGet($oObject, 'innerthml'), $sStringToCheck <> 0) Then
			Sleep(100)
		Else
			ExitLoop
		EndIf
	Wend
	;wait for the string to appear
	While 1
		If (StringInStr(_IEPropertyGet($oObject, 'innerthml'), $sStringToCheck) = 0) Then
			Sleep(100)
		Else
			ExitLoop
		EndIf
	Wend	
EndFunc


;====================================================================================================
;Function to get repair manual html
;====================================================================================================
Func GetRepairManual($sRM_Link)
	;get IE objet of TOYOTA OEM website
	Local $oIE = _IEAttach ('TIS')

	;create IE object with RM link
	;Local $oIE = _IECreate($sRM_Link)
	;
	;get frame object 
	ConsoleWrite('get navigation_frame' & @CRLF)
	Local $oFrame = _IEFrameGetObjByName ($oIE, 'navigation_frame')
	;
	;find the section 'Engine/Hybrid System' and click to the item to expand
	Local $oAs = _IETagNameGetCollection ($oFrame, 'a')		;get all tag "a"
	For $oA In $oAs											;browse each object in the collection
		If StringInStr($oA.innertext, 'engine') Then 		;if the object title contains the word "engine"
			_IEAction($oA, 'click')							;click the object
			ConsoleWrite('>>> Click: ' & $oA.innertext & @CRLF)	
			ExitLoop
		EndIf
	Next
	;
	;wait for the site finish loading
	Local $sID = _IEGetObjById($oFrame, 'staticDiv')	;get the object which contains the string to check
	_IECheckLoadDone ($sID, 'display: block;')			;call the function the check the loading 
	;
	;find the section 'Engine Control' and click to the item to expand
	Local $oAs = _IETagNameGetCollection ($oFrame, 'a')			;get all tag "a"
	For $oA In $oAs												;browse each object in the collection
		If StringInStr($oA.innertext, 'engine control') Then 	;if the object title contains the word "engine control"
			_IEAction($oA, 'click')								;click the object
			ConsoleWrite('>>> Click: ' & $oA.innertext & @CRLF)	
			ExitLoop 											;exit loop after finding the object
		EndIf
	Next
	;
	;wait for the site finish loading
	Local $sID = _IEGetObjById($oFrame, 'staticDiv')	;get the object which contains the string to check
	_IECheckLoadDone ($sID, 'display: block;')			;call the function the check the loading 
	exit
EndFunc







Func NameFunc()

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
		$sTempTxt = StringRegExpReplace($sTempTxt,	'[\/:*?"<>|]', '-')
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
;						FUNCTION DISCRIPTION: DOWNLOAD AN IMAGE A THE LINK
;					INPUT					: $sFilePath, $sLink
;						OUTPUT					: AN JPG IMAGE
;====================================================================================================================
Func FileDownload($sLink, $sFileName)
	Local $sFilePath = @ScriptDir
	;Download the file in the background with the selected option of 'force a reload from the remote site.'
	;Local $hDownload = InetGet($sLink, $sFilePath &"\"& $sFileName, $INET_FORCERELOAD, $INET_DOWNLOADBACKGROUND)
	;Wait for the download to complete by monitoring when the 2nd index value of InetGetInfo returns True.
	;Do
	;Sleep(250)
	;Until InetGetInfo($hDownload, $INET_DOWNLOADCOMPLETE)

	;Download the file by waiting for it to complete. The option of 'get the file from the local cache' has been selected.
	InetGet($sLink, $sFilePath &"\"& $sFileName, $INET_LOCALCACHE , $INET_DOWNLOADWAIT)
EndFunc


;====================================================================================================================
;						FUNCTION DISCRIPTION:
;					INPUT					:
;						OUTPUT					:
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
;						FUNCTION DISCRIPTION: CREATE HTML FILE FROM
;					INPUT					: $sFilePath, $sTxt_Title,$HTML_body
;						OUTPUT					: AN HTML FILE IN $sFilePath
;====================================================================================================================
Func Create_HTML ($html, $sFilePath)
	Local $hFileOpen = FileOpen ($sFilePath,$FO_OVERWRITE)
	FileWrite($hFileOpen, $html)
	FileClose($hFileOpen)
EndFunc



;====================================================================================================================
;						FUNCTION DISCRIPTION: CREATE LOG FILE OF DTC OR PROCEDURE
;					INPUT					: $sFilePath, $sWhichLogFile,	$sTxt, $sMode
;						OUTPUT					: AN LOG FILE IN $sFilePath
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
;						FUNCTION DISCRIPTION: LOAD FILE CONTENT
;						INPUT				:
; 					OUTPUT				:
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








#cs
Local $bAttach = 1
Local $bVisible = 1
Local $bWait = 1
Local $bTakeFocus = 1


Local $aModel =	["ACCORD", "ACCORD HYBRID", "ACCORD PLUG-IN", _
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

