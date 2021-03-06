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
#include <InetConstants.au3>

HotKeySet('{ESC}', 'Autoit_Exit')






;====================================================================================================================
;                  FUNCTION DISCRIPTION: LOAD FILE CONTENT
;                  INPUT               :
;                  OUTPUT              :
;====================================================================================================================

Global $oErrorHandler  = ObjEvent("AutoIt.Error","ErrFunc")
Global $bErrorOccurred = False
Global $sErrorPosition = ''
Global $bErrorHappened = False
Func ErrFunc()
  Local $sText = ("=====================================================================================" & @CRLF & _
             "COM Error!"    & @CRLF  & @CRLF & _
             "          err.description is : " & @TAB & $oErrorHandler.description  & @CRLF & _
             "          err.windescription : " & @TAB & $oErrorHandler.windescription & @CRLF & _
             "          err.number is      : " & @TAB & hex($oErrorHandler.number,8)  & @CRLF & _
             "          err.lastdllerror is: " & @TAB & $oErrorHandler.lastdllerror   & @CRLF & _
             "          err.scriptline is  : " & @TAB & $oErrorHandler.scriptline   & @CRLF & _
             "          err.source is      : " & @TAB & $oErrorHandler.source       & @CRLF & _
             "          err.helpfile is  : " & @TAB & $oErrorHandler.helpfile     & @CRLF & _
             "          err.helpcontext is : " & @TAB & $oErrorHandler.helpcontext & @CRLF & _
             "          link when error    : " & @TAB & $sErrorPosition)
   $sText = StringRegExpReplace ($sText, '\r\n\r\n+', @CRLF)
   $bErrorOccurred = True
   $bErrorHappened = True
   Write_Error (@CRLF & @CRLF & $sText)
Endfunc



Func Write_Error ($sText)
   Static Local $bFirst = True
   If $bFirst = True Then
      Local $hFileOpen = FileOpen(@ScriptDir & "/ErrorLog.txt",  $FO_OVERWRITE)
      FileWrite($hFileOpen, $sText)
      FileClose($hFileOpen)
      $bFirst = False
   Else
      Local $hFileOpen = FileOpen(@ScriptDir & "/ErrorLog.txt", $FO_APPEND)
      FileWrite($hFileOpen, $sText)
      FileClose($hFileOpen)
   EndIf
EndFunc













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
    Global $Form1 = GUICreate('Collector', 400, 200, -1, -1)
    GUISetOnEvent($GUI_EVENT_CLOSE, 'Autoit_Exit')
    GUISetBkColor(0xFFFFFF)
    ;-------------------------------------------
    ;Create Get Info button
    Global $Button_Info = GUICtrlCreateButton('Get Files', 150, 70, 100, 30)
    GUICtrlSetOnEvent($Button_Info, 'Set_GetHtml_Flag')
    ;-------------------------------------------
    ;CREATE GUI NOTIFICATION PLACE
    Global $Commu_Ctrl = GUICtrlCreateLabel('', 50, 110, 300, 55)
    Global $CopyRight = GUICtrlCreateLabel('Created by Thien Nguyen', 150, 180, 150, 50)
    ;-------------------------------------------
    ;Create input
    Global $Model = GUICtrlCreateInput("", 50, 40, 120, 21)
    GUICtrlCreateLabel("Enter the vehicle name:", 50, 20, 150, 17)
    Global $StartYear = GUICtrlCreateInput("", 230, 40, 120, 21)
    GUICtrlCreateLabel("Enter the beginning year:", 230, 20, 150, 17)


    ; If the directory exists the don't continue.
    If FileExists(@ScriptDir & "\Downloads") = False Then
        ; Create the directory.
        DirCreate(@ScriptDir & "\Downloads")
    EndIf



    ;SHOW GUI
    GUISetState(@SW_SHOW)
    GUICtrlSetData ($Commu_Ctrl, 'Please use IE to login Toyota OEM website' & @CRLF & _
                                 'Select tab "RM" in the website' & @CRLF & _
                                 'Enter Model name and Beginning Year' & @CRLF & _
                                 'Then press the button on the app to begin getting documents')
EndFunc

;====================================================================================================
;This function is to remove redundant enter characters in a string just keep 1 enter each line
;Also remove beginning and ending enters
;====================================================================================================
Func _StringRemoveRedundantEnter ($sString)
    ;process the string
    $sString = StringRegExpReplace($sString, '[\r\n]+', @CRLF)    ;replace 2 enters by 1 enter
    $sString = StringRegExpReplace($sString, '[\r\n]+$', '')    ;remove beginning enters
    $sString = StringRegExpReplace($sString, '^[\r\n]+', '')    ;remove ending enters
    ;return processed string
    Return $sString
EndFunc


;====================================================================================================
;This function is to get options in a form object and put in to an array
;====================================================================================================
Func _IEGetOptions (Byref $oIE, Byref $oObject)
    Local $shtml = _IEPropertyGet ($oObject, 'outerhtml')        ;get html of object
    $shtml = StringReplace($shtml, '><', '>' & @CRLF & '<')        ;replace '><' by enter in between to have multi-line string
    $shtml = StringRegExpReplace($shtml, '<.+".+">', '')        ;remove redundant text to get options text only
    $shtml = StringRegExpReplace($shtml, '<.+>', '')            ;remove redundant text to get options text only
    $shtml = StringRegExpReplace($shtml, '<.+>', '')            ;remove redundant text to get options text only
    Local $sAllOptions = _StringRemoveRedundantEnter ($shtml)    ;remove redundant enters
    Local $aOption = StringSplit($sAllOptions, @CRLF,  $STR_ENTIRESPLIT + $STR_NOCOUNT)    ;convert string options to an arrray
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
	Local $sModel = GUICtrlRead ($Model)
	Local $iStartYear = Number(GUICtrlRead ($StartYear))

    ;get IE objet of TOYOTA OEM website
    Local $oIE = _IEAttach ('TIS')
    ;
    GUICtrlSetData ($Commu_Ctrl, 'Selecting Make ...')
    ;get make object then set value to it, always set TOYOTA to the object
    Local $oMake = _IEGetObjByName($oIE, 'repairformwlw-select_key:{actionForm.division}')
    _IEFormElementSetValue($oMake,'TOYOTA')    ;set valuae to make
    Sleep(2000)                                ;wait for the selection to be applied
    ;
    ;get model object then get all options of the object and put in an array
    Local $oModel = _IEGetObjByName($oIE, 'repairformwlw-select_key:{actionForm.model}')
    Local $aModelOption = _IEGetOptions($oIE, $oModel)
    ;
    ;Select each model of the model array
;    For $i = 1 To UBound($aModelOption) - 1                    ;skip the first option "All"
        GUICtrlSetData ($Commu_Ctrl, 'Selecting Model ...')
;        _IEFormElementSetValue($oModel, $aModelOption[$i])    ;set value to model
        _IEFormElementSetValue($oModel, $sModel)    ;set value to model
        Sleep(2000)                                            ;wait for the selection to be applied
        ;
        ;get year object then get all options of the object and put in an array
        Local $oYear = _IEGetObjByName($oIE, 'repairformwlw-select_key:{actionForm.year}')
        Local $aYearOption = _IEGetOptions($oIE, $oYear)
        ;
        ;select each model of the model array
        For $j = 1 To UBound($aYearOption) - 1                    ;skip the first option "All"
            GUICtrlSetData ($Commu_Ctrl, 'Selecting Year ...')
            If ($aYearOption[$j] > $iStartYear) Then ContinueLoop

            Local $oYear = _IEGetObjByName($oIE, 'repairformwlw-select_key:{actionForm.year}')
            _IEFormElementSetValue($oYear, $aYearOption[$j])    ;set value to model
            Sleep(2000)                                            ;wait for the selection to be applied
            ;
            GUICtrlSetData ($Commu_Ctrl, 'Searching for Repair Manual ...')



            ;get button object then click the button     
            Local $oKeyWord = _IEGetObjById($oIE, 'keyword')     
            _IEFormElementSetValue ($oKeyWord, 'introduction')


            ;get button object then click the button     
            Local $oSearchButton = _IEGetObjById($oIE, 'searchButton')    ;get the search button object    
            _IEAction($oSearchButton, 'click')                            ;click the button        
            _IELoadWait($oIE)                                            ;wait for IE to load completely
            ;
            ;get all links of the current IE site
            Local $oLinks = _IELinkGetCollection ($oIE)
            ;
            ;check each link to find to Repair Manual Link
            For $oLink In $oLinks
                If (StringInStr($oLink.href, '/RM') <> 0) Then    ;if the current link has "/RM" in it => this link contains Repair Manual
                    Local $sRM_Link = $oLink.href                ;get the RM link
                    ExitLoop                                    ;exit loop after get the link 
                EndIf
            Next
            ;
            ;tranfer the link to function GetRepairManual
            Local $sModelYear = $sModel & ' ' & $aYearOption[$j]
            GetRepairManual($sRM_Link, $sModelYear)



            ;
            ;exit loop if the selected year is 1996
            If ($aYearOption[$j] = 1996) Then ExitLoop
        Next
;    Next
	GUICtrlSetData ($Commu_Ctrl, 'Done')
    $bGetHTML_Flag = False
EndFunc



;====================================================================================================
;check to see if the RM site has done loading document
;====================================================================================================
Func _IECheckLoadDone (Byref $oIE)

	Local $oCheckObject = null
	while 1
		ConsoleWrite('Checking Load ...' & @CRLF)
		Local $oFrame = _IEFrameGetObjByName ($oIE, 'navigation_frame')
	    ;find the section 'Engine Control' and click to the item to expand
	    Local $oAs = _IETagNameGetCollection ($oFrame, 'a')            ;get all tag "a"
	    For $oA In $oAs                                                ;browse each object in the collection
	        If StringInStr($oA.innertext, 'SFI') Or StringInStr($oA.innertext, 'ENGINE CONTROL') Then     ;if the object title contains the word "engine control" 
	            $oCheckObject = $oA
	            ExitLoop 2
	        EndIf
	    Next
	    Sleep(300)
	Wend
	ConsoleWrite('Done Checking Load ...' & @CRLF)
EndFunc


;====================================================================================================
;check to see if the RM site has done loading document
;====================================================================================================
Func _IECheckLoadDonezzzzz (Byref $oIE, $sStringToCheck)
    ;wait for the string to disappear
    While 1
        while 1
            ;Sleep(10)
            Local $oFrame = _IEFrameGetObjByName ($oIE, 'navigation_frame')
            ;get the object which contains the string to check
            Local $sID = _IEGetObjById($oFrame, 'staticDiv')    
            If (IsObj($sID)) Then ExitLoop
        Wend

        If (StringInStr(_IEPropertyGet($sID, 'outerhtml'), $sStringToCheck) <> 0) Then
            ConsoleWrite('checking load done: string ON' & @CRLF)
            ;Sleep(10)
        Else
            ExitLoop
        EndIf
    Wend
    ;wait for the string to appear
    While 1
        while 1
            ;Sleep(10)
            Local $oFrame = _IEFrameGetObjByName ($oIE, 'navigation_frame')
            ;get the object which contains the string to check
            Local $sID = _IEGetObjById($oFrame, 'staticDiv')    
            If (IsObj($sID)) Then ExitLoop
        Wend

        If (StringInStr(_IEPropertyGet($sID, 'outerhtml'), $sStringToCheck) = 0) Then
            ConsoleWrite('checking load done: string OFF' & @CRLF)
            ;Sleep(10)
        Else
            ExitLoop
        EndIf
    Wend    

    ConsoleWrite('checking load done: string back ON => done checking' & @CRLF)
EndFunc



;GetRepairManual('https://techinfo.toyota.com/t3Portal/resources/jsp/siviewer/index.jsp?dir=rm/RM1091U&href=xhtml/RM1091U_0004.html&locale=en&model=Tundra&MY=2004&keyWord=introduction&t3id=RM1091U_0004&User=false&publicationNumber=RM1091U&objType=rm&docid=en_rm_RM1091U_RM1091U_0004&context=ti', 'aaa')

;====================================================================================================
;Function to get repair manual html
;====================================================================================================
Func GetRepairManual($sRM_Link, $sModelYear)
    While 1
    	GUICtrlSetData ($Commu_Ctrl, 'Getting Repair Manual ...')
    	ConsoleWrite('Get Repair Manual' & @CRLF & $sRM_Link & @CRLF & @CRLF)
        ;get IE objet of TOYOTA OEM website
        ;Local $oIE = _IEAttach ('TIS')

        ;create IE object with RM link
        Local $oIE = _IECreate($sRM_Link)
        Sleep(500)



        If $bErrorHappened = True Then 
            Sleep(1000)
            _IEQuit($oIE)
            $bErrorHappened = False
            ContinueLoop
        EndIf



        ;get button object then click the button     
        Local $oLoginButton = _IEGetObjById($oIE, 'externalloginsubmit')    ;get the search button object    
        _IEAction($oLoginButton, 'click')                            ;click the button        
        _IELoadWait($oIE) 



        If $bErrorHappened = True Then 
            Sleep(1000)
            _IEQuit($oIE)
            $bErrorHappened = False
            ContinueLoop
        EndIf


        ;
        ;get frame object 
        ConsoleWrite('get navigation_frame' & @CRLF)
        Local $oFrame = _IEFrameGetObjByName ($oIE, 'navigation_frame')
        ;



        If $bErrorHappened = True Then 
            Sleep(1000)
            _IEQuit($oIE)
            $bErrorHappened = False
            ContinueLoop
        EndIf



        ;find the section 'Engine/Hybrid System' and click to the item to expand
        Local $oAs = _IETagNameGetCollection ($oFrame, 'a')        ;get all tag "a"
        For $oA In $oAs                                            ;browse each object in the collection
            If StringInStr($oA.innertext, 'engine') Then         ;if the object title contains the word "engine"
                _IEAction($oA, 'click')                            ;click the object
                ConsoleWrite('>>> Click: ' & $oA.innertext & @CRLF) 
                ExitLoop
            EndIf
        Next
        ;
        ;wait for the site finish loading ;display: block;
        _IECheckLoadDone ($oIE)            ;call the function to check the loading 
        Sleep(500)


        ;find the section 'Engine Control' and click to the item to expand
        Local $oAs = _IETagNameGetCollection ($oFrame, 'a')            ;get all tag "a"
        Local $oIE_RM = _IECreate('',0,1,1,0)  
        For $oA In $oAs                                                ;browse each object in the collection
            If StringInStr($oA.innertext, 'engine control') Then     ;if the object title contains the word "engine control"
                _IEAction($oA, 'click')                                ;click the object
                ConsoleWrite('>>> Click: ' & $oA.innertext & @CRLF)    
                ;wait for the site finish loading
                Sleep(500)
                
                Local $oEngineControl = $oA

    		    ;create SFI documents ID by SFI object
    		    Local $sEngineControlSub_ID = StringReplace($oEngineControl.id, 'i_txt', 'i_div')
    		    ;get object
    		    Local $oEngineControlSub = _IEGetObjById($oFrame, $sEngineControlSub_ID)
    		    ;get all document links



    		    ;find the section 'Engine Control' and click to the item to expand
    		    Local $oAs = _IETagNameGetCollection ($oEngineControlSub, 'a')            ;get all tag "a"
    		    For $oA In $oAs                                                ;browse each object in the collection
    		        If StringInStr($oA.innertext, 'SFI') Then     ;if the object title contains the word "engine control"
    		            _IEAction($oA, 'click')                                ;click the object
    		            ConsoleWrite('>>> Click: ' & $oA.innertext & @CRLF)    
    		            Local $oSFI = $oA
    		            ExitLoop
    		        EndIf
    		    Next
    		    ;
    		    ;wait for the site finish loading
    		    Sleep(500)
    		    ;
    		    ;GET ALL DOCUMENT LINKS
    		    ;Create Model Year Engine string to be used as a folder name
    		    Local $sModelYearEngine = $sModelYear & ' (' & StringLeft($oEngineControl.innertext, StringInStr($oEngineControl.innertext, ' ') - 1) & ')'
    		    ; If the directory exists the don't continue.
    		    If FileExists(@ScriptDir & "\Downloads\" & $sModelYearEngine) = False Then
    		        ; Create the directory.
    		        DirCreate(@ScriptDir & "\Downloads\" & $sModelYearEngine)
    		    EndIf
    		    $sFolderPath = @ScriptDir & "\Downloads\" & $sModelYearEngine


    		    ;create SFI documents ID by SFI object
    		    Local $sSFI_Documents_ID = StringReplace($oSFI.id, 'i_txt', 'i_div')
    		    ;get object
    		    Local $oSFI_Documents = _IEGetObjById($oFrame, $sSFI_Documents_ID)
    		    ;get all document links
    		    Local $oAs = _IETagNameGetCollection ($oSFI_Documents, 'a')   

    		           
    		    For $oA In $oAs        
    		        If ($oA.innertext <> '') Then
    		            ConsoleWrite($oA.href & @CRLF)
    		            HtmlDownload($oIE_RM, $oA.href, $sFolderPath, $oA.innertext)
    ;ExitLoop
    		        EndIf    
    		    Next
            EndIf
        Next
        _IEQuit($oIE_RM)
        _IEQuit($oIE)
        ExitLoop
    Wend
EndFunc




;====================================================================================================================
;                        FUNCTION DISCRIPTION: DOWNLOAD AN IMAGE A THE LINK
;                        INPUT                    : $sFilePath, $sLink
;                        OUTPUT                    : AN JPG IMAGE
;====================================================================================================================
Func HtmlDownload($oIE_RM, $sLink, $sFolderPath, $sFileName)
    GUICtrlSetData ($Commu_Ctrl, 'Downloading html documents:' & @CRLF & $sFileName)

    Local $sFileNameOrigin = $sFileName
	;Read Log file and check if the link has already downloaded
	Local $sLogData = LogFile ($sFolderPath, 'Read')
	If (StringInStr($sLogData, $sFileName)) = 0 Then

	    $sFileName = StringRegExpReplace($sFileName, '[\/:*?"<>|]', ' ')
	    Local $sFilePath = $sFolderPath & '\' & $sFileName & '.html'

	    If (stringlen($sFilePath) > 255) Then $sFilePath = StringLeft($sFilePath, 250) & '.html'


	    _IENavigate($oIE_RM, $sLink)
	    Sleep(500)
	    $shtml = _IEDocReadHTML ($oIE_RM)
	    $shtml = StringRegExpReplace($shtml, '<link href=.+\.css"', '<link href="repair_procedure.css"')
	    Create_HTML ($shtml, $sFilePath)

	    $sLink = StringRegExpReplace($sLink, '\.html?\?.+', '.html?')
	    LogFile ($sFolderPath, 'Write', $sFileNameOrigin & @CRLF)
	Else
		GUICtrlSetData ($Commu_Ctrl, 'The document below has been downloaded:' & @CRLF & $sFileName)
	EndIf
EndFunc




;====================================================================================================================
;                        FUNCTION DISCRIPTION:
;                    INPUT                    :
;                        OUTPUT                    :
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
;                        FUNCTION DISCRIPTION: CREATE HTML FILE FROM
;                    INPUT                    : $sFilePath, $sTxt_Title,$HTML_body
;                        OUTPUT                    : AN HTML FILE IN $sFilePath
;====================================================================================================================
Func Create_HTML ($html, $sFilePath)
    Local $hFileOpen = FileOpen ($sFilePath,$FO_OVERWRITE)
    FileWrite($hFileOpen, $html)
    FileClose($hFileOpen)
EndFunc



;====================================================================================================================
;                        FUNCTION DISCRIPTION: CREATE LOG FILE OF DTC OR PROCEDURE
;                    INPUT                    : $sFilePath, $sWhichLogFile,    $sTxt, $sMode
;                        OUTPUT                    : AN LOG FILE IN $sFilePath
;====================================================================================================================
Func LogFile ($sFilePath, $sMode, $sTxt = '')
    If $sMode = 'Read' Then
        Local $hFileOpen = FileOpen ($sFilePath & '\#LogFile.txt', $FO_READ )
        Local $sFileRead = FileRead($hFileOpen)
        FileClose($hFileOpen)
        Return $sFileRead
    Else
        Local $hFileOpen = FileOpen ($sFilePath & '\#LogFile.txt',$FO_APPEND)
        FileWrite($hFileOpen, $sTxt)
        FileClose($hFileOpen)
        Return 1
    EndIf
EndFunc




;====================================================================================================================
;                        FUNCTION DISCRIPTION: LOAD FILE CONTENT
;                        INPUT                :
;                     OUTPUT                :
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





