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

;HotKeySet('{ESC}', 'Autoit_Exit')
FileDelete (@ScriptDir & '\ErrorLog.txt')





;====================================================================================================================
;                  FUNCTION DISCRIPTION: LOAD FILE CONTENT
;                  INPUT               :
;                  OUTPUT              :
;====================================================================================================================

;Global $oErrorHandler  = ObjEvent("AutoIt.Error","ErrFunc")
;Global $bErrorOccurred = False
;Global $sErrorPosition = ''
;Global $bErrorHappened = False
;Func ErrFunc()
;  Local $sText = ("=====================================================================================" & @CRLF & _
;             "COM Error!"    & @CRLF  & @CRLF & _
;             "          err.description is : " & @TAB & $oErrorHandler.description  & @CRLF & _
;             "          err.windescription : " & @TAB & $oErrorHandler.windescription & @CRLF & _
;             "          err.number is      : " & @TAB & hex($oErrorHandler.number,8)  & @CRLF & _
;             "          err.lastdllerror is: " & @TAB & $oErrorHandler.lastdllerror   & @CRLF & _
;             "          err.scriptline is  : " & @TAB & $oErrorHandler.scriptline   & @CRLF & _
;             "          err.source is      : " & @TAB & $oErrorHandler.source       & @CRLF & _
;             "          err.helpfile is  : " & @TAB & $oErrorHandler.helpfile     & @CRLF & _
;             "          err.helpcontext is : " & @TAB & $oErrorHandler.helpcontext & @CRLF & _
;             "          link when error    : " & @TAB & $sErrorPosition)
;   $sText = StringRegExpReplace ($sText, '\r\n\r\n+', @CRLF)
;   $bErrorOccurred = True
;   $bErrorHappened = True
;   Write_Error (@CRLF & @CRLF & $sText)
;Endfunc



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





Global $oIE
Global $bButton_Flag = False

Func Set_Button_Flag ()
    $bButton_Flag = True
EndFunc

GUIInit()
;
While 1
    If $bButton_Flag = True Then GetCode()
WEnd

;====================================================================================================
;This function is Exit AutoIT
;====================================================================================================
Func Autoit_Exit ()
   _IEQuit($oIE)
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
    Global $Button_Info = GUICtrlCreateButton('Get Code', 230, 85, 100, 30)
    GUICtrlSetOnEvent($Button_Info, 'Set_Button_Flag')
    ;-------------------------------------------
    ;CREATE GUI NOTIFICATION PLACE
    Global $Commu_Ctrl = GUICtrlCreateLabel('', 50, 130, 300, 55)
    Global $CopyRight = GUICtrlCreateLabel('Created by Thien Nguyen', 150, 180, 150, 50)
    ;-------------------------------------------
    ;Create input
    Global $Character = GUICtrlCreateInput("", 50, 40, 80, 21)
    GUICtrlCreateLabel("First character:", 50, 20, 150, 17)


    Global $FromNumber = GUICtrlCreateInput("", 150, 40, 80, 21)
    GUICtrlCreateLabel("From numer:", 150, 20, 150, 17)

    Global $ToNumber = GUICtrlCreateInput("", 250, 40, 80, 21)
    GUICtrlCreateLabel("To number:", 250, 20, 150, 17)


    Global $FileName = GUICtrlCreateInput("", 50, 90, 120, 21)
    GUICtrlCreateLabel("File Name:", 50, 70, 150, 17)
    ;SHOW GUI
    GUISetState(@SW_SHOW)
    GUICtrlSetData ($Commu_Ctrl, 'Fill out the infor and press the button')
EndFunc





Func GetMultiCode ()
    CreateMultiIE()
    Sleep(1000)
    QuitMultiIE()
    ; return flag to stop the function
    $bButton_Flag = False
EndFunc


Func CreateMultiIE()
    ; open new ie with the link
    For $i = 0 To 1
        Local $oIE = _IECreate('https://app.radiocodeford.com/?',0,1)
        Assign ('oIE_' & $i, $oIE, $ASSIGN_FORCEGLOBAL )
    Next
EndFunc


Func QuitMultiIE()
    ; quit all ie
    For $i = 0 To 1
        Execute('_IEQuit($oIE_' & $i & ')')
    Next
EndFunc


;====================================================================================================
;create html file from input string
;filepath needs to include the html name with extension ".html"
;====================================================================================================
Func GetCode ()
   Local $sCharacter = GUICtrlRead ($Character)
   Local $iFromNumber = Number(GUICtrlRead ($FromNumber))
   Local $iToNumber = Number(GUICtrlRead ($ToNumber))
   Local $sFileName = GUICtrlRead ($FileName)

	While 1

        If ($sCharacter) <> 'M' And ($sCharacter) <> 'V' Then
            GUICtrlSetData ($Commu_Ctrl, 'Make sure the character must be "M" or "V"')
            ExitLoop
        EndIf

        If $iFromNumber > 999999 Then
            GUICtrlSetData ($Commu_Ctrl, 'Make sure the number is less then 999999')
            ExitLoop
        EndIf

        If $iToNumber > 999999 Then
            GUICtrlSetData ($Commu_Ctrl, 'Make sure the number is less then 999999')
            ExitLoop
        EndIf

        If $sFileName = '' Then
            GUICtrlSetData ($Commu_Ctrl, 'Please type a name for data file')
            ExitLoop
		 EndIf


        ; open new ie with the link
        $oIE = _IECreate('https://app.radiocodeford.com/?',0,0)


        ; this object is use for checking loading status
        Local $oCheckLoad = _IEGetObjById($oIE, 'unlock-form')


        For $i = $iFromNumber + 1000000 To $iToNumber + 1000000
            ; create serial string
            Local $sSerial = $sCharacter & StringRight($i, 6)
            GUICtrlSetData ($Commu_Ctrl, 'Getting code for Radio Serial: ' & $sSerial)
            Sleep(500)

            ; get iput object and fill the radio serial
            Local $oInput = _IEGetObjById($oIE, 'serial')
            _IEFormElementSetValue ($oInput, $sSerial)

            ; press search button
            Local $oButton = _IEGetObjById($oIE, 'unlock-btn')
            _IEAction($oButton, 'click')


            ; wait for the page fisnihing loading
            While ($oCheckLoad.GetAttribute('class') = 'unlock unlock--loading')
                Sleep(100)
            WEnd

            ; get code object then get innertext inside
            Local $oCode = _IEGetObjById($oIE, 'code')
            Local $sCode = _IEPropertyGet($oCode,'innertext')

            ; press search button
            Local $oButton = _IEGetObjById($oIE, 'unlock-btn')
            _IEAction($oButton, 'click')


            ; write to a data file
            Local $sFilePath = @ScriptDir & '\' & $sFileName & '.txt'
            WriteFile ($sFilePath, $sSerial & @TAB & $sCode & @CRLF)
        Next


		 GUICtrlSetData ($Commu_Ctrl, 'Done, Please check!')
        ; return flag to stop the function
        $bButton_Flag = False
	    _IEQuit($oIE)
        ExitLoop
    Wend

    ; return flag to stop the function

    $bButton_Flag = False
EndFunc







;====================================================================================================================
;                        FUNCTION DISCRIPTION: CREATE LOG FILE OF DTC OR PROCEDURE
;                    INPUT                    : $sFilePath, $sWhichLogFile,    $sTxt, $sMode
;                        OUTPUT                    : AN LOG FILE IN $sFilePath
;====================================================================================================================
Func WriteFile ($sFilePath, $sTxt = '')
    Local $hFileOpen = FileOpen ($sFilePath,$FO_APPEND)
    FileWrite($hFileOpen, $sTxt)
    FileClose($hFileOpen)
    Return 1
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





