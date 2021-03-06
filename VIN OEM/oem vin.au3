#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         Thien Nguyen

 Script Function:
   Copy data from bonbanh.com

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include <MsgBoxConstants.au3>
#include <Clipboard.au3>
#include < IE.au3 >
#include <Excel.au3>



; Set Hotkey for the program
; HotKeySet("{ESC}", "Autoit_Exit")
;HotKeySet("^z", "Excel_Open")



Demo()



Func Autoit_Exit ()
   Exit
EndFunc

Func Demo ()
    ; Open browser with basic example, get link collection,
    ; loop through items and display the associated link URL references

    #include <IE.au3>
    #include <MsgBoxConstants.au3>

    $sAllVin = Read_File(@ScriptDir & "/input.txt")
    $sData = Read_File(@ScriptDir & "/Data.txt")

    $aVin = StringSplit ($sAllVin, '|', $STR_NOCOUNT)


    Local $oIE = _IECreate("")
    Local $i
    $i = 1
    For $vin In $aVin
        ConsoleWrite($i & @CRLF)
        If (StringInStr($sData, $vin)) Then
            ConsoleWrite($vin & ' --> exists' & @CRLF)
            $i = $i +1
            ContinueLoop
        EndIf

        ConsoleWrite($vin & @CRLF)

        _IENavigate($oIE, "http://www.etis.ford.com/vehicleSelection.do")

        $oForm = _IEGetObjByName($oIE, "VehicleSelectionForm")

        $oVin = _IEFormElementGetObjByName($oForm, 'vin')
        _IEFormElementSetValue($oVin, $vin)

        _IEFormSubmit ($oForm)



        while 1
            $oInfo = _IEGetObjById($oIE, "vehicleDetails")
            If IsObj($oInfo) Then

                $txt = $oInfo.innerhtml
                If (StringInStr($txt, '<span')) Then 
                    
                    $txt = GetItemStringByMark ($txt, '<span>', '</span>')
                    Write_File ($vin & @TAB & $txt & @CRLF)
                    ExitLoop
                    

                EndIf

            Else
                If StringInStr(_IEPropertyGet ($oIE, "innerhtml"), "no information could be found for the VIN") Then
                    Write_File ($vin & @TAB & "No Info" & @CRLF)
                    ExitLoop
                EndIf

            EndIf
        Wend



        Sleep(1000)
        $i = $i +1
    Next
EndFunc




;====================================================================================================
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




Func Write_File($sText)
    Local $hFileOpen = FileOpen(@ScriptDir & "/Data.txt", $FO_APPEND)
    FileWrite($hFileOpen, $sText)
    FileClose($hFileOpen)
EndFunc





Func Read_File($fpath)
    Local $hFileOpen = FileOpen($fpath, $FO_READ)
    $sFileRead = FileRead($hFileOpen)
    FileClose($hFileOpen)

    Return $sFileRead
EndFunc