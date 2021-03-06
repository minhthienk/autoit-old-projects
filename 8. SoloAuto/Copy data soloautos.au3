#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         Thien Nguyen

 Script Function:
   Copy data from bonbanh.com

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include <MsgBoxConstants.au3>
#include <Clipboard.au3>
#include <IE.au3 >
#include <Excel.au3>



; Set Hotkey for the program
;~ HotKeySet("{ESC}", "Autoit_Exit")
;~ HotKeySet("^z", "Main")

Main ()

while (1)
WEnd


Func Main()

   Close_All_IE();
   MsgBox (0, "NOTE", "Click OK" & @CRLF & @CRLF & "PRESS the ESC button to exit the application if needed")
   $sLink = "https://soloautos.mx/Autos?q=%28and.tipoveh%C3%ADculo.autos%2c%20camionetas%20y%204x4._.servicio.soloautos_.mx._.%28or.marca.lincoln._.marca.mercury._.marca.honda._.marca.mazda._.marca.mercedes-benz._.marca.jaguar._.marca.land%20rover.%29_.ano.range%281996..2019%29.%29&sort=~year&s="
   $sMake = "ChinaGroup"
   $sFileName = "soloautosmx_" & $sMake

   ;Create a text file to save the content from the sites
   Local $sFilePath = @ScriptDir
   WriteTxtFile ($sFileName, "This is the collection of " & $sMake & " vehicles on soloautosmx website", "overwrite", $sFilePath)
   WriteTxtFile ($sFileName, @CRLF & "Vehicle" & @TAB & "Cylinder" & @TAB & "Capacity" & @TAB & "Transmission" & @TAB & "Fuel Type" & @TAB & "Door", "append", $sFilePath)

   ;Open an IE object
   Local $oIE = IECreateCheckError ("about:blank")

   ;Collect all vehicle links from all pages
   Local $aAllPagesLinks [0]
   For $j = 0 To 305
	  IENavigateCheckError ($oIE, $sLink & $j*15)
	  Local $aPageLinks = GetAllVehicleLinks ($oIE)
	  _ArrayConcatenate ($aAllPagesLinks, $aPageLinks, 1)
	  $aAllPagesLinks = _ArrayUnique ($aAllPagesLinks, 0, 0, 0, $ARRAYUNIQUE_NOCOUNT)
   Next

   ;Write links to a file
   WriteTxtFile ($sFileName & "_Links", _ArrayToString ($aAllPagesLinks, @CRLF), "overwrite", $sFilePath)

;~    Local $aAllPagesLinks = LoadFile ("soloautosmx_ChinaGroup_Links")
   ;Open each link and get data
   For $i = 0 To UBound($aAllPagesLinks) - 1
	  Do
		 IENavigateCheckError ($oIE, $aAllPagesLinks[$i])
		 If StringInStr (_IEPropertyGet ($oIE, "innertext"), "502 Bad Gateway") <> 0 Or StringInStr (_IEPropertyGet ($oIE, "innertext"), "You don't have permission to access") <> 0 Then Sleep (5000)
	  Until StringInStr (_IEPropertyGet ($oIE, "innertext"), "502 Bad Gateway") = 0 And StringInStr (_IEPropertyGet ($oIE, "innertext"), "You don't have permission to access") = 0


	  If StringInStr (_IEPropertyGet ($oIE, "innertext"), "¡Un error ha ocurrido mientras procesabamos tu petición!") <> 0 _
		 Or StringInStr (_IEPropertyGet ($oIE, "innertext"), "The website cannot display the page") <> 0 Then
		 WriteTxtFile ($sFileName, @CRLF & "N/A", "append", $sFilePath)
		 ContinueLoop
	  EndIf

	  Local $oContent = _IEGetObjById($oIE, "basic")
	  Local $sContent = $oContent.innertext



	  ;Form a line of data
	  $sTxt = GetItemStringByMarkTxt ($sContent, "Vehículo") & @TAB & _
			  GetItemStringByMarkTxt ($sContent, "Cilindros") & @TAB & _
			  GetItemStringByMarkTxt ($sContent, "Litros (motor)") & @TAB & _
			  GetItemStringByMarkTxt ($sContent, "Transmisión") & @TAB & _
			  GetItemStringByMarkTxt ($sContent, "Combustible") & @TAB & _
			  GetItemStringByMarkTxt ($sContent, "Puertas")

	  ;Write data into a file
	  WriteTxtFile ($sFileName, @CRLF & $sTxt, "append", $sFilePath)
   Next

   MsgBox ($MB_TOPMOST, "", "done")
   Exit
EndFunc


;====================================================================================================================
;                  FUNCTION DISCRIPTION:
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func GetItemStringByMarkTxt ($sContent, $sMark)
   If StringInStr ($sContent, $sMark, 0, 1, 1) <> 0 Then
	  Local $iStart = StringInStr ($sContent, $sMark, 0, 1, 1) + StringLen ($sMark)
	  Local $iEnd = StringInStr ($sContent, @CRLF, 0, 1, $iStart)
	  Local $sItemString = StringMid ($sContent, $iStart, $iEnd - $iStart)
   Else
	  Local $sItemString = ""
   EndIf
   Return $sItemString
EndFunc








;====================================================================================================================
;                  FUNCTION DISCRIPTION: GET ALL LINKS
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func GetAllVehicleLinks ($oIE)
   Local $oLinks = _IELinkGetCollection($oIE)
   Local $aLinks [0]
   Local $iNumLinks = @extended

   Local $sTxt = $iNumLinks & " links found" & @CRLF & @CRLF
   For $oLink In $oLinks
	  $sLink = $oLink.href
	  If StringInStr($sLink, "/SA-AD-") <> 0 Then
		 ReDim $aLinks [UBound($aLinks) + 1]
		 $aLinks [UBound($aLinks) - 1] = $sLink
	  EndIf
   Next
   $aLinks = _ArrayUnique ($aLinks)
   Return $aLinks
EndFunc

;====================================================================================================================
;                  FUNCTION DISCRIPTION: CREATE IE OBJECT AND CHECK ERROR
;                  RETURN			   :
;====================================================================================================================
Func IECreateCheckError ($sLink, $bAttach = 0, $bVisible = 1, $bWait = 1, $bTakeFocus = 0)
   Do
	  Local $oIE = _IECreate($sLink, $bAttach, $bVisible, $bWait, $bTakeFocus)
	  If @error <> 0 Then Sleep(1000)
   Until @error = 0
   Sleep (2000)
   Return $oIE
EndFunc

;====================================================================================================================
;                  FUNCTION DISCRIPTION: NAVIGATE IE OBJECT AND CHECK ERROR
;                  RETURN			   : A STRING OF PROCEDURE NAME
;====================================================================================================================
Func IENavigateCheckError (ByRef $oIE, $sLink)
;~    MsgBox (0, "", "New Site")

   Local $icount = 0
   Local $bFlag = True
	  Do
   ;~ $icount = $icount + 1
		 __IELockSetForegroundWindow($LSFW_LOCK)
		 _IENavigate ($oIE, $sLink, 1)
		 Sleep (500)

   If @error = 0 Then
	  $bFlag = True
   Else
	  $bFlag = False
   EndIf
   ;ConsoleWrite ("Navigate count: " & $icount & " --- Error code: " & @error & @CRLF)
   Until $bFlag = True
EndFunc


;====================================================================================================================
;                  FUNCTION DISCRIPTION: EXIT AUTOIT SCRIPT
;                  INPUT			   :
; 				   OUTPUT			   :
;====================================================================================================================
Func Autoit_Exit ()
   Exit
EndFunc


;====================================================================================================================
;                  FUNCTION DISCRIPTION: CLOSE ALL IE OBJECT
;                  INPUT			   :
; 				   OUTPUT			   :
;====================================================================================================================
Func Close_All_IE()
   $Proc = "iexplore.exe"
   While ProcessExists($Proc)
      ProcessClose($Proc)
   Wend
EndFunc




;====================================================================================================================
;                  FUNCTION DISCRIPTION: CLOSE ALL IE OBJECT
;                  INPUT			   :
; 				   OUTPUT			   :
;====================================================================================================================
Func WriteTxtFile ($sFileName, $sTxt, $sMode = "append", $sFilePath = @ScriptDir)
   If $sMode = "overwrite" Then
	  Local $hFileOpen = FileOpen ($sFilePath & "\" & $sFileName & ".txt",$FO_OVERWRITE)
   Else
	  Local $hFileOpen = FileOpen ($sFilePath & "\" & $sFileName & ".txt",$FO_APPEND)
   EndIf
   FileWrite($hFileOpen, $sTxt)
   FileClose($hFileOpen)
EndFunc




;====================================================================================================================
;                  FUNCTION DISCRIPTION: LOAD FILE CONTENT
;                  INPUT			   :
; 				   OUTPUT			   :
;====================================================================================================================
Func LoadFile ($sFileName, $sFilePath = @ScriptDir)
   ;Open YMME config file and get data
   Local $hFileOpen = FileOpen ($sFilePath & "\" & $sFileName & ".txt",$FO_READ )
   Local $sFileRead = FileRead($hFileOpen)
   FileClose($hFileOpen)
   ;String => Array
   Local $alConfigData = StringSplit ($sFileRead, @CRLF, $STR_ENTIRESPLIT + $STR_NOCOUNT)
   Return $alConfigData
EndFunc




