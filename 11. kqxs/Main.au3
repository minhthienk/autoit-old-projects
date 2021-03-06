#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         Thien Nguyen

 Script Function:
   Copy data from bonbanh.com

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include-once

#include <MsgBoxConstants.au3>
#include <Clipboard.au3>
#include <IE.au3 >
#include <Excel.au3>

#include "COMErrorHandler.au3"



Global $sFilePath = @ScriptDir
Global $sFileName = 'Value.xlsx'
Global $vWorksheet = 'Result Refined'


; Set Hotkey for the program
;~ HotKeySet("{ESC}", "Autoit_Exit")

Analyze ()

;====================================================================================================================
;                  FUNCTION DISCRIPTION:
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func Analyze ()
   ;First Massage ----------------------------------------------------------
   MsgBox (0, 'NOTE', 'Click OK' & @CRLF & @CRLF & 'PRESS the ESC button to exit the application if needed')
   ;Number to analyze
   Local $iNumber = 0
   ;Load data
   Local $aValue = LoadData()
   ;
   ;Variables Declaration for Analyze
   Local $iDayCount = 0
   Local $bFlag = False
   Local $aDayInterval[2000][100]
   Local $iRowCount = 0


;~    ;Create file to save result ---------------------------------------------
;~    ;Create application object
;~    Local $oExcel = _Excel_Open()
;~    ;Open an existing workbook and return its object identifier.
;~    Local $sWorkbook = $sFilePath & '\' & $sFileName
;~    Local $oWorkbook = _Excel_BookOpen($oExcel, $sWorkbook)



   For $iNumber = 0 To 0
	  $iDayCount = 0
	  For $i = UBound($aValue, $UBOUND_ROWS) - 1 To 1 Step -1
		 ;Flag to mark the existence of an number
		 $bFlag = False
		 $iDayCount += 1
		 For $j = 2 To UBound($aValue, $UBOUND_COLUMNS ) - 1
			If $iNumber = $aValue[$i][$j] Then
			   $bFlag = True
			   ExitLoop
			EndIf
		 Next
		 ;Write day interval into an array
		 If $bFlag = True Then
			$aDayInterval[$iRowCount][$iNumber] = $iDayCount
			$iRowCount += 1
;~ 			_Excel_RangeWrite ($oWorkbook, 'Intervals', $iDayCount, Chr(Asc('B')+$iNumber)  & ($i + 1))
			$iDayCount = 0
		 EndIf
	  Next
   Next

   _ArrayDisplay($aDayInterval)



   MsgBox ($MB_TOPMOST, "", "done")
EndFunc






;====================================================================================================================
;                  FUNCTION DISCRIPTION:
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func LoadData()
   ;Create application object
   Local $oExcel = _Excel_Open()
   ;Open an existing workbook and return its object identifier.
   Local $sWorkbook = $sFilePath & '\' & $sFileName

   Local $oWorkbook = _Excel_BookAttach ($sWorkbook)
   If @error <> 0 Then Local $oWorkbook = _Excel_BookOpen($oExcel, $sWorkbook)


   ;Read values from excel file
   Local $aValue = _Excel_RangeRead ($oWorkbook, $vWorksheet)

;~    _ArrayDisplay($aValue)
   Return $aValue
EndFunc








;====================================================================================================================
;                  FUNCTION DISCRIPTION:
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func GetData()
   ;First Massage ----------------------------------------------------------
   MsgBox (0, 'NOTE', 'Click OK' & @CRLF & @CRLF & 'PRESS the ESC button to exit the application if needed')

   ;Link formation ---------------------------------------------------------
   Local $sMainLink = 'https://www.xoso.net/tra-cuu-ket-qua-xo-so.html?mien=1&thu=0&'
   Local $sDate = 'ngay=1&thang=1&nam=2010'

   ;Create file to save result ---------------------------------------------
   ;Create application object
   Local $oExcel = _Excel_Open()
   ;Open an existing workbook and return its object identifier.
   Local $sWorkbook = $sFilePath & '\' & $sFileName
   Local $oWorkbook = _Excel_BookOpen($oExcel, $sWorkbook)

   ;Demo link

   Local $iRow = 800
   For $iYearCount = 2017 To 2008 Step -1
	  For $iMonthCount = 12 To 01 Step -1
		 For $iDateCount = 31 To 1 Step -1
			Local $sLink = $sMainLink & 'ngay=' & $iDateCount & '&thang=' & $iMonthCount & '&nam=' & $iYearCount
			Local $iStep = 0
			Local $aValue = GetInfoContent ($sLink, $iStep)
			If $aValue[0][0] <> 'No information' Then
			   _Excel_RangeWrite ($oWorkbook, $vWorksheet, $aValue, 'A' & $iRow)
			   $iRow += $iStep
			EndIf
		 Next
	  Next
   Next



   MsgBox ($MB_TOPMOST, "", "done")
EndFunc






;====================================================================================================================
;                  FUNCTION DISCRIPTION:
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func GetInfoContent ($sLink, Byref $iStep)
   Local $html = GetHtmlSourceUsingHttpRequest ($sLink)


   ;Get the result of the lottery
   Local $sTemp = ''
   ;Count number of province
   StringReplace ($html, '<td class="tinh">','')
   Local $iProvinceMax = @extended
   Local $aValue [$iProvinceMax][20]

   $iStep = $iProvinceMax

   If $iProvinceMax <> 0 Then
	  For $iProvinceCount = 1 To $iProvinceMax
		 ;Date Info -----------------------------------------------------------------------
		 $sTemp = GetItemStringByMark ($html, '<td class="ngay">' & @CRLF, '</a>')
		 $aValue[$iProvinceCount - 1][0] = GetItemStringByMark ($sTemp, '">', '</a>')

		 ;Province Info -------------------------------------------------------------------
		 $sTemp = GetItemStringByMark ($html, '<td class="tinh">' & @CRLF, '</a>', $iProvinceCount)
		 $aValue[$iProvinceCount - 1][1] = GetItemStringByMark ($sTemp, '">', '</a>')

		 ;Giai 8 --------------------------------------------------------------------------
		 $sTemp = GetItemStringByMark ($html, '<td class="giai8">' & @CRLF, '</td>', $iProvinceCount)
		 $aValue[$iProvinceCount - 1][2] = StringRight(GetItemStringByMark ($sTemp, '<div>', '</div>'),2)

		 ;Giai 7 --------------------------------------------------------------------------
		 $sTemp = GetItemStringByMark ($html, '<td class="giai7">' & @CRLF, '</td>', $iProvinceCount)
		 $aValue[$iProvinceCount - 1][3] = StringRight(GetItemStringByMark ($sTemp, '<div>', '</div>', $iProvinceCount),2)

		 ;Giai 6 --------------------------------------------------------------------------
		 $sTemp = GetItemStringByMark ($html, '<td class="giai6">' & @CRLF, '</td>', $iProvinceCount)
		 ;Number 1
		 $aValue[$iProvinceCount - 1][4] = StringRight(GetItemStringByMark ($sTemp, '<div>', '</div>', 1),2)
		 ;Number 2
		 $aValue[$iProvinceCount - 1][5] = StringRight(GetItemStringByMark ($sTemp, '<div>', '</div>', 2),2)
		 ;Number 3
		 $aValue[$iProvinceCount - 1][6] = StringRight(GetItemStringByMark ($sTemp, '<div>', '</div>', 3),2)

		 ;Giai 5 --------------------------------------------------------------------------
		 $sTemp = GetItemStringByMark ($html, '<td class="giai5">' & @CRLF, '</td>', $iProvinceCount)
		 $aValue[$iProvinceCount - 1][7] = StringRight(GetItemStringByMark ($sTemp, '<div>', '</div>'),2)

		 ;Giai 4 --------------------------------------------------------------------------
		 $sTemp = GetItemStringByMark ($html, '<td class="giai4">' & @CRLF, '</td>', $iProvinceCount)
		 ;Number 1
		 $aValue[$iProvinceCount - 1][8] = StringRight(GetItemStringByMark ($sTemp, '<div>', '</div>', 1),2)
		 ;Number 2
		 $aValue[$iProvinceCount - 1][9] = StringRight(GetItemStringByMark ($sTemp, '<div>', '</div>', 2),2)
		 ;Number 3
		 $aValue[$iProvinceCount - 1][10] = StringRight(GetItemStringByMark ($sTemp, '<div>', '</div>', 3),2)
		 ;Number 4
		 $aValue[$iProvinceCount - 1][11] = StringRight(GetItemStringByMark ($sTemp, '<div>', '</div>', 4),2)
		 ;Number 5
		 $aValue[$iProvinceCount - 1][12] = StringRight(GetItemStringByMark ($sTemp, '<div>', '</div>', 5),2)
		 ;Number 6
		 $aValue[$iProvinceCount - 1][13] = StringRight(GetItemStringByMark ($sTemp, '<div>', '</div>', 6),2)
		 ;Number 7
		 $aValue[$iProvinceCount - 1][14] = StringRight(GetItemStringByMark ($sTemp, '<div>', '</div>', 7),2)

		 ;Giai 3 --------------------------------------------------------------------------
		 $sTemp = GetItemStringByMark ($html, '<td class="giai3">' & @CRLF, '</td>', $iProvinceCount)
		 ;Number 1
		 $aValue[$iProvinceCount - 1][15] = StringRight(GetItemStringByMark ($sTemp, '<div>', '</div>', 1),2)
		 ;Number 2
		 $aValue[$iProvinceCount - 1][16] = StringRight(GetItemStringByMark ($sTemp, '<div>', '</div>', 2),2)

		 ;Giai 2 --------------------------------------------------------------------------
		 $sTemp = GetItemStringByMark ($html, '<td class="giai2">' & @CRLF, '</td>', $iProvinceCount)
		 $aValue[$iProvinceCount - 1][17] = StringRight(GetItemStringByMark ($sTemp, '<div>', '</div>'),2)

		 ;Giai 1 --------------------------------------------------------------------------
		 $sTemp = GetItemStringByMark ($html, '<td class="giai1">' & @CRLF, '</td>', $iProvinceCount)
		 $aValue[$iProvinceCount - 1][18] = StringRight(GetItemStringByMark ($sTemp, '<div>', '</div>'),2)

		 ;Giai db --------------------------------------------------------------------------
		 $sTemp = GetItemStringByMark ($html, @TAB & '<td class="giaidb">' & @CRLF, '</td>', $iProvinceCount)
		 $aValue[$iProvinceCount - 1][19] = StringRight(GetItemStringByMark ($sTemp, '<div>', '</div>'),2)
	  Next
   Else
	  Local $aValue[1][20]
	  $aValue[0][0] = 'No information'
   EndIf

   ;Return the information
   Return $aValue
EndFunc

















;====================================================================================================================
;                  FUNCTION DISCRIPTION:
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func GetHtmlSourceUsingHttpRequest ($sLink)
   Write_Log (@CRLF & $sLink)
   Sleep (300)
   ;the var $sErrorPosition is used to save the link when the error occurs
   $sErrorPosition = $sLink
   ;Get html source
   Local $html = ''
	  ;Get http object
	  ConsoleWrite ('Begin to get source: ' & $sLink & @CRLF)
	  While 1
		 ConsoleWrite ('   => Create COM')
		 Local $oHTTP = ObjCreate("winhttp.winhttprequest.5.1")
		 ;Check error
		 If @error <> 0 Then
			ConsoleWrite (' => Error: ' & @error & '/ Relink' & @CRLF)
			Sleep (500)
		 Else
			ExitLoop
		 EndIf
	  WEnd

	  While 1
		 ;Get source
		 ConsoleWrite (' => Open source')
		 $oHTTP.Open("GET", $sLink)
		 ConsoleWrite (' => Send Request')
		 $oHTTP.Send()
		 ;Check status to know if the network issue

		 If StringLen($oHTTP.Status) = 0 Then
			ConsoleWrite (' => Object failed/ Relink')
			Sleep (3000)
		 Else
			If $oHTTP.Status = 403 Then
			   ConsoleWrite ('   => Access denied')
			   Local $sText = ('=====================================================================================' & @CRLF & _
						 'Error 403: Access denied'   & @CRLF & _
						 '          The owner of this website has banned your IP address!'   & @CRLF & _
						 '          Current Link: ' & $sLink)
			   Write_Log (@CRLF & @CRLF & $sText)
			   Sleep (30000)

			ElseIf $oHTTP.Status = 200 Then
			   ExitLoop
			Else ;Other status
			   Local $sText = ('=====================================================================================' & @CRLF & _
						 'Error staus: ' &  $oHTTP.Status   & @CRLF  & @CRLF & _
						 '          An unexpected error has occured'   & @CRLF & _
						 '          Current Link: ' & $sLink)
			   Write_Log (@CRLF & @CRLF & $sText)
			   ExitLoop
			EndIf
		 EndIf
	  WEnd

	  ConsoleWrite (' => Status ' & $oHTTP.Status & ' => Get response')
	  $html = $oHTTP.Responsetext
	  ;Release object
	  ConsoleWrite (' => Replease COM')
	  $oHTTP = 0

;~    ;Check content
;~    If StringInStr ($html, 'The page you were looking for was not found') <> 0 Then
;~ 	  $html = ''
;~    EndIf

   ;Refined html source
   $html = StringReplace ($html, '><', '>' & @CRLF & '<')
   $html = StringReplace ($html, '> <', '>' & @CRLF & '<')
   ConsoleWrite (' => Completed' & @CRLF)
   Return $html
EndFunc








;====================================================================================================================
;                  FUNCTION DISCRIPTION:
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func GetItemStringByMark ($sString, $sStartMark, $sEndMark, $iOccurrenceStart = 1, $iOccurrenceEnd = 1)
   If StringInStr ($sString, $sStartMark, 0, 1, 1) <> 0 Then
	  Local $iStart = StringInStr ($sString, $sStartMark, 0, $iOccurrenceStart, 1) + StringLen ($sStartMark)
	  Local $iEnd = StringInStr ($sString, $sEndMark, 0, $iOccurrenceEnd, $iStart)
	  Local $sItemString = StringMid ($sString, $iStart, $iEnd - $iStart)
   Else
	  Local $sItemString = ""
   EndIf
   Return $sItemString
EndFunc



;====================================================================================================================
;                  FUNCTION DISCRIPTION: CLOSE ALL IE OBJECT
;                  INPUT			   :
; 				   OUTPUT			   :
;====================================================================================================================
Func WriteTxtFile ($sFileName, $sTxt, $sMode = "append")
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




Func Autoit_Exit ()
   Exit
EndFunc