#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include <Excel.au3>
#include <MsgBoxConstants.au3>
#include <StringConstants.au3>
#include <GUIConstantsEx.au3>
#include <Timers.au3>



#Region ### START GUI section ### Form=
   Opt('GUIOnEventMode', 1)
   $Form1 = GUICreate('DTC Processor', 300, 100, -1, -1)
   GUISetBkColor(0xFFFFFF)
   ;-------------------------------------------
   ;Create Labels
   $Commu_Ctrl = GUICtrlCreateLabel('', 70, 40, 200, 50)
#EndRegion ### END Koda GUI section ###

Local $sFolderPath = InputBox('Folder Path', 'Please input folder path of the file!')
Local $sMake = InputBox('Make', 'Please input make name of the file!')

;-------------------------------------------

;~ Local $sFolderPath = @ScriptDir
;~ Local $sMake = 'Honda'

Local $sFileName = $sMake & '.xlsx'
Local $sWorkbook = $sFolderPath & '\' & $sFileName


Local $vWorksheetRead = $sMake
Local $vWorksheetWrite = $sMake & ' Processed'


;Create application object
Local $oExcel = _Excel_Open()
;Open an existing workbook and return its object identifier.
Local $oWorkbook = _Excel_BookAttach ($sWorkbook)
If @error Then Local $oWorkbook = _Excel_BookOpen($oExcel, $sWorkbook)
;Check if the sheet to write data has alread existed or not, if not => add sheet
_Excel_RangeRead ($oWorkbook, $vWorksheetWrite, 'A1')
If @error = 2 Then _Excel_SheetAdd ($oWorkbook, 1, False, 1, $vWorksheetWrite)

;SHOW GUI
GUISetState(@SW_SHOW)

;Read all data from excel
GUICtrlSetData ($Commu_Ctrl, 'Loading ' & $sMake & ' data ...')
Local $aDataRead = _Excel_RangeRead ($oWorkbook, $vWorksheetRead)
Local $iRowDataRead = 1


;Array to save data which will be write into excel
Local $aDataWrite[0][5]
Local $iRowDataWrite = 0

;Check dup in the array
Local $aCheckDup[0]

;Process data by data from excel
For $iRowDataRead = 1 To UBound ($aDataRead,  $UBOUND_ROWS) - 1
   GUICtrlSetData ($Commu_Ctrl, 'Processing row #' & $iRowDataRead)
   ;Read DTC cell of a line
   Local $sDTCs = $aDataRead[$iRowDataRead][4]

   ;Split DTC string into array
   Local $aDTCs = StringSplit ($sDTCs, ',', $STR_NOCOUNT )

   For $sDTC In $aDTCs

	  If $sDTC <> '' Then

		 Local $aCheckDupString[5] = [$aDataRead[$iRowDataRead][0], _
									  $aDataRead[$iRowDataRead][1], _
									  $aDataRead[$iRowDataRead][2], _
									  $aDataRead[$iRowDataRead][3], _
									  $sDTC]
		 If _ArrayCheckDup ($aDataWrite, $aCheckDupString) = False Then
			ReDim $aDataWrite[$iRowDataWrite + 1][5]
			$aDataWrite[$iRowDataWrite][0] = $aDataRead[$iRowDataRead][0]
			$aDataWrite[$iRowDataWrite][1] = $aDataRead[$iRowDataRead][1]
			$aDataWrite[$iRowDataWrite][2] = $aDataRead[$iRowDataRead][2]
			$aDataWrite[$iRowDataWrite][3] = $aDataRead[$iRowDataRead][3]
			$aDataWrite[$iRowDataWrite][4] = $sDTC

			$iRowDataWrite += 1
		 EndIf
	  EndIf
   Next
Next
GUICtrlSetData ($Commu_Ctrl, 'Writing data to excel ...')
Sleep (3000)
_Excel_RangeWrite ($oWorkbook, $vWorksheetWrite, $aDataWrite, 'A1')

MsgBox ($MB_TOPMOST, 'Message', 'Done, Please Check!')




Func _ArrayCheckDup ($aArray, $aSubArray)
   Local $bCheckFlag = False
   ;Check last line backward
   For $iRow = UBound($aArray, $UBOUND_ROWS) - 1 To 0 Step -1
	  ;If they have same year => Continure if not => Exit
	  If $aArray[$iRow][1] <> $aSubArray[1] Then ExitLoop
	  If $aArray[$iRow][2] <> $aSubArray[2] Then ExitLoop
	  If $aArray[$iRow][3] <> $aSubArray[3] Then ExitLoop
	  If $aArray[$iRow][4] = $aSubArray[4] Then
		 $bCheckFlag = True
		 ExitLoop
	  EndIf
   Next
   Return $bCheckFlag
EndFunc



