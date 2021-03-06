#include-once

#include <MsgBoxConstants.au3>
#include <FileConstants.au3>
#include <InetConstants.au3>
#include <Clipboard.au3>
#include <IE.au3>
#include <Excel.au3>
#include <WinAPIFiles.au3>
#include <Array.au3>

#include "General_Library.au3"
#include "Create_JAVASCRIPT_Procedure.au3"
#include "Create_NORMAL_Procedure.au3"




;====================================================================================================================
;                  FUNCTION DISCRIPTION: ADD PROCEDURE
;====================================================================================================================
Func Add_Procedure ()
   Local $sDTC_Path = GUICtrlRead ($Input_DTC_Path)
   Local $hFileOpen = FileOpen($sDTC_Path, $FO_READ)
   Local $sDTC_HTML = FileRead ($hFileOpen)
   FileClose($hFileOpen)
   Local $aProcedure [0]
   Local $aTemp = StringSplit ($sDTC_HTML, @CRLF, $STR_ENTIRESPLIT)
   For $i = 1 to $aTemp[0]
	  If StringInStr ($aTemp[$i], "PROCEDURE_OTHER_") <> 0 Then
		 ReDim $aProcedure [UBound ($aProcedure) + 1]
		 $aProcedure [UBound ($aProcedure) - 1] = $aTemp[$i]
	  EndIf
   Next

   MsgBox (0, "", _ArrayToString($aProcedure, @CRLF))
   $bFind_Flag = False
EndFunc



