
;====================================================================================================================
;                  FUNCTION DISCRIPTION:
;                  INPUT			   :
; 				   OUTPUT			   :
;====================================================================================================================

Global $oErrorHandler  = ObjEvent("AutoIt.Error","ErrFunc")
Global $bErrorOccurred = False
Global $sErrorPosition = ''
Func ErrFunc()
  Local $sText = ("=====================================================================================" & @CRLF & _
			 "COM Error!"    & @CRLF  & @CRLF & _
             "          err.description is : " & @TAB & $oErrorHandler.description  & @CRLF & _
             "          err.windescription : " & @TAB & $oErrorHandler.windescription & @CRLF & _
             "          err.number is      : " & @TAB & hex($oErrorHandler.number,8)  & @CRLF & _
             "          err.lastdllerror is: " & @TAB & $oErrorHandler.lastdllerror   & @CRLF & _
             "          err.scriptline is  : " & @TAB & $oErrorHandler.scriptline   & @CRLF & _
             "          err.source is      : " & @TAB & $oErrorHandler.source       & @CRLF & _
             "          err.helpfile is	 : " & @TAB & $oErrorHandler.helpfile     & @CRLF & _
             "          err.helpcontext is : " & @TAB & $oErrorHandler.helpcontext & @CRLF & _
			 "          link when error    : " & @TAB & $sErrorPosition)
   $sText = StringRegExpReplace ($sText, '\r\n\r\n+', @CRLF)
   $bErrorOccurred = True
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

