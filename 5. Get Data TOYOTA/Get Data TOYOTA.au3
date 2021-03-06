#cs ----------------------------------------------------------------------------
#ce ----------------------------------------------------------------------------
#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <MsgBoxConstants.au3>
#include <StringConstants.au3>
#include <WinAPIFiles.au3>
#include <Clipboard.au3>
#include <IE.au3>

Global $sImagePath = @DesktopDir & "\TOYOTA\Land Cruiser TMV Apr2008_24\repair\png"
Main ()


Func Main ()
   ;Open the file for reading and store the handle to a variable.
   Local $sFilePath = @DesktopDir & "\TOYOTA\rm000002ppm007x.xml"
   ;Get XML source
   Local $sXMLData = OpenXML_GetData ($sFilePath)
   While StringInStr ($sXMLData, @CRLF & @CRLF) <> 0
	  $sXMLData = StringReplace ($sXMLData, @CRLF & @CRLF, @CRLF)
   WEnd
   ;Split line by line XML source and put into an array
   $aXMLData = StringSplit ($sXMLData, @CRLF, $STR_ENTIRESPLIT)

;~    MsgBox (0, "", GetTagName ($aXMLData[26]) & "|")
;~    Exit


   ;---------------------------------------------------------------
   ;Get the number of DTC to help determine the length of the header
   StringReplace ($sXMLData, "<dtccode>", "")
   Local $iNumOfDTC = @extended
   ;---------------------------------------------------------------
   ;Declare vars used in the Switch
   Global $sHTMLBody = "<br>"

   Local $iDTC_Count = 0
   Local $iHeaderEndPos = 9999
   Local $bTableFlag = False
   Local $iItem_Count_1 = 0
   Local $iItem_Count_2 = 0
   Local $aID [999][2]
   ;---------------------------------------------------------------
   For $i = 1 To $aXMLData [0]
	  Local $sString = $aXMLData [$i]
	  If GetTagName ($sString) = "<testgrp" Then
		 $iItem_Count_1 += 1
		 $aID [$iItem_Count_1][0] = GetID ($sString)
		 $aID [$iItem_Count_1][1] = "Go to step " & $iItem_Count_1
	  EndIf
   Next
   Local $iItem_Count_1 = 0



   For $i = 1 To $aXMLData [0]
	  Local $sString = $aXMLData [$i]
	  ;Devide xml source into lines and handle them by cases
	  Switch GetTagName ($sString)
		 ;------------------------------------
		 ;The line contains DTC Code
		 Case "<dtccode>"
			Local $sDTC = GetInnertext ($sString)
			If $iDTC_Count = 0 Then
			   $sHTMLBody &= @CRLF & "<h1>" & @CRLF & $sDTC & ": "
			Else
			   $sHTMLBody &= "<br>" & @CRLF & $sDTC & ": "
			EndIf
		 ;------------------------------------
		 ;The line contains DTC definition
		 Case "<dtcname>"
			Local $sDTC_Def = GetInnertext ($sString)
			$sHTMLBody &= $sDTC_Def
			$iDTC_Count += 1
			If $iDTC_Count = $iNumOfDTC Then $sHTMLBody &= @CRLF & "</h1>"
			Local $iHeaderEndPos = $i
		 ;------------------------------------
		 ;The line contains group name
		 Case "<name>"
			If $i > $iHeaderEndPos Then
			   Local $sGroup = GetInnertext ($sString)
			   $sHTMLBody &= "<br>" & @CRLF & "<h2><center>" & $sGroup & "</center></h2>"
			EndIf
		 ;------------------------------------
		 ;The line contains normal text
		 Case "<ptxt>"
			If $bTableFlag = False Then
			   Local $sNormalText = GetInnertext ($sString)
			   $sHTMLBody &= "<br>" & @CRLF & "<br>" & @CRLF &  $sNormalText
			Else
			   Local $sNormalText = GetInnertext ($sString)
			   $sHTMLBody &= "<br>" & @CRLF &  $sNormalText
			EndIf

		 ;------------------------------------
		 ;The line determines table title
		 Case "<title>"
			Local $sTableTitle = GetInnertext ($sString)
			$sHTMLBody &= "waiting for replace" & $sTableTitle
			$sHTMLBody = StringReplace($sHTMLBody, "<table>waiting for replace" & $sTableTitle, _
												   "<br>" & @CRLF & "<b>" & $sTableTitle & "</b>" & "<br>" & @CRLF & "<table>")

			$sHTMLBody = StringReplace($sHTMLBody, "waiting for replace" & $sTableTitle, _
												   "<br>" & @CRLF & "<br>" & @CRLF & "<b>" & $sTableTitle & "</b>" & @CRLF)

		 ;------------------------------------
		 ;The line determines the beginning of a table
		 Case "<table"
			$bTableFlag = True
			$sHTMLBody &= "<br>" & @CRLF & "<table>"


		 ;------------------------------------
		 ;The line determines the beginning of a table
		 Case "<table>"
			$bTableFlag = True
			$sHTMLBody &= "<br>" & @CRLF & "<table>"


		 ;------------------------------------
		 ;The line determines the beginning of a row
		 Case "<row>"
			$bTableFlag = True
			$sHTMLBody &= @CRLF & "<tr>"
		 ;------------------------------------
		 ;The line determines the beginning of text in table cell
		 Case "<entry>"
			$bTableFlag = True
			$sHTMLBody &= @CRLF & "<td>"
		 ;------------------------------------
		 ;The line determines the beginning of text in table cell
		 Case "<entry"
			$bTableFlag = True
			$sHTMLBody &= @CRLF & "<td>"

		 ;------------------------------------
		 ;The line determines the end of text in table cell
		 Case "</entry>"
			$bTableFlag = True
			$sHTMLBody &= @CRLF & "</td>"



		 ;------------------------------------
		 ;The line determines the end of a row
		 Case "</row>"
			$bTableFlag = True
			$sHTMLBody &= @CRLF & "</tr>"
		 ;------------------------------------
		 ;The line determines the end of a table
		 Case "</table>"
			$bTableFlag = False
			$sHTMLBody &= @CRLF & "</table>"
		 ;------------------------------------
		 ;The line determines graphic
		 Case "<graphic"
			Local $sImageName, $sImageWidth, $sImageHeight
			GetImageInfo ($sString, $sImageName, $sImageWidth, $sImageHeight)
			Local $sSource = $sImagePath & "/" & $sImageName & ".png"
			Local $sDes = @ScriptDir
			FileCopy ($sSource, $sDes, $FC_OVERWRITE)

			$sHTMLBody &= "<br>" & @CRLF & "<br>" & @CRLF & "<br>" & @CRLF &  "<img src=""" & $sImageName & ".png""" & " width=""" &  $sImageWidth*100 & """ height=""" & $sImageHeight*100 & """ border=""2"">" & "<br>" & @CRLF

		 ;------------------------------------
		 ;The line determines Test Item Name
		 Case "<testtitle>"
			$iItem_Count_2 = 0
			$iItem_Count_1 += 1
			Local $sTestTitle = GetInnertext ($sString)
			Local $sID_Line = $aXMLData [$i-1]
			$sHTMLBody &= "<br>" & @CRLF & "<br>" & @CRLF &  "<b id="""    & GetID ($sID_Line) &    """>" & $iItem_Count_1 & ". " & $sTestTitle & "</b>"
		 ;------------------------------------
		 ;The line determines OK, NG
		 Case "<down"
			Local $sResult = GetInnertext ($sString)
			Local $sID = GetID ($sString)
			For $j = 1 To 900
			   If $aID [$j][0] = $sID Then
				  Local $sStep = $aID [$j][1]
				  ExitLoop
			   EndIf
			Next


			$sHTMLBody &= "<br>" & @CRLF & "<br>" & @CRLF &  "<b>" & $sResult & "</b>" & " --- "
			$sHTMLBody &= "<a href=""#" & $sID & """>" & $sStep & "</a>"

		 ;------------------------------------
		 ;The line determines OK, NG
		 Case "<right"
			Local $sResult = GetInnertext ($sString)
			Local $sID = GetID ($sString)
			For $j = 1 To 900
			   If $aID [$j][0] = $sID Then
				  Local $sStep = $aID [$j][1]
				  ExitLoop
			   EndIf
			Next

			$sHTMLBody &= "<br>" & @CRLF & "<br>" & @CRLF &  "<b>" & $sResult & "</b>" & " --- "
			$sHTMLBody &= "<a href=""#" & $sID & """>" & $sStep & "</a>"

		 ;------------------------------------
		 Case "<test1>"
			$iItem_Count_2 += 1
			$sHTMLBody &= "<br>" & @CRLF & "<br>" & @CRLF & Chr ($iItem_Count_2 + 96) & ")waiting for replace"

		 ;------------------------------------
		 Case "<atten4>"
			$sHTMLBody &=  "<br>" & @CRLF & "<br>" & @CRLF & "<b> HINT: </b><br>" & @CRLF & "<ol>"

		 ;------------------------------------
		 Case "</atten4>"
			$sHTMLBody &= @CRLF & "</ol>" & @CRLF

		 ;------------------------------------
		 Case "<atten3>"
			$sHTMLBody &=  "<br>" & @CRLF & "<br>" & @CRLF & "<b> NOTICE: </b><br>" & @CRLF & "<ol>"

		 ;------------------------------------
		 Case "</atten3>"
			$sHTMLBody &= @CRLF & "</ol>" & @CRLF


		 Case Else

	  EndSwitch
   Next
   $sHTMLBody = StringReplace($sHTMLBody, "<td><br>", "<td>")
   $sHTMLBody = StringReplace($sHTMLBody, ")waiting for replace" & "<br>" & @CRLF & "<br>" & @CRLF, ") ")
   Create_HTML (@ScriptDir, "DTC Demo", "DTC Demo", $sHTMLBody)
EndFunc


;========================================================================================
;   FUNCTION DISCRIPTION: GET IMAGE INFORMATION
;	INPUT               : $sString
;   RETURN			    : THE TAG NAME OF A LINE
;========================================================================================
Func GetImageInfo ($sString, ByRef $sImageName, ByRef $sImageWidth, ByRef $sImageHeight)
   ;Get image name
   Local $iStart = StringInStr ($sString, """", 0, 1) + 1
   Local $iEnd = StringInStr ($sString, """", 0, 2) - 1
   Local $iCount = $iEnd - $iStart + 1
   $sImageName = StringMid ($sString, $iStart, $iCount)

   ;Get image width
   Local $iStart = StringInStr ($sString, """", 0, 3) + 1
   Local $iEnd = StringInStr ($sString, "in", 0, 1) - 2
   Local $iCount = $iEnd - $iStart + 1
   $sImageWidth = StringMid ($sString, $iStart, $iCount)

   ;Get image height
   Local $iStart = StringInStr ($sString, """", 0, 5) + 1
   Local $iEnd = StringInStr ($sString, "in", 0, 2) - 2
   Local $iCount = $iEnd - $iStart + 1
   $sImageHeight = StringMid ($sString, $iStart, $iCount)

EndFunc




;========================================================================================
;   FUNCTION DISCRIPTION: GET TAG NAME OF A LINE
;	INPUT               : $sString
;   RETURN			    : THE TAG NAME OF A LINE
;========================================================================================
Func GetTagName ($sString)
   If StringInStr ($sString, ">", 0 , 2) <> 0 Then
	  Local $sTagName = StringLeft ($sString, StringInStr ($sString, ">", 0 ,1))
	  If StringInStr ($sTagName, " ", 0) <> 0 Then $sTagName = StringLeft ($sString, StringInStr ($sString, " ", 0 ,1) - 1)

   ElseIf StringInStr ($sString, " ", 0) = 0 Then
	  Local $sTagName = $sString
   Else
	  Local $sTagName = StringLeft ($sString, StringInStr ($sString, " ", 0 ,1) - 1)
   EndIf

   Return $sTagName
EndFunc


;========================================================================================
;   FUNCTION DISCRIPTION: GET TAG NAME OF A LINE
;	INPUT               : $sString
;   RETURN			    : THE TAG NAME OF A LINE
;========================================================================================
Func GetInnertext ($sString)
   Local $iStart = StringInStr ($sString, ">", 0 ,1) + 1
   Local $iEnd = StringInStr ($sString, "<", 0 ,2) - 1
   Local $iCount = $iEnd - $iStart + 1
   Local $sInnertext = StringMid ($sString, $iStart, $iCount)
   Return $sInnertext
EndFunc


;========================================================================================
;   FUNCTION DISCRIPTION: GET ID
;	INPUT               : $sString
;   RETURN			    : ID String
;========================================================================================
Func GetID ($sString)
   Local $iStart = StringInStr ($sString, """", 0 ,1) + StringLen ("""")
   Local $iEnd = StringInStr ($sString, """", 0 ,2) - 1
   Local $iCount = $iEnd - $iStart + 1
   Local $sID = StringMid ($sString, $iStart, $iCount)
   Return $sID
EndFunc

;<testgrp id="RM000002PPM007X_05_0013" proc-id="RM0810E___000091S00000">

;========================================================================================
;   FUNCTION DISCRIPTION: DISPLAY INPUT BOXES FOR USER TO ENTER INITIAL DATA
;	INPUT               : $sFilePath OF THE XML FILE
;   RETURN			    : THE CONTENT OF THE XML FILE
;========================================================================================
Func OpenXML_GetData ($sFilePath)
   ;Open the file for reading and store the handle to a variable.
   Local $hFileOpen = FileOpen($sFilePath, $FO_READ)
   ;Read the contents of the file using the handle returned by FileOpen.
   Local $sFileRead = FileRead($hFileOpen)
   ;Close the handle returned by FileOpen.
   FileClose($hFileOpen)
   Return $sFileRead
EndFunc



;========================================================================================
;   FUNCTION DISCRIPTION: CREATE HTML FILE
;	INPUT               : $sFilePath, $sTxt_Title,$HTML_body
;   RETURN              : AN HTML FILE IN $sFilePath
;========================================================================================

Func Create_HTML ($sFilePath, $sName, $sTitle, $sHTMLBody)

   Local $sHTML = ""
	  $sHTML &= "<HTML>" & @CRLF
	  $sHTML &= "<HEAD>" & @CRLF
	  $sHTML &= "<TITLE>" & $sTitle & "</TITLE>" & @CRLF

	  $sHTML &= "<style type=""text/css"">" & @CRLF
	  $sHTML &= "table" & @CRLF & "{" & @CRLF & "border-collapse: collapse;" & @CRLF & "}"
	  $sHTML &= "table, th, td" & @CRLF & "{" & @CRLF & "border:1px solid black;" & @CRLF & "}"

	  $sHTML &= "</style>"

	  $sHTML &= "</HEAD>" & @CRLF
	  $sHTML &= "<BODY>"
	  $sHTML &= "<OL>"
	  $sHTML &= $sHTMLBody & @CRLF
	  $sHTML &= @CRLF & "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"
	  $sHTML &= "</OL>" & @CRLF
	  $sHTML &= "</BODY>" & @CRLF
	  $sHTML &= "</HTML>"

   Local $hFileOpen = FileOpen ($sFilePath & "\" & $sName & ".html",$FO_OVERWRITE)
   FileWrite($hFileOpen, $sHTML)
   FileClose($hFileOpen)
EndFunc
