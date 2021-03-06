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
HotKeySet("{ESC}", "Autoit_Exit")
;HotKeySet("^z", "Excel_Open")

MsgBox (0, "NOTE", "Click OK to get 5 first pages of Ford vehicles data from the site ""bonbanh.com"" " & @CRLF & @CRLF & "PRESS the ESC button to exit the application if needed")
Excel_Open ()


while (1)
WEnd



Func Autoit_Exit ()
   Exit
EndFunc

Func Demo ()
   ; Open browser with basic example, get link collection,
   ; loop through items and display the associated link URL references

   #include <IE.au3>
   #include <MsgBoxConstants.au3>

   Local $oIE = _IECreate("https://bonbanh.com/oto/ford/page,2")
   Local $oLinks = _IELinkGetCollection($oIE)
   Local $iNumLinks = @extended

   Local $sTxt = $iNumLinks & " links found" & @CRLF & @CRLF
   For $oLink In $oLinks
	   $sTxt &= $oLink.href & @CRLF
   Next

   _ClipBoard_SetData ($sTxt, $CF_UNICODETEXT)
   MsgBox (0, "","Done")
   Autoit_Exit ()
EndFunc


Func Excel_Open ()
   Local $oExcel = _Excel_Open()
   Local $oWorkbook = _Excel_BookNew ( $oExcel)
   Local $aExcel_Column [9] =  ["A", "B", "C", "D", "E", "F", "G", "H", "I"]
   Local $Excel_Row

   Local $aVehicle_Link_1Page [21]
   Local $aData [10]

   For $pagenum = 2 To 141
	  $aVehicle_Link_1Page = Get_Vehicle_Link_1Page ("https://bonbanh.com/oto/ford/page," & $pagenum)
	  For $j=1 To 20 Step 1
		 $aData = Get_Vehicle_Data ($aVehicle_Link_1Page [$j])
		 If $j = 1 Then
			$Excel_Row = $j + ($pagenum-1)*20 + $pagenum - 1
			_Excel_RangeWrite ($oWorkbook, $oWorkbook.Activesheet, "Page " & $pagenum, $aExcel_Column [0] & $Excel_Row)
			$Excel_Row = $j + ($pagenum-1)*20 + $pagenum
		 Else
			$Excel_Row = $j + ($pagenum-1)*20 + $pagenum
		 EndIf
		 For $i = 0 To 8 Step 1
			_Excel_RangeWrite ($oWorkbook, $oWorkbook.Activesheet, $aData [$i+1], $aExcel_Column[$i] & $Excel_Row)
		 Next
	  Next
   Next
   MsgBox (0, "NOTE", "Done")
   Autoit_Exit ()

EndFunc



Func Get_Vehicle_Link_1Page ($Link_Page)

   Local $oIE = _IECreate($Link_Page, 0, 0)
   Sleep (100)
   Local $oLinks = _IELinkGetCollection($oIE)
   Local $iNumLinks = @extended

   Local $aData_All [651]
   Local $aData_Select [21]
   Local $count = 0

   For $oLink In $oLinks
	  $count = $count + 1
	  $aData_All [$count] = $oLink.href
   Next


   Local $Mark = 0
   Local $Flag = 0

   For $i=650 To 1 Step -1
	  If $aData_All [$i] = "https://bonbanh.com/oto/ford-ranger" Then $Flag += 1

	  If $Flag = 2 Then
		 If ASC(StringRight($aData_All [$i],1)) >= 48 And ASC(StringRight($aData_All [$i],1)) <= 57 Then
			$Mark = $i
			ExitLoop
		 EndIf
	  EndIf

   Next

   $count = 0
   For $i = $Mark to ($Mark - 19) Step -1
	  $count += 1
	  $aData_Select [$count] = $aData_All [$i]
   Next

   _IEQuit($oIE)
   Return $aData_Select

EndFunc










Func Get_Vehicle_Data ($Vehicle_Link)

   Local $oIE = _IECreate($Vehicle_Link, 0, 0)
   Local $oSpans = _IETagNameGetCollection($oIE, "span")
   Local $sTxt = ""
   Local $aData_All [1000]
   Local $aData_Select [10]
   Local $count = 0


   For $oSpan In $oSpans
	  $count = $count + 1
	  $aData_All [$count] = $oSpan.innertext
   Next


   $aData_Select [1] = $aData_All [20]
   $aData_Select [2] = $aData_All [38]
   $aData_Select [3] = $aData_All [40]
   $aData_Select [4] = $aData_All [44]
   $aData_Select [5] = $aData_All [45]
   $aData_Select [6] = $aData_All [46]
   $aData_Select [7] = $aData_All [48]
   $aData_Select [8] = $aData_All [49]
   $aData_Select [9] = $aData_All [50]

   _IEQuit($oIE)

   Return $aData_Select
EndFunc



#cs
Func IE_Example_Zing_Search ()

   Local $oIE = _IECreate("zing.vn")
   Sleep (1000)
   Local $SearchBox = _IEGetObjById($oIE, "search_keyword")
   _IEAction($SearchBox, "focus")
   Send ("alo")
   MsgBox (0,"","")



EndFunc



Func IE_Example_HTML_RreadWrite ()

   Local $oIE = _IE_Example("basic")
   Local $sHTML = _IEBodyReadHTML($oIE)
   Sleep (2000)
   $sHTML = $sHTML & "<p><font color=red size=+5>Big RED text!</font>"
   _IEBodyWriteHTML($oIE, $sHTML)

EndFunc


Func IE_Example_ReadText ()

   Local $oIE = _IECreate ("zing.vn")
   Local $sText = _IEBodyReadText($oIE)

   Run("notepad.exe")

EndFunc

#ce
