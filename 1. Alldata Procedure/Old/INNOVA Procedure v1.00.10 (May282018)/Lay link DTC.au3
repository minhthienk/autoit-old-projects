#include <MsgBoxConstants.au3>
#include <FileConstants.au3>
#include <InetConstants.au3>

#include <Clipboard.au3>
#include <IE.au3>
#include <Excel.au3>
#include <WinAPIFiles.au3>

#include "General_Library.au3"
#include "Create_JAVASCRIPT_Procedure.au3"
#include "Create_NORMAL_Procedure.au3"


;Local $sLink = "http://repair.alldata.com/alldata/navigation/treedisplay.action?vehicleId=47006&componentId=0&iTypeId=0&nonStandardId=0"
Local $sLink = "http://repair.alldata.com/alldata/navigation/treedisplay.action?vehicleId=54276&componentId=1&iTypeId=0&nonStandardId=0&fromJs=true&openUrl=#ygtvlabelel1"
Local $oIE = _IECreate ("about:blank")
Sleep (1000)
_IENavigate ($oIE, $sLink)
_IEAction($oIE, "refresh")


Local $sID_Click
Local $sID_Sub

;-----------------------------------
;CLICK TREE VIEW LEVEL 1
;Wait for the text to appear
Do
   Sleep (500)
   Local $sHTML = _IEPropertyGet ($oIE, "outertext")
Until StringInStr ($sHTML, "A L L Diagnostic Trouble Codes ( DTC )") <> 0
Sleep (1000)
;Get object
$oIE_SubObject = $oIE
;Get ID by text
$sID_Click = GetIDByText ($oIE, "A L L")
;Get ID from the previous $sID_Click
$sID_Sub = StringReplace ($sID_Click, "vt", "vc")
;Get Object by ID
$oTreeView = _IEGetObjById ($oIE, $sID_Click)
;Click object
_IEAction($oTreeView, "click")
;-----------------------------------
;CLICK TREE VIEW LEVEL 2
;Wait for the text to appear
Do ;Loop until get the SubOject
   ;Get Object by ID
   $oIE_SubObject = _IEGetObjById ($oIE, $sID_Sub)
   Sleep (200)
Until @error = 0
;Get ID by text
$sID_Click = GetIDByText ($oIE_SubObject, "Information for A L L")
$sID_Sub = StringReplace ($sID_Click, "vt", "vc")
;Get Object by ID
$oTreeView = _IEGetObjById ($oIE, $sID_Click)
;Click object
_IEAction($oTreeView, "click")
;-----------------------------------
;CLICK TREE VIEW LEVEL 2
;Wait for the text to appear
Do ;Loop until get the SubOject
   ;Get Object by ID
   $oIE_SubObject = _IEGetObjById ($oIE, $sID_Sub)
   Sleep (200)
Until @error = 0
;Get ID by text
$sID_Click = GetIDByText ($oIE_SubObject, "Testing and Inspection")
$sID_Sub = StringReplace ($sID_Click, "vt", "vc")
;Get Object by ID
$oTreeView = _IEGetObjById ($oIE, $sID_Click)
;Click object
_IEAction($oTreeView, "click")
;-----------------------------------
;CLICK TREE VIEW LEVEL 3
Do ;Loop until get the SubOject
   ;Get Object by ID
   $oIE_SubObject = _IEGetObjById ($oIE, $sID_Sub)
   Sleep (200)
Until @error = 0
;Get ID by text
$sID_Click = GetIDByText ($oIE_SubObject, "P Code")
$sID_Sub = StringReplace ($sID_Click, "vt", "vc")
;Get Object by ID
$oTreeView = _IEGetObjById ($oIE, $sID_Click)
;Click object
_IEAction($oTreeView, "click")




Do ;Loop until get the SubOject
   ;Get Object by ID
   $oIE_SubObject = _IEGetObjById ($oIE, $sID_Sub)
   Sleep (200)
Until @error = 0
;------------------------------------------------------------------------------------------------------------------
;ĐOẠN CODE LẤY TEXT VÀ LINK PROCEDURE TRONG TAG <A>
   ;----------------------------------------------------
   Local $oTags = _IETagNameGetCollection($oIE_SubObject, "td")
   Local $Array [1000]
   Local $i = 0, $iMark = 0
   Local $Txt
   ;Loop until found the string

	  For $oTag In $oTags
		 $Txt &= $oTag.innertext & @CRLF
		 $i += 1
	  Next





MsgBox (0, "", $Txt)
Exit






Func GetIDByText ($oIE, $sTxt)
   ;----------------------------------------------------
   Local $oTags = _IETagNameGetCollection($oIE, "td")
   Local $Array [1000]
   Local $i = 0, $iMark = 0
   ;Loop until found the string
   Do
	  For $oTag In $oTags
		 $Array [$i] = $oTag
		 If StringInStr ($oTag.innertext, $sTxt) <> 0 Then
			$iMark = $i
		 EndIf
		 $i += 1
	  Next
   Until $iMark <> 0

   $iMark = $iMark - 1
   $oTemp = $Array [$iMark]
   Return $oTemp.id
EndFunc
Exit