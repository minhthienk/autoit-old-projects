#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here


#include <IE.au3>
#include <array.au3>
#include <clipboard.au3>
#include <StringConstants.au3>

$sUrl  = 'https://thuxe.vn/xe/trieu-hoi-xe/Mazda/3'
$oIE = _IECreate($sUrl)




Local $oTags = _IETagNameGetCollection($oIE, 'a')
Local $sTxt = ""
Local $aData[0]
For $oTag In $oTags
   If StringInStr($oTag.href, 'https://thuxe.vn/xe/trieu-hoi-xe/Mazda/3/') = 0 Then ContinueLoop

   Redim $aData[UBound($aData)+1]
   $aData[UBound($aData)-1] = $oTag.innertext
Next

_ArrayDisplay($aData)



Local $oTags = _IETagNameGetCollection($oIE, 'p')
Local $sTxt = ""
Local $aData2[0]
For $oTag In $oTags
   If StringInStr($oTag.innertext, 'Danh') <> 0 Or $oTag.innertext = '' Then ContinueLoop

   Redim $aData2[UBound($aData2)+1]
   $aData2[UBound($aData2)-1] = $oTag.innertext
Next
_ArrayDisplay($aData2)


Local $txt = ''
For $i=0 To UBound($aData2) - 1
   Local $aVIN = StringSplit ($aData2[$i], ' ', $STR_NOCOUNT + $STR_ENTIRESPLIT)

   For $j = 0 To UBound($aVIN) - 1
	  $txt &= $aData[$i] & @TAB & $aVIN[$j] & @CRLF
   Next
Next

MsgBox (0,0,$txt)
_ClipBoard_SetData($txt)

