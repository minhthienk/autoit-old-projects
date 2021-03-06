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


Global $sFilePath = @ScriptDir

; Set Hotkey for the program
HotKeySet("{ESC}", "Autoit_Exit")

Local $oIE = _IECreate ('http://www.autocats.ws/manual/chevrolet/tis0211/en/GMDE_TIS_START.html')
Local $oFrame = _IEFrameGetCollection($oIE, 1)

Local $sButton_Name = 'Matiz/Spark'
Local $oButton = _IEGetObjByName ($oFrame, $sButton_Name)

_IEAction ($oButton, 'click')

Sleep (1000)
MsgBox ($MB_TOPMOST, '', 'done')

Exit




