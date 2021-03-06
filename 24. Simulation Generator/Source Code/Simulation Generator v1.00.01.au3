#cs
v01.00.01
   Created first version
   VIN deocde

v01.00.02

v01.00.03


#ce

#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>
#include <StringConstants.au3>
#include <Array.au3>


Main()


;====================================================================================================================
Func Main()
   ;~ $sVIN = '1FADP3J2XJL222680'
   $sVIN = InputBox('Enter VIN', 'Please Enter a VIN!')
   $sVIN = StringReplace($sVIN, ' ', '')

   If StringLen($sVIN) <> 17 Then
	  MsgBox (0,'VIN Wrong','The VIN you entered is wrong' & @CRLF & 'Please check your VIN!')
	  Exit
   EndIf

   Local $sVIN_Hex = VIN_To_Hex($sVIN)
   Local $sSim = Add_Original_Txt()
   $sSim &= @CRLF & @CRLF & _
			'INFO_DATABASE = Req>1			000007DF 08 02 09 02 00 00 00 00 00 	NONE	0	0' & @CRLF & _
			'INFO_DATABASE = Res<1			000007E8 08 10 14 49 02 01 ' & StringMid($sVIN_Hex,1,8)  & ' 	NONE	0	0' & @CRLF & _
			'INFO_DATABASE = Res<1			000007E8 08 21 ' & StringMid($sVIN_Hex,10,20) & ' 	NONE	0	0' & @CRLF & _
			'INFO_DATABASE = Res<1			000007E8 08 22 ' & StringMid($sVIN_Hex,31,20) & ' 	NONE	0	0'

   $sFileName = InputBox('Enter File Name', 'Please Enter File Name!')
   Global $sFilePath = @DesktopDir
   WriteTxtFile ($sVIN & '_' & $sFileName & '.sim', $sSim, 'overwrite')
EndFunc



;====================================================================================================================
Func WriteTxtFile ($sFileName, $sTxt, $sMode = "append")
   If $sMode = "overwrite" Then
	  Local $hFileOpen = FileOpen ($sFilePath & "\" & $sFileName,$FO_OVERWRITE)
   Else
	  Local $hFileOpen = FileOpen ($sFilePath & "\" & $sFileName,$FO_APPEND)
   EndIf
   FileWrite($hFileOpen, $sTxt)
   FileClose($hFileOpen)
EndFunc




;====================================================================================================================
Func Autoit_Exit ()
   Exit
EndFunc




;====================================================================================================================
Func VIN_To_Hex($sVIN)

   Local $aVIN = StringSplit($sVIN, '')
   Local $sVIN_Hex
   For $i = 1 To $aVIN[0]
	  $sVIN_Hex &= Hex (Asc($aVIN[$i]), 2) & ' '
   Next

   Return $sVIN_Hex
EndFunc



;====================================================================================================================
Func Add_Original_Txt ()
   Local $sOriginal_Txt = '' & _
   '###########################################' & @CRLF & _
   '#         Auto Generated                  #' & @CRLF & _
   '###########################################' & @CRLF & _
   '<config sw> Protocol = 29' & @CRLF & _
   '<config sw> PIN_KRX_CANH = 6' & @CRLF & _
   '<config sw> TYPE_KRX_CANH = 0' & @CRLF & _
   '<config sw> VOLT_KRX_CANH = 3' & @CRLF & _
   '<config sw> PIN_KTX_CANH = 14' & @CRLF & _
   '<config sw> TYPE_KTX_CANH = 0' & @CRLF & _
   '<config sw> VOLT_KTX_CANH = 3' & @CRLF & _
   '<config sw> PIN_LRX_CANH =  6' & @CRLF & _
   '<config sw> TYPE_LTX_CANH = 0' & @CRLF & _
   '<config sw> VOLT_LTX_CANH = 3' & @CRLF & _
   '<config sw> VREF = 0' & @CRLF & _
   '<config sw> BAUDRATE = 500000' & @CRLF & _
   '<config sw> DATABIT = 0' & @CRLF & _
   '<config sw> PARITY = 0' & @CRLF & _
   '<config sw> TBYTE = 8' & @CRLF & _
   '<config sw> TFRAME = 10' & @CRLF & _
   '<config sw> F CAN NUMBER FRAME = 1' & @CRLF & _
   '<config sw> RANGE =   700,7ff;' & @CRLF & _
   '###########################################' & @CRLF & _
   '#         End of config                   #' & @CRLF & _
   '###########################################' & @CRLF & _
   '' & @CRLF & _
   '' & @CRLF & _
   '' & @CRLF & _
   '//MODE 1' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 02 01 xx xx xx xx xx xx 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E9 08 06 41 00 98 18 80 13 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 06 41 00 BE 3F A8 13 00 	NONE	0	0' & @CRLF & _
   '' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 02 01 01 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E9 08 06 41 01 01 04 65 65 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 06 41 01 00 07 65 00 00 	NONE	0	0' & @CRLF & _
   '' & @CRLF & _
   '//MODE 3' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 01 03 xx xx xx xx xx xx 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E9 08 04 43 01 12 33 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 02 43 01 F0 09 00 00 00 	NONE	0	0' & @CRLF & _
   '' & @CRLF & _
   '//MODE 7' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 01 07 00 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E9 08 04 47 01 F0 09 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 10 10 47 06 01 21 01 23 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E0 08 30 00 00 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 21 07 06 07 07 12 33 21 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 22 01 02 28 00 00 00 00 	NONE	0	0' & @CRLF & _
   '' & @CRLF & _
   '//MODE 0A' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 01 0A 00 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 04 4A 01 01 06 00 00 00 	NONE	0	0' & @CRLF & _
   '' & @CRLF & _
   '' & @CRLF & _
   '' & @CRLF & _
   '//MODE 2' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 00 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E9 08 07 42 00 00 58 18 80 03 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 07 42 00 00 FE 3F 88 03 	NONE	0	0' & @CRLF & _
   '' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 20 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E9 08 07 42 20 00 00 00 00 01 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 07 42 20 00 00 17 F0 11 	NONE	0	0' & @CRLF & _
   '' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 40 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E9 08 07 42 40 00 40 80 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 07 42 40 00 FE D0 00 00 	NONE	0	0' & @CRLF & _
   '' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 02 01 01 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E9 08 06 41 01 81 04 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 06 41 01 00 07 E5 E5 00 	NONE	0	0' & @CRLF & _
   '' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 01 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 07 42 01 00 00 07 E5 E5 	NONE	0	0' & @CRLF & _
   '' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 02 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E9 08 05 42 02 00 C1 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 05 42 02 00 12 33 00 00 	NONE	0	0' & @CRLF & _
   '' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 03 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 05 42 03 00 01 00 00 00 	NONE	0	0' & @CRLF & _
   '' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 04 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E9 08 04 42 04 00 75 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 04 42 04 00 FF 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 05 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 04 42 05 00 16 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E9 08 04 42 05 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 06 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 04 42 06 00 80 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 07 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 04 42 07 00 80 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 0B 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 04 42 0B 00 7B 00 00 00 	NONE	0	0' & @CRLF & _
   '' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 0C 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E9 08 05 42 0C 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 05 42 0C 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   '' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 0D 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E9 08 04 42 0D 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 04 42 0D 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 0E 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 04 42 0E 00 94 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 0F 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 04 42 0F 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 10 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 05 42 10 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 11 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E9 08 04 42 11 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 04 42 11 00 FB 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 15 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 05 42 15 00 FF FF 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 1F 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E9 08 05 42 1F 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 05 42 1F 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 2C 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 04 42 2C 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 2E 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 04 42 2E 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 2F 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 04 42 2F 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 30 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 04 42 30 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 31 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 05 42 31 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 32 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 05 42 32 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 33 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 04 42 33 00 5B 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 34 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 07 42 34 00 00 00 80 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 3C 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 05 42 3C 00 00 DE 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 41 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 07 42 41 00 00 76 00 E5 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 42 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E9 08 05 42 42 00 28 BE 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 05 42 42 00 29 57 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 43 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 05 42 43 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 44 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 05 42 44 00 86 66 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 45 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 04 42 45 00 DC 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 46 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 04 42 46 00 16 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 47 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 04 42 47 00 FD 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 49 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E9 08 04 42 49 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 04 42 49 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 4A 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 04 42 4A 00 01 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 03 02 4C 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 04 42 4C 00 18 00 00 00 	NONE	0	0' & @CRLF & _
   '' & @CRLF & _
   '//MODE 1 Live Data' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 02 01 0D 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E9 08 04 41 0D 00 FF 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 04 41 0D 00 FF 00 00 00 	NONE	0	0' & @CRLF & _
   '' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 02 01 05 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 04 41 05 00 16 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E9 08 04 41 05 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   '' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 02 01 11 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E9 08 04 41 11 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 04 41 11 00 FB 00 00 00 	NONE	0	0' & @CRLF & _
   '' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 02 01 0C 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E9 08 05 41 0C FF 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 05 41 0C FF 00 00 00 00 	NONE	0	0' & @CRLF & _
   '' & @CRLF & _
   '' & @CRLF & _
   '' & @CRLF & _
   '//MODE 9' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 02 01 01 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E9 08 06 41 01 81 04 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 06 41 01 00 07 E5 E5 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 02 09 00 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E9 08 06 49 00 14 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 06 49 00 55 40 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Req>1			000007DF 08 02 09 08 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 10 2B 49 08 14 02 04 06 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E0 08 30 00 00 00 00 00 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 21 B5 02 3C 02 04 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 22 00 00 02 A8 02 04 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 23 00 00 00 02 B9 02 04 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 24 00 00 00 00 00 B5 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 25 46 02 44 02 04 00 00 	NONE	0	0' & @CRLF & _
   'INFO_DATABASE = Res<1			000007E8 08 26 00 00 00 00 00 00 00 	NONE	0	0'
   Return $sOriginal_Txt
EndFunc

