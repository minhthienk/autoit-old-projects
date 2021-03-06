#cs ----------------------------------------------------------------------------
NOTE:
Làm file Log lưu lại hình bị lỗi khi tải (Nếu tải lâu hơn bao nhiêu giây thì phải vào function check mạng, note lại tốc độ mạng => tải lại)

http://repair.alldata.com/alldata/article/display.action?componentId=3926&iTypeId=423&nonStandardId=3951433&vehicleId=54277&windowName=mainADOnlineWindow
http://repair.alldata.com/alldata/article/display.action?componentId=3926&iTypeId=423&nonStandardId=3956429&vehicleId=54277&windowName=mainADOnlineWindow
http://repair.alldata.com/alldata/article/display.action?componentId=3926&iTypeId=423&nonStandardId=3952079&vehicleId=53841&windowName=mainADOnlineWindow

Link chứa DTC có Part:
http://repair.alldata.com/alldata/article/display.action?componentId=3926&iTypeId=423&nonStandardId=5349152&vehicleId=52950&windowName=mainADOnlineWindow

Link chứa EVAP
http://repair.alldata.com/alldata/article/display.action?componentId=3926&iTypeId=423&nonStandardId=5244558&vehicleId=47132

;Link thử nhiều procedure và có javascript
http://repair.alldata.com/alldata/article/display.action?componentId=3926&iTypeId=423&nonStandardId=5364910&vehicleId=52950&windowName=mainADOnlineWindow


Func Write_Log_File_Error ($sTxt)
	  Local $hFileOpen = FileOpen ("C:\Users\K\Desktop\Alldata DTC" & "\" & "Log File Error" & ".txt",$FO_APPEND)
	  FileWrite($hFileOpen, $sTxt & @CRLF & @CRLF)
EndFunc
#ce ----------------------------------------------------------------------------

#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>


#include "General_Library.au3"
#include "Create_JAVASCRIPT_Procedure.au3"
#include "Create_NORMAL_Procedure.au3"
#include "Create_DTC.au3"

HotKeySet ("{ESC}", Autoit_Exit)

Func Autoit_Exit ()

EndFunc

Global $bWeb_Attach = 0
Global $bWeb_Visible = 0
Global $bWeb_Wait = 1
Global $bWeb_TakeFocus = 0

Global $bImage_Download = 0
Global $iSubscription_Num

App_GUI ()

;Main_function ()



Func App_GUI ()
   #Region ### START Koda GUI section ### Form=
   $Form1 = GUICreate("Form1", 329, 270, 352, 197)
   $Input_Link = GUICtrlCreateInput("", 32, 40, 265, 21)
   $Label1 = GUICtrlCreateLabel("Input Alldata DTC Link ", 112, 16, 114, 17)
   $Button_Begin = GUICtrlCreateButton("Begin", 56, 128, 75, 25)
   $Button_Exit = GUICtrlCreateButton("Exit", 200, 128, 75, 25)
   $Radio_Visible = GUICtrlCreateRadio("Web Visible", 32, 80, 113, 17)
   $Radio_Invisible = GUICtrlCreateRadio("Web Invisible", 32, 104, 113, 17)
   $Combo_Subscription = GUICtrlCreateCombo("(License #)", 152, 96, 145, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
   GUICtrlSetData(-1, "# 1|# 2|# 3|# 4|# 5")
   $Label2 = GUICtrlCreateLabel("Select License", 180, 78, 120, 17)

   Global $Label_Notification = GUICtrlCreateLabel("", 40, 192, 250, 60)
   GUICtrlSetBkColor(-1, 0xFFFFFF)
   $Label5 = GUICtrlCreateLabel("", 32, 192, 8, 60)
   GUICtrlSetBkColor(-1, 0xFFFFFF)
   $Label6 = GUICtrlCreateLabel("", 290, 192, 8, 60)
   GUICtrlSetBkColor(-1, 0xFFFFFF)
   $Label4 = GUICtrlCreateLabel(" Notification", 32, 168, 60, 17)

   GUISetState(@SW_SHOW)
   #EndRegion ### END Koda GUI section ###

   GUICtrlSetState ($Radio_Invisible, $GUI_CHECKED)


   While 1
	  $nMsg = GUIGetMsg()
	  Switch $nMsg
		 Case $GUI_EVENT_CLOSE
			Exit
		 Case $Button_Begin
			;-------------------------------------------------
			;CHECK SUBSCRIPTION NUMBER
			Local $sCombo_Subscription_Val = GUICtrlRead ($Combo_Subscription)
			Local $bSub_Flag = 0
			Local $sNoti_Sub = ""
			Switch $sCombo_Subscription_Val
			   Case  "(License #)"
				  Local $sNoti_Sub = "Please select your subscrtion number" & @CRLF & @CRLF
				  $bSub_Flag = 0
			   Case  "# 1"
				  $iSubscription_Num = 1
				  $bSub_Flag = 1
			   Case  "# 2"
				  $iSubscription_Num = 2
				  $bSub_Flag = 1
			   Case  "# 3"
				  $iSubscription_Num = 3
				  $bSub_Flag = 1
			   Case  "# 4"
				  $iSubscription_Num = 4
				  $bSub_Flag = 1
			   Case  "# 5"
				  $iSubscription_Num = 5
				  $bSub_Flag = 1
			EndSwitch
			;-------------------------------------------------
			;CHECK ALLDATA LINK
			Local $Input_Link_Val = GUICtrlRead ($Input_Link)
			Local $bLink_Flag = 0
			Local $sNoti_Link = ""
			If StringInStr ($Input_Link_Val, "repair.alldata.com", 0, 1) = 0 Then
			   $sNoti_Link = "The Link is invalid" & @CRLF & "Please input valid link"
			   $bLink_Flag = 0
			Else
			   $bLink_Flag = 1
			EndIf
			;-------------------------------------------------
			;IF SUBSCRIPTION IS SELECTED AND THE LINK IS VALID => EXECUTE
			If ($bSub_Flag = 1) And ($bLink_Flag = 1) Then
			   Main_function ($Input_Link_Val)
			Else
			   Notification ($sNoti_Sub & $sNoti_Link)
			EndIf
			;-------------------------------------------------
		 Case $Button_Exit
			Exit
		 Case $Radio_Visible
			$bWeb_Visible = 1
		 Case $Radio_Invisible
			$bWeb_Visible = 0
		 Case $Combo_Subscription

	  EndSwitch
   WEnd


EndFunc



Func Notification ($sNoti)
   GUICtrlSetData ($Label_Notification, $sNoti)
EndFunc

