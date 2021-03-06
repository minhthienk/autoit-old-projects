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

Chú ý turn off script debugging trong IE
Thêm dòng đầu html
Tắt image tăng tốc dowload
#ce ----------------------------------------------------------------------------

#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <GuiEdit.au3>



#include "General_Library.au3"
#include "Create_JAVASCRIPT_Procedure.au3"
#include "Create_NORMAL_Procedure.au3"
#include "Create_DTC.au3"


HotKeySet ("^zas","Check_Initial_Input")



Global $bWeb_Attach = 0
Global $bWeb_Visible = 0
Global $bWeb_Wait = 1
Global $bWeb_TakeFocus = 0
Global $bImage_Download = 1

Global $iSubscription_Num
Global $sLink_DTC

Global $bRun_Flag = False


Func Close_All_IE()
   $Proc = "iexplore.exe"
   While ProcessExists($Proc)
      ProcessClose($Proc)
   Wend
EndFunc


Func Autoit_Exit ()
   Close_All_IE()
   Exit
EndFunc


Func Set_Run_Flag ()
   $bRun_Flag = True
EndFunc


Close_All_IE()


#Region ### START GUI section ### Form=
   Opt("GUIOnEventMode", 1)
   GUISetOnEvent($GUI_EVENT_CLOSE, "Autoit_Exit")
   $Form1 = GUICreate("Prepair Procedure Generator", 329, 460, -1, -1)
   GUISetBkColor(0xE4E4E4)
   ;-------------------------------------------
   ;CREATE GUI INPUT ELEMENTS
   $Input_Link = GUICtrlCreateInput("", 32, 40, 265, 21)
   $Label1 = GUICtrlCreateLabel("Input DTC Link ", 128, 16, 114, 17)

   $Radio_Visible = GUICtrlCreateRadio("Web Visible", 32, 80, 113, 17)
   $Radio_Invisible = GUICtrlCreateRadio("Web Invisible", 32, 104, 113, 17)
   GUICtrlSetState ($Radio_Invisible, $GUI_CHECKED)

   $Combo_Subscription = GUICtrlCreateCombo("(License #)", 152, 96, 145, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
   GUICtrlSetData(-1, "# 1|# 2|# 3|# 4|# 5")
   $Label2 = GUICtrlCreateLabel("Select License", 180, 78, 120, 17)
   ;-------------------------------------------
   ;CREATE GUI NOTIFICATION PLACE
   $Noti_Ctrl = GUICtrlCreateEdit("", 32, 192, 265, 250)
   GUICtrlSetBkColor(-1, 0xF0F0F0)
   _GUICtrlEdit_SetReadOnly ($Noti_Ctrl, True)
   _GUICtrlEdit_SetMargins ($Noti_Ctrl, BitOR($EC_LEFTMARGIN, $EC_RIGHTMARGIN), 7, 7)
   $Label4 = GUICtrlCreateLabel(" Notification", 28, 172, 60, 17)
   ;-------------------------------------------
   ;CREATE BUTTONS AND SET ON EVENT
   $Button_Begin = GUICtrlCreateButton("Begin", 32, 128, 75, 25)
   GUICtrlSetOnEvent($Button_Begin, "Set_Run_Flag")

   $Button_Clear = GUICtrlCreateButton("Clear Notif", 130, 128, 75, 25)
   GUICtrlSetOnEvent($Button_Clear, "Notification_Clear")

   $Button_Exit = GUICtrlCreateButton("Exit", 223, 128, 75, 25)
   GUICtrlSetOnEvent($Button_Exit, "Autoit_Exit")
   ;-------------------------------------------
   ;SHOW GUI
   GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###



While 1
   If $bRun_Flag = True Then Check_Initial_Input()
WEnd




Func Check_Initial_Input ()
   ;-------------------------------------------------
   ;CHECK WEB SHOW/HIDE
   If GUICtrlRead ($Radio_Visible) = $GUI_CHECKED Then $bWeb_Visible = 1
   If GUICtrlRead ($Radio_Invisible) = $GUI_CHECKED Then $bWeb_Visible = 0
   ;-------------------------------------------------
   ;CHECK SUBSCRIPTION NUMBER
   Local $sCombo_Subscription_Val = GUICtrlRead ($Combo_Subscription)
   Local $bSub_Flag = 0
   Local $sNoti_Sub = ""
   Switch $sCombo_Subscription_Val
   Case  "(License #)"
	  Notification ("Please select your LISENSE NUMBER", "Normal")
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
   Case Else
	  Notification ("Please select your LISENSE NUMBER", "Normal")
	  $bSub_Flag = 0
   EndSwitch
   ;-------------------------------------------------
   ;CHECK ALLDATA LINK
   $sLink_DTC = GUICtrlRead ($Input_Link)
   Local $bLink_Flag = 0
   Local $sNoti_Link = ""
   If StringLeft ($sLink_DTC, 26) <> "http://repair.alldata.com/" And StringLeft ($sLink_DTC, 26) <> "https://repair.alldata.com"  Then
	  Notification ("The link is INVALID. Please input VALID link", "Normal")
	  $bLink_Flag = 0
   Else
	  $bLink_Flag = 1
   EndIf
   ;-------------------------------------------------
   ;IF SUBSCRIPTION IS SELECTED AND THE LINK IS VALID => EXECUTE
   If ($bSub_Flag = 1) And ($bLink_Flag = 1) Then
	  Notification ("Please wait ...", "Normal")
	  GUI_Input_Ctrl_Set_State ("Disable")
	  If Main_function () = "The link is invalid" Then
		 Notification ("The link is not DTC Link", "Normal")
		 GUI_Input_Ctrl_Set_State ("Enable")
	  Else
		 Notification ("Done. Please Check!", "Normal")
		 GUI_Input_Ctrl_Set_State ("Enable")
	  EndIf
   EndIf
   $bRun_Flag = False
EndFunc



Func GUI_Input_Ctrl_Set_State ($sState)
   If $sState = "Disable" Then
	  GUICtrlSetState ($Button_Begin, $GUI_DISABLE)
	  GUICtrlSetState ($Radio_Visible, $GUI_DISABLE)
	  GUICtrlSetState ($Radio_Invisible, $GUI_DISABLE)
	  GUICtrlSetState ($Combo_Subscription, $GUI_DISABLE)
	  GUICtrlSetState ($Input_Link, $GUI_DISABLE)
   Else
	  GUICtrlSetState ($Button_Begin, $GUI_ENABLE)
	  GUICtrlSetState ($Radio_Visible, $GUI_ENABLE)
	  GUICtrlSetState ($Radio_Invisible, $GUI_ENABLE)
	  GUICtrlSetState ($Combo_Subscription, $GUI_ENABLE)
	  GUICtrlSetState ($Input_Link, $GUI_ENABLE)
   EndIf
EndFunc




Func Notification ($sNoti, $sMode) ;Mode = "Normal", "Replace Previous"
   Static Local $sPrevious_Noti
   If $sMode = "Normal" Then
	  _GUICtrlEdit_AppendText ($Noti_Ctrl,  @CRLF & $sNoti & @CRLF)
   Else
	  _GUICtrlEdit_SetReadOnly ( $Noti_Ctrl, False )
	  _GUICtrlEdit_Undo ($Noti_Ctrl)
	  _GUICtrlEdit_SetReadOnly ( $Noti_Ctrl, True )
	  _GUICtrlEdit_AppendText ($Noti_Ctrl,  @CRLF & $sNoti & @CRLF)
   EndIf
EndFunc


Func Notification_Clear ()
   GUICtrlSetData ($Noti_Ctrl, "")
EndFunc
