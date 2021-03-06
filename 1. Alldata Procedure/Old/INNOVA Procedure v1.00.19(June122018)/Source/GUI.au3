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

;Link DTC mẫu GM
http://repair.alldata.com/alldata/article/display.action?componentId=3926&iTypeId=423&nonStandardId=4391899&vehicleId=52299&windowName=mainADOnlineWindow

Link research lay link DTC
http://repair.alldata.com/alldata/navigation/treedisplay.action?nonStandardId=3844431&iTypeId=383&vehicleId=54276&openUrl=&fromJs=true&componentId=621&


;Link chứa text bên dưới hình
http://repair.alldata.com/alldata/article/display.action?componentId=3926&iTypeId=423&nonStandardId=3885481&vehicleId=39067&windowName=mainADOnlineWindow

;Link YMME test SCAN
;http://repair.alldata.com/alldata/navigation/treedisplay.action?vehicleId=54276&componentId=1&iTypeId=0&nonStandardId=0&fromJs=true&openUrl=#ygtvlabelel1

;Linik YMME la
http://repair.alldata.com/alldata/navigation/treedisplay.action?vehicleId=46884&componentId=0&iTypeId=0&nonStandardId=0&fromJs=false&openUrl=#ygtvlabelel621

;Lỗi loop
http://repair.alldata.com/alldata/navigation/treedisplay.action?vehicleId=49237&componentId=0&iTypeId=0&nonStandardId=0&fromJs=false&openUrl=


! Chú ý turn off script debugging trong IE
! Tắt image tăng tốc dowload
? Thêm dòng đầu html
? TĂng tốc download
? Sử dụng server trung gian để check busy của subscription
? Check lại Procedure có xổ ra mà vẫn không lấy link được
? Làm 3 tabs: Scan, DTC, Procedure
? Làm 1 tab để tải hình bị crash, add nhiều hình
#ce ----------------------------------------------------------------------------

#include-once

#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <GuiEdit.au3>
#include <GuiComboBox.au3>
#include <Clipboard.au3>
#include <Timers.au3>


#include "General_Library.au3"
#include "Create_JAVASCRIPT_Procedure.au3"
#include "Create_NORMAL_Procedure.au3"
#include "Create_DTC.au3"
#include "Get_All_DTC_Links.au3"
#include "Add_Procedure.au3"
#include "Permission.au3"



Global $bWeb_Attach = 0
Global $bWeb_Visible = 1
Global $bWeb_Wait = 1
Global $bWeb_TakeFocus = 0
Global $bImage_Download = 0

Global $iRemaining_Hours = 0

Global $iSubscription_Num = 2
Global $sLink_YMME
Global $sLink_DTC



Global $bDTC_Flag = False
Global $bScan_Flag = False
Global $bWriteConfig_Flag = False
Global $bFind_Flag = False
Global $bAdd_Flag = False
Global $bAdd_Allow_Flag = False
Global $bKeySubmit_Flag = False
Global $bTotal_Allow_Flag = False

Func Close_All_IE()
   $Proc = "iexplore.exe"
   While ProcessExists($Proc)
      ProcessClose($Proc)
   Wend
EndFunc



Func Set_DTC_Flag ()
   $bDTC_Flag = True
EndFunc

Func Set_Scan_Flag ()
   $bScan_Flag = True
EndFunc

Func Set_WriteConfig_Flag ()
   $bWriteConfig_Flag = True
EndFunc

Func Set_Find_Flag ()
   $bFind_Flag = True
EndFunc

Func Set_Add_Flag ()
   $bAdd_Flag = True
EndFunc

Func Set_KeySubmit_Flag ()
   $bKeySubmit_Flag = True
EndFunc

Func Autoit_Exit ()
   ;Close_All_IE()
   __IELockSetForegroundWindow($LSFW_UNLOCK)
   Exit
EndFunc

;Close_All_IE()

#Region ### START GUI section ### Form=
   Opt("GUIOnEventMode", 1)
   $Form1 = GUICreate("Prepair Procedure Generator", 329, 460, -1, -1)
   GUISetOnEvent($GUI_EVENT_CLOSE, "Autoit_Exit")
   GUISetBkColor(0xFFFFFF)
   ;-------------------------------------------
   ;CREATE GUI BASE TAB
   GUICtrlCreateTab(10, 10, 309, 120)
	  ;-------------------------------------------
	  ;CREATE PERMISSION TAB
	  GUICtrlCreateTabItem("Permission")
		 ;-------------------------------------------
		 ;Create input
		 $Input_Key = GUICtrlCreateInput("", 32, 96, 180, 23)
		 $Key_Warning = GUICtrlCreateLabel("Để tránh việc sử dụng app cho mục đích cá nhân. Bạn phải nhập vào một KEY được lấy từ Thuận Huỳnh. Mỗi KEY được sử dụng trong vòng 12 giờ kể từ lần đầu nhập", 32, 35, 260, 60)
		 ;-------------------------------------------
		 ;Create buttons and set on event
		 $Button_Submit = GUICtrlCreateButton("Submit", 224, 95, 75, 25)
		 GUICtrlSetOnEvent($Button_Submit, "Set_KeySubmit_Flag")
	  ;-------------------------------------------
	  ;CREATE SINGLE DTC TAB
	  GUICtrlCreateTabItem("Single DTC")
		 ;-------------------------------------------
		 ;Create input
		 $Input_DTC_Link = GUICtrlCreateInput("", 32, 60, 265, 21)
		 GUICtrlCreateLabel("Input DTC Link ", 128, 40, 114, 17)
		 ;-------------------------------------------
		 ;Create buttons and set on event
		 $Button_Begin_1 = GUICtrlCreateButton("Begin", 130, 90, 75, 25)
		 GUICtrlSetOnEvent($Button_Begin_1, "Set_DTC_Flag")
	  ;-------------------------------------------
	  ;CREATE SCAN DTC TAB
	  GUICtrlCreateTabItem("Scan DTCs")
		 ;-------------------------------------------
		 ;Create input
		 $Input_YMME_Link = GUICtrlCreateInput("", 32, 60, 265, 21)
		 GUICtrlCreateLabel("Input Vehicle Link ", 120, 40, 114, 17)
		 ;-------------------------------------------
		 ;Create buttons and set on event
		 $Button_Begin_2 = GUICtrlCreateButton("Begin", 80, 90, 75, 25)
		 GUICtrlSetOnEvent($Button_Begin_2, "Set_Scan_Flag")
		 $Button_WriteConfig = GUICtrlCreateButton("Write Config", 170, 90, 75, 25)
		 GUICtrlSetOnEvent($Button_WriteConfig, "Set_WriteConfig_Flag")
	  ;-------------------------------------------
	  ;CREATE PROCEDURE TAB
	  GUICtrlCreateTabItem("Add Procedure")
		 ;-------------------------------------------
		 ;Create input DTC Path
		 $Input_DTC_Path = GUICtrlCreateInput("", 90, 40, 210, 20)
		 GUICtrlCreateLabel("DTC Path", 30, 42, 114, 17)
		 ;-------------------------------------------
		 ;Create button Find PROCEDURE and set on event
		 $Button_Find = GUICtrlCreateButton("Find Proc", 30, 70, 50, 22)
		 GUICtrlSetOnEvent($Button_Find, "Set_Find_Flag")
		 ;-------------------------------------------
		 ;Select PROCEDURE
		 $Combo_Procedure = GUICtrlCreateCombo("(Select Procedure)", 90, 70, 210, 20, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
		 ;-------------------------------------------
		 ;Create input PROCEDURE link
		 $Input_Procedure_Link = GUICtrlCreateInput("", 90, 100, 150, 20)
		 GUICtrlCreateLabel("Proc Link", 32, 102, 114, 17)
		 ;-------------------------------------------
		 ;Create button ADD and set on event
		 $Button_Add = GUICtrlCreateButton("Add", 250, 99, 50, 22)
		 GUICtrlSetOnEvent($Button_Add, "Set_Add_Flag")
   GUICtrlCreateTabItem("") ; end tabitem definition
   ;-------------------------------------------
   ;CREATE SETTINGS GROUP
   GUICtrlCreateGroup("Settings", 10, 140, 309, 70)
	  ;Select IE visible or invisible
	  $Radio_Visible = GUICtrlCreateRadio("Web Visible", 32, 160, 113, 17)
	  $Radio_Invisible = GUICtrlCreateRadio("Web Invisible", 32, 180, 113, 17)
	  GUICtrlSetState ($Radio_Invisible, $GUI_CHECKED)
	  ;Select subscription
	  $Combo_Subscription = GUICtrlCreateCombo("(License #)", 152, 175, 145, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
	  GUICtrlSetData(-1, "# 1|# 2|# 3|# 4|# 5")
	  GUICtrlCreateLabel("Select License", 180, 155, 120, 17)
   GUICtrlCreateGroup("", -99, -99, 1, 1) ;close group
   ;-------------------------------------------
   ;CREATE GUI NOTIFICATION PLACE
	  $Commu_Ctrl = GUICtrlCreateEdit("", 10, 240, 309, 180)
	  GUICtrlSetBkColor(-1, 0xF0F0F0)
	  _GUICtrlEdit_SetReadOnly ($Commu_Ctrl, True)
	  _GUICtrlEdit_SetMargins ($Commu_Ctrl, BitOR($EC_LEFTMARGIN, $EC_RIGHTMARGIN), 7, 7)
	  GUICtrlCreateLabel("Communication", 10, 220, 80, 17)
	  ;Button clear notif
	  $Button_Clear_1 = GUICtrlCreateButton("Clear Notif", 130, 430, 75, 25)
	  GUICtrlSetOnEvent($Button_Clear_1, "Notification_Clear")


   ;-------------------------------------------
   ;SHOW GUI
   GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###



;====================================================================================================================
;                  DESCRIPTION: WAIT FOR VALID KEY
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================


$bTotal_Allow_Flag = True
$iRemaining_Hours = 12

Func Wait_Key ()
   $bTotal_Allow_Flag = False
   GUI_Input_Ctrl_Set_State ("Disable")
   GUICtrlSetData ($Key_Warning, "Để tránh việc sử dụng app cho mục đích cá nhân. Bạn phải nhập vào một KEY được lấy từ Thuận Huỳnh. Mỗi KEY được sử dụng trong vòng 12 giờ kể từ lần đầu nhập")

   MsgBox (0, "Message", "Please enter a VALID KEY to continue using the app!")
   While 1
	  If $bKeySubmit_Flag = True Then KeySubmit ()
	  If $bTotal_Allow_Flag = True Then
		 GUI_Input_Ctrl_Set_State ("Enable")
		 ExitLoop
	  EndIf
   WEnd
EndFunc




;====================================================================================================================
;                  FUNCTION DESCRIPTION: WAIT FOR BUTTON PRESSED AND COUNT TIME OUT
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================

Local $hStarttime  = _Timer_Init ()
Local $iSec_Count = 0
Local $iSec = 0
Local $iSec_Max = $iRemaining_Hours*3600
While 1
   ;NẾU HẾT THỜI GIAN SỬ DỤNG THÌ VÀO FUNCTION CHỜ KEY
   If $iSec >= $iSec_Max Then
	  Wait_Key ()
	  $iSec = 0
	  $iSec_Max = $iRemaining_Hours*3600
	  $hStarttime  = _Timer_Init ()
   ;NẾU CÒN THỜI GIAN SỬ DỤNG THÌ CHO PHÉP CHỜ NÚT NHẤN VÀ LẤY GIÁ TRỊ ĐẾM THỜI GIAN
   Else
	  $iSec = Round(_Timer_Diff($hStarttime)/1000, 0)
	  If $bDTC_Flag = True Then DTC_Generation_Begin()
	  If $bWriteConfig_Flag = True Then Scan_DTC_Write_Config()
	  If $bScan_Flag = True Then Scan_DTC_Begin ()
	  If $bFind_Flag = True Then Find_Procedure ()
	  If $bAdd_Flag = True Then Add_Procedure_Begin ()
   EndIf
WEnd



;====================================================================================================================
;                  FUNCTION DESCRIPTION: WRITE CONFIG FILE FOR THE FUNCTION SCAN DTCs
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func Add_Procedure_Begin ()
   If $bAdd_Allow_Flag = True Then
	  ;-------------------------------------------------
	  ;IF SUBSCRIPTION IS SELECTED AND THE LINK IS VALID => EXECUTE
	  If Check_Init_Vals ("Add Procedure") = True Then
		 If StringInStr (GUICtrlRead ($Combo_Procedure), "PROCEDURE_") <> 0 Then
			Notification ("Please wait ...", "Normal")
			GUI_Input_Ctrl_Set_State ("Disable")
			   Local $oIE = Add_Procedure ()
			   _IEQuit ($oIE)
			GUI_Input_Ctrl_Set_State ("Enable")
		 Else
			Notification ("Please select one PROCEDURE", "Normal")
		 EndIf
	  EndIf
   Else
	  Notification ("Please press button FIND PROCEDURE first", "Normal")
   EndIf
   $bAdd_Allow_Flag = False
   $bAdd_Flag = False
EndFunc




;====================================================================================================================
;                  FUNCTION DESCRIPTION: WRITE CONFIG FILE FOR THE FUNCTION SCAN DTCs
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func Scan_DTC_Begin ()
   ;-------------------------------------------------
   ;IF SUBSCRIPTION IS SELECTED AND THE LINK IS VALID => EXECUTE
   If Check_Init_Vals ("Scan DTCs") = True Then
	  Notification ("Please wait ...", "Normal")
	  GUI_Input_Ctrl_Set_State ("Disable")
		 Local $oIE = Scan_DTCs ()
		 _IEQuit ($oIE)
	  GUI_Input_Ctrl_Set_State ("Enable")
   EndIf
   $bScan_Flag = False
EndFunc




;====================================================================================================================
;                  FUNCTION DESCRIPTION: WRITE CONFIG FILE FOR THE FUNCTION SCAN DTCs
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func Scan_DTC_Write_Config ()
   Close_All_IE()
   ;-------------------------------------------------
   ;IF SUBSCRIPTION IS SELECTED AND THE LINK IS VALID => EXECUTE
   If Check_Init_Vals ("Scan DTCs") = True Then
	  Notification ("Please wait ...", "Normal")
	  GUI_Input_Ctrl_Set_State ("Disable")
			Local $oIE = Write_Config ()
			_IEQuit ($oIE)
	  GUI_Input_Ctrl_Set_State ("Enable")
   EndIf
   $bWriteConfig_Flag = False
EndFunc



;====================================================================================================================
;                  FUNCTION DESCRIPTION: GENERATE DTC PROCEDURE WHEN USER PRESS "BEGIN" ON TAB "SINGLE DTC"
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func DTC_Generation_Begin ()
   ;-------------------------------------------------
   ;IF SUBSCRIPTION IS SELECTED AND THE LINK IS VALID => EXECUTE
   If Check_Init_Vals ("Single DTC") = True Then
	  Notification ("Please wait ...", "Normal")
	  GUI_Input_Ctrl_Set_State ("Disable")
			Local $oIE = Main_function_DTC ()
			_IEQuit ($oIE)
	  GUI_Input_Ctrl_Set_State ("Enable")
   EndIf
   $bDTC_Flag = False
EndFunc



;====================================================================================================================
;                  FUNCTION DESCRIPTION: CHECK THE PROPER INITIIALS
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func Check_Init_Vals ($sWhich_Tab)
   ;-------------------------------------------------
   ;CHECK WEB SHOW/HIDE
   If GUICtrlRead ($Radio_Visible) = $GUI_CHECKED Then $bWeb_Visible = 1
   If GUICtrlRead ($Radio_Invisible) = $GUI_CHECKED Then $bWeb_Visible = 0
   ;-------------------------------------------------
   ;CHECK SUBSCRIPTION NUMBER
   Local $sCombo_Subscription_Val = GUICtrlRead ($Combo_Subscription)
   Local $bSub_Flag = False
   Switch $sCombo_Subscription_Val
	  Case  "(License #)"
		 Notification ("Please select your LISENSE NUMBER", "Normal")
		 $bSub_Flag = False
	  Case  "# 1"
		 $iSubscription_Num = 1
		 $bSub_Flag = True
	  Case  "# 2"
		 $iSubscription_Num = 2
		 $bSub_Flag = True
	  Case  "# 3"
		 $iSubscription_Num = 3
		 $bSub_Flag = True
	  Case  "# 4"
		 $iSubscription_Num = 4
		 $bSub_Flag = True
	  Case  "# 5"
		 $iSubscription_Num = 5
		 $bSub_Flag = True
	  Case Else
		 Notification ("Please select your LISENSE NUMBER", "Normal")
		 $bSub_Flag = False
   EndSwitch
   ;-------------------------------------------------
   ;SELECT LINK
   If $sWhich_Tab = "Single DTC" Then
	  $sLink_DTC = GUICtrlRead ($Input_DTC_Link)
	  Local $sLink = $sLink_DTC
   Elseif  $sWhich_Tab = "Scan DTCs" Then
	  $sLink_YMME = GUICtrlRead ($Input_YMME_Link)
	  Local $sLink = $sLink_YMME
   Else
	  Local $sLink = GUICtrlRead ($Input_Procedure_Link)
   EndIf
   ;-------------------------------------------------
   ;CHECK LINK
   Local $bLink_Flag = False
   If StringLeft ($sLink, 26) <> "http://repair.alldata.com/" And StringLeft ($sLink, 26) <> "https://repair.alldata.com"  Then
	  Notification ("The link is INVALID. Please input VALID link", "Normal")
	  $bLink_Flag = False
   Else
	  $bLink_Flag = True
   EndIf
   ;-------------------------------------------------
   ;SET FLAG
   If ($bSub_Flag = True) And ($bLink_Flag = True) Then
	  Local $bBegin_Flag = True
   Else
	  Local $bBegin_Flag = False
   EndIf
   Return $bBegin_Flag
EndFunc


















;====================================================================================================================
;                  FUNCTION DESCRIPTION: ENABLE OR DISABLE INPUT ELEMENTS
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func GUI_Input_Ctrl_Set_State ($sState)
   If $sState = "Disable" Then
	  GUICtrlSetState ($Button_Begin_1, $GUI_DISABLE)
	  GUICtrlSetState ($Radio_Visible, $GUI_DISABLE)
	  GUICtrlSetState ($Radio_Invisible, $GUI_DISABLE)
	  GUICtrlSetState ($Combo_Subscription, $GUI_DISABLE)
	  GUICtrlSetState ($Input_DTC_Link, $GUI_DISABLE)
	  GUICtrlSetState ($Input_YMME_Link, $GUI_DISABLE)
	  GUICtrlSetState ($Button_Begin_2, $GUI_DISABLE)
	  GUICtrlSetState ($Button_WriteConfig, $GUI_DISABLE)
	  GUICtrlSetState ($Input_DTC_Path, $GUI_DISABLE)
	  GUICtrlSetState ($Button_Find, $GUI_DISABLE)
	  GUICtrlSetState ($Combo_Procedure, $GUI_DISABLE)
	  GUICtrlSetState ($Input_Procedure_Link, $GUI_DISABLE)
	  GUICtrlSetState ($Button_Add, $GUI_DISABLE)
   Else
	  GUICtrlSetState ($Button_Begin_1, $GUI_ENABLE)
	  GUICtrlSetState ($Radio_Visible, $GUI_ENABLE)
	  GUICtrlSetState ($Radio_Invisible, $GUI_ENABLE)
	  GUICtrlSetState ($Combo_Subscription, $GUI_ENABLE)
	  GUICtrlSetState ($Input_DTC_Link, $GUI_ENABLE)
	  GUICtrlSetState ($Input_YMME_Link, $GUI_ENABLE)
	  GUICtrlSetState ($Button_Begin_2, $GUI_ENABLE)
	  GUICtrlSetState ($Button_WriteConfig, $GUI_ENABLE)
	  GUICtrlSetState ($Input_DTC_Path, $GUI_ENABLE)
	  GUICtrlSetState ($Button_Find, $GUI_ENABLE)
	  GUICtrlSetState ($Combo_Procedure, $GUI_ENABLE)
	  GUICtrlSetState ($Input_Procedure_Link, $GUI_ENABLE)
	  GUICtrlSetState ($Button_Add, $GUI_ENABLE)
   EndIf
EndFunc


;====================================================================================================================
;                  FUNCTION DESCRIPTION: WRITE TO NOTIFICATION SCREEN
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func Notification ($sNoti, $sMode) ;Mode = "Normal", "Replace Previous"
   Static Local $sPrevious_Noti
   If $sMode = "Normal" Then
	  _GUICtrlEdit_AppendText ($Commu_Ctrl,  @CRLF & $sNoti & @CRLF)
   Else
	  _GUICtrlEdit_SetReadOnly ( $Commu_Ctrl, False )
	  _GUICtrlEdit_Undo ($Commu_Ctrl)
	  _GUICtrlEdit_SetReadOnly ( $Commu_Ctrl, True )
	  _GUICtrlEdit_AppendText ($Commu_Ctrl,  @CRLF & $sNoti & @CRLF)
   EndIf
EndFunc


;====================================================================================================================
;                  FUNCTION DESCRIPTION: CLEAR NOTIFICATION SCREEN
;				   INPUT               :
;                  OUTPUT              :
;====================================================================================================================
Func Notification_Clear ()
   GUICtrlSetData ($Commu_Ctrl, "")
EndFunc
