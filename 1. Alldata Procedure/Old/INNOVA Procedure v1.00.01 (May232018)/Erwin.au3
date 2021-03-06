#include <MsgBoxConstants.au3>
#include <FileConstants.au3>
#include <InetConstants.au3>

#include <Clipboard.au3>
#include <IE.au3>
#include <Excel.au3>
#include <WinAPIFiles.au3>




Erwin()

Func Erwin ()
   ;----------------------------------------------------------
   ;MỞ TRANG VÀ LẤY OBJECT
   Local $oIE = _IECreate ("https://erwin.audiusa.com/erwin/performSearchArticleSimple.do")
   ;----------------------------------------------------------
   ;KHAI BÁO CÁC BIẾN ĐẾM CHO VÒNG LẶP
   Local $iCount_Year, $iCount_Model, $iCount_Cate1, $iCount_Cate2
   ;----------------------------------------------------------
   ;ĐOẠN CODE CHỌN YEAR
   ;Lấy các option year
   Local $oSelect_Year = _IEGetObjById ($oIE, "f_modelYear")
   Local $aYear_Options = Get_Option ($oSelect_Year)
   ;Vòng lặp chọn year
   $iCount_Year = 0
   While $aYear_Options [$iCount_Year] <> ""
	  ;Chọn year
	  $oSelect_Year = _IEGetObjById ($oIE, "f_modelYear")
	  _IEFormElementOptionSelect ($oSelect_Year, $aYear_Options [$iCount_Year])
	  _IELoadWait ($oIE)
	  ;----------------------------------------------------------
	  ;ĐOẠN CODE CHỌN MODEL
	  ;Lấy các option model
	  Local $oSelect_Model = _IEGetObjById ($oIE, "f_cartypeId")
	  Local $aModel_Options = Get_Option ($oSelect_Model)
	  ;Vòng lặp chọn model
	  $iCount_Model = 0
	  While $aModel_Options [$iCount_Model] <> ""
		 ;Chọn model
		 $oSelect_Model = _IEGetObjById ($oIE, "f_cartypeId")
		 _IEFormElementOptionSelect ($oSelect_Model, $aModel_Options [$iCount_Model])
		 _IELoadWait ($oIE)
		 ;----------------------------------------------------------
		 ;ĐOẠN CODE CHỌN CATEGORY 1
		 ;Lấy các option category 1
		 Local $oSelect_Cate1 = _IEGetObjByName ($oIE, "mainTopicCode")
		 Local $aCate1_Options = Get_Option ($oSelect_Cate1)
		 ;Vòng lặp chọn category 1
		 $iCount_Cate1 = 0
		 While $aCate1_Options [$iCount_Cate1] <> ""
			;Chọn category 1
			$oSelect_Cate1 = _IEGetObjByName ($oIE, "mainTopicCode")
			_IEFormElementOptionSelect ($oSelect_Cate1, $aCate1_Options [$iCount_Cate1])
			_IELoadWait ($oIE)
			;----------------------------------------------------------
			;ĐOẠN CODE CHỌN CATEGORY 2
			;Lấy các option category 2
			Local $oSelect_Cate2 = _IEGetObjByName ($oIE, "topicCode")
			If @error <> 7 Then ;No match
			   Local $aCate2_Options = Get_Option ($oSelect_Cate2)
			   ;Vòng lặp chọn category 2
			   $iCount_Cate2 = 0
			   While $aCate2_Options [$iCount_Cate2] <> ""
				  If $aCate2_Options [0] <> "" Then
					 ;Chọn category 2
					 $oSelect_Cate2 = _IEGetObjByName ($oIE, "topicCode")
					 _IEFormElementOptionSelect ($oSelect_Cate2, $aCate2_Options [$iCount_Cate2])
					 _IELoadWait ($oIE)
					 $iCount_Cate2 = $iCount_Cate2 + 1
					 Sleep (2000)
				  EndIf
			   WEnd
			EndIf
			$iCount_Cate1 = $iCount_Cate1 + 1
			Sleep (2000)
		 WEnd
		 $iCount_Model = $iCount_Model + 1
	  WEnd
	  $iCount_Year = $iCount_Year + 1
   WEnd

EndFunc





Func Get_Option ($oSelect)
   ;------------------------------------------------------------------------------------------------------------------
   ;ĐOẠN CODE LẤY TEXT VÀ LINK PROCEDURE PARTS TRONG TAG <A>
   Local $oOptions = _IETagNameGetCollection($oSelect, "option")
   Local $Text = ""
   Local $aOptions [1000]
   Local $iCount = 0
   For $oOption In $oOptions
	  If $oOption.value <> "" Then
		 $aOptions [$iCount] = $oOption.value
		 $iCount = $iCount + 1
	  EndIf
   Next
   Return $aOptions
EndFunc