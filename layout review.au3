#include <Excel.au3>
#include <Array.au3>
#include <GuiButton.au3>


#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GuiButton.au3>


Local $sFilePath1 = @ScriptDir & "\順序.xls" ;This file should already exist
Local $oAppl = _Excel_Open(False)
Local $oExcel = _Excel_BookOpen($oAppl, $sFilePath1, True)
Const $len = 6
Const $row = $oExcel.Application.WorksheetFunction.CountA ($oExcel.ActiveSheet.Columns(1))
Local $aArray1[$len]
For $i = 1 to $len
	$aArray1[$i-1] = _Excel_RangeRead($oExcel, Default, $oExcel.Sheets(1).Range($oExcel.Sheets(1).cells(2,$i),$oExcel.Sheets(1).cells($row,$i)),1,True)
Next

Local $type = $oExcel.Sheets(1).cells(1,1).value

_Excel_BookClose($oExcel,0)
_Excel_Close($oAppl,0)
$oExcel = 0 ; delete object

#Region ### START Koda GUI section ### Form=d:\sun\work\software\mycode\autoit\layout\review\form1.kxf
$Form1 = GUICreate("Layout Review", 241, 121, 0, 0, -1, BitOR($WS_EX_TOPMOST,$WS_EX_WINDOWEDGE))
$Run = GUICtrlCreateButton("Run(F5)", 120, 80, 120, 40)
GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
$Previous = GUICtrlCreateButton("Previous", 0, 80, 120, 20)
GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
$Message = GUICtrlCreateEdit("", 0, 0, 240, 80, BitOR($ES_CENTER,$ES_WANTRETURN))
GUICtrlSetData(-1, "Message")
GUICtrlSetFont(-1, 14, 400, 0, "MS Sans Serif")
GUICtrlSetState(-1, $GUI_DISABLE)
$next = GUICtrlCreateButton("next", 0, 100, 120, 20)
GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###


If Not WinExists("[REGEXPTITLE:PADS Layout]") Then
	MsgBox(0,"PADS 未開啟", "請確認是否已開啟 PADS")
	Exit
EndIf

Global $winname = "[REGEXPTITLE:(?i)" & $type & ".* - PADS Layout]"
While Not WinExists($winname)
	$type = InputBox("無法連結 PADS", "請輸入正確 Type", "C")
	$winname = "[REGEXPTITLE:(?i)" & $type & ".* - PADS Layout]"
WEnd

; different pads has different intance, like 7 or 4
Global $pads_win
If ControlGetHandle ( $winname, "", "[CLASS:AfxFrameOrView90; INSTANCE:7]" ) <> 0 Then
	$pads_win = "[CLASS:AfxFrameOrView90; INSTANCE:7]"
ElseIf ControlGetHandle ( $winname, "", "[CLASS:AfxFrameOrView90; INSTANCE:4]" ) <> 0 Then
	$pads_win = "[CLASS:AfxFrameOrView90; INSTANCE:4]"
EndIf

WinSetTitle($Form1, "", $type & " Layout Review")



Local $n = $aArray1[1]
Local $l = $aArray1[2]
Local $lo = $aArray1[3]
Local $m = $aArray1[4]
Local $o_cmd = $aArray1[5]
Local $preArray[0][2]
Local $start = 0

HotKeySet ( "{F5}", "HotKeyPressed" )
HotKeySet ( "{F2}", "HotKeyPressed" )
_GUICtrlButton_Enable($Previous, False)

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
		Case $Previous
			_GUICtrlButton_Enable($Run, True)
			_GUICtrlButton_Enable($next, True)
			$lastRow = UBound($preArray) - 1
			If $lastRow >= 0 Then
				$start = $preArray[$lastRow][0]
				GUICtrlSetData($Message, "回到" & @CRLF & $preArray[$lastRow][1])
				_ArrayDelete($preArray, $lastRow)
			Else
				$start = 0
				GUICtrlSetData($Message, "過去" & @CRLF & "已經回不去了")
				_GUICtrlButton_Enable($Previous, False)
			EndIf
		Case $next
			_GUICtrlButton_Enable($Previous, True)
			Local $ele = $start & "|"
			For $i = $start to UBound($aArray1[0],1)-1
				If($m[$i] <> "") Then
					$ele &= $m[$i]

					_ArrayAdd($preArray, $ele)
					$start = $i + 1
					ExitLoop
				EndIf
			Next
			If $start == UBound($aArray1[0],1) Then
				GUICtrlSetData($Message, "未來" & @CRLF & "仍在施工中")
				_GUICtrlButton_Enable($Run, False)
				_GUICtrlButton_Enable($next, False)
			Else
				For $i = $start to UBound($aArray1[0],1)-1
					If($m[$i] <> "") Then
						GUICtrlSetData($Message, "跳到" & @CRLF & $m[$i])
						ExitLoop
					EndIf
				Next
			EndIf
		Case $Run
			_GUICtrlButton_Enable($Previous, True)
			GUISetState(@SW_HIDE, $Form1)
			SplashTextOn("Warning", "執行中" & @CRLF & "放開那隻老鼠！", 270, 80, -1, 80, -1, -1, 20)
			WinActivate($winname)
			WinWaitActive($winname)
			show_net("")
			If $start == UBound($aArray1[0],1) Then
				show_layer(1)
				extral_cmd("spd")
				GUICtrlSetData($Message, "檢查鋪銅" & @CRLF & "所有層數")
				_GUICtrlButton_Enable($Run, False)
				_GUICtrlButton_Enable($next, False)
			Else
				Local $ele = $start & "|"
				For $i = $start to UBound($aArray1[0],1)-1
					show_net($n[$i])
					If($m[$i] <> "") Then
						show_layer($l[$i] & " " & $lo[$i])
						extral_cmd($o_cmd[$i])
						If StringRegExp ($o_cmd[$i], 'ss ' ) == 0 Then
							ControlSend($winname,"",$pads_win,"{Home}")
						EndIf
						GUICtrlSetData($Message, $m[$i] & @CRLF & "所在層數：" & $l[$i] & @CRLF & "上下層數：" & $lo[$i])
						$ele &= $m[$i]

						_ArrayAdd($preArray, $ele)
						$start = $i + 1
						ExitLoop
					EndIf
				Next
			EndIf
			SplashOff()
			GUISetState(@SW_SHOW, $Form1)
			ControlFocus ( $winname, "", $pads_win )

	EndSwitch
WEnd



Func extral_cmd($cmds)
	Local $cmd = StringSplit($cmds,",")
	For $i = 1 To $cmd[0]
		command(StringStripWS ($cmd[$i],1))
	Next
EndFunc

Func show_net($net)
	command("n " & $net)
EndFunc

Func show_layer($layer)
	command("z " & $layer)
EndFunc

Func command($command)
	Do
		ControlSend($winname,"",$pads_win, "s")
	Until WinWait("Modeless Command","",1000) <> 0
	ControlSetText("Modeless Command","","[CLASS:Edit; INSTANCE:1]",$command)
	ControlSend("Modeless Command","","[CLASS:Edit; INSTANCE:1]","{Enter}")
	Do
	Until WinExists("Modeless Command") == 0
EndFunc

Func HotKeyPressed()
    Switch @HotKeyPressed
		Case "{F5}"
			GUISetState(@SW_SHOW, $Form1)
			WinActivate($Form1)
			WinWaitActive($Form1)
            ControlClick($Form1, "", $Run, "left")
		Case "{F2}"
			ConsoleWrite(WinGetState($Form1))
			If BitAND(WinGetState($Form1), 2) Then
				GUISetState(@SW_HIDE, $Form1)
			Else
				GUISetState(@SW_SHOW, $Form1)
			EndIf
    EndSwitch
EndFunc   ;==>HotKeyPressed

