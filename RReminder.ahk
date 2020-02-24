#NoEnv ; recommended for performance and compatibility with future AutoHotkey releases
#Persistent ; keep the script permanently running
#SingleInstance force ; allow only one instance of this script to be running
#Warn ; enable warnings to assist with detecting common errors
SendMode Input ; recommended for new scripts due to its superior speed and reliability
SetWorkingDir %A_ScriptDir% ; ensure a consistent starting directory
FileEncoding, UTF-8 ; set the default encoding for FileRead, FileReadLine, Loop Read, FileAppend, and FileOpen

ScriptName := "RReminder"
ScriptVersion := "1.3.0.0"
CopyrightNotice := "Copyright (c) 2020 Chaohe Shi"

ConfigDir := A_AppData . "\" . ScriptName
ConfigFile := ConfigDir . "\" . ScriptName . ".ini"
ReminderTextFile := ConfigDir . "\" . ScriptName . ".txt"

TEXT_Reminder := "Reminder"
TEXT_ReminderText := "The time now is "
TEXT_Period := "."
TEXT_Settings := "Settings"
TEXT_TEXT := "Reminder Text"
TEXT_Style := "Reminder Style"
TEXT_Test := "Test"
TEXT_Preview := "Preview"
TEXT_OK := "OK"
TEXT_ErrorMsg := "Not a valid command!"
TEXT_ConfirmMsg := "Confirm before run button action"
TEXT_RunAction := "Run Action"
TEXT_RunActionMsg := "You are about to run "
TEXT_About := "About"
TEXT_Exit := "Exit"
TEXT_AboutMsg := ScriptName . " " . ScriptVersion . "`n`n" . CopyrightNotice

ReminderInterval := 30 ; this number needs to be one of the factors of 60min. e.g. when reminderInterval is 15, reminder occurs at 0, 15, 30, 45min of every hour
MsgBoxTimeout := 60
MsgBoxOption := 0 ; use 0, 1, or 3 to for 1, 2, or 3 buttons and enable the close button

FileRead, ReminderText, %ReminderTextFile%
IniRead, NumButton, %ConfigFile%, General, NumButton, 1
IniRead, Button1, %ConfigFile%, Button, Button1, OK
IniRead, Button2, %ConfigFile%, Button, Button2, Work
IniRead, Button3, %ConfigFile%, Button, Button3, Relax
IniRead, Action1, %ConfigFile%, Action, Action1, Notepad.exe
IniRead, Action2, %ConfigFile%, Action, Action2, https://www.google.com/
IniRead, Action3, %ConfigFile%, Action, Action3, ::{20d04fe0-3aea-1069-a2d8-08002b30309d}
IniRead, NeedConfirm, %ConfigFile%, Confirm, NeedConfirm, 1
Gosub, UpdateMsgBoxOption

;Menu, Tray, NoStandard ; remove the standard menu items
Menu, Tray, Tip, %ScriptName% ; change the tray icon's tooltip
Menu, Tray, Add, %TEXT_Reminder%, ShowReminder
Menu, Tray, Default, %TEXT_Reminder%
Menu, Tray, Add
Menu, Tray, Add, %TEXT_Settings%, ShowSettingsGUI
Menu, Tray, Add
Menu, Tray, Add, %TEXT_About%, ShowAboutMsg
Menu, Tray, Add
Menu, Tray, Add, %TEXT_Exit%, ExitProgram

SetTimer, Reminder, % GetTimerPeriod()

Return ; end of the auto-execute section

ShowSettingsGUI:
Process, Exist
DetectHiddenWindows, On
if WinExist(TEXT_Settings . " ahk_class AutoHotkeyGUI ahk_pid " . ErrorLevel) ; if the settings GUI already exists
{
	WinShow ; show the GUI window if it is hidden
	WinActivate
}
else ; else display the settings GUI
{
	Gui, New, +HwndGuiHwnd, %TEXT_Settings%
	Gui, %GuiHwnd%:Default

	Gui, Add, Tab3, , %TEXT_TEXT%|%TEXT_Style%

	Gui, Tab, 1
	Gui, Add, Edit, vReminderText gSettings r14 w574, %ReminderText%

	Gui, Tab, 2
	Gui, Add, Text, Section, Number of button(s)

	Position := NumButton + 1
	Gui, Add, DropDownList, vNumButton gSettings Choose%Position% w75, 0|1|2|3

	Gui, Add, Text, Section ys, Button1 Text
	Gui, Add, Edit, vButton1 gSettings r1 w75 Limit, %Button1%

	Gui, Add, Text, ys, Button1 Action
	Gui, Add, Edit, vAction1 gSettings r1 w300, %Action1%

	Gui, Add, Text, ys
	Gui, Add, Button, vTest1 gAction1 w75, %TEXT_Test%

	Gui, Add, Text, Section xs, Button2 Text
	Gui, Add, Edit, vButton2 gSettings r1 w75 Limit, %Button2%

	Gui, Add, Text, ys, Button2 Action
	Gui, Add, Edit, vAction2 gSettings r1 w300, %Action2%

	Gui, Add, Text, ys
	Gui, Add, Button, vTest2 gAction2 w75, %TEXT_Test%

	Gui, Add, Text, Section xs, Button3 Text
	Gui, Add, Edit, vButton3 gSettings r1 w75 Limit, %Button3%

	Gui, Add, Text, ys, Button3 Action
	Gui, Add, Edit, vAction3 gSettings r1 w300, %Action3%

	Gui, Add, Text, ys
	Gui, Add, Button, vTest3 gAction3 w75, %TEXT_Test%

	Gui, Add, Text
	Gui, Add, Checkbox, xs Checked%NeedConfirm% vNeedConfirm gSettings, %TEXT_ConfirmMsg%

	Gui, Tab
	Gui, Add, Button, Section xm vPreview gShowReminder w75 h23, %TEXT_Preview%
	Gui, Add, Button, vOK gOK ys w75 h23, %TEXT_OK%
	GuiControl, +Default, OK
	GuiControl, Focus, OK
	Gosub, UpdateGUI
	Gui, Show
}
Return

Action1:
Try
{
	Run, %Action1%
}
Catch
{
	ToolTip, %TEXT_ErrorMsg%
	SetTimer, RemoveToolTip, -3000
}
Return

Action2:
Try
{
	Run, %Action2%
}
Catch
{
	ToolTip, %TEXT_ErrorMsg%
	SetTimer, RemoveToolTip, -3000
}
Return

Action3:
Try
{
	Run, %Action3%
}
Catch
{
	ToolTip, %TEXT_ErrorMsg%
	SetTimer, RemoveToolTip, -3000
}
Return

Settings:
Gui, Submit, NoHide
Gosub, UpdateMsgBoxOption
Gosub, UpdateGUI
Gosub, EnsureConfigDirExists
FileDelete, %ReminderTextFile%
FileAppend, %ReminderText%, %ReminderTextFile%
IniWrite, %NumButton%, %ConfigFile%, General, NumButton
IniWrite, %Button1%, %ConfigFile%, Button, Button1
IniWrite, %Button2%, %ConfigFile%, Button, Button2
IniWrite, %Button3%, %ConfigFile%, Button, Button3
IniWrite, %Action1%, %ConfigFile%, Action, Action1
IniWrite, %Action2%, %ConfigFile%, Action, Action2
IniWrite, %Action3%, %ConfigFile%, Action, Action3
IniWrite, %NeedConfirm%, %ConfigFile%, Confirm, NeedConfirm
Return

OK:
GuiClose:
GuiEscape:
Gui, Submit, NoHide
Gui, Destroy
Return

ShowAboutMsg:
Process, Exist
DetectHiddenWindows, On
if WinExist(TEXT_About . " ahk_class #32770 ahk_pid " . ErrorLevel) ; if the about message already exists
{
	WinShow ; show the message window if it is hidden
	WinActivate
}
else ; else display the about message
{
	MsgBox, 0, %TEXT_About%, %TEXT_AboutMsg%
}
Return

ExitProgram:
ExitApp

UpdateMsgBoxOption:
if (NumButton == 1)
{
	MsgBoxOption := 0 ; OK
}
else if (NumButton == 2)
{
	MsgBoxOption := 1 ; OK/Cancel
}
else
{
	MsgBoxOption := 3 ; Yes/No/Cancel
}
Return

UpdateGUI:
if (NumButton == 0)
{
	GuiControl, Disable, Button1
	GuiControl, Disable, Action1
	GuiControl, Disable, Test1
	GuiControl, Disable, Button2
	GuiControl, Disable, Action2
	GuiControl, Disable, Test2
	GuiControl, Disable, Button3
	GuiControl, Disable, Action3
	GuiControl, Disable, Test3
	GuiControl, Disable, NeedConfirm
}
else if (NumButton == 1)
{
	GuiControl, Enable, Button1
	GuiControl, Enable, Action1
	GuiControl, Enable, Test1
	GuiControl, Disable, Button2
	GuiControl, Disable, Action2
	GuiControl, Disable, Test2
	GuiControl, Disable, Button3
	GuiControl, Disable, Action3
	GuiControl, Disable, Test3
	GuiControl, Enable, NeedConfirm
}
else if (NumButton == 2)
{
	GuiControl, Enable, Button1
	GuiControl, Enable, Action1
	GuiControl, Enable, Test1
	GuiControl, Enable, Button2
	GuiControl, Enable, Action2
	GuiControl, Enable, Test2
	GuiControl, Disable, Button3
	GuiControl, Disable, Action3
	GuiControl, Disable, Test3
	GuiControl, Enable, NeedConfirm
}
else
{
	GuiControl, Enable, Button1
	GuiControl, Enable, Action1
	GuiControl, Enable, Test1
	GuiControl, Enable, Button2
	GuiControl, Enable, Action2
	GuiControl, Enable, Test2
	GuiControl, Enable, Button3
	GuiControl, Enable, Action3
	GuiControl, Enable, Test3
	GuiControl, Enable, NeedConfirm
}
Return

; ensure ConfigDir exists
EnsureConfigDirExists:
if !InStr(FileExist(ConfigDir), "D")
{
	FileCreateDir, %ConfigDir%
}
Return

RemoveToolTip:
ToolTip
Return

Reminder:
if (IsTimeReached()) ; check the time before display the reminder to avoid the accumulated timer precision error https://www.autohotkey.com/docs/commands/SetTimer.htm#Precision
{
	Gosub, ShowReminder
}
SetTimer, Reminder, % GetTimerPeriod()
Return

WM_COMMNOTIFY(wParam)
{
	global ; use assume-global mode to access global variables
	if (wParam == 1027) ; AHK_DIALOG
	{
		Process, Exist
		DetectHiddenWindows, On
		if WinExist(TEXT_Reminder . " ahk_class #32770 ahk_pid " . ErrorLevel)
		{
			ControlSetText, Button1, &%Button1%
			ControlSetText, Button2, &%Button2%
			ControlSetText, Button3, &%Button3%
		}
	}
}

; get the period for the next reminder in millisecond, accurate to one second
GetTimerPeriod()
{
	global ReminderInterval
	Return Mod(3600 - A_Min * 60 - A_Sec, ReminderInterval * 60) * 1000
}

; check if the current minute is the desired minute
IsTimeReached()
{
	global ReminderInterval
	Return Mod(60 - A_Min, ReminderInterval) == 0
}

HideTrayTip()
{
	TrayTip ; attempt to hide the TrayTip in the normal way
	if (SubStr(A_OSVersion, 1, 3) == "10.") ; if the OS version is Windows 10
	{
		; temporarily removing the tray icon to hide the TrayTip
		Menu, Tray, NoIcon
		Sleep, 100
		Menu, Tray, Icon
	}
}

ShowReminder:
Process, Exist
DetectHiddenWindows, On
if WinExist(TEXT_Reminder . " ahk_class #32770 ahk_pid " . ErrorLevel) ; if the message already exists
{
	WinShow ; show the message window if it is hidden
	WinActivate
}
else ; else display the message
{
	if (NumButton == 0)
	{
		TrayTip, %TEXT_Reminder%, % TEXT_ReminderText . A_Hour . ":" . A_Min . TEXT_Period
		SetTimer, HideTrayTip, -60000
	}
	else
	{
		OnMessage(0x44, "WM_COMMNOTIFY") ; https://autohotkey.com/board/topic/56272-msgbox-button-label-change/?p=353457
		MsgBox, % MsgBoxOption, %TEXT_Reminder%, % TEXT_ReminderText . A_Hour . ":" . A_Min . TEXT_Period . ((ReminderText == "") ? "" : "`n`n" . ReminderText), % MsgBoxTimeout
		if (NumButton == 1) ; MsgBoxOption = 0: OK
		{
			IfMsgBox, OK
			{
				if (NeedConfirm && Action1 != "")
				{
					MsgBox, 1, %TEXT_RunAction%, % TEXT_RunActionMsg . Action1 . TEXT_Period
					IfMsgBox, OK
					{
						Gosub Action1
					}
					IfMsgBox, Cancel
					{
						Return
					}
				}
				else
				{
					Gosub Action1
				}
			}
		}
		else if (NumButton == 2) ; MsgBoxOption = 1: OK/Cancel
		{
			IfMsgBox, OK
			{
				if (NeedConfirm && Action1 != "")
				{
					MsgBox, 1, %TEXT_RunAction%, % TEXT_RunActionMsg . Action1 . TEXT_Period
					IfMsgBox, OK
					{
						Gosub Action1
					}
					IfMsgBox, Cancel
					{
						Return
					}
				}
				else
				{
					Gosub Action1
				}
			}
			IfMsgBox, Cancel
			{
				if (NeedConfirm && Action2 != "")
				{
					MsgBox, 1, %TEXT_RunAction%, % TEXT_RunActionMsg . Action2 . TEXT_Period
					IfMsgBox, OK
					{
						Gosub Action2
					}
					IfMsgBox, Cancel
					{
						Return
					}
				}
				else
				{
					Gosub Action2
				}
			}
		}
		else ; MsgBoxOption = 3: Yes/No/Cancel
		{
			IfMsgBox, Yes
			{
				if (NeedConfirm && Action1 != "")
				{
					MsgBox, 1, %TEXT_RunAction%, % TEXT_RunActionMsg . Action1 . TEXT_Period
					IfMsgBox, OK
					{
						Gosub Action1
					}
					IfMsgBox, Cancel
					{
						Return
					}
				}
				else
				{
					Gosub Action1
				}
			}
			IfMsgBox, No
			{
				if (NeedConfirm && Action2 != "")
				{
					MsgBox, 1, %TEXT_RunAction%, % TEXT_RunActionMsg . Action2 . TEXT_Period
					IfMsgBox, OK
					{
						Gosub Action2
					}
					IfMsgBox, Cancel
					{
						Return
					}
				}
				else
				{
					Gosub Action2
				}
			}
			IfMsgBox, Cancel
			{
				if (NeedConfirm && Action3 != "")
				{
					MsgBox, 1, %TEXT_RunAction%, % TEXT_RunActionMsg . Action3 . TEXT_Period
					IfMsgBox, OK
					{
						Gosub Action3
					}
					IfMsgBox, Cancel
					{
						Return
					}
				}
				else
				{
					Gosub Action3
				}
			}
		}
	}
}
Return
