#NoEnv ; recommended for performance and compatibility with future AutoHotkey releases
#Persistent ; keep the script permanently running
#SingleInstance force ; allow only one instance of this script to be running
#Warn ; enable warnings to assist with detecting common errors
SendMode Input ; recommended for new scripts due to its superior speed and reliability
SetWorkingDir %A_ScriptDir% ; ensures a consistent starting directory

ScriptName := "Reminder"
ScriptVersion := "1.1.0.0"
CopyrightNotice := "Copyright (c) 2020 Chaohe Shi"

ConfigDir := A_AppData . "\" . ScriptName
ConfigFile := ConfigDir . "\" . ScriptName . ".ini"

TEXT_ReminderText := "The time is now "
TEXT_Colon := ":"
TEXT_Period := "."
TEXT_Settings := "Settings"
TEXT_Test := "Test"
TEXT_Try := "Try"
TEXT_OK := "OK"
TEXT_ErrorMsg := "Not a valid command!"
TEXT_About := "About"
TEXT_Exit := "Exit"
TEXT_AboutMsg := ScriptName . " " . ScriptVersion . "`n`n" . CopyrightNotice

ReminderInterval := 30 ; this number needs to be one of the factors of 60min. e.g. when reminderInterval is 15, reminder occurs at 0, 15, 30, 45min of every hour
MsgBoxTimeout := 60

IniRead, MsgBoxOption, %ConfigFile%, General, MsgBoxOption, 0 ; use 0, 1, or 3 to for 1, 2, or 3 buttons and enable the close button
IniRead, NumButton, %ConfigFile%, General, NumButton, 1
IniRead, Button1, %ConfigFile%, Button, Button1, OK
IniRead, Button2, %ConfigFile%, Button, Button2, Work
IniRead, Button3, %ConfigFile%, Button, Button3, Relax
IniRead, Action1, %ConfigFile%, Action, Action1, Notepad.exe
IniRead, Action2, %ConfigFile%, Action, Action2, https://www.google.com/
IniRead, Action3, %ConfigFile%, Action, Action3, ::{20d04fe0-3aea-1069-a2d8-08002b30309d}

Menu, Tray, NoStandard ; remove the standard menu items
Menu, Tray, Tip, %ScriptName% ; change the tray icon's tooltip
Menu, Tray, Add, %ScriptName%, ShowReminder
Menu, Tray, Default, %ScriptName%
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

	Gui, Add, GroupBox, Section w795 h180, %TEXT_Settings%
	Gui, Add, Text, Section xp+10 yp+20, Number of button(s)

	;Gui, Add, Text, Section, Number of button(s)
	Gui, Add, DropDownList, vNumButton gSettings Choose%NumButton% w75, 1|2|3

	Gui, Add, Text, Section ys, Button1 Text
	Gui, Add, Edit, vButton1 gSettings r1 w75 Limit, %Button1%

	Gui, Add, Text, ys, Button1 Action
	Gui, Add, Edit, vAction1 gSettings r1 w500, %Action1%

	Gui, Add, Text, ys
	Gui, Add, Button, vTest1 gAction1 w75, %TEXT_Test%

	Gui, Add, Text, Section xs, Button2 Text
	Gui, Add, Edit, vButton2 gSettings r1 w75 Limit, %Button2%

	Gui, Add, Text, ys, Button2 Action
	Gui, Add, Edit, vAction2 gSettings r1 w500, %Action2%

	Gui, Add, Text, ys
	Gui, Add, Button, vTest2 gAction2 w75, %TEXT_Test%

	Gui, Add, Text, Section xs, Button3 Text
	Gui, Add, Edit, vButton3 gSettings r1 w75 Limit, %Button3%

	Gui, Add, Text, ys, Button3 Action
	Gui, Add, Edit, vAction3 gSettings r1 w500, %Action3%

	Gui, Add, Text, ys
	Gui, Add, Button, vTest3 gAction3 w75, %TEXT_Test%

	Gui, Add, Text
	Gui, Add, Button, Section xm vTry gShowReminder w75 h23, %TEXT_Try%
	Gui, Add, Button, vOK gOK ys w75 h23, %TEXT_OK%
	GuiControl, +Default, OK
	GuiControl, Focus, OK
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
if (NumButton == 1)
{
	MsgBoxOption := 0
}
else if (NumButton == 2)
{
	MsgBoxOption := 257 ; 1 + 256
}
else
{
	MsgBoxOption := 515 ; 3 + 512
}
Gosub, EnsureConfigDirExists
IniWrite, %MsgBoxOption%, %ConfigFile%, General, MsgBoxOption
IniWrite, %NumButton%, %ConfigFile%, General, NumButton
IniWrite, %Button1%, %ConfigFile%, Button, Button1
IniWrite, %Button2%, %ConfigFile%, Button, Button2
IniWrite, %Button3%, %ConfigFile%, Button, Button3
IniWrite, %Action1%, %ConfigFile%, Action, Action1
IniWrite, %Action2%, %ConfigFile%, Action, Action2
IniWrite, %Action3%, %ConfigFile%, Action, Action3
Return

OK:
GuiClose:
GuiEscape:
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
	MsgBox, 1, %TEXT_About%, %TEXT_AboutMsg%
}
Return

ExitProgram:
ExitApp

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
	ShowReminder()
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
		if WinExist(ScriptName . " ahk_class #32770 ahk_pid " . ErrorLevel)
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

ShowReminder()
{
	global
	OnMessage(0x44, "WM_COMMNOTIFY") ; https://autohotkey.com/board/topic/56272-msgbox-button-label-change/?p=353457
	MsgBox, % MsgBoxOption, %ScriptName%, % TEXT_ReminderText . A_Hour . TEXT_Colon . A_Min . TEXT_Period, % MsgBoxTimeout
	if (NumButton == 1) ; MsgBoxOption = 0: OK
	{
		IfMsgBox, OK
		{
			Gosub Action1
		}
	}
	else if (NumButton == 2) ; MsgBoxOption = 1 + 256: OK/Cancel
	{
		IfMsgBox, OK
		{
			Gosub Action1
		}
		IfMsgBox, Cancel
		{
			Gosub, Action2
		}
	}
	else ; MsgBoxOption = 3 + 512: Yes/No/Cancel
	{
		IfMsgBox, Yes
		{
			Gosub Action1
		}
		IfMsgBox, No
		{
			Gosub, Action2
		}
		IfMsgBox, Cancel
		{
			Gosub, Action3
		}
	}
}

