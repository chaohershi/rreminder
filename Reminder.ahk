#NoEnv ; recommended for performance and compatibility with future AutoHotkey releases
#Persistent ; keep the script permanently running
#SingleInstance force ; allow only one instance of this script to be running
#Warn ; enable warnings to assist with detecting common errors
SendMode Input ; recommended for new scripts due to its superior speed and reliability
SetWorkingDir %A_ScriptDir% ; ensures a consistent starting directory

ScriptName := "Reminder"
ScriptVersion := "1.0.0.0"
CopyrightNotice := "Copyright (c) 2020 Chaohe Shi"

reminderInterval := 30 ; this number needs to be one of the factors of 60min. e.g. when reminderInterval is 15, reminder occurs at 0, 15, 30, 45min of every hour
msgBoxTimeout := 60
msgBoxOption := 0

TEXT_Reminder := "Reminder"
TEXT_ReminderText := "The time is now "
TEXT_Colon := ":"
TEXT_Period := "."
TEXT_About := "About"
TEXT_Exit := "Exit"
TEXT_AboutMsg := ScriptName . " " . ScriptVersion . "`n`n" . CopyrightNotice

;Menu, Tray, NoStandard ; remove the standard menu items
Menu, Tray, Add, Test, ShowReminder
Menu, Tray, Default, Test

Menu, Tray, Tip, %ScriptName% ; change the tray icon's tooltip
Menu, Tray, Add, %TEXT_About%, ShowAboutMsg
Menu, Tray, Add
Menu, Tray, Add, %TEXT_Exit%, ExitProgram

SetTimer, Reminder, % GetTimerPeriod()

Return ; end of the auto-execute section

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

Reminder:
if (IsTimeReached()) ; check the time before display the reminder to avoid the accumulated timer precision error https://www.autohotkey.com/docs/commands/SetTimer.htm#Precision
{
	ShowReminder()
}
SetTimer, Reminder, % GetTimerPeriod()
Return

; get the period for the next reminder in millisecond, accurate to one second
GetTimerPeriod()
{
	global reminderInterval
	Return Mod(3600 - A_Min * 60 - A_Sec, reminderInterval * 60) * 1000
}

; check if the current minute is the desired minute
IsTimeReached()
{
	global reminderInterval
	Return Mod(60 - A_Min, reminderInterval) == 0
}

ShowReminder()
{
	global
	MsgBox, % msgBoxOption, %TEXT_Reminder%, % TEXT_ReminderText . A_Hour . TEXT_Colon . A_Min . TEXT_Period, % msgBoxTimeout
}

