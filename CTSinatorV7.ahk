#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

^#!x::
ExitApp

chromePageWait()
	{
	global
	Loop 50
		{
		while (A_Cursor = "AppStarting")
			continue
		Sleep, 100
		}
	}


^#!g:: ;generate message to user
	WinActivate, ahk_exe chrome.exe
	WinWaitActive, ahk_exe chrome.exe
	WinGetTitle, wintitle
	StringSplit, wintitlevars, wintitle, |, %A_Space%

	ticknum := wintitlevars1
	ticktype := wintitlevars2

	MsgBox, 4, , %ticktype% - %ticknum% , correct?



	if ticktype = Incident
	{
		MsgBox, INC detected
		clipboard :=
		Send, ^f	;SEARCHWINDOW_START
		Sleep, 300
		Send, {BackSpace}
		Sleep, 300
		Send, @
		Sleep, 300
		Send, {PgUp}{PgUp}{PgUp}
		PixelSearch, Px, Py, 50, 190, 1000, 700, 0xff9632, , Fast RGB
		if ErrorLevel
			{
			PixelSearch, Px, Py, 50, 190, 1000, 700, 0xff0000, , Fast RGB
			if ErrorLevel
				{
				Sleep, 1000
				Send, {BackSpace}
				Sleep, 1000
				Send, @
				Sleep, 1000
				PixelSearch, Px, Py, 50, 190, 1000, 700, 0xff9632, , Fast RGB
				if ErrorLevel
					{
					PixelSearch, Px, Py, 50, 190, 1000, 700, 0xff0000, , Fast RGB
					if ErrorLevel
						{
						MsgBox, email not found
						}
					}
				}
			}
		if Px
			{
				MouseMove, %Px%, %Py%
				Px =
				Py =
				Send, {LButton}{LButton}{LButton}
				Send, ^c
				ClipWait, 2
				Send, {Esc}
			}

		InputBox, email, type in the email, is the email address correct?, , , , , , , , %clipboard%

		Send, {PgDn}
		Sleep, 300
		Send, ^f ;SEARCHWINDOW_START
		Sleep, 300
		Send, {BackSpace}
		Sleep, 300
		Send, Short description
		Sleep, 600
		PixelSearch, Px, Py, 50, 140, 1000, 700, 0xffff00, , Fast RGB
		if ErrorLevel
			{
			PixelSearch, Px, Py, 50, 140, 1000, 700, 0xff9632, , Fast RGB
			if ErrorLevel
				{
				MsgBox, desc not found
				}
			}
		if Px
			{
				MouseMove, %Px%, %Py%
				Px =
				Py =
				Send, {LButton}
				Send, ^a
				Sleep, 200
				Send, ^c
				Sleep, 500
				ClipWait

			}
			Send, {Esc}


		InputBox, desc, description, is the description correct?, , , , , , , , %clipboard%
		finalmsg = Hello, I am writing in regard to %ticknum% - %desc%. Please let me know when would you be available for a remote session to resolve this issue.
		MsgBox, %finalmsg%

		return

	}
	else if ticktype = Catalog Task
	{
		MsgBox, SCTASK detected
		Sleep, 300
		Send, {PgUp}
		Send, ^f
		Sleep, 300
		Send, Requested for
		Send, {PgUp}{PgUp}{PgUp}
		Sleep, 300
		PixelSearch, Px, Py, 50, 190, 1000, 900, 0xff9632, , Fast RGB
		if ErrorLevel
			{
			PixelSearch, Px, Py, 50, 190, 1000, 900, 0xff0000, , Fast RGB
			if ErrorLevel
				{
				Sleep, 1000
				Send, ^a
				Send, {BackSpace}
				Sleep, 1000
				Send, Requested for
				Sleep, 1000
				Send, {PgUp}{PgUp}{PgUp}
				PixelSearch, Px, Py, 50, 190, 1000, 900, 0xff9632, , Fast RGB
				if ErrorLevel
					{
					PixelSearch, Px, Py, 50, 190, 1000, 900, 0xff0000, , Fast RGB
					if ErrorLevel
						{
						MsgBox, requested not found
						}
					}
				}
			}
		Send, {Esc}
		if Px
			{
				MouseMove, %Px%, %Py%
				Px =
				Py =
				Sleep, 300
				Send, {LButton}
				Sleep, 300
				Send, {Tab}{Tab}
				Send, {Enter}
				Sleep, 500
				Send, {Tab}
				clipboard :=
				Send, ^c
				Sleep, 300
				ClipWait
			}
			InputBox, email, type in the email, is the email address correct?, , , , , , , , %clipboard%
			Send, {Esc} ;close the user info window

			Sleep, 300
			Send, ^f
			Sleep, 300
			Send, Affected Software Version
			Sleep, 300
			Send, {PgUp}
			Sleep, 300
			PixelSearch, Px, Py, 50, 190, 1000, 700, 0xff9632, , Fast RGB
			if ErrorLevel
				{
				PixelSearch, Px, Py, 50, 190, 1000, 700, 0xff0000, , Fast RGB
				if ErrorLevel
					{
					Sleep, 1000
					Send, ^a
					Send, {BackSpace}
					Sleep, 1000
					Send, Affected Software Version
					Send, {PgUp}
					Sleep, 1000
					PixelSearch, Px, Py, 50, 190, 1000, 700, 0xff9632, , Fast RGB
					if ErrorLevel
						{
						PixelSearch, Px, Py, 50, 190, 1000, 700, 0xff0000, , Fast RGB
						if ErrorLevel
							{
							MsgBox, , ,software not found,1
							Send, ^f
							Sleep, 300
							Send, Affected Software Version
							Sleep, 300
							Send, {PgUp}
							Sleep, 300
							PixelSearch, Px, Py, 50, 190, 1000, 700, 0xff9632, , Fast RGB
							if ErrorLevel
								{
								PixelSearch, Px, Py, 50, 190, 1000, 700, 0xff0000, , Fast RGB
								if ErrorLevel
									{
									Sleep, 1000
									Send, ^a
									Send, {BackSpace}
									Sleep, 1000
									Send, Affected CI
									Send, {PgUp}
									Sleep, 1000
									PixelSearch, Px, Py, 50, 190, 1000, 700, 0xff9632, , Fast RGB
									if ErrorLevel
										{
										PixelSearch, Px, Py, 50, 190, 1000, 700, 0xff0000, , Fast RGB
										if ErrorLevel
											{
											MsgBox, , ,software not found,1
											}
										}
									}
								}
							}
						}
					}
				}
			Send, {Esc}
			if Px
				{
					MouseMove, %Px%, %Py%
					Px =
					Py =
					Sleep, 300
					Send, {LButton}
					clipboard :=
					Sleep, 300
					Send, ^a
					Sleep, 300
					Send, ^c
					Sleep, 300
					ClipWait

				}
			InputBox, softname, type in the softname, is the softname correct?, , , , , , , , %clipboard%

			MsgBox, 4, , Manual install?

			IfMsgBox Yes
				finalmsg = Hello, I am writing in regard to %ticknum% - %softname% install. Please let me know when would you be available for a remote session to install this software.
			else IfMsgBox No
				finalmsg = Hello, I am writing in regard to %ticknum% - %softname% install. I've added you to the proper AD group, the software should be available in your software center in 24h. Please let me know when it's available so I can close the ticket.

			MsgBox, %finalmsg%


		return
	}





^#!h:: ;hold the ticket
WinActivate, ahk_exe chrome.exe
WinWaitActive, ahk_exe chrome.exe
WinGetTitle, wintitle
StringSplit, wintitlevars, wintitle, |, %A_Space%

ticknum := wintitlevars1
ticktype := wintitlevars2

MsgBox, 4, , %ticktype% - %ticknum% , correct?

if ticktype = Incident
{
	InputBox, KeySth, hold reason?
	if (KeySth = "u")
	{
		Send, {PgUp}
		Send, ^f ;SEARCHWINDOW_START
		Sleep, 300
		Send, {BackSpace}
		Sleep, 300
		Send, State
		Sleep, 400
		Send, {PgUp}{PgUp}{PgUp}
		Sleep, 400
		PixelSearch, Px, Py, 908, 120, 1901, 577, 0xff9632, , Fast RGB
		if ErrorLevel
			{
			PixelSearch, Px, Py, 908, 120, 1901, 577, 0xffff00, , Fast RGB
			MouseMove, %Px%, %Py%
			}
		else
    	MouseMove, %Px%, %Py%
			Px =
			Py =
		Send, {LButton 2}
		clipboard :=
		Sleep, 100
		Send, ^c
		ClipWait
		state := clipboard
		isState := InStr(state, "State")
		if isState = 0
			{
				MsgBox, nopestate
				return
			}

		if ErrorLevel
			{
				MsgBox, nopeerror
				return
			}
		Sleep, 300
		Send, {Down}
		Sleep, 500
		Send, ^f ;SEARCHWINDOW_START
		Sleep, 300
		Send, {BackSpace}
		Sleep, 300
		Send, On Hold Reason
		Sleep, 300
		Send, {PgUp}
		Sleep, 300
		PixelSearch, Px, Py, 908, 120, 1901, 577, 0xff9632, , Fast RGB
		if ErrorLevel
			{
			PixelSearch, Px, Py, 908, 120, 1901, 577, 0xffff00, , Fast RGB
			MouseMove, %Px%, %Py%
			}
		else
    	MouseMove, %Px%, %Py%
			Px =
			Py =
		Send, {LButton 3}
		clipboard :=
		Sleep, 100
		Send, ^c
		ClipWait
		holdrsn := clipboard
		isholdrsn := InStr(holdrsn, "On Hold Reason")
		if isholdrsn = 0
			{
				MsgBox, nopeholdrsn
				return
			}

		if ErrorLevel
			{
				MsgBox, nopeerror
				return
			}
		Sleep, 300
		Send, {Down 5}
		Sleep, 500
		MsgBox, ticket set to hold?
		Send, ^f ;SEARCHWINDOW_START
		Send, {BackSpace}
		Send, Additional comments (Customer visible)
		Sleep, 100
		PixelSearch, Px, Py, 0, 150, 1095, 1038, 0x8b8c0e, , Fast RGB
		if ErrorLevel
			{
			Send, ^f
			Sleep, 300
			Send, {BackSpace}
			Sleep, 300
			Send, Notes
			Sleep, 300
			PixelSearch, Px, Py, 0, 150, 360, 1036, 0xff9632, , Fast RGB
			if ErrorLevel
				{
				PixelSearch, Px, Py, 0, 150, 360, 1036, 0xffff00, , Fast RGB
				if ErrorLevel
					{
						Send, ^f
						Sleep, 300
						Send, {Enter}
						Sleep, 300
						PixelSearch, Px, Py, 0, 150, 360, 1036, 0xff9632, , Fast RGB
						if ErrorLevel
							{
							PixelSearch, Px, Py, 0, 150, 360, 1036, 0xffff00, , Fast RGB
							}
					}
				}
			if Px
				{
	    	MouseMove, %Px%, %Py%
				}
				Px =
				Py =
			Send, {LButton}
			Sleep, 300
			Send, ^f
			Send, {BackSpace}
			Sleep, 100
			Send, Additional comments (Customer visible)
			Sleep, 100
			PixelSearch, Px, Py, 0, 126, 1095, 1038, 0x8b8c0e, , Fast RGB
			MouseMove, %Px%, %Py%
			Send, {LButton}
			MsgBox, enter if text box ok
			MouseMove, %Px%, %Py%
			Send, {LButton}
			Send, Contacted user on Teams, waiting for reply
			Send, {Esc}
			}
		else
			{
    	MouseMove, %Px%, %Py%
			Send, {LButton}
			Send, Contacted user on Teams, waiting for reply
			Send, {Esc}
			}
		return
	}
}
else if ticktype := "Catalog Task"
{
	InputBox, KeySth, hold reason?
	if (KeySth = "u")
	{

		Send, {PgUp}
		Send, ^f
		Send, {BackSpace}
		Sleep, 100
		Send, State
		Sleep, 300
		Send, {PgUp}
		Sleep, 300
		PixelSearch, Px, Py, 908, 120, 1901, 577, 0xff9632, , Fast RGB
		if ErrorLevel
			{
			PixelSearch, Px, Py, 908, 120, 1901, 577, 0xffff00, , Fast RGB
			MouseMove, %Px%, %Py%
			}
		else
    	MouseMove, %Px%, %Py%
			Sleep, 300
		Send, {LButton}{LButton}
		clipboard :=
		Sleep, 300
		Send, ^c
		ClipWait
		state := clipboard
		isState := InStr(state, "State")
		if isState = 0
			{
				MsgBox, nopestate
				return
			}

		if ErrorLevel
			{
				MsgBox, nopeerror
				return
			}
		Sleep, 300
		Send, {Down}
		Sleep, 300
		Send, {Tab}
		Sleep, 300
		Send, {Down}
		Sleep, 500
		Send, ^f
		Send, {BackSpace}
		Sleep, 600
		Send, Comments and Work Notes
		Sleep, 500
		PixelSearch, Px, Py, 0, 500, 1901, 1000, 0xff9632, , Fast RGB
		MouseMove, %Px%, %Py%
		Send, {LButton}
		Sleep, 500
		Send, ^f
		Send, {BackSpace}
		Sleep, 600
		Send, Additional comments (Customer visible)
		Sleep, 100
		PixelSearch, Px, Py, 0, 126, 1095, 1038, 0x8b8c0e, , Fast RGB
		MouseMove, %Px%, %Py%
		Sleep, 200
		Send, {LButton}
		MsgBox, is this the correct text box?
		Send, Contacted user on Teams, waiting for reply
		Send, {Esc}
		return
	}
}

;=============================================================================================================================================================================================
;==================================================================================quicktypes===========================================================================================================
;=============================================================================================================================================================================================

^#!v::
	Send, %finalmsg%
	return

^#!+v::
	clipboard := finalmsg

^#!e::
	Send, %email%
	return

	^#!c::
	Send, Contacted user on Teams, waiting for reply
	return

^#!i::
	Send, Please go to the start menu, type 'my computer info', open the application, click OK and when the information opens just ctrl+v in teams. please do not screenshot or copy, just paste in teams.
	return

^#!t::
	Send, user stated issue has been resolved already, closing ticket
	return

^#!s::
		Send, ok, I will be connecting shortly
		return

^#!b::
	Send, I will close the ticket now, have a great day :)
	return

^#!k::
	Send, user has confirmed - software installed successfully
		return

^#!d:: ;d jak debilu
	Send, I asked you not to screenshot. please read my message again.

^#!f:: ;search for string in SNOW
	clipboard :=
	Send, ^c
	ClipWait
	ticknum = %clipboard%
	MsgBox, %ticknum%
	Run, chrome.exe "wood.service-now.com/text_search_exact_match.do?sysparm_search=%ticknum%"

	return





^#!a::
	InputBox, KeySt, action?
	if (KeySt = "l")
		{
		InputBox, hostnamelaps, type in the hostname, is the hostname correct?, , , , , , , , %clipboard%
		InputBox, dcname, check the domain controller, is the domain controller correct?, , , , , , , , MIL1-GDC10

		run, powershell
		Sleep, 1000
		Send, Get-ADComputer -Identity "%hostnamelaps%" -Server "%dcname%" -properties ms-mcs-admpwd | select-object ms-mcs-admpwd
		Send, {Enter}
		}
	else if (KeySt = "d")
		{
		clipboard = ok, first off, please go into the start menu, search for "dell command update", open it, click "check for updates", then "install all", then reboot when it asks you to do so. when you're done with it we can continue troubleshooting.
		}
	else if (KeySt = "c")
		{
			InputBox, hostnamecmrc, type in the hostname, is the description correct?, , , , , , , , %clipboard%
		run, "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\i386\CmRcViewer.exe" %hostnamecmrc%
		}
	else if (KeySt = "r")
		{
			Send, #r
			WinWaitActive, ahk_class #32770
			MsgBox, test script for win+r

		}

	return
