#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;; Key Bindings for AutoHotkey v1 and PowerPoint
;; nothans.com/powerpoint

+F12::
    
    If WinExist("PowerPoint Slide Show")
    {
        WinActivate ;
        Send {PgDn}
        return
    }
    Else
    {
        ControlSend,mdiClass1,+{F5},ahk_exe POWERPNT.EXE
        return
    }

+F11::

    if WinExist("PowerPoint Slide Show")
    {
        WinActivate ;
        Send {PgUp}
        return
    }

    return

+F8::
       
    if WinExist("PowerPoint Slide Show")
    {
        WinActivate ;
        Send {Esc}
	return
    }

    return

+F9::
    
    If WinExist("PowerPoint Slide Show")
    {
        WinActivate ;
        Send {PgDn}
        return
    }
    Else
    {
        ControlSend,mdiClass1,{F5},ahk_exe POWERPNT.EXE
        return
    }

+F7::
    SetTitleMatchMode, 2

    if WinExist("| Microsoft Teams")
    {
        WinActivate ;
        WinMaximize ;
        return
    }

    return
