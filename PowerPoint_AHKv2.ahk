#Requires AutoHotkey v2.0

;; Key Bindings for AutoHotkey v2 and PowerPoint
;; nothans.com/powerpoint

; Bind to Shift+F12 - Advance slide or start presenting from the last slide
+F12::
{    
    if WinExist("PowerPoint Slide Show")
    {
        WinActivate
        Send "{PgDn}"
        return
    }
    else
    {
        if WinExist("ahk_exe POWERPNT.EXE") {
             WinActivate
	     Send "+{F5}"	
        }        
        return
    }
}

; Bind to Shift+F11 - Go back one slide
+F11::
{
    if WinExist("PowerPoint Slide Show")
    {
        WinActivate
        Send "{PgUp}"
        return
    }

    return
}

; Bind to Shift+F8 - Exit slideshow
+F8::
{       
    if WinExist("PowerPoint Slide Show")
    {
        WinActivate
        Send "{Esc}"
	return
    }

    return
}

; Bind to Shift+F7 - Advance slide or start presenting from the first slide
+F7::
{    
    if WinExist("PowerPoint Slide Show")
    {
        WinActivate
        Send "{PgDn}"
        return
    }
    else
    {
        if WinExist("ahk_exe POWERPNT.EXE") {
             WinActivate
	     Send "{F5}"	
        }        
        return
    }
}