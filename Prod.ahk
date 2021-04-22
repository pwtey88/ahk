; Create the popup menu by adding some items to it.
Menu, quickCopyer, Add, ITSM View Ticket         				Win+1, MenuHandler
Menu, quickCopyer, Add, ITSM Create CRQ         				Win+2, MenuHandler
Menu, quickCopyer, Add, ITSM Create INC          				Win+3, MenuHandler
Menu, quickCopyer, Add, ITSM Create WO         					Win+4, MenuHandler
Menu, quickCopyer, Add, ITSM Search Asset       				Win+5, MenuHandler
Menu, quickCopyer, Add, ITSM KBA search         				Win+6, MenuHandler
Menu, quickCopyer, Add, TFS Search                    				Win+7, MenuHandler
Menu, quickCopyer, Add, Send email Skype team                     		Win+8, MenuHandler
Menu, quickCopyer, Add, Send email to Nadia 	                     		                                              Win+9, MenuHandler
Menu, quickCopyer, Add  ; Add a separator line.
Menu, quickCopyer, Add, Google Search              				Win+g, MenuHandler
Menu, quickCopyer, Add, XOM Search                  				Win+x, MenuHandler
Menu, quickCopyer, Add, Copy to Address Bar    					Win+c, MenuHandler 
Menu, quickCopyer, Add, Putty SSH Logon          				Win+z, MenuHandler
Menu, quickCopyer, Add, RDP Logon                    				Win+d, MenuHandler
Menu, quickCopyer, Add  ; Add a separator line.
Menu, quickCopyer, Add, Ping Nslookup and Tracert     				Win+a, MenuHandler 
Menu, quickCopyer, Add, Reboot Laptop 		     			        Win+b, MenuHandler 
Menu, quickCopyer, Add, Ping     						Win+p, MenuHandler 
Menu, quickCopyer, Add, Empty Trash                   				Win+del, MenuHandler
Menu, quickCopyer, Add  ; Add a separator line.
Menu, quickCopyer, Add, Notepad                       				F1, MenuHandler
Menu, quickCopyer, Add, Word                       				F2, MenuHandler
Menu, quickCopyer, Add, Excel                       				F4, MenuHandler
Menu, quickCopyer, Add  ; Add a separator line.
Menu, quickCopyer, Add, Close all windows                       		Ctrl+Win+Z, MenuHandler




return  ; End of script's auto-execute section.

MenuHandler:

if A_ThisMenuItem = ITSM View Ticket         Win+1 ; itsm view  
{
    Send {ctrldown}c{ctrlup}
    Sleep 100
    Run, http://itsm/?%selectedText%
    Return
}

if A_ThisMenuItem = Google Search              Win+g ; 
{
    Send {ctrldown}c{ctrlup}
    Sleep 100
    Run, http://www.google.com/search?q=%selectedText%
    Return
}

if A_ThisMenuItem = Ping Nslookup and Tracert         Win+a ; ping 30 times          
{
    Send {ctrldown}c{ctrlup}
    Sleep 100
    Run C:\Windows\System32\ping.exe -n 30 %selectedText%
    Run C:\Windows\System32\cmd.exe /K "C:\Windows\System32\tracert.exe %selectedText%"
    Run C:\Windows\System32\cmd.exe /K "C:\Windows\System32\nslookup.exe %clipboard%"
    Return
}

if A_ThisMenuItem = Ping        Win+p ; continous ping          
{
    Send {ctrldown}c{ctrlup}
    Sleep 100
    Run C:\Windows\System32\ping.exe -t %selectedText%
    Return
}


if A_ThisMenuItem = Copy to Address Bar    Win+c ;        
{
    Send {ctrldown}c{ctrlup}
    Sleep 100
    Run, C:\Program Files\Internet Explorer\iexplore.exe %selectedText%
    Return
}

if A_ThisMenuItem = ITSM Create CRQ         Win+2 
{
    Send {ctrldown}c{ctrlup}
    Sleep 100
    Run, http://crq?new
    Return
}

if A_ThisMenuItem = ITSM Create INC          Win+3 
{
    Send {ctrldown}c{ctrlup}
    Sleep 100
    Run, http://inc?new
    Return
}

if A_ThisMenuItem = ITSM Create WO          Win+4 
{
    Send {ctrldown}c{ctrlup}
    Sleep 100
    Run, http://workorder?new
    Return
}

if A_ThisMenuItem = ITSM Search Asset       Win+5 
{
    Send {ctrldown}c{ctrlup}
    Sleep 100
    Run, http://ast/?%clipboard%
    Return
}

if A_ThisMenuItem = TFS Search                    Win+7 
{
    Send {ctrldown}c{ctrlup}
    Sleep 100
    Run, https://tfs.xom.com/tfs/APPS2/DWS/_backlogs/backlog/UC_East/Features
    Return
}

if A_ThisMenuItem = ITSM KBA search         Win+6 
{
    Send {ctrldown}c{ctrlup}
    Sleep 100
    Run, http://kba?search 
    Sleep 5000
    send %clipboard%
    Sleep 100
    send {enter}
    Return 
}

if A_ThisMenuItem = Send email Skype team    Win+8 
{
    Send {ctrldown}c{ctrlup}
    Sleep 100
    Run, C:\Program Files (x86)\microsoft office\Office15\outlook.EXE /c ipm.note /m "EMIT.CUST.INF.COLLAB.RT.REALTIME.ALLWTYPES.UG1@exxonmobil.com" 
    Sleep 5000
    send %clipboard%
    Sleep 100
    send {enter}
    Return 
}


if A_ThisMenuItem = Send email to Nadia   Win+9 
{
    Send {ctrldown}c{ctrlup}
    Sleep 100
    Run, C:\Program Files (x86)\microsoft office\Office15\outlook.EXE /c ipm.note /m "nadia@exxonmobil.com@exxonmobil.com" 
    Sleep 5000
    send %clipboard%
    Sleep 100
    send {enter}
    Return 
}


if A_ThisMenuItem = RDP Logon                    Win+d  
{
    Send {ctrldown}c{ctrlup}
    Sleep 100
;    Run, C:\WINDOWS\System32\mstsc.exe /v:%clipboard%
    Run, C:\WINDOWS\System32\mstsc.exe 
    Return
}

if A_ThisMenuItem = XOM Search                  Win+x ; malaysia
{
    Send {ctrldown}c{ctrlup}
    Sleep 100
    Run, https://es.na.xom.com/search?q=%clipboard% 
    Return
}

return

if A_ThisMenuItem = Notepad                  F1; notepad
{
    Send {ctrldown}c{ctrlup}
    Sleep 100
    F1::Run "%WINDIR%\notepad.exe"
    Return
}

if A_ThisMenuItem = Word                  F2; word
{
    Send {ctrldown}c{ctrlup}
    Sleep 100
    F2::Run "C:\Program Files (x86)\Microsoft Office\Office15\WINWORD.EXE"
    Return
}

if A_ThisMenuItem = Word                  F4; excel
{
    Send {ctrldown}c{ctrlup}
    Sleep 100
    F4::Run "C:\Program Files (x86)\Microsoft Office\Office15\EXCEL.EXE"
    Return
}

return

if A_ThisMenuItem = Empty Trash                  Win+del; 
{
    Send {ctrldown}c{ctrlup}
    Sleep 100
    #Del::FileRecycleEmpty
    Return
}

return

if A_ThisMenuItem = Reboot Laptop                  Win+b; 
{
    Send {ctrldown}c{ctrlup}
    Sleep 100
    #b::Run, %comspec% /c "shutdown.exe /r /t 0"
    ;/r = restart
    ;/t 0 = wait 0 seconds
    Return
}

return


;getting contents from selection, and combine with some related information, like url.
;win+`

#`::
;got selection
clipboard =
send,^c
ClipWait,0.2
selectedText:=Clipboard ;plain text
;export to clipboard
clipboard = %selectedText%

;call the menu
Menu, quickCopyer, Show  ;press the Win+` hotkey to show the menu.

return

#s::
Send {ctrldown}c{ctrlup}
Sleep 1
Run, C:\Program Files\Centrify\Centrify PuTTY\putty.exe %clipboard%
Return

#g::
Send {ctrldown}c{ctrlup}
Sleep 100
Run, http://www.google.com/search?q=%clipboard%
Return

#1::
Send {ctrldown}c{ctrlup}
Sleep 100
Run, http://itsm/?%clipboard%
Return

#c::
Send {ctrldown}c{ctrlup}
Sleep 100
Run, C:\Program Files\Internet Explorer\iexplore.exe %clipboard%
Return

#a::
Send {ctrldown}c{ctrlup}
Sleep 100
Run, C:\Windows\System32\ping.exe -n 30 %clipboard%
Run, C:\Windows\System32\cmd.exe /K "C:\Windows\System32\tracert.exe %clipboard%
Run, C:\Windows\System32\cmd.exe /K "C:\Windows\System32\nslookup.exe %clipboard%"
Return

#p::
Send {ctrldown}c{ctrlup}
Sleep 100
Run, C:\Windows\System32\ping.exe -t %clipboard%
Return


#d::
Send {ctrldown}c{ctrlup}
Sleep 100
Run, C:\WINDOWS\System32\mstsc.exe /v:%clipboard%
Return

#4::
Send {ctrldown}c{ctrlup}
Sleep 100
Run, http://workorder?new
Return

#3::
Send {ctrldown}c{ctrlup}
Sleep 100
Run, http://inc?new
Return

#2::
Send {ctrldown}c{ctrlup}
Sleep 100
Run, http://crq?new
Return 

#5::
Send {ctrldown}c{ctrlup}
Sleep 100
Run, http://ast/?%clipboard%
Return 

#6::
Send {ctrldown}c{ctrlup}
Sleep 100
Run, http://kba?search 
Sleep 5000
send %clipboard%
Sleep 100
send {enter}
Return 

#7::
Send {ctrldown}c{ctrlup}
Sleep 100
Run, https://tfs.xom.com/tfs/APPS2/DWS/_backlogs/backlog/UC_East/Features
Return

#8::
Send {ctrldown}c{ctrlup}
Sleep 100
Run, C:\Program Files (x86)\microsoft office\Office15\outlook.EXE /c ipm.note /m "EMIT.CUST.INF.COLLAB.RT.REALTIME.ALLWTYPES.UG1@exxonmobil.com" 
Return

#9::
Send {ctrldown}c{ctrlup}
Sleep 100
Run, C:\Program Files (x86)\microsoft office\Office15\outlook.EXE /c ipm.note /m "nadia@exxonmobil.com" 
Return


#x::
Send {ctrldown}c{ctrlup}
Sleep 100
Run, https://es.na.xom.com/search?q=%clipboard% 
Return

#IfWinActive ahk_class ConsoleWindowClass
^V::
SendInput {Raw}%clipboard%
return
#IfWinActive

 
;Insert special character
!q::SendInput {™} ; Alt + Q
return

;; Start SSH session using the currently selected text as the target hostname.
#z::

    ; Putty copies selected text to the clipboard so you don't need to copy it
    ; doing Ctrl-Insert throws away what you already have in the Clipboard.
    ; Can't use Ctrl-C in putty, because it sends that to your session as ^C
    WinGet, Active_ID, ID, A
    WinGet, Active_Process, ProcessName, ahk_id %Active_ID%
    if ( Active_Process ="putty.exe" )
    {

        host = %Clipboard% 

    }
    else
    {

        host_tmp := SelectedViaClipboard()
        host := ValidateHostname(host_tmp)

    }

    title_msg := "Please enter the hostname"
    prompt_msg := "Didn't get host from selected text:"
 
    if (!host)
    {
        InputBox, new_host, %title_msg%, %prompt_msg%,, 220, 111, , , , , %host_tmp%
        if (ErrorLevel)
        {
            return
        }
        else
        {
            host := ValidateHostname(new_host)
        }
    }

    if (host)
    {
        ; MsgBox, Doing ssh %host%
        Run "C:\Program Files\Centrify\Centrify PuTTY\putty.exe" "-ssh" "%host%"
        ; NB: You could add your username as "user@%host%"
    }
    else
    {
        MsgBox, %prompt_msg% [%host_tmp%]
    }
    return

;; utility functions

;; Get selected text without clobbering Clipboard
SelectedViaClipboard()
{
    old_clipboard = %ClipboardAll% ; save current Clipboard
    Clipboard := "" ; clears Clipboard
    Send, ^{Insert}
    selection = %Clipboard% ; save the content of the clipboard
    Clipboard = %old_clipboard% ; restore old content of the clipboard
    return selection

}

;; Capture a hostname out of some selected text
ValidateHostname(str)
{
    if (RegExMatch(str, "^([ \'\[@]*)([\w\-\.]+)([ \'\]:/]*)$", match))
        return match2
    return ""
}



; Parameter #1: Pass 1 instead of 0 to hibernate rather than suspend.
; Parameter #2: Pass 1 instead of 0 to suspend immediately rather than asking each application for permission.
; Parameter #3: Pass 1 instead of 0 to disable all wake events.
^!s::DllCall("PowrProf\SetSuspendState", "int", 0, "int", 1, "int", 1) ;  Suspend
^!h::DllCall("PowrProf\SetSuspendState", "int", 1, "int", 0, "int", 1) ;  Hibernate
^!d::shutdown, 8 ;Powerdown ; 

^#z::
WinGet, id, list, , , Program Manager
Loop, %id%
{
 StringTrimRight, this_id, id%a_index%, 0
 WinGetTitle, this_title, ahk_id %this_id%
 winclose,%this_title%
}
Return