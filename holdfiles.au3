
Run("C:\Users\ajafri\Desktop\EOD\EODMonitor\EodMonitor.exe")

Sleep( 30000 )

WinActivate("C:\Users\ajafri\Desktop\EOD\EODMonitor\EodMonitor.exe")
WinActivate("EOD Monitor")
Send("!th")  ; open hold window

Send("{TAB}cftz*{TAB 3}{ENTER}")    ; cftz* add to hold list
Send("{TAB}fles*{TAB 4}{ENTER}")    ; fles* add to hold list
Send("{TAB}jtrd*{TAB 4}{ENTER}")    ; jtrd* add to hold list

Send("{TAB 2}{ENTER}")  ; Send the instruction

Send("{TAB 2}{ENTER}")  ; Exit from window