Set obj = CreateObject("wscript.shell")
dim result
dim main
dim username
dim tpygt
dim time
dim actime
dim acactime

result = msgbox("Hi and welcome to Box assistant do you want to proceed?", 4+32 , "Box assistant")

If result=6 then
msgbox "hi"
msgbox "please note that this is a beta version, please report any bugs to my Discord"
a = msgbox("Do you want to see a list of things i can do?", 4+32, "Box asssistant")
If a=6 then
msgbox "ok"
obj.run "notepad.exe"
WScript.sleep(1000)
obj.sendkeys "1. r_shutdown, shuts down the computer"
obj.sendkeys "{ENTER}"
obj.sendkeys "2. r_restart, restarts the computer"
obj.sendkeys "{ENTER}"
obj.sendkeys "3. r_winver, shows the windows version"
obj.sendkeys "{ENTER}"
obj.sendkeys "4. r_systemi, shows the system info"
obj.sendkeys "{ENTER}"
obj.sendkeys "5. r_taskl, show the running softwares"
obj.sendkeys "{ENTER}"
obj.sendkeys "6. r_killt, closes a app"
obj.sendkeys "{ENTER}"
obj.sendkeys "7. r_dtime, shows the exact time"
obj.sendkeys "{ENTER}"
obj.sendkeys "8. r_treetype, Graphically displays the folder structure of a drive or path"
obj.sendkeys "{ENTER}"
obj.sendkeys "9. r_run, runs a certain program"
obj.sendkeys "{ENTER}"
obj.sendkeys "10. r_alarm, sets a timer"
obj.sendkeys "{ENTER}"
obj.sendkeys "11. r_hi, says hi"
obj.sendkeys "{ENTER}"
obj.sendkeys "12. r_cal, opens calculator"
obj.sendkeys "{ENTER}"
obj.sendkeys "end, ends the app"
obj.sendkeys "{ENTER}"
obj.sendkeys "If you need any help please send a message to Mr_pop#5941 on Discord"
wscript.sleep(3000)
msgbox "if done press ok"
call asli
else
msgbox "ok"
call asli
sub asli
main = inputbox("tell me what to do", "Box assistant", "write a command here")
if main = "r_alarm" then 
msgbox "ok"
time= inputbox("how many seconds to you want to wait", "Box assistant", "..")
actime = time*1000
WScript.sleep(actime)
msgbox "timer done"
call asli
Elseif main = "r_treetype" then
msgbox "ok"
tpygt = inputbox("please write the exact location of you file/floder with no spaces", "box assistant", "leave blank for the whole pc")
obj.run "cmd.exe"
WScript.sleep(1000)
obj.sendkeys "tree "
obj.sendkeys tpygt
obj.sendkeys "{ENTER}"
call asli
Elseif main = "r_shutdown" then
msgbox "ok"
obj.run "cmd.exe"
WScript.sleep(1000)
obj.sendkeys "shutdown /s"
obj.sendkeys "{ENTER}"
call asli
Elseif main = "r_restart" then 
msgbox "ok"
obj.run "cmd.exe"
WScript.sleep(1000)
obj.sendkeys "shutdown /r"
obj.sendkeys "{ENTER}"
call asli
Elseif main = "r_winver" then 
obj.run "cmd.exe"
WScript.sleep(1000)
obj.sendkeys "ver"
obj.sendkeys "{ENTER}"
call asli
Elseif main = "r_systemi" then
msgbox "ok"
username = inputbox("please write your username", "box assistant", "..")
obj.run "cmd.exe"
WScript.sleep(1000)
obj.sendkeys "systeminfo /s "
obj.sendkeys username
obj.sendkeys "{ENTER}"
call asli
Elseif main = "r_dtime" then
msgbox "ok"
obj.run "cmd.exe"
Wscript.sleep(1000)
obj.sendkeys "time"
obj.sendkeys "{ENTER}"
call asli
Elseif main = "r_taskl" then
msgbox "ok"
obj.run "cmd.exe"
WScript.sleep(1000)
obj.sendkeys "tasklist"
obj.sendkeys "{ENTER}"
call asli
Elseif main = "r_killt" then
msgbox "ok"
username=inputbox("please write the name of the app + the file", "box assistant", "for example notepad.exe")
obj.run "cmd.exe"
WScript.sleep(1000)
obj.sendkeys "taskkill /im "
obj.sendkeys username
obj.sendkeys "{ENTER}"
call asli
Elseif main = "r_cal" then 
obj.run "calc.exe"
call asli
Elseif main = "r_hi" then
msgbox "hi"
call asli
Elseif main = "end" then
msgbox "bye"
wscript.quit
end if
end sub
end if
else
msgbox "bye"
end if