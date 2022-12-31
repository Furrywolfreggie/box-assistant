Set obj = CreateObject("wscript.shell")
Set WshShell = CreateObject("WScript.shell")
Dim result
Dim main
Dim username
Dim tpygt
Dim time
Dim actime
Dim acactime
Dim URL

result = msgbox("Hi and welcome to Box assistant do you want to proceed?", 4+32 , "Box assistant")

If result=6 Then
	msgbox "hi"
	msgbox "please note that this is a beta version, please report any bugs to my Discord"
	a = msgbox("Do you want to see a list of commands i can run?", 4+32, "Box assistant")
	If a=6 Then
		Sub help
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
		obj.sendkeys "13. end, ends the app"
		obj.sendkeys "{ENTER}"
		obj.sendkeys "14. help, opens help menu"
		obj.sendkeys "{ENTER}"
		obj.sendkeys "15. r_opweb, opens a certain URL"
		obj.sendkeys "{ENTER}"
		obj.sendkeys "If you need any help please send a message to Mr_pop#5941 on Discord"
		wscript.sleep(3000)
		msgbox "if done press ok"
		Call asli
		End sub
	Else
		msgbox "ok"
		Call asli
		Sub asli
			main = inputbox("tell me what to do", "Box assistant", "write a command here")
			If main = "r_alarm" Then 
				msgbox "ok"
				time= inputbox("how many seconds to you want to wait", "Box assistant", "..")
				actime = time*1000
				WScript.sleep(actime)
				msgbox "timer done"
				Call asli
			ElseIf main = "r_treetype" Then
				msgbox "ok"
				tpygt = inputbox("please write the exact location of you file/floder with no spaces", "box assistant", "leave blank for the whole pc")
				obj.run "cmd.exe"
				WScript.sleep(1000)
				obj.sendkeys "tree "
				obj.sendkeys tpygt
				obj.sendkeys "{ENTER}"
				Call asli
			ElseIf main = "r_shutdown" Then
				msgbox "ok"
				obj.run "cmd.exe"
				WScript.sleep(1000)
				obj.sendkeys "shutdown /s"
				obj.sendkeys "{ENTER}"
				Call asli
			ElseIf main = "r_restart" Then 
				msgbox "ok"
				obj.run "cmd.exe"
				WScript.sleep(1000)
				obj.sendkeys "shutdown /r"
				obj.sendkeys "{ENTER}"
				Call asli
			ElseIf main = "r_winver" Then 
				obj.run "cmd.exe"
				WScript.sleep(1000)
				obj.sendkeys "ver"
				obj.sendkeys "{ENTER}"
				Call asli
			ElseIf main = "r_systemi" Then
				msgbox "ok"
				username = inputbox("please write your username", "box assistant", "..")
				obj.run "cmd.exe"
				WScript.sleep(1000)
				obj.sendkeys "systeminfo /s "
				obj.sendkeys username
				obj.sendkeys "{ENTER}"
				Call asli
			ElseIf main = "r_dtime" Then
				msgbox "ok"
				obj.run "cmd.exe"
				Wscript.sleep(1000)
				obj.sendkeys "time"
				obj.sendkeys "{ENTER}"
				Call asli
			ElseIf main = "r_taskl" Then
				msgbox "ok"
				obj.run "cmd.exe"
				WScript.sleep(1000)
				obj.sendkeys "tasklist"
				obj.sendkeys "{ENTER}"
				Call asli
			ElseIf main = "r_killt" Then
				msgbox "ok"
				username=inputbox("please write the name of the app + the file", "box assistant", "for example notepad.exe")
				obj.run "cmd.exe"
				WScript.sleep(1000)
				obj.sendkeys "taskkill /im "
				obj.sendkeys username
				obj.sendkeys "{ENTER}"
				Call asli
			ElseIf main = "r_cal" Then 
				obj.run "calc.exe"
				Call asli
			ElseIf main = "r_hi" Then
				msgbox "hi"
				Call asli
			ElseIf main = "end" Then
				msgbox "bye"
				wscript.Quit
			ElseIf main = "help" Then
				Call help
			ElseIf main = "r_opweb" Then
				URL = inputbox("write the URL that you want to go to", "Box assistant", "write a URL here")
				WshShell.run "CMD /C start chrome.exe " & URL & "",0,False
			ElseIf main = "" Then
				MsgBox "nothing entered"
				Call asli
			ElseIf main = "write a command here" Then
				MsgBox "nothing entered"
				Call asli
			End If
		End Sub
	End If
Else
	msgbox "bye"
End If
'pop studios 2022 