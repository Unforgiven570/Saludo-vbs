Dim speaks, speech 
speaks="Hola Cristobal" 
Set speech=CreateObject("sapi.spvoice") 
speech.Speak speaks

Set objshell = createobject("wscript.shell")
Set ping = objshell.exec("cmd /c ipconfig")
Msgbox ping.stdout.readall
