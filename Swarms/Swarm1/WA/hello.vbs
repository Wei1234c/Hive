dim oController, oProcess
set oController = CreateObject("WSHController") 
set oProcess = oController.CreateScript("C:\Users\Wei\Documents\Dropbox\專案\通用工具\HIVE\Swarms\Swarm1\WA\remote.vbs", "\\192.168.0.105")
'WScript.ConnectObject oProcess, "\\192.168.0.105"

oProcess.Execute
While oProcess.Status <> 2
   WScript.Sleep 100
WEnd
WScript.Echo "Done!!!"