
Option Explicit

'變數宣告 ____________________________
Dim oExcel
Dim oXMLhttp
Dim result
Dim randomNumber

Dim Hive_fileName
Dim theHiveFile
Dim theHive


Dim swarm_id
Dim theQueen

Dim worker_id
Dim theWorker

Dim hostName
Dim theHost

Dim theTask
Dim taskID
Dim targetedURL
Dim httpMethod
Dim parameters
Dim callBackFunction



'變數賦值 ____________________________





Hive_fileName = "- HIVE.xls"




swarm_id = 1


worker_id = 1


hostName = "- 資料蒐集.xls"



taskID = "4711863218537"
targetedURL = "http://ipac.library.taichung.gov.tw/toread/opac/Advancedsearch.page?classification_type=all&level=all&limit=20&location=3&material_type=all&q=isbn_issn%3A4711863218537&source=local&view=CONTENT&wi=false"
httpMethod = "GET"

callBackFunction = "saveResult taskID, result"






'主要邏輯 ____________________________

'抓取 Excel.Application 物件
Set oExcel = GetObject(, "Excel.Application")
oExcel.visible = True

'建立物件

Set theHost = oExcel.Workbooks(hostName)
'Set theHive = theHiveFile.getTheHive()
'Set theHive = theHost.getTheHive()
'Set theQueen = theHive.getTheQueenBySwarmID(swarm_id)
'Set theWorker = theQueen.getWorkerByID(worker_id)
'Set theWorker.Host = theHost
'theWorker.data.setProperty "targetedURL", targetedURL
'theWorker.data.setProperty "httpMethod", httpMethod
'theWorker.data.update





'Report status
'theWorker.jobIsStarted()

'發出 request
Set oXMLhttp = Wscript.CreateObject("MSXML2.XMLHTTP")
oXMLhttp.Open httpMethod, targetedURL, False
Wscript.Sleep 50
oXMLhttp.send
Wscript.Sleep 50

'讀取 response
result = oXMLhttp.responseText

'等一下
randomNumber = Int(Rnd * (800 + 1 - 350)) + 350
Wscript.Sleep randomNumber

'Call back
'theWorker.result = result

theHost.saveResult taskID, result

'Clean up
Set oXMLhttp = Nothing
'Set theQueen = Nothing
'Set theHive = Nothing
Set theHost = Nothing

Set theHiveFile = Nothing
Set oExcel = Nothing

'Report status
'theWorker.JobIsDone()
'Set theWorker = Nothing




