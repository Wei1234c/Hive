
Option Explicit

'�ܼƫŧi ____________________________
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



'�ܼƽ�� ____________________________





Hive_fileName = "- HIVE.xls"




swarm_id = 1


worker_id = 1


hostName = "- ��ƻ`��.xls"



taskID = "4711863218537"
targetedURL = "http://ipac.library.taichung.gov.tw/toread/opac/Advancedsearch.page?classification_type=all&level=all&limit=20&location=3&material_type=all&q=isbn_issn%3A4711863218537&source=local&view=CONTENT&wi=false"
httpMethod = "GET"

callBackFunction = "saveResult taskID, result"






'�D�n�޿� ____________________________

'��� Excel.Application ����
Set oExcel = GetObject(, "Excel.Application")
oExcel.visible = True

'�إߪ���

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

'�o�X request
Set oXMLhttp = Wscript.CreateObject("MSXML2.XMLHTTP")
oXMLhttp.Open httpMethod, targetedURL, False
Wscript.Sleep 50
oXMLhttp.send
Wscript.Sleep 50

'Ū�� response
result = oXMLhttp.responseText

'���@�U
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




