# Hive
Hive, for distributing tasks from Excel VBA to Windows Scripting Host.
Hive 是一個蜂巢，結構為Hive-Swarms-Swarm-Queen-Worker。
可使用Excel VBA (Host)透過 Hive來將工作分散，產生許多Wsh scripts，透過Windows Scripting Host產生多個processes來執行，再將結果彙整到Hive中，是一種類似Map-Reduce的模式。
使用範例可見於 Test/TestHost。
