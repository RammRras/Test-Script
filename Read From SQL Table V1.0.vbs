
Dim sServer,sDataBaseName ,sConn,oConn,oRS
Dim sUserName,sPassWord 
Dim mRecordstr
Dim field
Dim i,j 

sServer = "192.168.40.3"
sDataBaseName = "REGISTROPLC"
sUserName = "plcindustrial"
sPassWord = "plcindustrial"
 

'sConn="DRIVER={SQL Server};SERVER=" & sServer & ";DATABASE=" & sDataBaseName & ";Encrypt=NO;"
sConn="Driver={SQL Server Native Client 10.0};Server=tcp:ve8o13q17o.database.windows.net,1433;Database=az_database;Uid=Tecnico@ve8o13q17o;Pwd=az_Merda123;Encrypt=yes;Connection Timeout=30;"

Set oConn = CreateObject("ADODB.Connection")
'oConn.CommandTimeout = 36000
oConn.Open sConn', sUserName, sPassWord
 
WScript.Echo "oConn.Open = " & CStr(sConn)

Set FetchData = CreateObject("ADODB.Recordset")
FetchData.open "SELECT * FROM INDUSTRIALLLENADO", oConn, 3



mRecordstr = "" 

WScript.Echo "Total Fields = " & CStr(FetchData.Fields.Count)

For i = 0 To FetchData.Fields.Count -1
	mRecordstr = mRecordstr & FetchData.Fields.Item(i).Name & vbTab
Next

mRecordstr = mRecordstr & vbNewLine

While not FetchData.eof
'	For i = 0 To FetchData.Fields.Count -1
'		mRecordstr = mRecordstr & FetchData.Fields.Item(i) & vbtab
'	Next
	
	For each field in FetchData.Fields
		mRecordstr = mRecordstr & field.value & vbtab
	Next 

	mRecordstr = mRecordstr & vbnewline
	
	FetchData.MoveNext 
Wend

	WScript.Echo mRecordstr


FetchData.Close
oConn.Close
