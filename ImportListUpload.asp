<!--#include file="connect.inc" -->
<%Server.ScriptTimeout=600%>
<html>
<head>
<title>Import CSV file to SQL</title>
</head>
<body>
<Center>

<%
Dim objFSO,CSVfile,sRows,arrRows
Dim strSQL,objExec
dim strFileYear, strFileMonth, strFileSPEC, PathToFile, testing, filepath 


'** Getfile Name ***'

PathToFile = request.form("file1")

'*** Create Object ***
Set objFSO = CreateObject("Scripting.FileSystemObject")

'*** Open Files ***'
Set CSVfile = objFSO.OpenTextFile(PathToFile,1)

'*** Strip month and year from file name ***
  strFileYear = Mid(Mid(PathToFile, InStrRev(PathToFile, "_") + 1), 1, 4) 'Right(CSVfile, Len(CSVfile) - InStrRev(strPath, "\"))
  strFileMonth = Mid(Mid(PathToFile, InStrRev(PathToFile, "_") + 1), 5, 2)


Set objExec = BillingCon.Execute("DELETE From ListTEMP;")




On Error Resume Next
'Set CSVfile = objFSO.OpenTextFile(PathToFile,1,2)
sRows = CSVfile.readLine
Do while not CSVfile.AtEndOfStream '.ReadLine = lngLineCount - 100000 '
	sRows = CSVfile.readLine
	arrRows = Split(sRows,";")

'*** Insert to table ListTEMP ***'
	strSQL = ""
	strSQL = strSQL &"INSERT INTO ListTEMP "
	strSQL = strSQL &"(field2, field3, field4, field5, field6, field7, field8, field9, field10, field11, field12, field13, field14, field15, field16, field17, field18, field19, MonthP, YearP) "
	strSQL = strSQL &"VALUES "
	strSQL = strSQL &"("&replace(arrRows(1),",",".")&","&replace(arrRows(2),",",".")&","&replace(arrRows(3),",",".")&","&replace(arrRows(4),",",".")&","&replace(arrRows(5),",",".")&","&replace(arrRows(6),",",".")&" "
	strSQL = strSQL &","&replace(arrRows(7),",",".")&","&replace(arrRows(8),",",".")&","&replace(arrRows(9),",",".")&","&replace(arrRows(10),",",".")&","&replace(arrRows(11),",",".")&","&replace(arrRows(12),",",".")&" "
	strSQL = strSQL &","&replace(arrRows(13),",",".")&","&replace(arrRows(14),",",".")&","&replace(arrRows(15),",",".")&","&replace(arrRows(16),",",".")&","&replace(arrRows(17),",",".")&","&replace(arrRows(18),",",".")&" "
	strSQL = strSQL &",'"& strFileMonth &"','"& strFileYear &"') "

	Set objExec = BillingCon.Execute(strSQL)
	Set objExec = Nothing
	i=i+1

Loop
On Error GoTo 0
'response.write "Imported rows from CSV file: " & i & " of " & lngLineCount & "<br><br>"



'Set objExec = BillingCon.Execute("DELETE From ImportTEMP;")

'Response.write ("CSV import completed.")


CSVfile.Close()
BillingCon.Close()
Set CSVfile = Nothing
Set BillingCon = Nothing

'*** Delete uploaded file ***
dim fs
Set fs=Server.CreateObject("Scripting.FileSystemObject")
if fs.FileExists(PathToFile)=true then
  fs.DeleteFile(PathToFile)
'response.write "Imported file deleted!"
else
'response.write "Imported file NOT deleted!"
end if
set fs=nothing

Response.AddHeader "REFRESH","0;URL=ImportListView.asp"

%>
</body>
</html>
