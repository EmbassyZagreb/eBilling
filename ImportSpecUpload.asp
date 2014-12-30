<!--#include file="connect.inc" -->
<%Server.ScriptTimeout=800%>
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

Const adTypeBinary = 1
Const adTypeText = 2

' accept a string and convert it to Bytes array in the selected Charset
Function StringToBytes(Str,Charset)
  Dim Stream : Set Stream = Server.CreateObject("ADODB.Stream")
  Stream.Type = adTypeText
  Stream.Charset = Charset
  Stream.Open
  Stream.WriteText Str
  Stream.Flush
  Stream.Position = 0
  ' rewind stream and read Bytes
  Stream.Type = adTypeBinary
  StringToBytes= Stream.Read
  Stream.Close
  Set Stream = Nothing
End Function

' accept Bytes array and convert it to a string using the selected charset
Function BytesToString(Bytes, Charset)
  Dim Stream : Set Stream = Server.CreateObject("ADODB.Stream")
  Stream.Charset = Charset
  Stream.Type = adTypeBinary
  Stream.Open
  Stream.Write Bytes
  Stream.Flush
  Stream.Position = 0
  ' rewind stream and read text
  Stream.Type = adTypeText
  BytesToString= Stream.ReadText
  Stream.Close
  Set Stream = Nothing
End Function
'  line 50
' This will alter charset of a string from 1-byte charset(as windows-1252)
' to another 1-byte charset(as windows-1251)
Function AlterCharset(Str, FromCharset, ToCharset)
  Dim Bytes
  Bytes = StringToBytes(Str, FromCharset)
  AlterCharset = BytesToString(Bytes, ToCharset)
End Function

'** Getfile Name ***'

'filepath = request.form("file1")
PathToFile = request.form("file1")


'*** Create Object ***
Set objFSO = CreateObject("Scripting.FileSystemObject")

'*** Open Files ***'
Set CSVfile = objFSO.OpenTextFile(PathToFile,1)


    strFileYear = Mid(Mid(PathToFile, InStrRev(PathToFile, "_") + 1), 1, 4) 
    strFileMonth = Mid(Mid(PathToFile, InStrRev(PathToFile, "_") + 1), 5, 2)


'response.write strFileYear & "<br>"
'response.write strFileMonth & "<br>"

Set objExec = BillingCon.Execute("DELETE From importRAW;")

On Error Resume Next
'---------- Create Temp Table -----------------
Set objExec = BillingCon.Execute("Drop Table ImportTEMP")
Set objExec = BillingCon.Execute("spCreateTEMPtable")
On Error GoTo 0

  dim i : i=0
 'Loop through counting the lines
       Dim lngLineCount : lngLineCount = 0   '   line 100
       Do While Not CSVfile.AtEndOfStream
           lngLineCount = lngLineCount + 1
           CSVfile.SkipLine()
       Loop

Set CSVfile = Nothing

'On Error Resume Next
Set CSVfile = objFSO.OpenTextFile(PathToFile,1,2)
'sRows = CSVfile.readLine + 170000
Do Until CSVfile.AtEndOfStream '.ReadLine = lngLineCount - 100000 '
	sRows = CSVfile.readLine
	arrRows = Split(sRows,";")

'*** Insert to table importRAW ***'
	strSQL = ""
	strSQL = strSQL &"INSERT INTO importRAW "
	strSQL = strSQL &"(MonthP, YearP, PhoneNumber,DialedDateTime,CallDuration,Cost,Discount,CallType,DialedNumber) "
	strSQL = strSQL &"VALUES "
	strSQL = strSQL &"('" & strFileMonth & "','" & strFileYear & "',"&arrRows(3)&",'"&(Mid(arrRows(4),4,2) & "/" & Left(arrRows(4),2) & "/" & Mid(arrRows(4),7,4) & " " & Mid(arrRows(4),12,8))&"','"&arrRows(5)&"' "
	strSQL = strSQL &","&replace(arrRows(7), ",", ".")&","&replace(arrRows(8), ",", ".")&" "
	strSQL = strSQL &",'"& AlterCharset(arrRows(11),"windows-1252","UTF-8") &"','"&arrRows(12)&"') "

	Set objExec = BillingCon.Execute(strSQL)
	Set objExec = Nothing
	i=i+1

Loop
On Error GoTo 0
'response.write "Imported rows from CSV file: " & i & " of " & lngLineCount & "<br><br>"

'  Update MultiSIM data in ImportRaw table
Set objExec = BillingCon.Execute("spMultiSIMUpdate")

' Copy to ImportEMP table for further processing
Set objExec = BillingCon.Execute("spCopyToImportTEMP")

CSVfile.Close()
BillingCon.Close()
Set CSVfile = Nothing
Set BillingCon = Nothing

dim fs
Set fs=Server.CreateObject("Scripting.FileSystemObject")
if fs.FileExists(PathToFile)=true then
  fs.DeleteFile(PathToFile)
'response.write "Imported file deleted!"
else
'response.write "Imported file NOT deleted!"
end if
set fs=nothing

Response.AddHeader "REFRESH","0;URL=ImportSpecView.asp"

%>
</body>
</html>
