<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<!--#include file="clsUpload.asp" -->
<!DOCUMENT html>

<html>
<head>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
<style type="text/css">
.tblMain { background-color:white;border-collapse:collapse;width:250px }
.tblMain td, .tblMain th {padding:10px;border:0px solid #000;font-size:13px }
.body {font-family:"Tahoma";font-size:16px }
</style>

</HEAD>

<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">Import New Bill</TD>
   </TR>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>

<body>
<table class="tblMain">
<%
Dim objFSO, objFile, objFolder
Dim rs, objExec 

Dim Upload, Folder, FileFullPath

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(Server.MapPath("uploads"))

For Each objFile in objFolder.Files
objFile.delete
Next
Set objFolder = Nothing
Set objFSO = Nothing

Set objExec = BillingCon.Execute("DELETE From ListTEMP;")
%>
<FORM method="post" encType="multipart/form-data" action="ImportListSave.asp">
<tr><td colspan="2"><b>Upload list_YYYMM.csv file here:</b></td></tr>
			<tr><td colspan="2"><INPUT type="File" name="File1">
</td></tr>
<tr><td colspan="2" align="left">
			<INPUT type="Submit" value="Upload"></td>

</td>
</tr>
</TABLE>
</FORM>

</BODY>
</html>