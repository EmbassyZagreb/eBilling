<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<!--#include file="clsUpload.asp" -->
<html>
<head>

<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
<style type="text/css">
.tblMain { background-color:white;border-collapse:collapse;width:270px }
.tblMain td, .tblMain th {padding:10px;border:0px solid #000;font-size:13px }
.body {font-family:"Tahoma";font-size:16px }
</style>

</HEAD>
<%
dim user_ 
 dim user1_  
 dim rst 
 dim strsql
 
user_ = request.servervariables("remote_user")
user1_ = user_  'user1_ = right(user_,len(user_)-4)
strsql = "select * from Users where loginId='" & user1_ & "'"
set RS_Query = server.createobject("adodb.recordset")
'response.write strsql & "<br>"
set RS_Query = BillingCon.execute(strsql)

if not RS_Query.eof then
	UserRole_ = RS_Query("RoleID")
end if
%>
<!--#include file="Header.inc" -->
<body>
  <TR>
		<TD COLSPAN="4" ALIGN="center" Class="title">Import New Bill</TD>
   </TR>
	<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
	</tr>
  <TR>
		<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
  <%if (UserRole_= "Admin") then %>
  
<table class="tblMain">
<%
Dim objFSO, objFile, objFolder
Dim rs

Dim Upload, Folder, FileFullPath

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(Server.MapPath("uploads"))

For Each objFile in objFolder.Files
objFile.delete
Next
Set objFolder = Nothing
Set objFSO = Nothing

%>
<FORM method="post" encType="multipart/form-data" action="ImportSpecSave.asp">
<tr><td colspan="2">Upload <b>spec_YYYMM.csv</b> file here:</td></tr>
			<tr><td colspan="2"><INPUT type="File" name="File1">
</td></tr>
<tr><td colspan="2" align="left">
			<INPUT type="Submit" value="Upload"></td>

</td>
</tr>
</TABLE>
</FORM>
<%
else
%>
<br><br>
<!--#include file="NoAccess.asp" -->
<%end if %>
</BODY>
</html>