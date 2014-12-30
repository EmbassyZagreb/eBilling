<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<!--#include file="clsUpload.asp"-->
<html>
<head>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
<style type="text/css">
.tblMain { background-color:white;border-collapse:collapse;width:370px }
.tblMain td, .tblMain th {padding:5px;border:0px solid #000;font-size:13px }
.body {font-family:"Tahoma";font-size:16px }
</style>
<script language="Javascript">
function validateForm() {
    var x = document.forms["Uploading"]["txtFile"].value;
    if (x == null || x == "") {
        alert("Please select file for upload!");
        return false;
    }
}
EnableSubmit = function()
{
    var sbmt = document.getElementById("Disable");
    {
        sbmt.disabled = true;
		document.getElementById("Upload").submit();
    }
}

</script>
</HEAD>
<!--#include file="Header.inc" -->
<BODY>
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


<div>
<table class="tblMain">	

<%

Dim Upload
Dim Folder, FileFullPath


			Set Upload = New clsUpload

			Folder = Server.MapPath("Uploads") & "\"
			
			Upload("File1").SaveAs Folder & Upload("File1").FileName


'			Set Upload = Nothing
    Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(Server.MapPath("uploads"))
FilePath=Upload("File1").FileName
FileFullPath=objFolder & "\" & FilePath

%>
<form method="POST" name="Form1" id="Upload" action="ImportSpecUpload.asp"> 
<input type="Hidden" value="<%=FileFullPath%>" name="file1">
<%
    strFileYear = Mid(Mid(FilePath, InStrRev(FilePath, "_") + 1), 1, 4) 
    strFileMonth = Mid(Mid(FilePath, InStrRev(FilePath, "_") + 1), 5, 2)
    strFileName = Mid(Mid(FilePath, InStrRev(FilePath, "_") - 4), 1, 4) 


  if strFileName = "spec" then
    Set Rs = Server.CreateObject("ADODB.Recordset")
    Rs.open ("SELECT MonthP, YearP FROM dbo.CellPhoneHd where MonthP = '" & strFileMonth & "' and YearP = '" & strFileYear & "' GROUP BY MonthP, YearP"), BillingCon,1,3
    if rs.recordcount <> 0 then
      response.write "<tr><td colspan='2'>You uploaded CSV file for month and year that alredy exist in the system.</td></tr>"
      response.write "<tr><td colspan='2'>Please use different CSV file.</td></tr>"
      response.write "<tr><td colspan='2'><i>This page will refresh in 10 seconds.</i></td></tr>"
      response.write "<td><button type='cancel' onclick=""window.location='ImportSpec.asp';return false;"">Go Back</button></td>"
      rs.close()
      Response.AddHeader "REFRESH","8;URL=ImportSpec.asp"
    else
	response.write "<tr><td colspan='2'>You uploaded CSV file with following information:</td></tr>"
        response.write "<tr><td colspan='2'>File name:  " & strFileName & "</td></tr>"
        response.write "<tr><td colspan='2'>Month: " & strFileMonth & "</td></tr>"
        response.write "<tr><td colspan='2'>Year:  " & strFileYear & "</td></tr>"
        response.write "<tr><td colspan='2'>If this information is correct, please click on "%><button type="submit" id="Disable" form="Upload" onclick="EnableSubmit();" value="Submit">next</button> to continue.</td></tr><%
        response.write "<tr><td colspan='2'>Please be patient. This process can take up to 10 minutes to complete.</td></tr>"
        response.write "<tr><td colspan='2'><i>Do not close this page. It will refresh once import is finished.</i></td></tr>"
        rs.close()
    end if
  else
    Response.write "<tr><td colspan='2'>You uploaded wrong file name. Please upload <b>'spec_YYYYMM.csv'</b> file only.</td></tr>"
    response.write "<tr><td colspan='2'><i>This page will refresh in 5 seconds.</i></td></tr>"
    response.write "<td><button type='cancel' onclick=""window.location='ImportSpec.asp';return false;"">Go Back</button></td>"
    Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(Server.MapPath("uploads"))
    objFSO.deletefile objFolder & "/" & FilePath
    Response.AddHeader "REFRESH","6;URL=ImportSpec.asp"
  end if
%>
</table>
</div>
</form>
</body>
</html>