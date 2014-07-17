<%@ Language=VBScript %>
<!--#include file="connect.inc" -->

<html>
   <head>
   <script type="text/javascript">
	function refreshForm()
	{
		opener.location.reload();
		window.close();
	}
</script>
<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  NAME="CDO for Windows 2000 Library" --> 
<%
EmpID_ = trim(request.form("txtEmpID"))
if EmpID_  = "" then
	EmpID_ = 0
end if
State_ = trim(request.form("State"))
ShuttleDate_ = trim(request.form("txtShuttleDate"))
AM_ = trim(request.form("txtAM"))
PM_ = trim(request.form("txtPM"))

'response.write LoginID_ & "<br>"
'response.write UserRole_ & "<br>"

%>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">Shuttle Bus Schedule Update</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
<table border=0 width=100%>
<%
       strsql = "Update ShuttleUserSchedule Set AM=" & AM_ & ", PM =" & PM_ & " Where EmpID= " & EmpID_ & " And ShuttleDate='" & ShuttleDate_ & "'"
      'response.write strsql 
       BillingCon.execute strsql
%>               
<tr><td align=center>Your data has already been saved. Thank you.</td></tr>
<tr><td>&nbsp;</td>
<tr>
	<td align="center"><input type="button" value="Close" name="btnClose" onclick="refreshForm();"></td>
</tr>
</table>

   </body>
</html>