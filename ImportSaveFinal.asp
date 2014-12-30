<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<!--#include file="Header.inc" -->
<html>
<head>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">

</HEAD>

<BODY>
<body>	<TR>
		<TD COLSPAN="4" ALIGN="center" Class="title">Import New Bill</TD>
	</TR>
	<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
	</tr>
	<TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
  </TR>
</TABLE>
<%
Dim objExec

' Transfer list data to CellPhoneHd
Set objExec = BillingCon.Execute("spCopyListFinal")

' Transfer list data to CellPhoneDt
Set objExec = BillingCon.Execute("spCopySpecFinal")

'Cleanup of temp tables
Set objExec = BillingCon.Execute("DELETE From importRAW;")
Set objExec = BillingCon.Execute("DELETE From ListTEMP;")
Set objExec = BillingCon.Execute("Drop Table ImportTEMP")

%>
<div>
<center>
<table>	
<tr><td>You have successfully import new data in to zBilling application.</td></tr>
<tr><td>Next step is to generate monthly bill for this newly imported month.</td></tr>
</table>
<br>
<table>	
<tr><button type="submit" onclick="window.location='GenerateMonthlyBill.asp'">Generate monthly bill</button></tr>
</table>
</div>
</form>
</body>
</html>