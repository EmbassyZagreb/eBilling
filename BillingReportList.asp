<%@ Language=VBScript%>
<% ' VI 6.0 Scripting Object Model Enabled %> 
<html>
<head>

<META name=VI60_defaultClientScript content=VBScript>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<!--#include file="connect.inc" -->


<script language="JavaScript" src="calendar.js"></script>
</head>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">BILLING REPORT LIST</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>

<%

dim rs 
dim strsql
dim tombol
dim hlm
%>

<%
Dim user_ , user1_

user_ = request.servervariables("remote_user")
user1_ = right(user_,len(user_)-4)
'response.write user1_ & "<br>"

strsql = "select RoleID from Users where loginId ='" & user1_ & "'"
set UserRS = server.createobject("adodb.recordset")
'response.write strsql & "<br>"
set UserRS = BillingCon.execute(strsql)
if not UserRS.eof then
	UserRole_ = UserRS("RoleID")
Else
	UserRole_ = ""
end if

If (UserRole_ <> "") Then
%>
<table align="center" cellpadding="1" cellspacing="0" width="100%">
<tr>
	<td>
	   <UL><UL TYPE="square">
		<LI>&nbsp;<A HREF="HomePaymentBillingReport.asp"><b>Home Payment Billing Report</b></A>
		<LI>&nbsp;<A HREF="LongDistanceCallsReport.asp"><b>Long Distance Calls Report (Office phone)</b></A>
		<LI>&nbsp;<A HREF="CellphoneCallsReport.asp"><b>Cell phone Calls Report (Office phone)</b></A>
		<LI>&nbsp;<A HREF="LineUsagesReport.asp"><b>Line Usages Report (Office phone)</b></A>
		<LI>&nbsp;<A HREF="NoLineUsagesReport.asp"><b>No Line Usages Report (Office phone)</b></A>
		<LI>&nbsp;<A HREF="MonthlyBillingByExtReport.asp"><b>Comparison Monthly Bill by Ext. (Office phone)</b></A>
	   </UL></UL>
	</td>
</tr>
</table>
<%Else%>
	<table>
		<tr>
			<td>You do not have permission to access this site.</td>
		</tr>
		<tr>
			<td>Please <a href="http://zagrebws03.eur.state.sbu/WebPASS/eservices/MainPage.asp">Submit Request </a> or contact Zagreb ISC Helpdesk at ext.3333.</td>
		</tr>
	</table>
<% end if %>

</body> 

</html>

