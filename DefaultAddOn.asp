<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<html>
<head>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
<form method="post"> 
<% 
 dim user_ 
 dim user1_  
 dim rst 
 dim strsql
 'response.end
 user_ = request.servervariables("remote_user") 
'response.write user_ & "<br>"
  user1_ = right(user_,len(user_)-4)

if (session("Month") = "") or (session("Year") = "") then
	strsql = "Select MonthP, YearP From Period"
	'response.write strsql & "<br>"
	set rsData = server.createobject("adodb.recordset") 
	set rsData = BillingCon.execute(strsql)
	if not rsData.eof then
		session("Month") = rsData("MonthP")
		session("Year") = rsData("YearP")
	end if
end if

strsql = "select RoleID from Users where loginId ='" & user1_ & "'"
set UserRS = server.createobject("adodb.recordset")
'response.write strsql & "<br>"
set UserRS = BillingCon.execute(strsql)
if not UserRS.eof then
	UserRole_ = UserRS("RoleID")
end if
%>       

<table cellspadding="1" cellspacing="1" width="500px">  
<tr>
	<td align="left" valign="top">
	   <UL><UL TYPE="square">
		<LI class="normal"><B>Employee</b>
			<UL TYPE="">
				<li><A HREF="MonthlyBillingAddOn.asp"><B>Monthly Billing</B></A></li>
				<li><A HREF="MonthlyBillingTrackingAddOn.asp"><B>Tracking Status</B></A></li>
				<li><A HREF="MonthlyBillListAddOn.asp"><B>Monthly Billing Report</B></A></li>
				<li><A HREF="UserSettingAddOn.asp"><B>Supervisor Setting</B></A></li>
				<li><A HREF="PersonalPhoneListAddOn.asp"><B>Personal Phone List</B></A></li>
			</ul>	
		</li>
		   </UL>
		</UL>
	</td>
	<td align="left" valign="top">
	   <UL>
		<UL TYPE="square">
		<LI class="normal"><B>Supervisor</B>
			<UL TYPE="">
				<li><A HREF="BillingApprovalListAddOn.asp"><B>Approval List</B></A></li>
			</ul>	
		</li>
		   </UL></UL>
	</td>
</tr>

</table>

</form>
</BODY>
</html>