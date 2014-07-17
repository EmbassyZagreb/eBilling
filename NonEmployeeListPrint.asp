<%@ Language=VBScript%>
<% ' VI 6.0 Scripting Object Model Enabled %> 
<html>
<head>
<%
Response.ContentType ="application/vnd.ms-excel" 
Response.Buffer  =  True 
Response.Clear() 
%>  
<META name=VI60_defaultClientScript content=VBScript>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<!--#include file="connect.inc" -->


<script language="JavaScript" src="calendar.js"></script>
<%

Name_ = Trim(request("Name"))
Status_ = Trim(request("Status"))
%>

<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
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

	strsql = "Select * from vwNonEmployeeList Where NonEmpName like '%" & Name_ & "%' and (Status='" & Status_ & "' or '" & Status_ & "'='X') order by NonEmpName"
	set DataRS = server.createobject("adodb.recordset")
	'response.write strsql & "<br>"
'	DataRS.CursorLocation = 3
	DataRS.open strsql,BillingCon

%>
<form method="post" name="frmPaymentList" action="">
<table align="center" cellpadding="1" cellspacing="0" width="90%" border="1" bordercolor="black"> 
<TR align="center" cellpadding="0" cellspacing="0" >
	<TD width="30px"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
	<TD><strong><label STYLE=color:#FFFFFF>Emp ID</label></strong></TD>
	<TD><strong><label STYLE=color:#FFFFFF>Name</label></strong></TD>
        <TD><strong><label STYLE=color:#FFFFFF>Agency Funding</label></strong></TD>
	<TD><strong><label STYLE=color:#FFFFFF>Email</label></strong></TD>
	<TD><strong><label STYLE=color:#FFFFFF>Remark</label></strong></TD>
       	<TD width="80px"><strong><label STYLE=color:#FFFFFF>Status</label></strong></TD>
</TR>    
<% 
	dim no_  
	no_ = 1 

	do while not DataRS.eof
   	if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
%>
   	<TR bgcolor="<%=bg%>">
		<td align="right"><%=No_%></td>
	        <td><FONT color=#330099 size=2><%= DataRS("NonEmpId") %></font></td>
	        <td><FONT color=#330099 size=2><%= DataRS("NonEmpName") %></font></td>
        	<td><FONT color=#330099 size=2><%=DataRS("AgencyDesc")%></font></td> 
        	<td><FONT color=#330099 size=2><%=DataRS("Email")%></font></td> 
        	<td><FONT color=#330099 size=2><%=DataRS("Remark")%></font></td> 
        	<td><FONT color=#330099 size=2><%=DataRS("StatusName")%></font></td>
		</td> 
  	 </TR>
<%   
		Count=Count +1
 		DataRS.movenext
   		no_ = no_ + 1
	loop
%>
</table>
</form>
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


