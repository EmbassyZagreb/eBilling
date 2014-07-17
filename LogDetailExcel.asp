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

MonthP_ = Request("MonthP")
YearP_ = Request("YearP")
ProgressID_ = Request("ProgressID")

%>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY BGCOLOR="#FFFFFF" TEXT="#000000" LINK="#0080FF" ALINK="0080FF" VLINK="#0055AA" MARGINWIDTH="8" MARGINHEIGHT="0" LEFTMARGIN="8" TOPMARGIN="0">

<%

dim rs 
dim strsql
If (UserRole_ = "Admin") or (UserRole_ = "Voucher") or (UserRole_ = "FMC") or (UserRole_ = "Cashier")  Then
%>
<%

strsql = "Exec spLogView '" & MonthP_ & "','" & YearP_ & "'," & ProgressID_ 
'response.write strsql & "<br>"

set DataRS = server.createobject("adodb.recordset")
DataRS.CursorLocation = 3
DataRS.Open strsql,BillingCon
'set DataRS=BillingCon.execute(strsql)

'response.write strsql

if not DataRS.eof Then
%>
<table bordercolor="#000000" cellpadding="2" cellspacing="0" width="650px">
<tr>
         <TD colspan="2"><strong>Log Type :</strong></TD>
	 <td colspan="2"><%=DataRS("ProgressDesc") %></td>
</tr>
</table>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" width="650px"  class="FontText">
    <TR  align="center">
         <TD width="3%" align="right"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
         <TD width="100px"><strong><label STYLE=color:#FFFFFF>Billing Period</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Employee ID</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Employee Name</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Phone Number</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Balance</label></strong></TD>
    </TR>
<% 
   dim no_  
   no_ = 1
   Count=1
   TotRecordNo_ = 0
   TotBillingAmount_ = 0
   do while not DataRS.eof
	   if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
%>    
	   <TR bgcolor="<%=bg%>">
	        <TD align="right"><%=no_ %></TD>
	        <TD align="right"><%= DataRS("MonthP")%>-<%= DataRS("YearP")%>&nbsp;</TD>
<!--	        <TD><%=DataRS("ProgressDesc") %></TD> -->
	        <TD><%=DataRS("EmpID") %></TD>
	        <TD><%=DataRS("EmpName") %></TD>
	        <TD>&nbsp;<%=DataRS("MobilePhone") %></TD>
	        <TD align="right"><%=formatnumber(DataRS("TotalBillingRp"),-1) %></TD>
	   </TR>

<%   
		TotRecordNo_ = cdbl(TotRecordNo_) + 1
		TotBillingAmount_ = cdbl(TotBillingAmount_) + cdbl(DataRS("TotalBillingRp"))
		Count=Count +1
	   DataRS.movenext
	   no_ = no_ + 1 
	if DataRS.eof then
%>
	<TR bgcolor="<%=bg%>">
	        <TD align="center" colspan="5"><b>Total</b></TD>
<!--	        <TD align="right"><b><%=formatnumber(TotRecordNo_,-1) %></b></TD> -->
	        <TD align="right"><b><%=formatnumber(TotBillingAmount_,-1) %></b></TD>
	   </TR>
<%	
	end if   
   loop 
%>
</table>
<%
else 
%>
	<table cellspadding="1" cellspacing="0" width="100%">  
	<tr>
        	<td><br></TD>
	</tr>
	<tr>
		<td align="center">There is no data.</td>
	</tr>
	</table>
<% end if %>
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


