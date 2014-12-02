<%@ Language=VBScript%>
<% ' VI 6.0 Scripting Object Model Enabled %>
<%
Response.ContentType ="application/vnd.ms-excel" 
Response.Buffer  =  True 
Response.Clear() 
%> 
<html>
<head>

<META name=VI60_defaultClientScript content=VBScript>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<!--#include file="connect.inc" -->

<style type="text/css">
<!--

.FontText {
	font-size: small;
}


.Hint{
	color: gray;
	font-size: x-small;
}
.style5 {color: #FFFFFF;};
-->
</style>
<%
Dim user_ , user1_

user_ = request.servervariables("remote_user")
user1_ = user_  'user1_ = right(user_,len(user_)-4)
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

curMonth_ = month(date())
curYear_ = year(date())
if len(curMonth_)= 1 then
	curMonth_ = "0" & curMonth_
end if

sMonthP = request("sMonthP")
'response.write sMonthP

sYearP = request("sYearP")

eMonthP = request("eMonthP")

eYearP = request("eYearP")

Agency_ = request("Agency")

Section_ = request("Section")

EmpID_ = request("EmpID")

Status = request("Status")
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

strsql = "Exec spRptAging '" & sMonthP & "','" & sYearP & "','" & eMonthP & "','" & eYearP & "','" & Agency_ & "','" & Section_ & "','" & EmpID_ & "'," & Status
'response.write strsql & "<br>"
set DataRS = server.createobject("adodb.recordset")
'DataRS.CursorLocation = 3
'DataRS.Open strsql,BillingCon
set DataRS=BillingCon.execute(strsql)

if not DataRS.eof Then

%>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" width="100%"  class="FontText">
    <TR BGCOLOR="#330099" align="center">
         <TD rowspan="2" width="3%" align="right" class="style5">No.</TD>
         <TD rowspan="2" width="15%" class="style5">Employee Name</TD>
         <TD rowspan="2" width="10%" class="style5">Billing Period</TD>
         <TD rowspan="2" width="8%" class="style5">Section</label></strong></TD>
	 <TD colspan="5" class="style5">Billing Amount (Kn.)</TD>
         <TD rowspan="2" class="style5">Aging</TD>
         <TD rowspan="2" class="style5">Status</TD>
    </TR>
    <tr BGCOLOR="#330099" align="center">
         <TD width="10%" class="style5">Home Phone</TD>
         <TD width="10%" class="style5">Office Phone</TD>
         <TD width="10%" class="style5">Mobile Phone</TD>
         <TD width="10%" class="style5">Shuttle Bus</TD>
         <TD width="8%" class="style5">Total</TD>
    </tr>

<% 
   dim no_  
   no_ = 1
   do while not DataRS.eof
	   if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
%>
      
	   <TR bgcolor="<%=bg%>">
	        <TD align="right"><FONT color=#330099 size=2><%=no_ %></font></TD>
	        <TD><%=DataRS("EmpName") %></TD>
	        <TD align="right"><FONT color=#330099 size=2><%= DataRS("MonthP")%>-<%= DataRS("YearP")%></font>&nbsp;</TD>
	        <TD><FONT color=#330099 size=2><%=DataRS("Office") %> </font></TD>
<!--	        <TD align="right"><FONT color=#330099 size=2><%= formatnumber(DataRS("HomePhoneBillRp"),0) %></font></TD>-->
		<td align="right">
<%		If CDbl(DataRS("HomePhonePrsBillRp")) > 0 Then %>
			<%= formatnumber(DataRS("HomePhonePrsBillRp"),-1) %>
<%		Else %>
			-
<%		End If %>
		</td>
		<td align="right">
<%		If CDbl(DataRS("OfficePhonePrsBillRp")) > 0 Then %>
			<%= formatnumber(DataRS("OfficePhonePrsBillRp"),-1) %>
<%		Else %>
			-
<%		End If %>
		</td>
		<td align="right">
<%		If CDbl(DataRS("CellPhonePrsBillRp")) > 0 Then %>
			<%= formatnumber(DataRS("CellPhonePrsBillRp"),-1) %>
<%		Else %>
			-
<%		End If %>
		</td>
		<td align="right">
<%		If CDbl(DataRS("TotalShuttleBillRp")) > 0 Then %>
			<%= formatnumber(DataRS("TotalShuttleBillRp"),-1) %>
<%		Else %>
			-
<%		End If %>
		</td>
	        <TD align="right"><FONT color=#330099 size=2><%= formatnumber(DataRS("TotalBillingRp"),-1) %></font></TD>
		<TD><FONT color=#330099 size=2><%= DataRS("Aging") %></font></TD>
		<TD><FONT color=#330099 size=2><%= DataRS("ProgressDesc") %></font></TD>
	   </TR>

<%   
	   DataRS.movenext
	   no_ = no_ + 1 
   loop 
	PageNo=1
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
	<tr>
        	<td><br></TD>
	</tr>
	<tr>
		<td align="center"><a href="Default.asp"><img src="images/Back.gif" border="0" alt="Go..Back" WIDTH="83" HEIGHT="25"></a></td>
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


