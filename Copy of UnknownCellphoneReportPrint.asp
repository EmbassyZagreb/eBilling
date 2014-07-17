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

-->
</style>
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

sMonthP = request("sMonthP")
sYearP = request("sYearP")
eMonthP = request("eMonthP")
eYearP = request("eYearP")

%>
<TITLE>U.S. Mission Jakarta e-Billing</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">

<STYLE TYPE="text/css"><!--
  A:ACTIVE { color:#003399; font-size:8pt; font-family:Verdana; }
  A:HOVER { color:#003399; font-size:8pt; font-family:Verdana; }
  A:LINK { color:#003399; font-size:8pt; font-family:Verdana; }
  A:VISITED { color:#003399; font-size:8pt; font-family:Verdana; }
  body {scrollbar-3dlight-color:#FFFFFF; scrollbar-arrow-color:#E3DCD5; scrollbar-base-color:#FFFFFF; scrollbar-darkshadow-color:#FFFFFF;	scrollbar-face-color:#FFFFFF; scrollbar-highlight-color:#E3DCD5; scrollbar-shadow-color:#E3DCD5; }
  p { font-family: verdana; font-size: 12px; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; color: #003399; text-decoration: none}
  h3 { font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 16px; font-style: normal; line-height: normal; font-weight: bold; color: #003399; letter-spacing: normal; word-spacing: normal; font-variant: small-caps}
  td { font-family: verdana; font-size: 10px; font-style: normal; font-weight: normal; color: #000000}
  .title { font-size:14px; font-weight:bold; color:#000080; }
  .SubTitle { font-size:16px; font-weight:bold; color:#000080;  }
  A.menu { text-decoration:none; font-weight:bold; }
  A.mmenu { text-decoration:none; color:#FFFFFF; font-weight:bold; }
  .normal { font-family:Verdana,Arial; color:black}
  .style5 {color: #FFFFFF;}
  .ActivePage {color: red; font-weight:bold; }
--></STYLE>
</HEAD>
<BODY BGCOLOR="#FFFFFF" TEXT="#000000" LINK="#0080FF" ALINK="0080FF" VLINK="#0055AA" MARGINWIDTH="8" MARGINHEIGHT="0" LEFTMARGIN="8" TOPMARGIN="0">
<%

dim rs 
dim strsql
If (UserRole_ = "Admin") or (UserRole_ = "Voucher") or (UserRole_ = "FMC") or (UserRole_ = "Cashier")  Then

sPeriod = sYearP&sMonthP
ePeriod = eYearP&eMonthP
'response.write sPeriod & ePeriod 
strsql = "Select * From vwUnknownCellphoneBill Where YearP+MonthP>='" & sPeriod & "' and YearP+MonthP<='" & ePeriod & "'"

strsql = strsql  & strFilter & " order by PhoneNumber"
'response.write strsql & "<br>"

set DataRS = server.createobject("adodb.recordset")
DataRS.CursorLocation = 3
DataRS.Open strsql,BillingCon
'set DataRS=BillingCon.execute(strsql)

if not DataRS.eof Then
%>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" width="60%"  class="FontText">
    <TR align="center">
         <TD width="3%" align="right"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
         <TD width="100px"><strong><label STYLE=color:#FFFFFF>Billing Period</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Phone Number</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Employee Name</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Office</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Current Balance</label></strong></TD>
    </TR>
<% 
   dim no_  
   no_ = 1
   Count=1
   do while not DataRS.eof
	   if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
%>    
	   <TR bgcolor="<%=bg%>">
	        <TD align="right"><FONT color=#330099 size=2><%=no_ %></font></TD>
	        <TD align="right"><FONT color=#330099 size=2><%= DataRS("MonthP")%>-<%= DataRS("YearP")%></font></TD>
	        <TD><FONT color=#330099 size=2><%=DataRS("PhoneNumber") %> </font></TD>
	        <TD><FONT color=#330099 size=2><%=DataRS("EmpName") %> </font></TD>
	        <TD><FONT color=#330099 size=2><%=DataRS("Office") %> </font></TD>
	        <TD align="right"><FONT color=#330099 size=2><%= formatnumber(DataRS("CurrentBalance"),0) %></font></TD>
	   </TR>

<%   
		Count=Count +1
	   DataRS.movenext
	   no_ = no_ + 1 
   loop 
%>
</table>
<%
else 
%>
	<table cellspadding="1" cellspacing="0" width="100%">  
	<tr>
		<td align="center">There is no data</td>
	</tr>
	</table>
<% end if %>
<%Else%>
	<table>
		<tr>
			<td>You do not have permission to access this site.</td>
		</tr>
		<tr>
			<td>Please <a href="http://jakartaws01.eap.state.sbu/CSC">Submit Request </a> or contact Jakarta CSC Helpdesk at ext.9111.</td>
		</tr>
	</table>
<% end if %>
</body> 

</html>

