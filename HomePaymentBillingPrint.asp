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
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
<!--#include file="connect.inc" -->

<script language="JavaScript">
function ClearFilter()
{
	document.forms['frmSearch'].elements['PostList'].value ='';
	document.forms['frmSearch'].elements['SortList'].value ='Homephone';
}
</script>


<%
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

Post_ = Request("Post")	
PhoneNumber_ = Request("PhoneNumber")
if PhoneNumber_ <> "A" Then
	PhoneNumber_= "+" & trim(PhoneNumber_)
end if
MonthP_ = Request("Month")
YearP_ = Request("Year")
SortBy_ = Request("SortBy")
Order_ = Request("Order")
if Post_ = "" then
	PostRpt_ = "All Post"
else
	PostRpt_ = Post_ 
end if
%>
</head>
<BODY>
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
	strsql = "spHomeBillingReport '1','" & Post_ & "','" & PhoneNumber_ & "','" & MonthP_ & "','" & YearP_ & "','" & SortBy_ & "','" & Order_ & "'"
	set DataRS = server.createobject("adodb.recordset")
	'response.write strsql & "<br>"
	set DataRS = BillingCon.execute(strsql)

%>	
	<form method="post" name="frmHomephonePrint">
	<table align="center" cellpadding="1" cellspacing="0" width="180%" >
	<tr align="center">
		<td><h3>Home - Payment Billing Report</h3></td>
	</tr>
	<tr align="center">
		<td><b>Period : </b><%=MonthP_ %> - <%=YearP_ %></td>
	</tr>
	<tr align="center">
		<td><b>Post : </b><%=PostRpt_ %></td>
	</tr>	
	</table>
	<table align="center" cellpadding="1" cellspacing="0" width="180%" border="1" bordercolor="black"> 
	<TR BGCOLOR="#000099" align="center" cellpadding="0" cellspacing="0" >
		<TD width="2%"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
	       	<TD><strong><label STYLE=color:#FFFFFF>HomePhone</label></strong></TD>
		<TD><strong><label STYLE=color:#FFFFFF>Customer</label></strong></TD>
	       	<TD><strong><label STYLE=color:#FFFFFF>Billing No.</label></strong></TD>
	       	<TD><strong><label STYLE=color:#FFFFFF>Monthly Fee</label></strong></TD>
	       	<TD><strong><label STYLE=color:#FFFFFF>Local call</label></strong></TD>
	       	<TD><strong><label STYLE=color:#FFFFFF>SLJJ</label></strong></TD>
	       	<TD><strong><label STYLE=color:#FFFFFF>STB</label></strong></TD>
	       	<TD><strong><label STYLE=color:#FFFFFF>JAPATI</label></strong></TD>
	       	<TD><strong><label STYLE=color:#FFFFFF>SLI007</label></strong></TD>
	       	<TD><strong><label STYLE=color:#FFFFFF>001+008</label></strong></TD>
	       	<TD><strong><label STYLE=color:#FFFFFF>17</label></strong></TD>
	       	<TD><strong><label STYLE=color:#FFFFFF>Operator</label></strong></TD>
	       	<TD><strong><label STYLE=color:#FFFFFF>Air Time</label></strong></TD>
	       	<TD><strong><label STYLE=color:#FFFFFF>Quota</label></strong></TD>
	       	<TD><strong><label STYLE=color:#FFFFFF>Miscellaneous</label></strong></TD>
	       	<TD><strong><label STYLE=color:#FFFFFF>Tax</label></strong></TD>
	       	<TD><strong><label STYLE=color:#FFFFFF>Stamp fee</label></strong></TD>
		<TD><strong><label STYLE=color:#FFFFFF>Total (Kn)</label></strong></TD>
		<TD><strong><label STYLE=color:#FFFFFF>Status Payment</label></strong></TD>
	</TR>    
	<% 
		dim no_  
		no_ = 1 
		do while not DataRS.eof
	   	if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
	%>
	   	<TR bgcolor="<%=bg%>">
			<td align="right"><%=No_%>&nbsp;</td>
	        	<td><FONT color=#330099 size=2>&nbsp;<%=DataRS("HomePhone")%></font></td> 
	        	<td><FONT color=#330099 size=2>&nbsp;<%=DataRS("EmpName")%></font></td> 
	        	<td><FONT color=#330099 size=2>&nbsp;<%=DataRS("NoTagihan")%></font></td> 
	        	<td align="right"><FONT color=#330099 size=2><%=formatnumber(DataRS("Abonemen"),-1)%></font></td> 
	        	<td align="right"><FONT color=#330099 size=2><%=formatnumber(DataRS("Lokal"),-1)%></font></td> 
	        	<td align="right"><FONT color=#330099 size=2><%=formatnumber(DataRS("SLJJ"),-1)%></font></td> 
	        	<td align="right"><FONT color=#330099 size=2><%=formatnumber(DataRS("STB"),-1)%></font></td> 
	        	<td align="right"><FONT color=#330099 size=2><%=formatnumber(DataRS("JAPATI"),-1)%></font></td> 
	        	<td align="right"><FONT color=#330099 size=2><%=formatnumber(DataRS("SLI007"),-1)%></font></td> 
	        	<td align="right"><FONT color=#330099 size=2><%=formatnumber(DataRS("001+008"),-1)%></font></td> 
	        	<td align="right"><FONT color=#330099 size=2><%=formatnumber(DataRS("17"),-1)%></font></td> 
	        	<td align="right"><FONT color=#330099 size=2><%=formatnumber(DataRS("Operator"),-1)%></font></td> 
	        	<td align="right"><FONT color=#330099 size=2><%=formatnumber(DataRS("AirTime"),-1)%></font></td> 
	        	<td align="right"><FONT color=#330099 size=2><%=formatnumber(DataRS("Quota"),-1)%></font></td> 
	        	<td align="right"><FONT color=#330099 size=2><%=formatnumber(DataRS("Lain2"),-1)%></font></td> 
	        	<td align="right"><FONT color=#330099 size=2><%=formatnumber(DataRS("PPN"),-1)%></font></td> 
	        	<td align="right"><FONT color=#330099 size=2><%=formatnumber(DataRS("Meterai"),-1)%></font></td> 
	        	<td align="right"><FONT color=#330099 size=2><%=formatnumber(DataRS("Total"),-1)%></font></td> 
	        	<td><FONT color=#330099 size=2><%=DataRS("Status")%></font></td> 
	  	 </TR>
	<%   
			Count=Count +1
	 		DataRS.movenext
	   		no_ = no_ + 1
		loop
	%>
	</table>
<%
	strsql = "spHomeBillingReport '2','" & Post_ & "','" & PhoneNumber_ & "','" & MonthP_ & "','" & YearP_ & "','" & SortBy_ & "','" & Order_ & "'"
	set DataRS = server.createobject("adodb.recordset")
	'response.write strsql & "<br>"
	set DataRS = BillingCon.execute(strsql)
%>
	<table cellpadding="1" cellspacing="0" width="180%">
	<tr>
		<td width="85%" align="right"><b>Payment to Telkom</b></td>
		<td width="1%" align="right"><b>=</b></td>
		<td width="20%" align="left">Kn. <%=formatnumber(DataRS("TotalBill"),-1)%> - <%=formatnumber(DataRS("PersonalBill"),-1)%></td>
	</tr>
	<tr align="right">
		<td>&nbsp;</td>
		<td width="1%" align="right"><b>=</b></td>
		<td align="left"><b><u>Kn. <%=formatnumber(DataRS("Payment"),-1)%></u></b></td>
	</tr>
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


