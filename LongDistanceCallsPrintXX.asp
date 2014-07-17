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
	font-size: small;}

.style1 {color: #990000;}
.style4 {color: #FFFFFF; font-weight: bold;}
.style5 {color: #FFFFFF;}
.style6 { font-size: 10px;}

-->
</style>
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
Extension_ = Request("Extension")
'response.write Post_
StartDate_ = Request("StartDate")
'response.write StartDate_ 
EndDate_ = Request("EndDate")
CallType_ = Request("CallType")
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
	strsql = "spLongDistanceCallsReport '1','" & Post_ & "','" & Extension_ & "','" & StartDate_ & "','" & EndDate_ & "','" & CallType_ & "','" & SortBy_ & "','" & Order_ & "'"
	set DataRS = server.createobject("adodb.recordset")
	'response.write strsql & "<br>"
	set DataRS = BillingCon.execute(strsql)

%>	
	<form method="post" name="frmLongdistanceCallsPrint">
	<table align="center" cellpadding="1" cellspacing="0" width="100%" >
	<tr align="center">
		<td><h3>Long Distance calls Report</h3></td>
	</tr>
	<tr align="center">
		<td><b>Period : </b><%=StartDate_  %> - <%=EndDate_%></td>
	</tr>
	<tr align="center">
		<td><b>Post : </b><%= PostRpt_ %></td>
	</tr>	
	</table>
	<table align="center" cellpadding="1" cellspacing="0" width="100%" border="1" bordercolor="black"> 
	<TR BGCOLOR="#000099" align="center" cellpadding="0" cellspacing="0" >
		<TD width="4%" class="style5">No.</TD>
	       	<TD width="14%" class="style5">Dialed Number</TD>
		<TD class="style5">Employee Name</TD>
		<TD width="10%" class="style5">Phone Ext.</TD>
	       	<TD width="14%" class="style5">Call date & time</TD>
		<TD width="10%" class="style5">Duration (Second)</TD>
		<TD width="10%" class="style5">Cost</TD>
		<TD width="10%" class="style5">Call Type</TD>
	</TR>  
	<% 
		dim no_  
		no_ = 1 
		do while not DataRS.eof 
	   	if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
	%>
	   	<TR bgcolor="<%=bg%>">
			<td align="right"><%=No_%>&nbsp;</td>
	        	<td><FONT color=#330099 size=2>&nbsp;<%=DataRS("DialedNumber")%></font></td> 
	        	<td><FONT color=#330099 size=2>&nbsp;<%=DataRS("EmpName")%></font></td>
			<td><FONT color=#330099 size=2>&nbsp;<%=DataRS("Extension")%></font></td> 
	        	<td><FONT color=#330099 size=2>&nbsp;<%=DataRS("DialedDateTime")%></font></td> 
	        	<td><FONT color=#330099 size=2>&nbsp;<%=DataRS("CallDuration")%></font></td> 
	        	<td align="right"><FONT color=#330099 size=2><%=formatnumber(DataRS("Cost"),0)%>&nbsp;</font></td> 
	        	<td><FONT color=#330099 size=2>&nbsp;<%=DataRS("CallType")%></font></td> 
	  	 </TR>
	<%   
	 		DataRS.movenext
	   		no_ = no_ + 1
		loop
	%>
	</table>
<%
	strsql = "spLongDistanceCallsReport '2','" & Post_ & "','" & Extension_ & "','" & StartDate_ & "','" & EndDate_ & "','" & CallType_ & "','" & SortBy_ & "','" & Order_ & "'"
	set sumRS = server.createobject("adodb.recordset")
	'response.write strsql & "<br>"
	set sumRS = BillingCon.execute(strsql)
%>	
	<table align="center" cellpadding="1" cellspacing="0" width="100%"> 
	<tr>
		<td width="70%" align="right"><b>Total</b>&nbsp;&nbsp;&nbsp;:&nbsp;</td>
<!--		<td width="10%" align="right"><b><%=formatnumber(sumRS("TotalDuration"),0)%></b>&nbsp;second(s)&nbsp;</td> -->
		<td width="10%" align="right"><b>Rp. <%=formatnumber(sumRS("TotalCost"),0)%></b>&nbsp;</td>
		<td width="10%">&nbsp;</td>
	</tr>
	</table>
	</form>
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


