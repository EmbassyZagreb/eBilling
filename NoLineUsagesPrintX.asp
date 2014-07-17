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
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
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
response.write Post_
StartDate_ = Request("StartDate")
'response.write StartDate_ 
EndDate_ = Request("EndDate")
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
	strsql = "spNoLineUsagesReport '" & Post_ & "','" & StartDate_ & "','" & EndDate_ & "','" & SortBy_ & "','" & Order_ & "'"
	set DataRS = server.createobject("adodb.recordset")
	'response.write strsql & "<br>"
	set DataRS = BillingCon.execute(strsql)
%>	
	<form method="post" name="frmNoLineUsagePrint">
	<table align="center" cellpadding="1" cellspacing="0" width="100%" >
	<tr align="center">
		<td><h3>No Line usages Report</h3></td>
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
	       	<TD width="14%" class="style5">Extension</TD>
		<TD class="style5"><strong>Employee Name</TD>
	       	<TD width="14%" class="style5">Post</TD>
	</TR>  
	<% 
		dim no_  
		no_ = 1 
		do while not DataRS.eof 
	   	if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
	%>
   	<TR bgcolor="<%=bg%>">
			<td align="right"><%=No_%>&nbsp;</td>
	        	<td><FONT color=#330099 size=2>&nbsp;<%=DataRS("PhoneNumber")%></font></td> 
	        	<td><FONT color=#330099 size=2>&nbsp;<%=DataRS("Location")%></font></td> 
	        	<td><FONT color=#330099 size=2>&nbsp;<%=DataRS("Post")%></font>&nbsp;</td> 
	  	 </TR>
	<%   
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


