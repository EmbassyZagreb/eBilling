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
</HEAD>
<BODY BGCOLOR="#FFFFFF" TEXT="#000000" LINK="#0080FF" ALINK="0080FF" VLINK="#0055AA" MARGINWIDTH="8" MARGINHEIGHT="0" LEFTMARGIN="8" TOPMARGIN="0">
<%

dim rs 
dim strsql
dim tombol
dim hlm
%>

<%
user_ = request.servervariables("remote_user")
user1_ = right(user_,len(user_)-4)

Post_ = Request("Post")
Section_ = Request("Section")
EmpName_ = Request("EmpName")
PhoneNumber_ = Request("PhoneNumber")	
PhoneType_ = Request("PhoneType")
Charge_ = Request("Charge")
SortBy_ = Request("SortBy")
Order_ = Request("Order")

strsql = "select * from Users where loginId='" & user1_ & "'"
set RS_Query = server.createobject("adodb.recordset")
'response.write strsql & "<br>"
set RS_Query = BillingCon.execute(strsql)

  if not RS_Query.eof then 
     if (trim(RS_Query("RoleID")) = "Admin") or (trim(RS_Query("RoleID")) = "IM") or (mid(RS_Query("RoleID"),1,3) = "FMC") then     
	strsql = "Select ID, PhoneNumber, CustNo, CustName, EmpName, Address, City, Post, Office, Case When Len(isNull(EmailAddress,''))<4 Then AlternateEmail Else isNull(EmailAddress,'') End As EmailAddress, PhoneTypeName, BillFlag "_
		  &"From vwHomePhoneNumberList Where (PhoneType = '" & PhoneType_ & "' or '" & PhoneType_ & "'='A') "_
		  &"And (Post='" & Post_ & "' or '" & Post_ & "'='A') "_
		  &"And (Office='" & Section_ & "' or '" & Section_ & "'='A') "_
		  &"And (PhoneNumber like '" & PhoneNumber_  & "%' or '" & PhoneNumber_ & "'='') "_
		  &"And (BillFlag = '" & Charge_ & "' or '" & Charge_ & "'='A') "_
		  &"And (EmpName like '%" & EmpName_  & "%' or '" & EmpName_ & "'='') "_
		  &"Order by " & SortBy_ & " " & Order_ 
		
	'response.write strsql & "<br>"
	set rs = server.createobject("adodb.recordset") 
	set rs = BillingCon.execute(strsql) 


%>	
     <table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" width=100%>
     <TR bgcolor='#000099' align="center">
         <TD class="style5" width="5%" align="center">No.</TD>
         <TD class="style5" width="15%">Phone Number</TD>
<!--         <TD class="style5" width="10%">Customer No.</TD>
         <TD class="style5">Customer Name</TD>
-->
         <TD class="style5">Employee Name</TD>
         <TD class="style5" width="12%">Post</TD>
	 <TD width="12%" class="style5">Office</TD>
	 <TD class="style5">Email Address</TD>
<!--         <TD class="style5" width="12%">Phone Type</TD> -->
	 <TD class="style5">Charged</strong>
     </TR>    
<% 
	   dim no_  
	   no_ = 1 
	   do while not rs.eof
	   if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
%>
      
     <TR bgcolor="<%=bg%>">
        <TD align=center ><FONT color=#330099 size=2> <%= no_ %> </font></TD>
        <TD><FONT color=#330099 size=2><%= cstr(rs("PhoneNumber")) %></A></font></TD>
<!--        <TD><FONT color=#330099 size=2><%= rs("CustNo") %></font></TD>
        <TD><FONT color=#330099 size=2><%= rs("CustName") %></font></TD>
-->
        <TD><FONT color=#330099 size=2><%= rs("EmpName") %></font>   </TD>
        <TD><FONT color=#330099 size=2><%= rs("Post") %></font>   </TD>
        <TD><FONT color=#330099 size=2><%= rs("Office") %></font>   </TD>
        <TD><FONT color=#330099 size=2><%= rs("EmailAddress") %></font>   </TD>
<!--        <TD><FONT color=#330099 size=2><%= rs("PhoneTypeName") %></font>   </TD> -->
        <TD><FONT color=#330099 size=2><%If rs("BillFlag")="Y" Then %>Yes<%else%>No<%End If%></font></TD>
      </TR>

<%   
	   rs.movenext
	   no_ = no_ + 1 
	   loop
%>
     </TABLE>
<%
   else 
%>
      <table>
		<tr>
			<td>You do not have permission to access this site.</td>
		</tr>
		<tr>
			<td>Please <a href="http://zagrebws03.eur.state.sbu/WebPASS/eservices/MainPage.asp">Submit Request </a> or contact Zagreb ISC Helpdesk at ext.3333.</td>
		</tr>
	</table>
<%   end if 
else %>
	<table>
		<tr>
			<td>You do not have permission to access this site.</td>
		</tr>
		<tr>
			<td>Please <a href="http://zagrebws03.eur.state.sbu/WebPASS/eservices/MainPage.asp">Submit Request </a> or contact Zagreb ISC Helpdesk at ext.3333.</td>
		</tr>
	</table>

<%
end if 
%>
</body> 

</html>


