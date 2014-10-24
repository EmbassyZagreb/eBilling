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
user1_ = user_  'user1_ = right(user_,len(user_)-4)
Post_ = Request("Post")
Section_ = Request("Section")
SectionGroup_ = Request("SectionGroup")
EmpName_ = Request("EmpName")
PhoneNumber_ = Request("PhoneNumber")	
Charge_ = Request("Charge")
Discontinued_ = Request("Discontinued")	
SortBy_ = Request("SortBy")	
Order_ = Request("Order")	

strsql = "select * from Users where loginId='" & user1_ & "'"
set RS_Query = server.createobject("adodb.recordset")
'response.write strsql & "<br>"
set RS_Query = BillingCon.execute(strsql)

  if not RS_Query.eof then 
     if (trim(RS_Query("RoleID")) = "Admin") or (trim(RS_Query("RoleID")) = "IM") or (mid(RS_Query("RoleID"),1,3) = "FMC")  then     
	strsql = "Select ID, PhoneNumber, PhoneTypeName, EmpId, EmpName, Post, Office, OwnerName, Case When Len(isNull(EmailAddress,''))<4 Then AlternateEmail Else isNull(EmailAddress,'') End As EmailAddress, BillFlag, Discontinued, DiscontinuedDesc, DiscontinuedDate "_
		  &"From vwCellPhoneNumberList Where (PhoneNumber like '" & PhoneNumber_  & "%' or '" & PhoneNumber_ & "'='') "_
		  &"And (EmpName like '%" & EmpName_  & "%' or '" & EmpName_ & "'='') "_
		  &"And (Post='" & Post_ & "' or '" & Post_ & "'='A') "_
		  &"And (Office='" & Section_ & "' or '" & Section_ & "'='A') "_
		  &"And (SectionGroup='" & SectionGroup_ & "' or '" & SectionGroup_ & "'='A') "_
		  &"And (BillFlag = '" & Charge_ & "' or '" & Charge_ & "'='A') "_
		  &"Order by " & SortBy_ & " " & Order_ 
		
	'response.write strsql & "<br>"
	set rs = server.createobject("adodb.recordset") 
	set rs = BillingCon.execute(strsql) 
%>	
     <table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" width=100%>
     <TR bgcolor='#000099' align="center">
         <TD class="style5" width="5%" align="center">No.</TD>
         <TD class="style5" width="15%">Phone Number</TD>
         <TD class="style5">Employee Name</TD>
<!--         <TD class="style5">Phone Type</TD>-->
	 <TD width="12%" class="style5">Post</TD>
         <TD width="12%" class="style5">Office</TD>
         <TD class="style5">Email Address / Alternate</TD>
         <TD class="style5" width="12%">Owner</TD>
	 <TD class="style5">Charged</strong>
         <TD class="style5">Discontinued</TD>
         <TD class="style5">Discontinued Date</TD>
     </TR>    
<% 
	   dim no_  
	   no_ = 1 
	   do while not rs.eof
	   if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
%>
      
     <TR bgcolor="<%=bg%>">
        <TD align=center > <%= no_ %> </TD>
        <TD><%= cstr(rs("PhoneNumber")) %></TD>
        <TD><%= rs("EmpName") %></TD>
<!--        <TD><%= rs("PhoneTypeName") %></TD> -->
	<TD><%= rs("Post") %>   </TD>
        <TD><%= rs("Office") %>   </TD>
        <TD><%= rs("EmailAddress") %>   </TD>
        <TD><%= rs("OwnerName") %></TD>
        <TD><%If rs("BillFlag")="Y" Then %>Yes<%else%>No<%End If%></TD>
        <TD>&nbsp;<%= rs("Discontinued") %>   </TD>
        <TD>&nbsp;<%= rs("DiscontinuedDate") %>   </TD>
      </TR>

<%   
	   Count=Count +1
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


