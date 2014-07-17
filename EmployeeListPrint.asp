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

dim x , ticket_, user_ , user1_
dim intPageSize,PageIndex,TotalPages 
dim RecordCount,RecordNumber,Count 
intpageSize=50 
PageIndex=request("PageIndex")


user_ = request.servervariables("remote_user")
user1_ = right(user_,len(user_)-4)

Post_ = Request("Post")	
Section_ = Request("Section")	
SectionGroup_ = Request("SectionGroup")
EmpName_ = Request("EmpName")
AgencyFundingCode_ = request("AgencyFundingCode")
Status_ = Request("Status")	
SortBy_ = Request("SortBy")
request("AgencyFundingCode")
Order_ = Request("OrderList")

strsql = "select * from Users where loginId='" & user1_ & "'"
set RS_Query = server.createobject("adodb.recordset")
'response.write strsql & "<br>"
set RS_Query = BillingCon.execute(strsql)

  if not RS_Query.eof then 
     if (trim(RS_Query("RoleID")) = "Admin") or (trim(RS_Query("RoleID")) = "IM") or (trim(RS_Query("RoleID")) = "FMC") then     
	strsql = "Select EmpID, EmpName, PostName, StatusName, Agency, Office, Type, ReportToID, ReportToName, EmailAddress, AgencyFunding, Remark "_
		  &"From vwDirectReport Where (PostName='" & Post_ & "' or '" & Post_ & "'='A') "_
		  &"And (Office='" & Section_ & "' or '" & Section_ & "'='A') "_
		  &"And (SectionGroup='" & SectionGroup_ & "' or '" & SectionGroup_ & "'='A') "_
		  &"And (EmpName like '%" & EmpName_  & "%' or '" & EmpName_ & "'='') "_
		  &"And (AgencyId = " & AgencyFundingCode_ & " or " & AgencyFundingCode_ & "=0) "_
		  &"And (Status='" & Status_ & "' or '" & Status_ & "'='A') "_
		  &"Order by " & SortBy_ & " " & Order_ 

'		  &"And EmpID in (Select EmpID From MsCellPhoneNumber Where BillFlag='Y' ) "_
		
'	response.write strsql & "<br>"
	set rs = server.createobject("adodb.recordset") 
	rs.CursorLocation = 3
	rs.open strsql,BillingCon

'	set rs = BillingCon.execute(strsql) 

%>
	
     <form name="frmHomePhoneList" Action="HomePhoneListAll.asp" onSubmit="return ValidateForm()">
     <table border="1" bordercolor="#EEEEEE" cellpadding="2" cellspacing="0" width="100%">
     <TR BGCOLOR="#330099" align="center">
         <TD class="style5"  width="3%" align="center"><strong>NO</strong></TD>
         <TD class="style5" ><strong>&nbsp;Employee Name</strong></TD>
         <TD class="style5"  width="12%"><strong>&nbsp;Post</strong></TD>
         <TD class="style5" ><strong>&nbsp;Status</strong></TD>
	 <TD class="style5"  width="12%"><strong>&nbsp;Office</strong></TD>
         <TD class="style5" ><strong>Type</strong></TD>
         <TD class="style5" ><strong>&nbsp;Email Address</strong></TD>
         <TD class="style5" ><strong>&nbsp;Supervisor</strong></TD>
         <TD class="style5" ><strong>&nbsp;Agency Funding</strong></TD>
	 <TD class="style5" ><strong>&nbsp;Remark</strong></TD>
     </TR>   
<% 
	   dim no_  
	   'no_ = 1 
	   no_ = 1 
	   do while not rs.eof 
	   if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
%>
      
     <TR bgcolor="<%=bg%>">
        <TD align=center > <%= no_ %> </font>   </TD>
        <TD>&nbsp;<%= rs("EmpName") %></font>   </TD>
        <TD>&nbsp;<%= rs("PostName") %></font>   </TD>
        <TD>&nbsp;<%= rs("StatusName") %></font>   </TD>
        <TD>&nbsp;<%= rs("Office") %></font>   </TD>
        <TD>&nbsp;<%= rs("Type") %></font>   </TD>
        <TD>&nbsp;<%= rs("EmailAddress") %></font>   </TD>
        <TD>&nbsp;<%= rs("ReportToName") %></font>   </TD>
        <TD>&nbsp;<%= rs("AgencyFunding") %></font>   </TD>
	<TD>&nbsp;<%= rs("Remark") %></font>   </TD>
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
	</form>
</body> 

</html>


