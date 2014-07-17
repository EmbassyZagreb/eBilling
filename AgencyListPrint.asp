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
<!--#include file="connect.inc" -->




<link href="style.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
function ClearFilter()
{
	document.forms['frmSearch'].elements['txtFundingAgency'].value ='';
	document.forms['frmSearch'].elements['SortList'].value ='AgencyFundingCode';
}
</script>
</head>
<%

dim rs 
dim strsql
dim tombol
dim hlm
%>

<%
Dim user_ , user1_, UserRole_

user_ = request.servervariables("remote_user")
user1_ = right(user_,len(user_)-4)


strsql = "select * from Users where loginId='" & user1_ & "'"
set RS_Query = server.createobject("adodb.recordset")
'response.write strsql & "<br>"
set RS_Query = BillingCon.execute(strsql)
'response.write RS_Query("RoleID") & "<br>"



FundingAgency_ = Request("FundingAgency")	
SortBy_ = Request("SortBy")	
Order_ = Request("OrderList")	

     if (trim(RS_Query("RoleID")) = "Admin") or (trim(RS_Query("RoleID")) = "Voucher") or (trim(RS_Query("RoleID")) = "FMC") then    
'	strsql = "select * from AgencyFunding"
	strsql = "select * from AgencyFunding"_
		  &" Where (AgencyDesc like '%" & FundingAgency_  & "%' or '" & FundingAgency_ & "'='') "_
		  &"Order by " & SortBy_ & " " & Order_ 

	'response.write strsql & "<br>"
	set rsAgency = server.createobject("adodb.recordset")
	set rsAgency = BillingCon.execute(strsql)
%>	

	
	<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" width="90%"  class="FontText">
	    <TR align="center">
		 <TD width=3%><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
	         <TD width=8%><strong><label STYLE=color:#FFFFFF>Agency Code</label></strong></TD>
	         <TD width="15%"><strong><label STYLE=color:#FFFFFF>Funding Agency</label></strong></TD>
	         <TD><strong><label STYLE=color:#FFFFFF>Fiscal Strip VAT</label></strong></TD>
	         <TD><strong><label STYLE=color:#FFFFFF>Fiscal Strip Non VAT</label></strong></TD>
	         <TD width="5%"><strong><label STYLE=color:#FFFFFF>Disabled</label></strong></TD>
	    </TR>    

<% 
	   dim no_  
	   no_ = 1 
	   do while not rsAgency.eof  
		   if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 

		If rsAgency("Disabled")="Y" Then
			AgencyType_ ="Yes"
		ElseIf rsAgency("Disabled")="N" Then
			AgencyType_ ="No"
		Else
			AgencyType_ =""
		End If
%>

	   <TR bgcolor="<%=bg%>">
       		<TD align="right"> <%= no_ %> </TD>
        	<TD align="right"><%= rsAgency("AgencyFundingCode")%>&nbsp;</TD>
	        <TD><%= rsAgency("AgencyDesc") %></TD>
	        <TD><%= rsAgency("FiscalStripVAT") %></TD>
	        <TD><%= rsAgency("FiscalStripNonVAT") %></TD>
	        <TD align="right"><%= AgencyType_ %>&nbsp;</TD>
	   </TR>

<%   
	   	rsAgency.movenext
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

<%
end if 
%>
</body> 

</html>


