<%@ Language=VBScript%>
<% ' VI 6.0 Scripting Object Model Enabled %>
 
 
<html>
<head>

<META name=VI60_defaultClientScript content=VBScript>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<!--#include file="connect.inc" -->
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
   <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">SHUTTLE BUS RATE</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
<%

dim rs 
dim strsql
dim tombol
dim hlm
%>

<%

dim x , ticket_, user_ , user1_


user_ = request.servervariables("remote_user")
user1_ = right(user_,len(user_)-4)
strsql = "select * from Users where loginId='" & user1_ & "'"
set RS_Query = server.createobject("adodb.recordset")
'response.write strsql & "<br>"
set RS_Query = BillingCon.execute(strsql)

  if not RS_Query.eof then 
     if (trim(RS_Query("RoleID")) = "Admin") or (trim(RS_Query("RoleID")) = "Cashier") or (trim(RS_Query("RoleID")) = "FMC") then     
	strsql = "select * from ShuttleBusRate order by ShuttleBusRateDate Desc"
	'response.write strsql & "<br>"
	set rs = server.createobject("adodb.recordset") 
	set rs = BillingCon.execute(strsql) 


%>
     <table width="35%" align="center">
     <tr>
		<td><a href="ShuttleBusRateEdit.asp?State=I">Add New Rate</a></td>
     </tr>
     </table>
     <table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" width="35%" align="center">
     <TR BGCOLOR="#330099" align="center">
         <TD width=5% align="center"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
         <TD width="35%"><strong><label STYLE=color:#FFFFFF>&nbsp;Rate Date</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>&nbsp;Shuttle Bus Rate($)</label></strong></TD>
         <TD width="10%">&nbsp;</TD>       
     </TR>    
<% 
	   dim no_  
	   no_ = 1 
	   do while not rs.eof  
	   if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
%>
      
     <TR bgcolor="<%=bg%>">
        <TD width=5% align=center ><FONT color=#330099 size=2> <%= no_ %> </font>   </TD>
        <TD width=20% ><FONT color=#330099 size=2><A HREF="ShuttleBusRateEdit.asp?ShuttleBusRateID=<%= rs("ShuttleBusRateID")%>&State=U"> <%= rs("ShuttleBusRateDate") %></A></font>   </TD>
        <TD width=20% align="right"><FONT color=#330099 size=2>&nbsp;<%= formatnumber(rs("ShuttleBusRate"),-1) %></font>&nbsp;</TD>
	<TD width=5% ><FONT color=#330099 size=2><A HREF="ShuttleBusRateDeleteConfirm.asp?ShuttleBusRateID=<%= rs("ShuttleBusRateID")%>&State=D" >Delete</A></font>   </TD>
      </TR>

<%   
	   rs.movenext
	   no_ = no_ + 1 
	   loop
%>
     </TABLE>
     <table width="35%" align="center">
	<tr>
		<td><br></td>
	</tr>
	<tr ><td align="Center">
	<a href="default.asp"><img src="images/Back.gif" border="0" alt="Go..Back" WIDTH="83" HEIGHT="25"></a>
	</td></tr>
     </table>
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


