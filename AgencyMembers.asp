<%@ Language=VBScript%>
<% ' VI 6.0 Scripting Object Model Enabled %>
 
 
<html>
<head>

<META name=VI60_defaultClientScript content=VBScript>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<!--#include file="connect.inc" -->


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

AgencyID_ = Request("AgencyID")
MonthP_ = Request("MonthP")
YearP_ = Request("YearP")
AgencyFundingCode_ = Request("AgencyFundingCode")
AgencyFundingDesc_ = Request("AgencyFundingDesc")
FiscalStripVAT_ = Request("FiscalStripVAT")
FiscalStripNonVAT_ = Request("FiscalStripNonVAT") 

%>
<TITLE>U.S. Embassy Zagreb - zBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">Agency Members</TD>
   </TR>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>

<%

dim rs 
dim strsql
If (UserRole_ = "Admin") or (UserRole_ = "Voucher") or (UserRole_ = "FMC") or (UserRole_ = "Cashier")  Then
%>
<%




strsql = "Select * From vwMonthlyBilling Where AgencyID = '" & AgencyID_ & "' And MonthP = '" & MonthP_ & "' And YearP = '" & YearP_ & "' And AgencyFundingCode = '" & AgencyFundingCode_ & "' And AgencyFundingDesc = '" & AgencyFundingDesc_ & "' And FiscalStripVAT = '" & FiscalStripVAT_ & "' And FiscalStripNonVAT = '" & FiscalStripNonVAT_ & "' Order by EmpName"
'strsql = "Exec spLogView '" & MonthP_ & "','" & YearP_ & "'," & ProgressID_ 
'response.write strsql & "<br>"

set DataRS = server.createobject("adodb.recordset")
DataRS.CursorLocation = 3
DataRS.Open strsql,BillingCon
'set DataRS=BillingCon.execute(strsql)

'response.write strsql

if not DataRS.eof Then
%>
<table bordercolor="#EEEEEE" cellpadding="2" cellspacing="0" width="650px">
<tr>
         <TD width="70px"><strong>&nbsp;</strong></TD>
	 
	 <td align="right"><input type="button" name="btnExport" value="Export to Excel" onClick="javascript:document.location.href('AgencyMemberExcel.asp?MonthP=<%=MonthP_%>&YearP=<%=YearP_%>&AgencyID=<%=AgencyID_%>');"/></td>
</tr>
</table>
<table border="1" bordercolor="#EEEEEE" cellpadding="2" cellspacing="0" width="650px"  class="FontText">
    <TR BGCOLOR="#330099" align="center">
         <TD width="30px" align="right"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Billing Period</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Employee Name</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Phone Number</label></strong></TD>
		 <TD><strong><label STYLE=color:#FFFFFF>Agency Funding</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Balance</label></strong></TD>
    </TR>
<% 
   dim no_  
   no_ = 1
   Count=1
   TotRecordNo_ = 0
   TotBillingAmount_ = 0
   do while not DataRS.eof
	   if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
%>    
	   <TR bgcolor="<%=bg%>">
	        <TD align="right"><%=no_ %>&nbsp;</font></TD>
	        <TD align="right">&nbsp;<%= DataRS("MonthP")%>-<%= DataRS("YearP")%></font>&nbsp;</TD>
	        <TD>&nbsp;<%=DataRS("EmpName") %></font></TD>
	        <TD>&nbsp;<%=DataRS("MobilePhone") %></font></TD>			
			<TD>&nbsp;<%=DataRS("AgencyFundingDesc") %></font></TD>
	        <TD align="right"><%=formatnumber(DataRS("TotalBillingRp"),-1) %>&nbsp;</font></TD>
	   </TR>

<%   
		TotRecordNo_ = cdbl(TotRecordNo_) + 1
		TotBillingAmount_ = cdbl(TotBillingAmount_) + cdbl(DataRS("TotalBillingRp"))
		Count=Count +1
	   DataRS.movenext
	   no_ = no_ + 1 
	if DataRS.eof then
%>
	<TR bgcolor="<%=bg%>">
	        <TD align="center" colspan="5"><strong>Total</strong></font></TD>
<!--	        <TD align="right"><strong><%=formatnumber(TotRecordNo_,-1) %></strong>&nbsp;</font></TD> -->
	        <TD align="right"><strong><%=formatnumber(TotBillingAmount_,-1) %>&nbsp;</strong></font></TD>
	   </TR>
<%	
	end if   
   loop 
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


