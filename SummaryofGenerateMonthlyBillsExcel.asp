<%@ Language=VBScript%>
<% ' VI 6.0 Scripting Object Model Enabled %>
 
 
<html>
<head>
   <%
Response.ContentType ="application/vnd.ms-excel" 
Response.Buffer  =  True 
Response.Clear() 
%> 
<META name=VI60_defaultClientScript content=VBScript>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<!--#include file="connect.inc" -->


<%
Dim user_ , user1_

user_ = request.servervariables("remote_user")
user1_ = user_  'user1_ = right(user_,len(user_)-4)
'response.write user1_ & "<br>"

sMonthP = request("sMonthP")
sYearP = request("sYearP")
eMonthP = request("eMonthP")
eYearP = request("eYearP")

%>
<TITLE>U.S. Embassy Zagreb - zBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY BGCOLOR="#FFFFFF" TEXT="#000000" LINK="#0080FF" ALINK="0080FF" VLINK="#0055AA" MARGINWIDTH="8" MARGINHEIGHT="0" LEFTMARGIN="8" TOPMARGIN="0">
<%
sPeriod = sYearP&sMonthP
ePeriod = eYearP&eMonthP
'response.write sPeriod & ePeriod 
strsql = "Select distinct MonthP, YearP, ProgressDesc, ProgressID, TotalRecord, TotalBill From vwProgressSummary Where YearP+MonthP>='" & sPeriod & "' and YearP+MonthP<='" & ePeriod & "' "

strsql = strsql  & strFilter & " order by YearP,MonthP,ProgressDesc"
'response.write strsql & "<br>"

set DataRS = server.createobject("adodb.recordset")
DataRS.CursorLocation = 3
DataRS.Open strsql,BillingCon
'set DataRS=BillingCon.execute(strsql)

'response.write strsql

if not DataRS.eof Then
%>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" width="600px"  class="FontText">
    <TR align="center">
         <TD width="3%" align="right"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
         <TD width="100px"><strong><label STYLE=color:#FFFFFF>Billing Period</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Description</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Total Record</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Total Bill</label></strong></TD>
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
	        <TD align="right"><%=no_ %>&nbsp;</TD>
	        <TD align="right">&nbsp;<%= DataRS("MonthP")%>-<%= DataRS("YearP")%>&nbsp;</TD>
	        <TD>&nbsp;<%=DataRS("ProgressDesc") %> </TD>
	        <TD align="right"><%=formatnumber(DataRS("TotalRecord"),-1) %></TD>
	        <TD align="right"><%=formatnumber(DataRS("TotalBill"),-1) %></TD>
	   </TR>

<%   
		TotRecordNo_ = cdbl(TotRecordNo_) + DataRS("TotalRecord")
		TotBillingAmount_ = cdbl(TotBillingAmount_) + cdbl(DataRS("TotalBill"))
		Count=Count +1
	   DataRS.movenext
	   no_ = no_ + 1 
	if DataRS.eof then
%>
	<TR bgcolor="<%=bg%>">
	        <TD align="center" colspan="3"><strong>Total</strong></TD>
	        <TD align="right"><strong><%=formatnumber(TotRecordNo_,-1) %></strong></TD>
	        <TD align="right"><strong><%=formatnumber(TotBillingAmount_,-1) %></strong></TD>
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
	<tr>
        	<td><br></TD>
	</tr>
	<tr>
		<td align="center"><a href="Default.asp"><img src="images/Back.gif" border="0" alt="Go..Back" WIDTH="83" HEIGHT="25"></a></td>
	</tr>	
	</table>
<% end if %>
</body> 

</html>


