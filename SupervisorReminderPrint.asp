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

sMonthP = request("sMonthP")
'response.write sMonthP

sYearP = request("sYearP")

eMonthP = request("eMonthP")

eYearP = request("eYearP")

Agency_ = request("Agency")

Section_ = request("Section")

SPVEmail_ = request("SPVEmail")

Status = request("Status")

SortBy_ = Request("SortBy")

Order_ = Request("OrderList")
%>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY BGCOLOR="#FFFFFF" TEXT="#000000" LINK="#0080FF" ALINK="0080FF" VLINK="#0055AA" MARGINWIDTH="8" MARGINHEIGHT="0" LEFTMARGIN="8" TOPMARGIN="0">
<%

dim rs 
dim strsql
If (UserRole_ = "Admin") or (UserRole_ = "Voucher") or (UserRole_ = "FMC") or (UserRole_ = "Cashier")  Then
sPeriod = sYearP&sMonthP
ePeriod = eYearP&eMonthP
'response.write sPeriod & ePeriod 
'strsql = "Select * From vwMonthlyBilling Where ProgressID=5 and YearP+MonthP>='" & sPeriod & "' and YearP+MonthP<='" & ePeriod & "'"
strsql = "Select * From vwSupervisorReminderRpt Where YearP+MonthP>='" & sPeriod & "' and YearP+MonthP<='" & ePeriod & "'"
strFilter=""
If Agency_ <> "X" then
	strFilter=strFilter & " and Agency='" & Agency_ & "'"
End If

If Section_ <> "X" then
	strFilter=strFilter & " and Office='" & Section_ & "'"
End If

If SPVEmail_ <> "X" then
	strFilter =strFilter & " and SupervisorEmail='" & SPVEmail_ & "'"
End If

If Status <> "0" then
	strFilter =strFilter & " and ProgressID=" & Status
End If

strsql = strsql  & strFilter & " Order By " & SortBy_ & " "& Order_
'response.write strsql & "<br>"

set DataRS = server.createobject("adodb.recordset")
'DataRS.CursorLocation = 3
'DataRS.Open strsql,BillingCon
set DataRS=BillingCon.execute(strsql)

'response.write strsql


if not DataRS.eof Then

%>
<form method="post" name="frmSupervisorReminderPrint">
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" width="100%"  class="FontText">
    <TR BGCOLOR="#330099" align="center">
         <TD width="3%" align="right" class="style5">No.</TD>
         <TD width="15%" class="style5">Employee Name</TD>
         <TD width="10%" class="style5">Number</TD>
         <TD width="10%" class="style5">Billing Period</TD>
         <TD width="10%" class="style5">Section</TD>
	 <TD width="10%" class="style5">Billing Amount (Kn.)</TD>
	 <TD width="10%" class="style5">Personal Amount (Kn.)</TD>
	 <TD class="style5">Aging</TD>
	 <TD class="style5">Supervisor</TD>
	 <TD class="style5">Supervisor Email</TD>
    </TR>
<!--    <tr BGCOLOR="#330099" align="center">
         <TD width="10%" class="style5">Home Phone</TD> 
         <TD width="10%" class="style5">Office Phone</TD>
         <TD width="10%" class="style5">Mobile Phone</TD>
        <TD width="10%" class="style5">Shuttle Bus</TD> 
        <TD width="8%" class="style5">Total</TD> 
    </tr> -->

<% 
   dim no_  
   no_ = 1

   do while not DataRS.eof
	   if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
%>
      
	   <TR bgcolor="<%=bg%>">
	        <TD align="right"><%=no_ %></font></TD>
	        <TD><%=DataRS("EmpName") %></TD>
	        <TD><%=DataRS("MobilePhone") %></TD>
	        <TD align="right"><%= DataRS("MonthP")%>-<%= DataRS("YearP")%></font></TD>
	        <TD><%=DataRS("Office") %> </font></TD>
<!--		<td align="right">
<%		If CDbl(DataRS("HomePhonePrsBillRp")) > 0 Then %>
			<%= formatnumber(DataRS("HomePhonePrsBillRp"),2) %>
<%		Else %>
			-
<%		End If %>
		</td>
		<td align="right">
<%		If CDbl(DataRS("OfficePhonePrsBillRp")) > 0 Then %>
			<%= formatnumber(DataRS("OfficePhonePrsBillRp"),2) %>
<%		Else %>
			-
<%		End If %>
		</td> -->
		<td align="right">
<%		If CDbl(DataRS("CellPhoneBillRp")) > 0 Then %>
			<%= formatnumber(DataRS("CellPhoneBillRp"),2) %>
<%		Else %>
			0
<%		End If %>
		</td>
		<td align="right">
<%		If CDbl(DataRS("CellPhonePrsBillRp")) > 0 Then %>
			<%= formatnumber(DataRS("CellPhonePrsBillRp"),2) %>
<%		Else %>
			0
<%		End If %>
		</td>
<!--		<td align="right">
<%		If CDbl(DataRS("TotalShuttleBillRp")) > 0 Then %>
			<%= formatnumber(DataRS("TotalShuttleBillRp"),2) %>
<%		Else %>
			-
<%		End If %>
		</td>
	        <TD align="right"><%= formatnumber(DataRS("TotalBillingAmountPrsRp"),2) %></font></TD> -->
	        <TD><%=DataRS("Aging") %> </font></TD>
	        <TD><%=DataRS("Supervisor") %> </font></TD>
	        <TD><%=DataRS("SupervisorEmail") %> </font></TD>
	    </TR>

<%   
	   DataRS.movenext
	   no_ = no_ + 1 
   loop 
	PageNo=1
%>
</table>
</form>
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

