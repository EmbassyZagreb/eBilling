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

sMonthP = request("sMonthP")
sYearP = request("sYearP")
eMonthP = request("eMonthP")
eYearP = request("eYearP")
Agency_ = request("Agency")
Section_ = request("Section")
EmpID_ = request("EmpID")
PaymentType_ = request("PaymentType")
%>
<TITLE>U.S. Embassy Zagreb - zBilling Application</TITLE>
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
strsql = "Select * From vwARPaymentRpt Where YearP+MonthP>='" & sPeriod & "' and YearP+MonthP<='" & ePeriod & "'"
strFilter=""
If Agency_ <> "X" then
	strFilter=strFilter & " and Agency='" & Agency_ & "'"
End If

If Section_ <> "X" then
	strFilter=strFilter & " and Office='" & Section_ & "'"
End If

If EmpID_ <> "X" then
	strFilter =strFilter & " and EmpID='" & EmpID_ & "'"
End If

If PaymentType_ <> "X" then
	strFilter =strFilter & " and PaymentType='" & PaymentType_ & "'"
End If

strsql = strsql  & strFilter & " order by EmpName"
'response.write strsql & "<br>"

set DataRS = server.createobject("adodb.recordset")
DataRS.CursorLocation = 3
'DataRS.Open strsql,BillingCon
set DataRS=BillingCon.execute(strsql)
'response.write strsql


if not DataRS.eof Then

%>
<!-- <div align="right"><input type="submit" name="btnApproval" value="Approve" /></div> -->
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" width="100%"  class="FontText">
    <TR align="center">
         <TD width="3%" align="right"><strong>No.</strong></TD>
         <TD width="15%"><strong>Employee Name</strong></TD>
         <TD width="6%"><strong>Number</strong></TD>
         <TD width="6%"><strong>Billing Period</strong></TD>
         <TD width="6%"><strong>Section</strong></TD>
	 <TD width="6%"><strong>Billing Amount (Kn.)</strong></TD>
         <TD width="6%"><strong>Paid Date</strong></TD>
         <TD width="6%"><strong>Receipt No.</strong></TD>
         <TD width="6%"><strong>Receipt Type</strong></TD>
         <TD width="6%"><strong>Receipt Amount</strong></TD>
<!--         <TD><strong>Receipt Currency</strong></TD> -->
         <TD width="6%"><strong>Age (day)</strong></TD>
         <TD width="15%"><strong>Funding Agency</strong></TD>
    </TR>
<!--     <tr align="center">
        <TD width="10%"><strong>Home Phone</strong></TD>
         <TD width="10%"><strong>Office Phone</strong></TD>
         <TD width="10%"><strong>Mobile Phone</strong></TD>
         <TD width="10%"><strong>Shuttle Bus</strong></TD>
         <TD width="8%"><strong>Total</strong></TD>
    </tr> -->

<% 
   dim no_  
   no_ = 1
   Count=1 
   do while not DataRS.eof
	   if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
%>
      
	   <TR bgcolor="<%=bg%>">
	        <TD align="right"><%=no_ %>&nbsp;</font></TD>
	        <TD>&nbsp;<%=DataRS("EmpName") %></TD>
		<TD>&nbsp;<%=DataRS("MobilePhone") %></TD>
	        <TD align="right">&nbsp;<%= DataRS("MonthP")%>-<%= DataRS("YearP")%></font>&nbsp;</TD>
	        <TD>&nbsp;<%=DataRS("Office") %> </font></TD>
<!--		<td align="right">
<%		If CDbl(DataRS("HomePhonePrsBillRp")) > 0 Then %>
			<%= formatnumber(DataRS("HomePhonePrsBillRp"),-1) %>
<%		Else %>
			-
<%		End If %>
		&nbsp;</td>
		<td align="right">
<%		If CDbl(DataRS("OfficePhonePrsBillRp")) > 0 Then %>
			<%= formatnumber(DataRS("OfficePhonePrsBillRp"),-1) %>
<%		Else %>
			-
<%		End If %>
		&nbsp;</td>   -->
		<td align="right">
<%		If CDbl(DataRS("CellPhonePrsBillRp")) > 0 Then %>
			<%= formatnumber(DataRS("CellPhonePrsBillRp"),-1) %>
<%		Else %>
			0
<%		End If %>
		&nbsp;</td>
<!--		<td align="right">
<%		If CDbl(DataRS("TotalShuttleBillRp")) > 0 Then %>
			<%= formatnumber(DataRS("TotalShuttleBillRp"),-1) %>
<%		Else %>
			-
<%		End If %>
		&nbsp;</td>
	        <TD align="right"><%= formatnumber(DataRS("TotalBillingAmountPrsRp"),-1) %>&nbsp;</font></TD>  -->
		<TD>&nbsp;<%= DataRS("PaidDate") %></font></TD>
		<TD>&nbsp;<%= DataRS("ReceiptNo") %></font></TD>
		<TD>&nbsp;<%= DataRS("PaymentTypeName") %></font></TD> 
	        <TD align="right"><%= formatnumber(DataRS("PaidAmount"),-1) %>&nbsp;</font></TD>
<!--		<TD>&nbsp;<%= DataRS("Currency") %></font></TD>  -->
		<TD>&nbsp;<%= DataRS("Aging") %></font></TD> 
		<TD>&nbsp;<%= DataRS("AgencyFunding") %></font></TD> 
	   </TR>

<%   
	   DataRS.movenext
	   no_ = no_ + 1 
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


