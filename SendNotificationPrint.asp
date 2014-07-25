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
SectionGroup_ = request("SectionGroup")
EmpID_ = request("EmpID")
BillType = request("BillType")
Status_ = request("Status")
SentStatus_ = request("SentStatus")
SortBy_ = request("SortBy")
Order_ = request("Order")
%>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY BGCOLOR="#FFFFFF" TEXT="#000000" LINK="#0080FF" ALINK="0080FF" VLINK="#0055AA" MARGINWIDTH="8" MARGINHEIGHT="0" LEFTMARGIN="8" TOPMARGIN="0">
<%

dim rs 
dim strsql
If (UserRole_ = "Admin") or (UserRole_ = "Voucher") or (UserRole_ = "FMC") Then

sPeriod = sYearP&sMonthP
ePeriod = eYearP&eMonthP
'response.write sPeriod & ePeriod 
strsql = "Select * From vwMonthlyBilling Where YearP+MonthP>='" & sPeriod & "' and YearP+MonthP<='" & ePeriod & "'"
strFilter=""
If Agency_ <> "X" then
	strFilter=strFilter & " and Agency='" & Agency_ & "'"
End If

If Section_ <> "X" then
	strFilter=strFilter & " and Office='" & Section_ & "'"
End If

If SectionGroup_ <> "X" then
	strFilter=strFilter & " and SectionGroup='" & SectionGroup_ & "'"
End If

If EmpID_ <> "X" then
	strFilter =strFilter & " and EmpID='" & EmpID_ & "'"
End If

'response.write EmpID_ 
If Status_ <> "0" then
	strFilter =strFilter & " and ProgressID=" & Status_
End If

If SentStatus_ <> "0" then
	strFilter =strFilter & " and SendMailStatusID=" & SentStatus_
End If
'response.write SentStatus_ 

If BillType = "O" then
	strFilter =strFilter & " and OfficePhoneBillRp>0 "
ElseIf BillType = "H" then
	strFilter =strFilter & " and HomePhoneBillRp>0 "
ElseIf BillType = "C" then
	strFilter =strFilter & " and CellPhoneBillRp>0 "
ElseIf BillType = "S" then
	strFilter =strFilter & " and TotalShuttleBillRp >0 "
End If

'strsql = strsql  & strFilter
strsql = strsql  & strFilter & " Order By " & SortBy_ & " "& Order_
'response.write strsql & "<br>"

set DataRS = server.createobject("adodb.recordset")
set DataRS=BillingCon.execute(strsql)


'response.write strsql

if not DataRS.eof Then

%>
<%If BillType = "X" then %>
	<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" width="100%"  class="FontText">
	    <TR BGCOLOR="#330099" align="center">
        	 <TD rowspan="2" width="3%" align="right" class="style5"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
	         <TD rowspan="2" width="15%" class="style5"><strong><label STYLE=color:#FFFFFF>EmpID</label></strong></TD>
	         <TD rowspan="2" width="15%" class="style5"><strong><label STYLE=color:#FFFFFF>Employee Name</label></strong></TD>
        	 <TD rowspan="2" width="10%" class="style5"><strong><label STYLE=color:#FFFFFF>Billing Period</label></strong></TD>
        	 <TD rowspan="2" width="10%" class="style5"><strong><label STYLE=color:#FFFFFF>Phone Number</label></strong></TD>
	         <TD rowspan="2" width="8%" class="style5"><strong><label STYLE=color:#FFFFFF>Section</label></strong></TD>
		 <TD colspan="5" class="style5"><strong><label STYLE=color:#FFFFFF>Billing Amount (Kn.)</label></strong></TD>
	         <TD rowspan="2" width="50px" class="style5"><strong><label STYLE=color:#FFFFFF>Sent Status</label></strong></TD>
	         <TD rowspan="2" width="80px" class="style5"><strong><label STYLE=color:#FFFFFF>Sent Date</label></strong></TD>
	         <TD rowspan="2" class="style5"><strong><label STYLE=color:#FFFFFF>Email</label></strong></TD>
	         <TD rowspan="2" class="style5"><strong><label STYLE=color:#FFFFFF>Status</label></strong></TD>
	    </TR>
	    <tr BGCOLOR="#330099" align="center">
        	 <TD width="10%" class="style5"><strong><label STYLE=color:#FFFFFF>Home Phone</label></strong></TD>
	         <TD width="10%" class="style5"><strong><label STYLE=color:#FFFFFF>Office Phone</label></strong></TD>
        	 <TD width="10%" class="style5"><strong><label STYLE=color:#FFFFFF>Mobile Phone</label></strong></TD>
	         <TD width="10%" class="style5"><strong><label STYLE=color:#FFFFFF>Shuttle Bus</label></strong></TD>
        	 <TD width="8%" class="style5"><strong><label STYLE=color:#FFFFFF>Total</label></strong></TD>
	    </tr>
<%else%>
	<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" width="100%"  class="FontText">
	    <TR BGCOLOR="#330099" align="center">
        	 <TD  width="3%" align="right" class="style5"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
	<!--         <TD  width="15%" class="style5"><strong><label STYLE=color:#FFFFFF>EmpID</label></strong></TD>  -->
	         <TD  class="style5"><strong><label STYLE=color:#FFFFFF>Employee Name</label></strong></TD>
        	 <TD  width="12%" class="style5"><strong><label STYLE=color:#FFFFFF>Billing Period</label></strong></TD>
	         <TD  width="15%" class="style5"><strong><label STYLE=color:#FFFFFF>Section</label></strong></TD>
        	 <TD  width="10%" class="style5"><strong><label STYLE=color:#FFFFFF>Phone Number</label></strong></TD>
 	<!--     <TD  width="12%" class="style5"><strong><label STYLE=color:#FFFFFF>Bill Type</label></strong></TD>  -->
		 <TD  class="style5"><strong><label STYLE=color:#FFFFFF>Billing Amount (Kn.)</label></strong></TD>
 		 <TD class="style5"><strong><label STYLE=color:#FFFFFF>Personal Usage (Kn.)</label></strong></TD>
	         <TD  class="style5"><strong><label STYLE=color:#FFFFFF>Sent Status</label></strong></TD>
	         <TD  class="style5"><strong><label STYLE=color:#FFFFFF>Sent Date</label></strong></TD>
	         <TD  class="style5"><strong><label STYLE=color:#FFFFFF>Email</label></strong></TD>
	         <TD  class="style5"><strong><label STYLE=color:#FFFFFF>Status</label></strong></TD>
	    </TR>
	<!--	    <tr BGCOLOR="#330099" align="center">
        	 <TD width="10%" class="style5"><strong><label STYLE=color:#FFFFFF>In Kuna (Kn.)</label></strong></TD>
	         <TD width="10%" class="style5"><strong><label STYLE=color:#FFFFFF>in US Dollar ($)</label></strong></TD>
	    </tr>  -->
<%end if%>
<% 
   dim no_  
   no_ = 1
   do while not DataRS.eof
		TotalBillingRp_ = 0
	   if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
%>
      
	   <TR bgcolor="<%=bg%>">
	        <TD align="right"><%=no_ %></TD>
	        <TD><%=DataRS("EmpName") %></TD>
	        <TD align="right"><%= DataRS("MonthP")%>-<%= DataRS("YearP")%></TD>
	        <TD><%=DataRS("Office") %> </TD>	        
		<TD>&nbsp;<%=DataRS("MobilePhone") %></TD>
<%If BillType <> "X" then 
		If BillType ="O" then
			BillTypeDesc ="Office Phone"
		ElseIf BillType ="C" then
			BillTypeDesc ="Cell Phone"
		ElseIf BillType ="H" then
			BillTypeDesc ="Home Phone"
		ElseIf BillType ="S" then
			BillTypeDesc ="Shuttle Bus"
		End If
%>
	    <!--      <TD><%=BillTypeDesc %></TD>  -->
<%end if%>
<%		If (BillType = "H") or (BillType = "X") Then 
			If CDbl(DataRS("HomePhoneBillRp")) > 0 then %>
				<td align="right">
					<%= formatnumber(DataRS("HomePhonePrsBillRp"),-1) %>
				</td>
<%				TotalBillingRp_ = TotalBillingRp_ + cdbl(DataRS("HomePhonePrsBillRp")) %>
<%			Else %>
				<td align="right">-</td>
<%			End If%>
<%		End If %>
<%		If (BillType = "O")  or (BillType = "X") Then 
			If CDbl(DataRS("OfficePhoneBillRp")) > 0 then %>
				<td align="right">
					<%= formatnumber(DataRS("OfficePhonePrsBillRp"),-1) %>
				</td>
<%			TotalBillingRp_ = cdbl(TotalBillingRp_ )+cdbl(DataRS("OfficePhonePrsBillRp"))%>
<%			Else %>
				<td align="right">-</td>
<%			End If%>
<%		End If %>
<%		If (BillType = "C")  or (BillType = "X") Then 
			If CDbl(DataRS("CellPhoneBillRp")) > 0 then %>
				<td align="right">
					<%= formatnumber(DataRS("CellPhoneBillRp"),-1) %>
				</td>
<%			TotalBillingRp_ = cdbl(TotalBillingRp_ )+cdbl(DataRS("CellPhoneBillRp"))%>
<%			Else %>
				<td align="right">0</td>
<%			End If%>
<%		End If %>
<%		If (BillType = "S")  or (BillType = "X") Then 
			If CDbl(DataRS("TotalShuttleBillRp")) > 0 then %>
				<td align="right">
					<%= formatnumber(DataRS("TotalShuttleBillRp"),-1) %>
				</td>
<%			TotalBillingRp_ = cdbl(TotalBillingRp_ )+cdbl(DataRS("TotalShuttleBillRp"))%>
<%			Else %>
				<td align="right">-</td>
<%			End If%>
<%		End If %>


<%		If (BillType = "H") Then 
			If CDbl(DataRS("HomePhonePrsBillRp")) > 0 Then %>
				<td align="right">
					<%= formatnumber(DataRS("HomePhonePrsBillRp"),-1) %>
				</td>
<%			Else %>
				<td align="right">-</td>
<%			End If %>
<%		End If %>

<%		If (BillType = "O") Then
			If CDbl(DataRS("OfficePhonePrsBillRp")) > 0 Then %>
			<td align="right">
				<%= formatnumber(DataRS("OfficePhonePrsBillRp"),-1) %>
			</td>
<%			Else %>
				<td align="right">-</td>
<%			End If %>
<%		End If %>
<%		If (BillType = "C") Then 
			If CDbl(DataRS("CellPhonePrsBillRp")) > 0 Then %>
				<td align="right">
					<%= formatnumber(DataRS("CellPhonePrsBillRp"),-1) %>
				</td>
<%			Else%>
				<td align="right">0</td>
<%			End If %>
<%		End If %>
<%		If (BillType = "S") Then
			If CDbl(DataRS("TotalShuttleBillRp")) > 0 Then %>
				<td align="right">
					<%= formatnumber(DataRS("TotalShuttleBillRp"),-1) %>
				</td>
<%			Else%>
				<td align="right">-</td>
<%			End If %>
<%		End If %>
<%		If CDbl(DataRS("HomePhonePrsBillDlr")) > 0 Then
			If (BillType = "H") Then %>
				<td align="right">
					<%= formatnumber(DataRS("HomePhonePrsBillDlr"),-1) %>
				</td>
<%			End If %>
<%		ElseIf (CDbl(DataRS("HomePhonePrsBillDlr")) = 0) and (BillType = "H") Then %>
			<td align="right">-</td>
<%		End If %>
<%		If CDbl(DataRS("OfficePhonePrsBillDlr")) > 0 Then 
			If (BillType = "O") Then %>
			<td align="right">
				<%= formatnumber(DataRS("OfficePhonePrsBillDlr"),-1) %>
			</td>
<%			End If %>
<%		ElseIf (CDbl(DataRS("OfficePhonePrsBillDlr")) = 0) and (BillType = "O") Then %>
			<td align="right">-</td>
<%		End If %>
<!-- <%		If CDbl(DataRS("CellPhonePrsBillDlr")) > 0 Then
			If (BillType = "C") Then %>
			<td align="right">
				<%= formatnumber(DataRS("CellPhonePrsBillDlr"),-1) %>
			</td>
<%			End If %>
<%		ElseIf (CDbl(DataRS("CellPhonePrsBillDlr")) = 0) and (BillType = "C")  Then %>
			<td align="right">-</td>
<%		End If %>   -->
<%		If CDbl(DataRS("TotalShuttleBillDlr")) > 0 Then
			If (BillType = "S") Then %>
				<td align="right">
					<%= formatnumber(DataRS("TotalShuttleBillDlr"),-1) %>
				</td>
<%			End If %>
<%		ElseIf (CDbl(DataRS("TotalShuttleBillDlr")) = 0) and (BillType = "S") Then %>
			<td align="right">-</td>
<%		End If %>

<%If BillType = "X" then 
%>
	        <TD align="right"><FONT color=#330099 size=2><%= formatnumber(TotalBillingRp_,-1) %></font></TD>
<%End if%>
	        <TD><%=DataRS("SendMailStatusDesc") %></TD>
	        <TD>&nbsp;<%=DataRS("SendMailDate") %></TD>
	        <TD>&nbsp;<%=DataRS("EmailAddress") %></TD>
	        <TD><%=DataRS("ProgressDesc") %></TD>
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


