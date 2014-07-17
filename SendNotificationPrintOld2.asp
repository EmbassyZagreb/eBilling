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

<style type="text/css">
<!--

.FontText {
	font-size: small;
}


.Hint{
	color: gray;
	font-size: x-small;
}
-->
</style>
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
sYearP = request("sYearP")
eMonthP = request("eMonthP")
eYearP = request("eYearP")
Agency_ = request("Agency")
Section_ = request("Section")
EmpID_ = request("EmpID")
BillType = request("BillType")
%>
<TITLE>U.S. Mission Jakarta e-Billing</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<STYLE TYPE="text/css"><!--
  A:ACTIVE { color:#003399; font-size:8pt; font-family:Verdana; }
  A:HOVER { color:#003399; font-size:8pt; font-family:Verdana; }
  A:LINK { color:#003399; font-size:8pt; font-family:Verdana; }
  A:VISITED { color:#003399; font-size:8pt; font-family:Verdana; }
  body {scrollbar-3dlight-color:#FFFFFF; scrollbar-arrow-color:#E3DCD5; scrollbar-base-color:#FFFFFF; scrollbar-darkshadow-color:#FFFFFF;	scrollbar-face-color:#FFFFFF; scrollbar-highlight-color:#E3DCD5; scrollbar-shadow-color:#E3DCD5; }
  p { font-family: verdana; font-size: 12px; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; color: #003399; text-decoration: none}
  h3 { font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 16px; font-style: normal; line-height: normal; font-weight: bold; color: #003399; letter-spacing: normal; word-spacing: normal; font-variant: small-caps}
  td { font-family: verdana; font-size: 10px; font-style: normal; font-weight: normal; color: #000000}
  .title { font-size:14px; font-weight:bold; color:#000080; }
  .SubTitle { font-size:16px; font-weight:bold; color:#000080;  }
  A.menu { text-decoration:none; font-weight:bold; }
  A.mmenu { text-decoration:none; color:#FFFFFF; font-weight:bold; }
  .normal { font-family:Verdana,Arial; color:black}
  .style5 {color: #FFFFFF;}
  .ActivePage {color: red; font-weight:bold; }
--></STYLE>
</HEAD>
<BODY BGCOLOR="#FFFFFF" TEXT="#000000" LINK="#0080FF" ALINK="0080FF" VLINK="#0055AA" MARGINWIDTH="8" MARGINHEIGHT="0" LEFTMARGIN="8" TOPMARGIN="0">
<%

dim rs 
dim strsql
If (UserRole_ = "Admin") or (UserRole_ = "Voucher") Then

sPeriod = sYearP&sMonthP
ePeriod = eYearP&eMonthP
'response.write sPeriod & ePeriod 
strsql = "Select * From vwMonthlyBilling Where Status='Pending' and YearP+MonthP>='" & sPeriod & "' and YearP+MonthP<='" & ePeriod & "'"
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

If BillType = "O" then
	strFilter =strFilter & " and OfficePhoneBillRp>0 "
ElseIf BillType = "H" then
	strFilter =strFilter & " and HomePhoneBillRp>0 "
ElseIf BillType = "C" then
	strFilter =strFilter & " and CellPhoneBillRp>0 "
ElseIf BillType = "S" then
	strFilter =strFilter & " and TotalShuttleBillRp >0 "
End If

strsql = strsql  & strFilter
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
	         <TD rowspan="2" width="15%" class="style5"><strong><label STYLE=color:#FFFFFF>Employee Name</label></strong></TD>
        	 <TD rowspan="2" width="10%" class="style5"><strong><label STYLE=color:#FFFFFF>Billing Period</label></strong></TD>
	         <TD rowspan="2" width="8%" class="style5"><strong><label STYLE=color:#FFFFFF>Section</label></strong></TD>
		 <TD colspan="5" class="style5"><strong><label STYLE=color:#FFFFFF>Billing Amount (Rp.)</label></strong></TD>
	         <TD rowspan="2" class="style5"><strong><label STYLE=color:#FFFFFF>Email</label></strong></TD>
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
        	 <TD rowspan="2" width="3%" align="right"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
	         <TD rowspan="2"><strong><label STYLE=color:#FFFFFF>Employee Name</label></strong></TD>
        	 <TD rowspan="2" width="12%"><strong><label STYLE=color:#FFFFFF>Billing Period</label></strong></TD>
	         <TD rowspan="2" width="15%"><strong><label STYLE=color:#FFFFFF>Section</label></strong></TD>
	         <TD rowspan="2" width="12%"><strong><label STYLE=color:#FFFFFF>Bill Type</label></strong></TD>
		 <TD rowspan="2" ><strong><label STYLE=color:#FFFFFF>Billing Amount (Rp.)</label></strong></TD>
		 <TD colspan="2"><strong><label STYLE=color:#FFFFFF>Should be paid (Rp.)</label></strong></TD>
	         <TD rowspan="2"><strong><label STYLE=color:#FFFFFF>Email</label></strong></TD>
	    </TR>
	    <tr BGCOLOR="#330099" align="center">
        	 <TD width="10%"><strong><label STYLE=color:#FFFFFF>In Rupiah (Rp.)</label></strong></TD>
	         <TD width="10%"><strong><label STYLE=color:#FFFFFF>in US Dollar ($)</label></strong></TD>
	    </tr>
<%end if%>
<% 
   dim no_  
   no_ = 1
   do while not DataRS.eof
	   if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
%>
      
	   <TR bgcolor="<%=bg%>">
	        <TD align="right"><%=no_ %>&nbsp;</TD>
	        <TD>&nbsp;<%=DataRS("EmpName") %></TD>
	        <TD align="right"><%= DataRS("MonthP")%>-<%= DataRS("YearP")%>&nbsp;</TD>
	        <TD>&nbsp;<%=DataRS("Office") %> </TD>
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
	        <TD>&nbsp;<%=BillTypeDesc %></TD>
<%end if%>
<%		If (BillType = "H") Then 
			If CDbl(DataRS("HomePhoneBillRp")) > 0 then %>
				<td align="right"><%= formatnumber(DataRS("HomePhoneBillRp"),0) %></td>
<%			Else %>
				<td align="right">-&nbsp;</td>
<%			End If%>
<%		End If %>
<%		If (BillType = "O") Then 
			If CDbl(DataRS("OfficePhoneBillRp")) > 0 then %>
				<td align="right"><%= formatnumber(DataRS("OfficePhoneBillRp"),0) %></td>
<%			Else %>
				<td align="right">-&nbsp;</td>
<%			End If%>
<%		End If %>
<%		If (BillType = "C") Then 
			If CDbl(DataRS("CellPhoneBillRp")) > 0 then %>
				<td align="right"><%= formatnumber(DataRS("CellPhoneBillRp"),0) %></td>
<%			Else %>
				<td align="right">-&nbsp;</td>
<%			End If%>
<%		End If %>
<%		If (BillType = "S") Then 
			If CDbl(DataRS("TotalShuttleBillRp")) > 0 then %>
				<td align="right"><%= formatnumber(DataRS("TotalShuttleBillRp"),0) %></td>
<%			Else %>
				<td align="right">-&nbsp;</td>
<%			End If%>
<%		End If %>


<%		If CDbl(DataRS("HomePhonePrsBillRp")) > 0 Then
			If ((BillType = "X") or (BillType = "H")) Then %>
				<td align="right"><%= formatnumber(DataRS("HomePhonePrsBillRp"),0) %></td>
<%			End If %>
<%		ElseIf (CDbl(DataRS("HomePhonePrsBillRp")) = 0) and ((BillType = "X") or (BillType = "H")) Then %>
			<td align="right">-&nbsp;</td>
<%		End If %>
<%		If CDbl(DataRS("OfficePhonePrsBillRp")) > 0 Then 
			If ((BillType = "X") or (BillType = "O")) Then %>
			<td align="right"><%= formatnumber(DataRS("OfficePhonePrsBillRp"),0) %></td>
<%			End If %>
<%		ElseIf (CDbl(DataRS("OfficePhonePrsBillRp")) = 0) and ((BillType = "X") or (BillType = "O")) Then %>
			<td align="right">-&nbsp;</td>
<%		End If %>
<%		If CDbl(DataRS("CellPhonePrsBillRp")) > 0 Then
			If ((BillType = "X") or (BillType = "C")) Then %>
			<td align="right"><%= formatnumber(DataRS("CellPhonePrsBillRp"),0) %></td>
<%			End If %>
<%		ElseIf (CDbl(DataRS("CellPhonePrsBillRp")) = 0) and ((BillType = "X") or (BillType = "C"))  Then %>
			<td align="right">-&nbsp;</td>
<%		End If %>
<%		If CDbl(DataRS("TotalShuttleBillRp")) > 0 Then
			If ((BillType = "X") or (BillType = "S")) Then %>
				<td align="right"><%= formatnumber(DataRS("TotalShuttleBillRp"),0) %></td>
<%			End If %>
<%		ElseIf (CDbl(DataRS("TotalShuttleBillRp")) = 0) and ((BillType = "X") or (BillType = "S")) Then %>
			<td align="right">-&nbsp;</td>
<%		End If %>

<%		If CDbl(DataRS("HomePhonePrsBillDlr")) > 0 Then
			If (BillType = "H") Then %>
				<td align="right"><%= formatnumber(DataRS("HomePhonePrsBillDlr"),0) %></td>
<%			End If %>
<%		ElseIf (CDbl(DataRS("HomePhonePrsBillDlr")) = 0) and (BillType = "H") Then %>
			<td align="right">-&nbsp;</td>
<%		End If %>
<%		If CDbl(DataRS("OfficePhonePrsBillDlr")) > 0 Then 
			If (BillType = "O") Then %>
			<td align="right"><%= formatnumber(DataRS("OfficePhonePrsBillDlr"),0) %></td>
<%			End If %>
<%		ElseIf (CDbl(DataRS("OfficePhonePrsBillDlr")) = 0) and (BillType = "O") Then %>
			<td align="right">-&nbsp;</td>
<%		End If %>
<%		If CDbl(DataRS("CellPhonePrsBillDlr")) > 0 Then
			If (BillType = "C") Then %>
			<td align="right"><%= formatnumber(DataRS("CellPhonePrsBillDlr"),0) %></td>
<%			End If %>
<%		ElseIf (CDbl(DataRS("CellPhonePrsBillDlr")) = 0) and (BillType = "C")  Then %>
			<td align="right">-&nbsp;</td>
<%		End If %>
<%		If CDbl(DataRS("TotalShuttleBillDlr")) > 0 Then
			If (BillType = "S") Then %>
				<td align="right"><%= formatnumber(DataRS("TotalShuttleBillDlr"),0) %></td>
<%			End If %>
<%		ElseIf (CDbl(DataRS("TotalShuttleBillDlr")) = 0) and (BillType = "S") Then %>
			<td align="right">-&nbsp;</td>
<%		End If %>

<%If BillType = "X" then %>
	        <TD align="right"><FONT color=#330099 size=2><%= formatnumber(DataRS("TotalBillingRp"),0) %>&nbsp;</font></TD>
<%End if%>
	        <TD>&nbsp;<%=DataRS("EmailAddress") %></TD>
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
		<td align="center">not data.</td>
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
			<td>Please <a href="http://jakartaws01.eap.state.sbu/CSC">Submit Request </a> or contact Jakarta CSC Helpdesk at ext.9111.</td>
		</tr>
	</table>
<% end if %>
</body> 

</html>


