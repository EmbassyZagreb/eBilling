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

curMonth_ = month(date())
curYear_ = year(date())
if len(curMonth_)= 1 then
	curMonth_ = "0" & curMonth_
end if

Agency_ = request("Agency")
Section_ = request("Section")
SectionGroup_ = request("SectionGroup")
EmpID_ = request("EmpID")
EmailAddress_ = request("EmailAddress")
BillType = request("BillType")
Status_ = request("Status")
SortBy_ = Request("SortBy")	
Order_ = Request("OrderList")	

NoRecord_ = request("NoRecord")
if NoRecord_ = "" Then NoRecord_ = Request.Form("cmbNoRecord")
if NoRecord_ = "" then
	NoRecord_ = 1
end if

rbPeriod_= request.form("rbPeriod")
if rbPeriod_ = "" then
	rbPeriod_ = "X"
End if
%>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<body>
<%

dim rs 
dim strsql
If (UserRole_ = "Admin") or (UserRole_ = "Voucher") or (UserRole_ = "FMC") Then

strsql = "Select CeilingAmount From PaymentDueDate"
set DataRS=BillingCon.execute(strsql)

CeilingAmount_ = DataRS("CeilingAmount")

sPeriod = sYearP&sMonthP
ePeriod = eYearP&eMonthP
'response.write sPeriod & ePeriod 
'strsql = "Select * From vwMonthlyBilling Where ProgressID=5 and YearP+MonthP>='" & sPeriod & "' and YearP+MonthP<='" & ePeriod & "'"
'strsql = "Select * From vwMonthlyBilling Where ProgressID=1 and YearP+MonthP>='" & sPeriod & "' and YearP+MonthP<='" & ePeriod & "'"

strsql = "Select EmpID, Min(YearP+MonthP) As StartPeriod, Max(YearP+MonthP) As EndPeriod, EmpName, Office, WorkPhone, MobilePhone, HomePhone "_
	&", SUM(HomePhonePrsBillRp) As HomePhonePrsBillRp, SUM(HomePhonePrsBillDlr) As HomePhonePrsBillDlr, SUM(OfficePhonePrsBillRp) As OfficePhonePrsBillRp, SUM(OfficePhonePrsBillDlr) As OfficePhonePrsBillDlr, SUM(CellPhonePrsBillRp) As CellPhonePrsBillRp, SUM(CellPhonePrsBillDlr) As CellPhonePrsBillDlr "_
	&", SUM(TotalShuttleBillRp) As TotalShuttleBillRp, SUM(TotalShuttleBillDlr) As TotalShuttleBillDlr, SUM(TotalBillingRp) As TotalBillingRp, SUM(TotalBillingDlr) As TotalBillingDlr"_
	&", EmailAddress, SUM(TotalBillingAmountPrsRp) As TotalBillingAmountPrsRp"_
	&", SUM(TotalBillingAmountPrsDlr) As TotalBillingAmountPrsDlr From vwMonthlyBilling Where ProgressID in(4,8)"

'Where YearP+MonthP>='201301' and YearP+MonthP<='201310' and ProgressID=4 and SendMailStatusID=1 and CellPhoneBillRp>0 
GroupBY_ = "Group By EmpID,  EmpName, Office, WorkPhone, MobilePhone, HomePhone, EmailAddress"

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
'response.write "EmailAddress_ :" & EmailAddress_ 
If EmailAddress_ <> "" then
	If EmailAddress_ = "-" then
		strFilter =strFilter & " and EmailAddress=''"
	Else
		strFilter =strFilter & " and EmailAddress like '%" & EmailAddress_  & "%'"
	End If
End If

'response.write EmpID_ 
If Status_ <> "0" then
'	if Status_ ="99" Then
'		strFilter =strFilter & " and ProgressID=6 and CellPhonePrsBillRp<=" & CeilingAmount_ 
'	else
		strFilter =strFilter & " and ProgressID=" & Status_

'	end if
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

'strsql = strsql  & strFilter & " Order By EmpName,YearP,MonthP"
strsql = strsql  & strFilter & GroupBY_ & " Order By " & SortBy_ & " "& Order_
'response.write strsql & "<br>"

set DataRS = server.createobject("adodb.recordset")
DataRS.CursorLocation = 3
DataRS.Open strsql,BillingCon
'set DataRS=BillingCon.execute(strsql)

'response.write "NoRecord_" & NoRecord_

dim intPageSize,PageIndex,TotalPages 
dim RecordCount,RecordNumber,Count 
intpageSize = NoRecord_ 
PageIndex=request("PageIndex")

if PageIndex ="" then PageIndex=1 

if not DataRS.eof then
	RecordCount = DataRS.RecordCount   
	'response.write RecordCount & "<br>"
	RecordNumber=(intPageSize * PageIndex) - intPageSize 
	'response.write RecordNumber
	DataRS.PageSize =intPageSize 
	DataRS.AbsolutePage = PageIndex
	TotalPages=DataRS.PageCount 
	'response.write TotalPages & "<br>"
End If
'response.write strsql

dim intPrev,intNext 	
intPrev=PageIndex - 1 
intNext=PageIndex +1 


if not DataRS.eof Then

%>
<%If BillType = "X" then %>
	<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" width="120%"  class="FontText">
	    <TR align="center">
        	 <TD rowspan="2" width="35px" align="right"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
	         <TD rowspan="2"><strong><label STYLE=color:#FFFFFF>Employee Name</label></strong></TD>
        	 <TD colspan="2"><strong><label STYLE=color:#FFFFFF>Period</label></strong></TD>
	         <TD rowspan="2"><strong><label STYLE=color:#FFFFFF>Section</label></strong></TD>
	         <TD rowspan="2"><strong><label STYLE=color:#FFFFFF>Phone Number</label></strong></TD>
		 <TD colspan="5"><strong><label STYLE=color:#FFFFFF>Billing Amount (Kn.)</label></strong></TD>
	         <TD rowspan="2"><strong><label STYLE=color:#FFFFFF>Email</label></strong></TD>
	    </TR>
	    <tr align="center">
        	 <TD><strong><label STYLE=color:#FFFFFF>Start</label></strong></TD>
        	 <TD><strong><label STYLE=color:#FFFFFF>End</label></strong></TD>
        	 <TD width="100px"><strong><label STYLE=color:#FFFFFF>Home Phone</label></strong></TD>
	         <TD width="100px"><strong><label STYLE=color:#FFFFFF>Office Phone</label></strong></TD>
        	 <TD width="100px"><strong><label STYLE=color:#FFFFFF>Mobile Phone</label></strong></TD>
	         <TD width="100px"><strong><label STYLE=color:#FFFFFF>Shuttle Bus</label></strong></TD>
        	 <TD width="100px"><strong><label STYLE=color:#FFFFFF>Total</label></strong></TD>
	    </tr>
<%else%>
	<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" width="120%"  class="FontText">
	    <TR align="center">
        	 <TD rowspan="2" width="35px" align="right"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
	         <TD rowspan="2"><strong><label STYLE=color:#FFFFFF>Employee Name</label></strong></TD>
        	 <TD colspan="2"><strong><label STYLE=color:#FFFFFF>Period</label></strong></TD>
	         <TD rowspan="2"><strong><label STYLE=color:#FFFFFF>Section</label></strong></TD>
	         <TD rowspan="2"><strong><label STYLE=color:#FFFFFF>Phone Number</label></strong></TD>
	         <TD rowspan="2" width="100px"><strong><label STYLE=color:#FFFFFF>Bill Type</label></strong></TD>
		 <TD rowspan="2" ><strong><label STYLE=color:#FFFFFF>Billing Amount (Kn.)</label></strong></TD>
		 <TD colspan="2"><strong><label STYLE=color:#FFFFFF>Personal Usage (Kn.)</label></strong></TD>
	         <TD rowspan="2"><strong><label STYLE=color:#FFFFFF>Email</label></strong></TD>
	    </TR>
	    <tr align="center">
        	 <TD><strong><label STYLE=color:#FFFFFF>Start</label></strong></TD>
        	 <TD><strong><label STYLE=color:#FFFFFF>End</label></strong></TD>
        	 <TD width="10%"><strong><label STYLE=color:#FFFFFF>In Kuna (Kn.)</label></strong></TD>
	         <TD width="10%"><strong><label STYLE=color:#FFFFFF>in US Dollar ($)</label></strong></TD>
	    </tr>
<%end if%>
<%   
 dim no_  
  
if (NoRecord_ = 1) then
 no_ = 1
 do while not DataRS.eof

	   TotalBillingRp_ = 0
	   if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
%>
      
	   <TR bgcolor="<%=bg%>">
	        <TD align="right"><FONT color=#330099 size=2><%=no_ %>&nbsp;</font></TD>
	        <TD>&nbsp;<%=DataRS("EmpName") %></TD>
	        <TD align="right"><FONT color=#330099 size=2><%=Right(DataRS("StartPeriod"),2) %> - <%=Left(DataRS("StartPeriod"),4) %></font>&nbsp;</TD>
	        <TD align="right"><FONT color=#330099 size=2><%=Right(DataRS("EndPeriod"),2) %> - <%=Left(DataRS("EndPeriod"),4) %></font>&nbsp;</TD>
	        <TD><FONT color=#330099 size=2>&nbsp;<%=DataRS("Office") %> </font></TD>
	        <TD><FONT color=#330099 size=2>&nbsp;<%=DataRS("MobilePhone") %> </font></TD>
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
	        <TD><FONT color=#330099 size=2>&nbsp;<%=BillTypeDesc %> </font></TD>
<%end if%>
<%		If (BillType = "H") or (BillType = "X") Then 
			If CDbl(DataRS("HomePhonePrsBillRp")) > 0 then %>
				<td align="right"><%= formatnumber(DataRS("HomePhonePrsBillRp"),-1) %></td>
<%				TotalBillingRp_ = TotalBillingRp_ + cdbl(DataRS("HomePhonePrsBillRp")) %>
<%			Else %>
				<td align="right">-&nbsp;</td>
<%			End If%>
<%		End If %>
<%		If (BillType = "O")  or (BillType = "X") Then 
			If CDbl(DataRS("OfficePhonePrsBillRp")) > 0 then %>
				<td align="right"><%= formatnumber(DataRS("OfficePhonePrsBillRp"),-1) %></td>
<%			TotalBillingRp_ = cdbl(TotalBillingRp_ )+cdbl(DataRS("OfficePhonePrsBillRp"))%>
<%			Else %>
				<td align="right">-&nbsp;</td>
<%			End If%>
<%		End If %>
<%		If (BillType = "C")  or (BillType = "X") Then 
			If CDbl(DataRS("CellPhonePrsBillRp")) > 0 then %>
				<td align="right"><%= formatnumber(DataRS("CellPhonePrsBillRp"),-1) %></td>
<%			TotalBillingRp_ = cdbl(TotalBillingRp_ )+cdbl(DataRS("CellPhonePrsBillRp"))%>
<%			Else %>
				<td align="right">-&nbsp;</td>
<%			End If%>
<%		End If %>
<%		If (BillType = "S")  or (BillType = "X") Then 
			If CDbl(DataRS("TotalShuttleBillRp")) > 0 then %>
				<td align="right"><%= formatnumber(DataRS("TotalShuttleBillRp"),-1) %></td>
<%			TotalBillingRp_ = cdbl(TotalBillingRp_ )+cdbl(DataRS("TotalShuttleBillRp"))%>
<%			Else %>
				<td align="right">-&nbsp;</td>
<%			End If%>
<%		End If %>


<%		If (BillType = "H") Then 
			If CDbl(DataRS("HomePhonePrsBillRp")) > 0 Then %>
				<td align="right"><%= formatnumber(DataRS("HomePhonePrsBillRp"),-1) %></td>
<%			Else %>
				<td align="right">-&nbsp;</td>
<%			End If %>
<%		End If %>

<%		If (BillType = "O") Then
			If CDbl(DataRS("OfficePhonePrsBillRp")) > 0 Then %>
			<td align="right"><%= formatnumber(DataRS("OfficePhonePrsBillRp"),-1) %></td>
<%			Else %>
				<td align="right">-&nbsp;</td>
<%			End If %>
<%		End If %>
<%		If (BillType = "C") Then 
			If CDbl(DataRS("CellPhonePrsBillRp")) > 0 Then %>
				<td align="right"><%= formatnumber(DataRS("CellPhonePrsBillRp"),-1) %></td>
<%			Else%>
				<td align="right">-&nbsp;</td>
<%			End If %>
<%		End If %>
<%		If (BillType = "S") Then
			If CDbl(DataRS("TotalShuttleBillRp")) > 0 Then %>
				<td align="right"><%= formatnumber(DataRS("TotalShuttleBillRp"),-1) %></td>
<%			Else%>
				<td align="right">-&nbsp;</td>
<%			End If %>
<%		End If %>
<%		If CDbl(DataRS("HomePhonePrsBillDlr")) > 0 Then
			If (BillType = "H") Then %>
				<td align="right"><%= formatnumber(DataRS("HomePhonePrsBillDlr"),-1) %></td>
<%			End If %>
<%		ElseIf (CDbl(DataRS("HomePhonePrsBillDlr")) = 0) and (BillType = "H") Then %>
			<td align="right">-&nbsp;</td>
<%		End If %>
<%		If CDbl(DataRS("OfficePhonePrsBillDlr")) > 0 Then 
			If (BillType = "O") Then %>
			<td align="right"><%= formatnumber(DataRS("OfficePhonePrsBillDlr"),-1) %></td>
<%			End If %>
<%		ElseIf (CDbl(DataRS("OfficePhonePrsBillDlr")) = 0) and (BillType = "O") Then %>
			<td align="right">-&nbsp;</td>
<%		End If %>
<%		If CDbl(DataRS("CellPhonePrsBillDlr")) > 0 Then
			If (BillType = "C") Then %>
			<td align="right"><%= formatnumber(DataRS("CellPhonePrsBillDlr"),-1) %></td>
<%			End If %>
<%		ElseIf (CDbl(DataRS("CellPhonePrsBillDlr")) = 0) and (BillType = "C")  Then %>
			<td align="right">-&nbsp;</td>
<%		End If %>
<%		If CDbl(DataRS("TotalShuttleBillDlr")) > 0 Then
			If (BillType = "S") Then %>
				<td align="right"><%= formatnumber(DataRS("TotalShuttleBillDlr"),-1) %></td>
<%			End If %>
<%		ElseIf (CDbl(DataRS("TotalShuttleBillDlr")) = 0) and (BillType = "S") Then %>
			<td align="right">-&nbsp;</td>
<%		End If %>

<%If BillType = "X" then 
%>
	        <TD align="right"><FONT color=#330099 size=2><%= formatnumber(TotalBillingRp_,-1) %>&nbsp;</font></TD>
<%End if%>

	        <TD><FONT color=#330099 size=2>&nbsp;<%=DataRS("EmailAddress") %> </font></TD>
	    </TR>

<%   
	   DataRS.movenext
	   no_ = no_ + 1 
   loop 
	PageNo=1
%>
</table>
<%
Else
   no_ = 1 + ((PageIndex*intPageSize)-intPageSize)
   Count=1 
   'response.write "intPageSize :" & intPageSize
   do while not DataRS.eof and cdbl(Count)<= cdbl(intPageSize)

	   TotalBillingRp_ = 0
	   if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
%>
      
	   <TR bgcolor="<%=bg%>">
	        <TD align="right"><FONT color=#330099 size=2><%=no_ %>&nbsp;</font></TD>
	        <TD>&nbsp;<%=DataRS("EmpName") %></TD>
	        <TD align="right"><FONT color=#330099 size=2><%=Right(DataRS("StartPeriod"),2) %> - <%=Left(DataRS("StartPeriod"),4) %></font>&nbsp;</TD>
	        <TD align="right"><FONT color=#330099 size=2><%=Right(DataRS("EndPeriod"),2) %> - <%=Left(DataRS("EndPeriod"),4) %></font>&nbsp;</TD>
	        <TD><FONT color=#330099 size=2>&nbsp;<%=DataRS("Office") %> </font></TD>
	        <TD><FONT color=#330099 size=2>&nbsp;<%=DataRS("MobilePhone") %> </font></TD>
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
	        <TD><FONT color=#330099 size=2>&nbsp;<%=BillTypeDesc %> </font></TD>
<%end if%>
<%		If (BillType = "H") or (BillType = "X") Then 
			If CDbl(DataRS("HomePhonePrsBillRp")) > 0 then %>
				<td align="right"><%= formatnumber(DataRS("HomePhonePrsBillRp"),-1) %></td>
<%				TotalBillingRp_ = TotalBillingRp_ + cdbl(DataRS("HomePhonePrsBillRp")) %>
<%			Else %>
				<td align="right">-&nbsp;</td>
<%			End If%>
<%		End If %>
<%		If (BillType = "O")  or (BillType = "X") Then 
			If CDbl(DataRS("OfficePhonePrsBillRp")) > 0 then %>
				<td align="right"><%= formatnumber(DataRS("OfficePhonePrsBillRp"),-1) %></td>
<%			TotalBillingRp_ = cdbl(TotalBillingRp_ )+cdbl(DataRS("OfficePhonePrsBillRp"))%>
<%			Else %>
				<td align="right">-&nbsp;</td>
<%			End If%>
<%		End If %>
<%		If (BillType = "C")  or (BillType = "X") Then 
			If CDbl(DataRS("CellPhonePrsBillRp")) > 0 then %>
				<td align="right"><%= formatnumber(DataRS("CellPhonePrsBillRp"),-1) %></td>
<%			TotalBillingRp_ = cdbl(TotalBillingRp_ )+cdbl(DataRS("CellPhonePrsBillRp"))%>
<%			Else %>
				<td align="right">-&nbsp;</td>
<%			End If%>
<%		End If %>
<%		If (BillType = "S")  or (BillType = "X") Then 
			If CDbl(DataRS("TotalShuttleBillRp")) > 0 then %>
				<td align="right"><%= formatnumber(DataRS("TotalShuttleBillRp"),-1) %></td>
<%			TotalBillingRp_ = cdbl(TotalBillingRp_ )+cdbl(DataRS("TotalShuttleBillRp"))%>
<%			Else %>
				<td align="right">-&nbsp;</td>
<%			End If%>
<%		End If %>


<%		If (BillType = "H") Then 
			If CDbl(DataRS("HomePhonePrsBillRp")) > 0 Then %>
				<td align="right"><%= formatnumber(DataRS("HomePhonePrsBillRp"),-1) %></a></td>
<%			Else %>
				<td align="right">-&nbsp;</td>
<%			End If %>
<%		End If %>

<%		If (BillType = "O") Then
			If CDbl(DataRS("OfficePhonePrsBillRp")) > 0 Then %>
			<td align="right"><%= formatnumber(DataRS("OfficePhonePrsBillRp"),-1) %></td>
<%			Else %>
				<td align="right">-&nbsp;</td>
<%			End If %>
<%		End If %>
<%		If (BillType = "C") Then 
			If CDbl(DataRS("CellPhonePrsBillRp")) > 0 Then %>
				<td align="right"><%= formatnumber(DataRS("CellPhonePrsBillRp"),-1) %></td>
<%			Else%>
				<td align="right">-&nbsp;</td>
<%			End If %>
<%		End If %>
<%		If (BillType = "S") Then
			If CDbl(DataRS("TotalShuttleBillRp")) > 0 Then %>
				<td align="right"><%= formatnumber(DataRS("TotalShuttleBillRp"),-1) %></td>
<%			Else%>
				<td align="right">-&nbsp;</td>
<%			End If %>
<%		End If %>
<%		If CDbl(DataRS("HomePhonePrsBillDlr")) > 0 Then
			If (BillType = "H") Then %>
				<td align="right"><%= formatnumber(DataRS("HomePhonePrsBillDlr"),-1) %></td>
<%			End If %>
<%		ElseIf (CDbl(DataRS("HomePhonePrsBillDlr")) = 0) and (BillType = "H") Then %>
			<td align="right">-&nbsp;</td>
<%		End If %>
<%		If CDbl(DataRS("OfficePhonePrsBillDlr")) > 0 Then 
			If (BillType = "O") Then %>
			<td align="right"><%= formatnumber(DataRS("OfficePhonePrsBillDlr"),-1) %></td>
<%			End If %>
<%		ElseIf (CDbl(DataRS("OfficePhonePrsBillDlr")) = 0) and (BillType = "O") Then %>
			<td align="right">-&nbsp;</td>
<%		End If %>
<%		If CDbl(DataRS("CellPhonePrsBillDlr")) > 0 Then
			If (BillType = "C") Then %>
			<td align="right"><%= formatnumber(DataRS("CellPhonePrsBillDlr"),-1) %></td>
<%			End If %>
<%		ElseIf (CDbl(DataRS("CellPhonePrsBillDlr")) = 0) and (BillType = "C")  Then %>
			<td align="right">-&nbsp;</td>
<%		End If %>
<%		If CDbl(DataRS("TotalShuttleBillDlr")) > 0 Then
			If (BillType = "S") Then %>
				<td align="right"><%= formatnumber(DataRS("TotalShuttleBillDlr"),-1) %></td>
<%			End If %>
<%		ElseIf (CDbl(DataRS("TotalShuttleBillDlr")) = 0) and (BillType = "S") Then %>
			<td align="right">-&nbsp;</td>
<%		End If %>

<%If BillType = "X" then 
%>
	        <TD align="right"><FONT color=#330099 size=2><%= formatnumber(TotalBillingRp_,-1) %>&nbsp;</font></TD>
<%End if%>
	        <TD><FONT color=#330099 size=2>&nbsp;<%=DataRS("EmailAddress") %> </font></TD>
	    </TR>

<%   
		Count=Count +1
		'response.write Count
	   DataRS.movenext
	   no_ = no_ + 1 
   loop 
	PageNo=1
%>
</table>

<table width="100%">
	<tr>
		<td align="right">
<%
		Do while PageNo<=TotalPages 
			if trim(pageNo) = trim(PageIndex) Then
%>		
				<label class="ActivePage"><%=PageNo%></label>&nbsp;
<%			Else%>
				<a href="SendNotificationPayGov.asp?PageIndex=<%=PageNo%>&NoRecord=<%=NoRecord_%>&sMonthP=<%=sMonthP%>&sYearP=<%=sYearP%>&eMonthP=<%=eMonthP%>&eYearP=<%=eYearP%>&Agency=<%=Agency_%>&Section=<%=Section_%>&SectionGroup=<%=SectionGroup_%>&EmpID=<%=EmpID_%>&BillType=<%=BillType%>&Status=<%=Status_%>&SentStatus=<%=SentStatus_%>&SortBy=<%=SortBy_%>&Order=<%=Order_%>"><%=PageNo%></a>&nbsp;
<%	
			End If						
			PageNo=PageNo+1
		Loop
%>
		</td>
	</tr>
</table>
<%end if%>
	<input type="hidden" Name="txtsMonthP" value="<%=sMonthP%>" />
	<input type="hidden" Name="txteMonthP" value="<%=eMonthP%>" />
	<input type="hidden" Name="txtsYearP" value="<%=sYearP%>" />
	<input type="hidden" Name="txteYearP" value="<%=eYearP%>" />
	<%'response.write "rbPeriod_:" & rbPeriod_%>
	<input type="hidden" Name="txtPeriod" value="<%=rbPeriod_%>" />
	<input type="hidden" Name="txtProgressID" value="<%=Status_%>" />
	<input type="hidden" Name="txtAgency" value="<%=Agency_ %>" />
	<input type="hidden" Name="txtSection" value="<%=Section_ %>" />
	<input type="hidden" Name="txtSectionGroup" value="<%=SectionGroup_%>" />
	<input type="hidden" Name="txtEmpID" value="<%=EmpID_ %>" />
	<input type="hidden" Name="txtEmailAddress" value="<%=EmailAddress_ %>" />
	<input type="hidden" Name="txtBillType" value="<%=BillType%>" />
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


