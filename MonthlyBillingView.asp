<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<html>
<head>

<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">Monthly Billing</TD>
   </TR>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
<% 
 dim user_ 
 dim user1_  
 dim rst 
 dim strsql
 
user1_ = Request("LoginID")
'response.write user1_ & "<br>"

MonthP = Request("Month")
YearP = Request("Year")

Period_ = MonthP &"-"&YearP 
%>  
<table cellspadding="1" cellspacing="0" width="60%" bgColor="white">  
<%
strsql = "Select * from vwMonthlyBilling Where LoginID='" & user1_ & "' And MonthP='" & MonthP & "' And YearP='" & YearP & "'"
'response.write strsql & "<br>"
set rsData = server.createobject("adodb.recordset") 
set rsData = BillingCon.execute(strsql) 
Period_ = MonthP & " - " & YearP
'response.write Period_  & "<br>"
if not rsData.eof then
	EmpName_ = rsData("EmpName")
	Office_ = rsData("Agency") & " - " & rsData("Office")
	Position_ = rsData("WorkingTitle")
	OfficePhone_ = rsData("WorkPhone")
	HomePhone_ = rsData("HomePhone")
	MobilePhone_ = rsData("MobilePhone")
	ExchangeRate_ = rsData("ExchangeRate")
	HomePhoneBillRp_ = rsData("HomePhoneBillRp")
	HomePhoneBillDlr_ = rsData("HomePhoneBillDlr")
	OfficePhonePrsBillRp_ = rsData("OfficePhonePrsBillRp")
	OfficePhonePrsBillDlr_ = rsData("OfficePhonePrsBillDlr")
	OfficePhoneBillRp_ = rsData("OfficePhoneBillRp")
	OfficePhoneBillDlr_ = rsData("OfficePhoneBillDlr")
	TotalShuttleBillRp_ = rsData("TotalShuttleBillRp")
	TotalShuttleBillDlr_ = rsData("TotalShuttleBillDlr")
	TotalBillingRp_ = rsData("TotalBillingRp")
	TotalBillingDlr_ = rsData("TotalBillingDlr")
	TotalBilling_ = CDbl(HomePhoneBillRp_) + CDbl(OfficePhoneBillRp_)+ CDbl(TotalShuttleBillRp_)
	ProgressID_ = rsData("ProgressID")
	ProgressStatus_ = rsData("ProgressDesc")
'response.write Period_  & "<br>"
%>

<tr>
	<td colspan="6" align="center"><u>Billing Period (Month - Year) : <a class="FontContent"><%=Period_%></a></u></td>
</tr>
<tr>
          <td align="Left"><u><b>Personal Info<b></u></TD>
</tr>  
<tr>
	<td width="20%">Employee Name</td>
	<td width="1%">:</td>
	<td class="FontContent"><%=EmpName_%></td>
	<td>Agency / Office</td>
	<td width="1%">:</td>
	<td class="FontContent"><%=Office_%></td>
</tr>
<tr>
	<td>Position</td>
	<td width="1%">:</td>
	<td class="FontContent"><%=Position_ %></td>
	<td>Office Phone/Ext.</td>
	<td width="1%">:</td>
	<td class="FontContent"><%=OfficePhone_ %></td>
</tr>
<tr>
	<td>Homephone</td>
	<td width="1%">:</td>
	<td class="FontContent"><%=HomePhone_ %></td>
</tr>
<tr>
	<td>Mobile Phone</td>
	<td width="1%">:</td>
	<td class="FontContent"><%=MobilePhone_ %></td>
	<td>Exchange Rate</td>
	<td width="1%">:</td>
	<td class="FontContent">Kn. <%= FormatNumber(ExchangeRate_,-1) %> / Dollar</td>

</tr>
<tr>
	<td>Payment Status</td>
	<td width="1%">:</td>
	<td class="FontContent" colspan="4">
	<%If (ProgressID_ = 1) or (ProgressID_ = 3) then %>
		<a href="OfficePhoneDetail.asp?Extension=<%=OfficePhone_ %>&MonthP=<%=MonthP%>&YearP=<%=YearP%>" target="_blank"><%=ProgressStatus_%></a>
	<%Else%>
		<%=ProgressStatus_%>
	<%End if%>
	</td>
</tr>
<tr>
	<td colspan="6"><hr></td>
</tr>

<tr>
	<td align="Left" colspan="5"><u><b>Billing detail :<b></u></TD>
</tr>
<tr>
	<td colspan="6">*Click on each billing type for more detail</td>
</tr>
<tr>
	<td align="Left" colspan="6">
	<table cellspadding="1" border="1" bordercolor="black" cellspacing="0" width="100%" bgColor="white" border="0">  
	<tr align="center">
		<td rowspan="2"><b>Type</b></td>
		<td rowspan="2"><b>Billing (Kn.)</b></td>
		<td colspan="2"><b>Should be paid</b></td>
	</tr>
	<tr>
		<td align="center"><b>In Kuna (Kn.)</b></td>
		<td align="center"><b>In US Dollar ($)</b></td>
	</tr>
	<tr>
		<td><a href="OfficePhoneDetail.asp?Extension=<%=OfficePhone_ %>&MonthP=<%=MonthP%>&YearP=<%=YearP%>" target="_blank">Office Phone</a></td>
		<td width="20%" class="FontContent" align="right"><%=formatnumber(OfficePhoneBillRp_,-1) %>&nbsp;</td>
		<td width="20%" class="FontContent" align="right"><%=formatnumber(OfficePhonePrsBillRp_ ,-1) %>&nbsp;</td>
		<td width="20%" class="FontContent" align="right"><%=formatnumber(OfficePhonePrsBillDlr_,-1) %>&nbsp;</td>		
	</tr>
	<tr>
		<td><a href="HomePhoneDetail.asp?HomePhone=<%=HomePhone_%>&MonthP=<%=MonthP%>&YearP=<%=YearP%>" target="_blank">Home Phone</a></td>		
		<td class="FontContent" align="right"><%=formatnumber(HomePhoneBillRp_ ,-1) %>&nbsp;</td>
		<td class="FontContent" align="right"><%=formatnumber(HomePhoneBillRp_ ,-1) %>&nbsp;</td>
		<td class="FontContent" align="right"><%=formatnumber(HomePhoneBillDlr_ ,-1) %>&nbsp;</td>
	</tr>
	<tr>
		<td>Mobile Phone</td>
		<td class="FontContent" align="right">- &nbsp;</td>
		<td class="FontContent" align="right">- &nbsp;</td>
		<td class="FontContent" align="right">-&nbsp;</td>
	</tr>
	<tr>
		<td><a href="ShuttleBusBillDetail.asp?Username=<%=user1_ %>&MonthP=<%=MonthP%>&YearP=<%=YearP%>" target="_blank">Shuttle Bus</a></td>
		<td class="FontContent" align="right"><%=formatnumber(TotalShuttleBillRp_ ,-1) %>&nbsp;</td>
		<td class="FontContent" align="right"><%=formatnumber(TotalShuttleBillRp_ ,-1) %>&nbsp;</td>
		<td class="FontContent" align="right"><%=formatnumber(TotalShuttleBillDlr_,-1) %>&nbsp;</td>
	</tr>
	</table>
	</TD>
</tr>
<tr>
	<td colspan="6">
	<table cellspadding="1" cellspacing="0" width="100%" bgColor="white" border="0">
	<tr>
		<td align="center"><b>Total</b></td>
		<td width="20%" class="FontContent" align="right"><!-- <b><u><%=formatnumber(TotalBilling_ , -1) %></u></b> -->&nbsp;</td>
		<td width="20%" class="FontContent" align="right"><b><u><%=formatnumber(TotalBillingRp_ , -1) %></u></b>&nbsp;</td>
		<td width="20%" class="FontContent" align="right"><b><u><%=formatnumber(TotalBillingDlr_ ,-1) %></u></b>&nbsp;</td>
	</tr>
	</table>	
	</td>
</tr>
<%Else%>
<tr>
	<td colspan="6" align="center">there is no data.</td>	
</tr>
<% end if %>
</table>
</BODY>
</html>