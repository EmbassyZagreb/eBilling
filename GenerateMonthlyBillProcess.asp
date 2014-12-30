<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<html>
<head>
<script language="JavaScript" src="calendar.js"></script>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</head>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">GENERATE MONTHLY BILL</TD>
   </TR>
<tr>
        <td colspan="3" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
        <td align="right"><FONT color=#330099 size=2><A HREF="GenerateMonthlyBill.asp">Back</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
<% 
BillingCon.CommandTimeout = 120


if AlwaysExemptedCallType_ = "" or IsNull(AlwaysExemptedCallType_) Then AlwaysExemptedCallType_ = "t#is $tR!ng !$ uN&e|ivea&|e"
if ExemptedIfOfficialCallType_ = "" or IsNull(ExemptedIfOfficialCallType_) Then ExemptedIfOfficialCallType_ = "t#is $tR!ng !$ uN&e|ivea&|e"

 dim user_ 
 dim user1_  
 dim rst 
 dim strsql
 
user_ = request.servervariables("remote_user")
user1_ = user_  'user1_ = right(user_,len(user_)-4)
strsql = "select * from Users where loginId='" & user1_ & "'"
set RS_Query = server.createobject("adodb.recordset")
'response.write strsql & "<br>"
set RS_Query = BillingCon.execute(strsql)

if not RS_Query.eof then
	UserRole_ = RS_Query("RoleID")
end if

MonthP = Request.Form("MonthList")
'response.write MonthP
'MonthP="08"
YearP = Request.Form("YearList")
'response.write YearP
'YearP ="2013"
'EmpID = Request.Form("cmbEmp")
MobilePhone = Request.Form("cmbMobilePhone")
 
strsql = "select * from ExchangeRate where ExchangeMonth='" & MonthP & "' and ExchangeYear='" & YearP & "'"
set RS_Query = BillingCon.execute(strsql)
If RS_Query.eof then

'********************  added to avoid exchange rate info  ******************** 
	State_ = "I"
	ExchangeID_ = 0
	ExchangeMonth_ = MonthP
	ExchangeYear_ = YearP
	ExchangeRate_ = 5.55
        strsql = "Exec spExchangeRate_IUD '" & State_ & "'," & ExchangeID_ & ",'" & ExchangeMonth_ & "','" & ExchangeYear_ & "'," & ExchangeRate_
       'response.write strsql 
        BillingCon.execute strsql
'*****************************************************************************

	'response.write "<div class='Hint'>Please input exchange rate for period :<b> " & MonthP & " - " & YearP & "</b>, before generates monthly bill !!!</div>"
'else
end if

	strsql = "Exec spGenerateMonthlyBilling '" & MonthP & "','" & YearP & "','" & MobilePhone & "','%" & AlwaysExemptedCallType_ & "%','%" & ExemptedIfOfficialCallType_ & "%','" & user1_ & "'"
	set RS_Query = BillingCon.execute(strsql)
	response.write "<div class='Hint2'>Process generates monthly bill completed !!!</div>"


	strsql = "Select * from vwProgressLog Where MonthP='" & MonthP & "' and YearP='" & YearP & "' order by Description"
	set SummaryRS = server.createobject("adodb.recordset")
	set SummaryRS = BillingCon.execute(strsql)
'	response.write strsql 
	if not SummaryRS.eof then
%>
		<table align="center" cellpadding="2" cellspacing="0" width="65%" border="1" bordercolor="#EEEEEE"> 
		<TR BGCOLOR="#000099" align="center" cellpadding="0" cellspacing="0" >
			<TD width="3px"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
	       		<TD><strong><label STYLE=color:#FFFFFF>Billing Status</label></strong></TD>
			<TD><strong><label STYLE=color:#FFFFFF>Number of Record</label></strong></TD>
		       	<TD><strong><label STYLE=color:#FFFFFF>Total Billing</label></strong></TD>
		</TR>    
<% 
		dim no_  
   		no_ = 1
   		TotalRecord_ = 0
  		TotalBilling_ = 0
   		do while not SummaryRS.eof
	   		if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
%>
		 	<TR bgcolor="<%=bg%>">
		        	<TD align="right"><%=no_ %>&nbsp;</font></TD>
	        		<TD>&nbsp;<a href="LogDetail.asp?MonthP=<%=SummaryRS("MonthP")%>&YearP=<%=SummaryRS("YearP")%>&ProgressID=<%=SummaryRS("ProgressID")%> " target="blank"><%=SummaryRS("Description") %></a></TD>

				<%
		 		if SummaryRS("ProgressID") >10 then
				%>
	        			<TD align="right"><font STYLE=color:#999999><%=formatnumber(SummaryRS("TotalRecord"),0) %>&nbsp;</font></TD>
	        			<TD align="right"><font STYLE=color:#999999><%=formatnumber(SummaryRS("TotalBill"),2) %>&nbsp;</font></TD>
				<%else
					TotalRecord_ = formatnumber(cdbl(TotalRecord_)+cdbl(SummaryRS("TotalRecord")),0)
					TotalBilling_ = formatnumber(cdbl(TotalBilling_)+cdbl(SummaryRS("TotalBill")),-1)
				%>
	        			<TD align="right"><%=formatnumber(SummaryRS("TotalRecord"),0) %>&nbsp;</font></TD>
	        			<TD align="right"><%=formatnumber(SummaryRS("TotalBill"),-1) %>&nbsp;</font></TD>	
				<%End If%>
			</TR>
<%

			SummaryRS.movenext
		   	no_ = no_ + 1 
		loop 
%>
		<TR cellpadding="0" cellspacing="0" >
	       		<TD colspan="2" align="center"><strong>Total</strong></TD>
			<TD align="right"><strong><%=TotalRecord_ %></strong>&nbsp;</TD>
		       	<TD align="right"><strong><%=TotalBilling_ %></strong>&nbsp;</TD>
		</TR>    
		</table>
		<br>
		<div align="center"><a href="GenerateMonthlyBillProcessExport.asp?MonthP=<%=MonthP%>&YearP=<%=YearP%>">Save Log</a></div>
<%
	end if
	'Close the connection with the database and free all database resources
	Set RS_Query = Nothing
	BillingCon.Close
	Set BillingCon = Nothing
%>
</BODY>
</html>