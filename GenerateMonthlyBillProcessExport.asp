<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
   <%
Response.ContentType ="application/vnd.ms-excel" 
Response.Buffer  =  True 
Response.Clear() 
%> 
<html>
<head>
<script language="JavaScript" src="calendar.js"></script>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</head>
<BODY BGCOLOR="#FFFFFF" TEXT="#000000" LINK="#0080FF" ALINK="0080FF" VLINK="#0055AA" MARGINWIDTH="8" MARGINHEIGHT="0" LEFTMARGIN="8" TOPMARGIN="0">
<% 
	MonthP = Request("MonthP")
	'response.write MonthP
	MonthP="09"
	YearP = Request("YearP")
	'response.write YearP
	YearP ="2013"

	strsql = "Select * from vwProgressLog Where MonthP='" & MonthP & "' and YearP='" & YearP & "' order by Description"
	set SummaryRS = server.createobject("adodb.recordset")
	set SummaryRS = BillingCon.execute(strsql)
'	response.write strsql 
	if not SummaryRS.eof then
%>
		<table align="center" cellpadding="1" cellspacing="0" width="600px" border="1" bordercolor="black"> 
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
	        	<TD align="right"><FONT color=#330099 size=2><%=no_ %>&nbsp;</font></TD>
        		<TD>&nbsp;<%=SummaryRS("Description") %></TD>
        		<TD align="right"><%=formatnumber(SummaryRS("TotalRecord"),-1) %>&nbsp;</TD>
	        	<TD align="right"><%=formatnumber(SummaryRS("TotalBill"),-1) %>&nbsp;</TD>
		</TR>
<%
			TotalRecord_ = formatnumber(cdbl(TotalRecord_)+cdbl(SummaryRS("TotalRecord")),-1)
			TotalBilling_ = formatnumber(cdbl(TotalBilling_)+cdbl(SummaryRS("TotalBill")),-1)
			SummaryRS.movenext
		   	no_ = no_ + 1 
		loop 
%>
		<TR cellpadding="0" cellspacing="0" >
	       		<TD colspan="2" align="center"><strong>Total</strong></TD>
			<TD align="right"><strong><%=TotalRecord_ %></strong>&nbsp;</TD>
		       	<TD align="right"><strong><%=TotalBilling_ %></strong>&nbsp;</TD>
		</TR>    

<%
	end if

%>
</BODY>
</html>