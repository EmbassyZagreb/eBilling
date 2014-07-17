<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<html>
<head>

<% 
 
userName_ = Request("UserName")
'response.write "User :" & userName_ & "<br>"
MonthP_ = Request("MonthP")
'response.write MonthP_ & "<br>"
YearP_ = Request("YearP")
'response.write YearP_ & "<br>"
%> 

<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
   <TR>
  	<TD COLSPAN="4" class="title" align="center">Shuttle Bus Detail</TD>
   </TR>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
           
<%
strsql = "Exec spGetShuttleBusList '" & userName_ & "','" & MonthP_ & "','" & YearP_ & "'"
'response.write strsql & "<br>"
set rsShuttle = server.createobject("adodb.recordset") 
set rsShuttle = BillingCon.execute(strsql) 

%>

<table cellspadding="1" cellspacing="0" width="100%" bgColor="white">  
<%if not rsShuttle.eof then%>
  <tr>
	<td colspan="6" align="center"><u>Billing period (Month - Year) : <a class="FontContent"><%=MonthP_ %> - <%=YearP_ %> </a></u></td>
  </tr>
<tr>
	<td colspan="6">
	<table align="center" cellpadding="1" cellspacing="0" width="40%" border="1" bordercolor="black"> 
	<TR BGCOLOR="#000099" align="center" cellpadding="0" cellspacing="0" >
		<TD width="6%"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
		<TD><strong><label STYLE=color:#FFFFFF>Date</label></strong></TD>
		<TD width="10%"><strong><label STYLE=color:#FFFFFF>AM</label></strong></TD>
		<TD width="10%"><strong><label STYLE=color:#FFFFFF>PM</label></strong></TD>
		<TD width="20%"><strong><label STYLE=color:#FFFFFF>Tot. Shuttle Qty</label></strong></TD>
		<TD width="20%"><strong><label STYLE=color:#FFFFFF>Tot. Shuttle Bill($)</label></strong></TD>
	</TR>    
<% 
		dim no_ , TotalQty_ , TotalAmount_ 
		no_ = 1 
		TotalQty_ = 0
		TotalAmount_ = 0
		do while not rsShuttle.eof 
	   	if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
			TotalQty_ = TotalQty_ + rsShuttle("TotalPerDay")
%> 
 	<TR bgcolor="<%=bg%>">
		<td align="right"><%=No_%>&nbsp;</td>
        	<td><FONT color=#330099 size=2><%=rsShuttle("ShuttleDate")%>&nbsp;</font></td> 
        	<td align="right"><FONT color=#330099 size=2><%=rsShuttle("AM")%>&nbsp;</font></td> 
		<td align="right"><FONT color=#330099 size=2><%=rsShuttle("PM")%>&nbsp;</font></td> 
        	<td align="right"><FONT color=#330099 size=2><%=rsShuttle("TotalPerDay")%>&nbsp;</font></td> 
		<td align="right"><FONT color=#330099 size=2><%=rsShuttle("TotalAmountPerDay")%>&nbsp;</font></td>
  	 </TR>
<%   
		TotalAmount_ = TotalAmount_ + formatnumber(rsShuttle("TotalAmountPerDay"),-1)
		'response.write rsShuttle("TotalAmountPerDay")
		'response.write TotalAmount_ 
 		rsShuttle.movenext
   		no_ = no_ + 1
	loop
%>	
	<tr>
		<td align="right" colspan="4"><b>Total&nbsp;</b></td>
		<td width="10%" class="FontContent" align="right"><b><%=formatnumber(TotalQty_ ,-1)%></b></td>
		<td width="10%" class="FontContent" align="right"><b><%=formatnumber(TotalAmount_  ,-1)%></b></td>
	</tr>
	</table>
	</td>
</tr>
<tr>
	<td colspan="6"><hr></td>
</tr>
<%else%>
<tr>
	<td colspan="6" align="center">&nbsp;</td>	
</tr>
<tr>
	<td colspan="6" align="center">there is not data.</td>	
</tr>
<%end if%>
</table>
</BODY>
</html>