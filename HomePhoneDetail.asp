<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<html>
<head>

<script type="text/javascript">
function checkall(obj)
{
	var c = document.frmHomePhoneBilling.elements.length
	for (var x=0; x<frmHomePhoneBilling.elements.length; x++)
	{
		cbElement = frmHomePhoneBilling.elements[x]
		if (cbElement.type == "checkbox")
		{
			cbElement.checked= obj.checked?true:false
		}
	}
}

</script>
<% 
 
HomePhone_ = trim(Request("HomePhone"))
'response.write "HomePhone_  :" & HomePhone_ & "<br>"
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
  	<TD COLSPAN="4" class="title" align="center">Homephone Bill Detail</TD>
   </TR>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
<%
TotalHomePhoneBillRp_ = 0
TotalHomePhonePrsBillRp_ = 0

strsql = "Select * From vwMonthlyBilling Where HomePhone ='" & HomePhone_ & "' and MonthP='" & MonthP_ & "' and YearP='" & YearP_ & "'"
'response.write strsql & "<br>"
set rsHomePhone = server.createobject("adodb.recordset") 
set rsHomePhone = BillingCon.execute(strsql)
if not rsHomePhone.eof then
	EmpID_ = rsHomePhone("EmpID")
	EmpName_ = rsHomePhone("EmpName")
	Office_ = rsHomePhone("Office")
	SupervisorEmail_ = rsHomePhone("SupervisorEmail")
	Notes_ = rsHomePhone("Notes")
	SpvRemark_ = rsHomePhone("SupervisorRemark")
	TotalHomePhoneBillRp_ = rsHomePhone("HomePhoneBillRp")
	TotalHomePhonePrsBillRp_ = rsHomePhone("HomePhonePrsBillRp")
	ProgressID_ = rsHomePhone("ProgressID")
	'response.write "TotalHomePhoneBillRp_ :" & TotalHomePhoneBillRp_ 
	'response.write "TotalHomePhonePrsBillRp_ :" & TotalHomePhonePrsBillRp_ 
end if
%>           
<form method="post" action="HomePhoneDetailSave.asp" name="frmHomePhoneBilling"> 
<table cellspadding="1" cellspacing="0" width="60%" bgColor="white" align="center">  
<%if not rsHomePhone.eof then%>
  <tr>
	<td colspan="6" align="center"><u><b>Billing period (Month - Year) :</b> <a class="FontContent"><%=MonthP_ %> - <%=YearP_ %> </a></u></td>
  </tr>
  <tr>
	<td colspan="6" align="Center">
		<table cellspadding="0" cellspacing="0" bordercolor="black" border="1" width="80%" bgColor="white">  
		<tr bgcolor="#330099" align="center" cellpadding="0" cellspacing="0" >
			<TD width="5%"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
		       	<TD><strong><label STYLE=color:#FFFFFF>Dialed Date/time</label></strong></TD>
			<TD width="20%"><strong><label STYLE=color:#FFFFFF>Dialed Number</label></strong></TD>
			<TD width="15%"><strong><label STYLE=color:#FFFFFF>Amount (Kn.)</label></strong></TD>
			<TD width="15%"><strong><label STYLE=color:#FFFFFF>Personal used</label></strong><br>
<%			if (ProgressID_ = 1) or (ProgressID_ = 3) then %>
				<input type="checkbox" name="cbAll" value="true" onclick="checkall(this)" />
<%			end if %>
			</TD>
		</tr>
		<%
		strsql = "Exec spGetHomephone '2','" & HomePhone_ & "','" & MonthP_ & "','" & YearP_ & "'"
		'response.write strsql & "<br>"
		set rsHomePhone = BillingCon.execute(strsql) 
		no_ = 1 
		do while not rsHomePhone.eof
   			if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
		%>
			<tr bgcolor="<%=bg%>">
				<td align="right"><%=No_%>&nbsp;</td>
			        <td><FONT color=#330099 size=2>&nbsp;<%=rsHomePhone("DialedDatetime")%></font></td> 
		        	<td><FONT color=#330099 size=2>&nbsp;<%=rsHomePhone("DialedNumber")%></font></td> 
			        <td align="right"><FONT color=#330099 size=2><%=formatnumber(rsHomePhone("Cost"),-1)%>&nbsp;</font></td> 
<%'			if (ProgressID_ = 1) or (ProgressID_ = 3) then %>
<%			if cdbl(ProgressID_ )<4 then %>
			        <td align="center">
				<%if rsHomePhone("isPersonal") = "Y" then%>
					<Input type="Checkbox" Id="cbPersonal" name="cbPersonal" Value='<%=rsHomePhone("CallRecordID")%>' Checked>				
				<%else%>
					<Input type="Checkbox" Id="cbPersonal" name="cbPersonal" Value='<%=rsHomePhone("CallRecordID")%>' >
				<%end if%>
				</td>
<%			else%>
			        <td align="center">
				<%if rsHomePhone("isPersonal") = "Y" then%>
					<Input type="Checkbox" Id="cbPersonal" name="cbPersonal" Value='<%=rsHomePhone("CallRecordID")%>' Checked disabled>				
				<%else%>
					<Input type="Checkbox" Id="cbPersonal" name="cbPersonal" Value='<%=rsHomePhone("CallRecordID")%>'  disabled>
				<%end if%>
				</td>
<%			end if %>
		<%   
			rsHomePhone.movenext
			no_ = no_ + 1
		loop
		%>
		</tr>
		</table>
	</td>
  </tr>
  <tr>
	<td colspan="6" align="center">
		<table cellspadding="0" cellspacing="0" bordercolor="black" width="80%" bgColor="white">  
		<tr>
			<td align="right" colspan="3"><b>Sub Total (Kn.) </b>&nbsp;</td>
			<td width="15%" class="FontContent" align="right"><b><%=formatnumber(TotalHomePhoneBillRp_ ,-1)%></b>&nbsp;</td>
			<td width="15%" class="FontContent" align="right"><b><u><%=formatnumber(TotalHomePhonePrsBillRp_ ,-1)%></u></b>&nbsp;</td>
		</tr>
		</table>
	</td>
  </tr>
  <tr>
	<td colspan="6" align="center">&nbsp;</td>
  </tr>
<%'		if (ProgressID_ = 1) or (ProgressID_ = 3) then%>
<%		if cdbl(ProgressID_)< 4 then%>
  <tr>
	<td colspan="6" align="center">	
		<input type="submit" name="btnSubmit" Value="Update Change(s)" />&nbsp;&nbsp;
		<input type="hidden" name="txtHomePhone" value='<%=HomePhone_ %>' />
		<input type="hidden" name="txtMonthP" value='<%=MonthP_%>' />
		<input type="hidden" name="txtYearP" value='<%=YearP_%>' />
		<input type="hidden" name="txtEmpID" value='<%=EmpID_ %>' />
	<td>
  </tr>
	<%end if%>
<%else%>
<tr>
	<td colspan="6" align="center">&nbsp;</td>	
</tr>
<tr>
	<td colspan="6" align="center">there is no data.</td>	
</tr>
<%end if%>
</table>
</form>
</BODY>
</html>