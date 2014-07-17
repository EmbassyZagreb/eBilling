<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<html>
<head>

<script type="text/javascript">
function checkall(obj)
{
	var c = document.frmOfficePhoneBilling.elements.length
	for (var x=0; x<frmOfficePhoneBilling.elements.length; x++)
	{
		cbElement = frmOfficePhoneBilling.elements[x]
		if (cbElement.type == "checkbox")
		{
			cbElement.checked= obj.checked?true:false
		}
	}
}

function validate_form()
{
	valid = true;
	msg = ""


	if (document.frmOfficePhoneBilling.txtSpvEmail.value != "" )
	{
		var alnum="a-zA-Z0-9";
		exp="^[^@\\s]+@(["+alnum+"+\\-]+\\.)+["+alnum+"]["+alnum+"]["+alnum+"]?$";
		emailregexp = new RegExp(exp);

		result = document.frmOfficePhoneBilling.txtSpvEmail.value.match(emailregexp);
		if (result == null)
		{
			msg = msg + "Invalid data type for supervisor email address !!!\n"
			valid = false;
		}
	}
	else
	{
		msg = "Please fill in your supervisor mail !!!\n"		
		valid = false;
	}

	if (valid == false)
	{
		alert(msg)
	}
	return valid;
}
</script>
<% 
 
Extension_ = Request("Extension")
'response.write "User :" & userName_ & "<br>"
MonthP_ = Request("MonthP")
'response.write MonthP_ & "<br>"
YearP_ = Request("YearP")
'response.write YearP_ & "<br>"
Period_ = MonthP_ & " - " & YearP_ 
%> 

<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
   <TR>
  	<TD COLSPAN="4" class="title" align="center">Office Phone Bill Detail</TD>
   </TR>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
           
<%
strsql = "Select * From vwMonthlyBilling Where WorkPhone ='" & Extension_ & "' and MonthP='" & MonthP_ & "' and YearP='" & YearP_ & "'"
'response.write strsql & "<br>"
set rsOfficePhone = server.createobject("adodb.recordset") 
set rsOfficePhone = BillingCon.execute(strsql)
if not rsOfficePhone.eof then
	EmpID_ = rsOfficePhone("EmpID")
	EmpName_ = rsOfficePhone("EmpName")
	Office_ = rsOfficePhone("Office")
	SupervisorEmail_ = rsOfficePhone("SupervisorEmail")
	Notes_ = rsOfficePhone("Notes")
	SpvRemark_ = rsOfficePhone("SupervisorRemark")
	TotalOffBill_ = rsOfficePhone("OfficePhoneBillRp")
	TotalOfficePhonePrsBillRp_ = rsOfficePhone("OfficePhonePrsBillRp")
	ProgressID_ = rsOfficePhone("ProgressID")
end if
'response.write SupervisorEmail_ & "<br>"
'response.write ProgressID_ & "<br>"
%>
<form method="post" action="UpdateOfficePhoneBilling.asp" name="frmOfficePhoneBilling" onSubmit="return validate_form();"> 
<table cellspadding="1" cellspacing="0" width="80%" bgColor="white" align="center">  
<%if not rsOfficePhone.eof then%>
<tr>
	<td colspan="6" align="center"><u><b>Billing period (Month - Year) :</b> <a class="FontContent"><%=MonthP_ %> - <%=YearP_ %> </a></u></td>
</tr>
<tr>
	<td colspan="6" align="Center">
		<table cellspadding="0" cellspacing="0" bordercolor="black" border="1" width="90%" bgColor="white">  
		<tr bgcolor="#330099" align="center" cellpadding="0" cellspacing="0" >
			<TD width="5%"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
		       	<TD><strong><label STYLE=color:#FFFFFF>Dialed Date/time</label></strong></TD>
			<TD width="20%"><strong><label STYLE=color:#FFFFFF>Dialed Number</label></strong></TD>
			<TD><strong><label STYLE=color:#FFFFFF>Area Dialed</label></strong></TD>
			<TD><strong><label STYLE=color:#FFFFFF>Duration</label></strong></TD>
			<TD width="10%"><strong><label STYLE=color:#FFFFFF>Amount (Kn.)</label></strong></TD>
			<TD width="10%"><strong><label STYLE=color:#FFFFFF>Personal used</label></strong><br>
<%			if (ProgressID_ = 1) or (ProgressID_ = 3) then %>
				<input type="checkbox" name="cbAll" value="true" onclick="checkall(this)" />
<%			end if %>
			</TD>
		</tr>
		<%
		strsql = "Exec spGetBilling 'Detail','" & Extension_ & "','" & MonthP_ & "','" & YearP_ & "'"
		'response.write strsql & "<br>"
		set rsOfficePhone = BillingCon.execute(strsql) 
		Dim no_
		no_ = 1 
		do while not rsOfficePhone.eof
   			if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
'			response.write no_
		%>
			<tr bgcolor="<%=bg%>">
				<td align="right"><%=No_%>&nbsp;</td>
			        <td><FONT color=#330099 size=2>&nbsp;<%=rsOfficePhone("DialedDatetime")%></font></td> 
		        	<td><FONT color=#330099 size=2>&nbsp;<%=rsOfficePhone("DialedNumber")%></font></td> 
		        	<td><FONT color=#330099 size=2>&nbsp;<%=rsOfficePhone("AreaDialed")%></font></td> 
		        	<td><FONT color=#330099 size=2>&nbsp;<%=rsOfficePhone("CallDuration")%></font></td> 
			        <td align="right"><FONT color=#330099 size=2><%=formatnumber(rsOfficePhone("Cost"),-1)%>&nbsp;</font></td> 
<%'			if (ProgressID_ = 1) or (ProgressID_ = 3) then %>
<%			if cdbl(ProgressID_ )<4 then %>
			        <td align="center">
				<%if rsOfficePhone("isPersonal") = "Y" then%>
					<Input type="Checkbox" Id="cbPersonal" name="cbPersonal" Value='<%=rsOfficePhone("CallRecordID")%>' Checked>				
				<%else%>
					<Input type="Checkbox" Id="cbPersonal" name="cbPersonal" Value='<%=rsOfficePhone("CallRecordID")%>' >
				<%end if%>
				</td>
<%			else%>
			        <td align="center">
				<%if rsOfficePhone("isPersonal") = "Y" then%>
					<Input type="Checkbox" Id="cbPersonal" name="cbPersonal" Value='<%=rsOfficePhone("CallRecordID")%>' Checked disabled>				
				<%else%>
					<Input type="Checkbox" Id="cbPersonal" name="cbPersonal" Value='<%=rsOfficePhone("CallRecordID")%>'  disabled>
				<%end if%>
				</td>
<%			end if %>
		<%   
			rsOfficePhone.movenext
			no_ = no_ + 1
		loop
		%>
		</table>
	</td>
</tr>
<tr>
	<td colspan="6" align="center">
		<table cellspadding="0" cellspacing="0" bordercolor="black" width="90%" bgColor="white">  
		<tr>
			<td align="right" colspan="3"><b>Sub Total (Kn.) </b>&nbsp;</td>
			<td width="15%" class="FontContent" align="right"><b><%=formatnumber(TotalOffBill_,-1)%></b>&nbsp;</td>
			<td width="10%" class="FontContent" align="right"><b><u><%=formatnumber(TotalOfficePhonePrsBillRp_ ,-1)%></u></b>&nbsp;</td>
		</tr>
<!--
		<tr>
			<td align="right"><b>Payment Status</b>&nbsp;</td>
			<td width="1%">:</td>
			<td width="1%">&nbsp;kn.</td>
			<td class="FontContent" align="right"><b><%=Status_ %></b></td>
			<td width="15%">&nbsp;</td>
		</tr>
-->
		</table>
	</td>
</tr>
<tr>
	<td colspan="6" align="center">&nbsp;</td>
</tr>
<%'		if (ProgressID_ = 1) or (ProgressID_ = 3) then%>
<%		if cdbl(ProgressID_)<4 then%>
<tr>
	<td colspan="6" align="center">	
		<input type="submit" name="btnSubmit" Value="Update Change(s)" />&nbsp;&nbsp;
		<input type="hidden" name="txtExtension" value='<%=Extension_ %>' />
		<input type="hidden" name="txtMonthP" value='<%=MonthP_%>' />
		<input type="hidden" name="txtYearP" value='<%=YearP_%>' />
		<input type="hidden" name="txtEmpID" value='<%=EmpID_ %>' />
		<input type="hidden" name="txtEmpName" value='<%=EmpName_%>' />
		<input type="hidden" name="txtPeriod" value='<%=Period_%>' />
		<input type="hidden" name="txtOffice" value='<%=Office_%>' />
		<input type="hidden" name="txtTotalCost" value='<%=TotalOffBill_ %>' />
	<td>
</tr>
<%		end if%>
<%else%>
<tr>
	<td colspan="6" align="center">&nbsp;</td>	
</tr>
<tr>
	<td colspan="6"><hr></td>
</tr>
<tr>
	<td colspan="6" align="center">there is no data.</td>	
</tr>
<%end if%>
</table>
</form>
</BODY>
</html>