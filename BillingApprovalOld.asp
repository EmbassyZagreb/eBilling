<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<html>
<head>

<script type="text/javascript">
function ValidateForm()
{
	valid = true;
	nRec = 0;

	for (var x=0; x<frmBillingApproval.SupervisorSign.length; x++)
	{
		if (frmBillingApproval.SupervisorSign[x].checked)
		{
			nRec++;
		}
	}

	if (nRec == 0)
	{
		alert("Please select data that you want to approve !!!");
		valid = false;
	}
	return valid;
}
</script>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY BGCOLOR="#FFFFFF" TEXT="#000000" LINK="#0080FF" ALINK="0080FF" VLINK="#0055AA" MARGINWIDTH="8" MARGINHEIGHT="0" LEFTMARGIN="8" TOPMARGIN="0">
  <Center><FONT COLOR=#009900><B>SENSITIVE BUT UNCLASSIFIED</Center></FONT></B>
  <BR>
<CENTER>
  <IMG SRC="images/embassytitle2.jpeg" WIDTH="661" HEIGHT="80" BORDER="0"> 
  <TABLE WIDTH="65%" BORDER="0" CELLPADDING="0" CELLSPACING="0">
  <CAPTION><H3 STYLE="font-size:17px;color:#000040">Mission Jakarta - Billing Application</H3></CAPTION>
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">OFFICE PHONE BILLING</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
<% 
 dim user_ 
 dim user1_  
 dim rst 
 dim strsql
 
 'user_ = request.servervariables("remote_user") 
  'user1_ = right(user_,len(user_)-4)
 'user1_ = "pranataw"
'response.write user1_ & "<br>"
LoginID = request("LoginID")
EmpID = request("EmpID") 
MonthP = request("Month")
YearP = request("Year")
'response.write EmpID_ & "<br>" & MonthP_ & "<br>"& YearP_ 


'if (session("Month") = "") or (session("Year") = "") then
'	strsql = "Select MonthP, YearP From Period"
'	'response.write strsql & "<br>"
'	set rsData = server.createobject("adodb.recordset") 
'	set rsData = BillingCon.execute(strsql)
'	if not rsData.eof then
'		session("Month") = rsData("MonthP")
'		session("Year") = rsData("YearP")
'	end if
'end if

'MonthP = session("Month")
'YearP = session("Year")
%>  
           
<table align="center" cellspadding="1" cellspacing="0" width="100%">  
<tr>
	<td colspan="2" align="center">Billing Period : <Label style="color:blue"><%=MonthP%> - <%=YearP%></label></td>
</tr>
</table>
<%
'strsql = "Exec spGetBilling 'Header','" & LoginID & "','" & MonthP & "','" & YearP & "'"
strsql = "Select * From vwMonthlyBilling Where EmpID='" & EmpID & "' and MonthP='" & MonthP & "' and YearP='" & YearP & "'"
'response.write strsql & "<br>"
set rsData = server.createobject("adodb.recordset") 
set rsData = BillingCon.execute(strsql) 
if not rsData.eof then
	EmpName_ = rsData("EmpName")
	Period_ = rsData("MonthP") & " - " & rsData("YearP")
	Office_ = rsData("Office")
	PersonalCost_ = rsData("OfficePhonePrsBillRp")
	TotalCost_ = rsData("TotalBillingRp")
	Ext_ = rsData("WorkPhone")
	ProgressID_ = rsData("ProgressID")
	ProgressDesc_ = rsData("ProgressDesc")
	EmpEmail_ = rsData("EmailAddress")
	SupervisorEmail_ = rsData("SupervisorEmail")
	Note_ = rsData("Notes")
end if
'response.write ProgressID_ 
%>
<table cellspadding="1" cellspacing="0" width="100%" bgColor="white">  
<tr>
          <td align="Left"><u><b>Personal Info<b></u></TD>
</tr>  
<tr>
	<td width="12%">Employee Name</td>
	<td width="1%">:</td>
	<td class="FontContent"><%=EmpName_%></td>
	<td width="20%">Billing Period (Month - Year)</td>
	<td width="1%">:</td>
	<td class="FontContent"><%=Period_%></td>
</tr>
<tr>
	<td>Agency / Office</td>
	<td width="1%">:</td>
	<td class="FontContent"><%=Office_%></td>
	<td>Personal Cost</td>
	<td width="1%">:</td>
	<td class="FontContent">Kn. <%= formatnumber(PersonalCost_ ,-1) %></td>
</tr>
<tr>
	<td>Phone Ext.</td>
	<td width="1%">:</td>
	<td class="FontContent"><%= Ext_ %></td>
	<td>Total Cost</td>
	<td width="1%">:</td>
	<td class="FontContent">Kn. <%= formatnumber(TotalCost_ ,-1) %></td>
</tr>
<tr>
	<td width="12%">Supervisor Email</td>
	<td width="1%">:</td>
	<td class="FontContent"><%=SupervisorEmail_ %></td>
	<td>Status</td>
	<td width="1%">:</td>
	<td class="FontContent"><%=ProgressDesc_ %></td>
</tr>
<tr>
	<td valign="top">Note</td>
	<td valign="top" width="1%">:</td>
	<td colspan="4">
		<TextArea name="txtNotes" Rows="5" Cols="80" Wrap readonly><%=Note_%></textarea>
	</td>
</tr>
<tr>
	<td align="Left" colspan="6"><u><b>Approval<b></u></TD>
</tr>
<tr>
	<td colspan="6">
	<form method="post" name="frmBillingApproval" action="BillingApprovalSave.asp" onsubmit="return ValidateForm()">	
	<table class="FontComment" width="100%">
	<tr>
		<td width="12%">Your decision</td>
		<td width="1%">:</td>
		<td class="FontContent">
<!--
			<select name="SupervisorSign">
				<option value="">-- Select --</option>
				<%if Status_ ="Approved" Then %>	
					<option value="A" Selected>Approve</option>
				<%Else%>
					<option value="A">Approve</option>
				<% End If %>
				<%if Status_ ="Correction" Then %>	
					<option value="C" Selected>Need Correction</option>
				<%Else%>
					<option value="C">Need Correction</option>
				<% End If %>
			</select>
-->
			<input type="radio" name="SupervisorSign" value="A">Approve</input>
			<input type="radio" name="SupervisorSign" value="C">Need Correction</input>
		</td>
	</tr>
	<tr>
		<td width="12%" valign="top">Remark/Correction(s)</td>
		<td width="1%" valign="top">:</td>
		<td><TextArea name="txtRemark" Rows="5" Cols="70" Wrap maxlength="500"></textarea></td>
  	</tr>
	<%if ProgressID_ = 2 then%>
	<tr align="center">
		<td align="center" colspan="3">
	        	<input type="submit" value="Submit">
        		&nbsp;<input type="button" value="Cancel" onClick="javascript:location.href='BillingApprovalList.asp'">
			<input type="hidden" name="txtExtension" value='<%=Ext_ %>' />
			<input type="hidden" name="txtMonthP" value='<%=MonthP%>' />
			<input type="hidden" name="txtYearP" value='<%=YearP%>' />
			<input type="hidden" name="txtEmpEmail" value='<%=EmpEmail_ %>' />

			<input type="hidden" name="txtEmpID" value='<%=EmpID %>' />
			<input type="hidden" name="txtEmpName" value='<%=EmpName_%>' />
			<input type="hidden" name="txtPeriod" value='<%=Period_%>' />
			<input type="hidden" name="txtOffice" value='<%=Office_%>' />
			<input type="hidden" name="txtPersonalCost" value='<%=PersonalCost_%>' />
			<input type="hidden" name="txtTotalCost" value='<%=TotalCost_%>' />
			<input type="hidden" name="txtNote" value='<%=Note_%>' />
	    	</td>


	</tr>
	<%End If%>
	</table>	
	</form>
	</td>
</tr>
<tr>
	<td colspan="6"><hr></td>
</tr>
<tr>
	<td align="Left" colspan="6"><u><b>Office Phone<b></u></TD>
</tr>
<tr>
	<td valign="top">&nbsp</td>
	<td valign="top" width="1%">&nbsp;</td>
	<td colspan="4" align="left">
		<table cellspadding="0" cellspacing="0" bordercolor="black" border="1" width="70%" bgColor="white">  
		<tr bgcolor="#330099" align="center" cellpadding="0" cellspacing="0" >
			<TD width="5%"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
		       	<TD><strong><label STYLE=color:#FFFFFF>Dialed Date/time</label></strong></TD>
			<TD width="20%"><strong><label STYLE=color:#FFFFFF>Dialed Number</label></strong></TD>
			<TD width="15%"><strong><label STYLE=color:#FFFFFF>Amount (Kn.)</label></strong></TD>
			<TD width="15%"><strong><label STYLE=color:#FFFFFF>Personal used</label></strong></TD>
		</tr>
		<%
		strsql = "Exec spGetBilling 'Detail','" & LoginID & "','" & MonthP & "','" & YearP & "'"
		'response.write strsql & "<br>"
		set rsData = BillingCon.execute(strsql) 
		no_ = 1 
		do while not rsData.eof
   			if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
		%>
			<tr bgcolor="<%=bg%>">
				<td align="right"><%=No_%>&nbsp;</td>
			        <td><FONT color=#330099 size=2>&nbsp;<%=rsData("DialedDatetime")%></font></td> 
		        	<td><FONT color=#330099 size=2>&nbsp;<%=rsData("DialedNumber")%></font></td> 
			        <td align="right"><FONT color=#330099 size=2><%=formatnumber(rsData("Cost"),-1)%>&nbsp;</font></td> 
			        <td align="center">
				<%if rsData("isPersonal") = "Y" then%>
					<Input type="Checkbox" Id="cbPersonal" name="cbPersonal" Value='<%=rsData("CallRecordID")%>' Checked disabled>
				<%else%>
					<Input type="Checkbox" Id="cbPersonal" name="cbPersonal" Value='<%=rsData("CallRecordID")%>' disabled>
				<%end if%>
				</td>
		<%   
			rsData.movenext
			no_ = no_ + 1
		loop
		%>
		</tr>
		</table>
	</td>
</tr>

</table>
</BODY>
</html>