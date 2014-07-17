<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<html>
<head>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
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
</head>
<!--#include file="Header.inc" -->
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
 
 user_ = request.servervariables("remote_user") 
  user1_ = right(user_,len(user_)-4)
user1_ = "agusaa"
'response.write user1_ & "<br>"

curMonth_ = month(date())
curYear_ = year(date())
if len(curMonth_)= 1 then
	curMonth_ = "0" & curMonth_
end if

MonthP = Request.Form("MonthList")
if MonthP ="" then
	MonthP = curMonth_ 
end if

YearP = Request.Form("YearList")
if YearP ="" then
	YearP = curYear_ 
end if

%>  
        
<%
strsql = "Exec spGetBilling 'Header','" & user1_ & "','" & MonthP & "','" & YearP & "'"
'response.write strsql & "<br>"
set rsData = server.createobject("adodb.recordset") 
set rsData = BillingCon.execute(strsql) 
if not rsData.eof then
	EmpName_ = rsData("EmpName")
	Period_ = rsData("MonthP") & " - " & rsData("YearP")
	Office_ = rsData("OfficeLocation")
	TotalCost_ = rsData("TotalCost")
	Ext_ = rsData("Extension")
	Status_ = rsData("Status")
end if
%>
<table cellspadding="1" cellspacing="0" width="60%" border="1" align="center">
<tr align="Center">
	<td colspan="2" align="center">
		<form method="post" name="frmSearch" Action="OfficePhoneBilling.asp">
		<table  width="100%">
		<tr bgcolor="#000099">
			<td height="25" colspan="7"><strong>&nbsp;<span class="style5">Search &amp; Sort By </span></strong></td>
		</tr>
		<tr>
			<td>&nbsp;Period&nbsp;</td>				
			<td>:</td>
			<td>
				<Select name="MonthList">
					<Option value="01" <%if MonthP ="01" then %>Selected<%End If%> >January</Option>
					<Option value="02" <%if MonthP ="02" then %>Selected<%End If%> >February</Option>
					<Option value="03" <%if MonthP ="03" then %>Selected<%End If%> >March</Option>
					<Option value="04" <%if MonthP ="04" then %>Selected<%End If%> >April</Option>
					<Option value="05" <%if MonthP ="05" then %>Selected<%End If%> >May</Option>
					<Option value="06" <%if MonthP ="06" then %>Selected<%End If%> >June</Option>
					<Option value="07" <%if MonthP ="07" then %>Selected<%End If%> >July</Option>
					<Option value="08" <%if MonthP ="08" then %>Selected<%End If%> >August</Option>
					<Option value="09" <%if MonthP ="09" then %>Selected<%End If%> >Sepetember</Option>
					<Option value="10" <%if MonthP ="10" then %>Selected<%End If%> >October</Option>
					<Option value="11" <%if MonthP ="11" then %>Selected<%End If%> >November</Option>
					<Option value="12" <%if MonthP ="12" then %>Selected<%End If%> >December</Option>
				</Select>&nbsp;
<%
				Year_ = Year(Date()) - 1
'					response.write Year_
%>

				<Select name="YearList">
<% 				Do While Year_ <= Year(Date()) %>
				<Option value='<%=Year_%>' <%if trim(Year_) = trim(YearP) then %>Selected<%End If%> ><%=Year_%></Option>		
<% 
					Year_ = Year_ + 1
				Loop %>	
				</Select>										
			</td>
			<td height="30" align="center">
				<input type="Button" name="btnBack" value="Back" onClick="Javascript:document.location.href('Default.asp');">
				<input type="submit" name="Submit" value="Search">
			</td>
		</tr>			
		</table>
		</form>
	</td>
</tr>	
</table><br>
<form method="post" action="UpdateOfficePhoneBilling.asp" name="frmOfficePhoneBilling" onSubmit="return validate_form();"> 
<table cellspadding="1" cellspacing="0" width="100%" bgColor="white">  
<%if not rsData.eof then%>
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
	<td>Total Cost</td>
	<td width="1%">:</td>
	<td class="FontContent">Kn. <%= formatnumber(TotalCost_ ,-1) %></td>
</tr>
<tr>
	<td>Phone Ext.</td>
	<td width="1%">:</td>
	<td class="FontContent"><%=rsData("Extension")%></td>
	<td>Status</td>
	<td width="1%">:</td>
	<td class="FontContent"><%=rsData("Status")%></td>

</tr>
<tr>
	<td colspan="6">
		<table cellspadding="1" cellspacing="0" bgColor="white" width="100%">  
		<tr>
			<td width="12%">Supervisor Email</td>
			<td width="1%">:</td>
			<td> <input type="input" name="txtSpvEmail" size="50" value='<%=rsData("SpvEmail")%>' /></td>
		</tr>
		<tr>
			<td valign="top">Note</td>
			<td valign="top" width="1%">:</td>
			<td>
				<TextArea name="txtNotes" Rows="5" Cols="90" Wrap <% if (Status_ <> "Pending") and (Status_ <> "Correction") then%>ReadOnly<%End If%> ><%=rsData("Notes")%></textarea>
			</td>
		</tr>
<%		if (Status_ <> "Pending")then%>
		<tr>
			<td colspan="3"><b>Remarks/Correction(s) :</b></td>
		</tr>
		<tr>
			<td colspan="3">
				<TextArea name="txtRemark" Rows="5" Cols="90" Wrap <% if (Status_ <> "Pending") or (Status_ <> "Correction") then%>ReadOnly<%End If%>><%=rsData("SpvRemark")%></textarea>
			</td>
		</tr>
<%		end if%>
		</table>
	</td>
</tr>
<tr>
	<td colspan="6"><hr></td>
</tr>
<tr>
	<td align="Left" colspan="6"><u><b>Billing Detail<b></u></TD>
</tr>
<tr>
	<td colspan="6" align="Center">
		<table cellspadding="0" cellspacing="0" bordercolor="black" border="1" width="80%" bgColor="white">  
		<tr bgcolor="#330099" align="center" cellpadding="0" cellspacing="0" >
			<TD width="5%"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
		       	<TD><strong><label STYLE=color:#FFFFFF>Dialed Date/time</label></strong></TD>
			<TD width="20%"><strong><label STYLE=color:#FFFFFF>Dialed Number</label></strong></TD>
			<TD width="15%"><strong><label STYLE=color:#FFFFFF>Amount (kn.)</label></strong></TD>
			<TD width="15%"><strong><label STYLE=color:#FFFFFF>Personal used</label></strong><br>
<%			if (Status_ = "Pending") or (Status_ = "Correction") then %>
				<input type="checkbox" name="cbAll" value="true" onclick="checkall(this)" />
<%			end if %>
			</TD>
		</tr>
		<%
		strsql = "Exec spGetBilling 'Detail','" & user1_ & "','" & MonthP & "','" & YearP & "'"
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
<%			if (Status_ = "Pending") or (Status_ = "Correction") then %>
			        <td align="center">
				<%if rsData("isPersonal") = "Y" then%>
					<Input type="Checkbox" Id="cbPersonal" name="cbPersonal" Value='<%=rsData("CallRecordID")%>' Checked>				
				<%else%>
					<Input type="Checkbox" Id="cbPersonal" name="cbPersonal" Value='<%=rsData("CallRecordID")%>' >
				<%end if%>
				</td>
<%			else%>
			        <td align="center">
				<%if rsData("isPersonal") = "Y" then%>
					<Input type="Checkbox" Id="cbPersonal" name="cbPersonal" Value='<%=rsData("CallRecordID")%>' Checked disabled>				
				<%else%>
					<Input type="Checkbox" Id="cbPersonal" name="cbPersonal" Value='<%=rsData("CallRecordID")%>'  disabled>
				<%end if%>
				</td>
<%			end if %>
		<%   
			rsData.movenext
			no_ = no_ + 1
		loop
		%>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td colspan="6" align="center">&nbsp;</td>	
</tr>
<%		if (Status_ = "Pending") or (Status_ = "Correction") then%>
<tr>
	<td colspan="6" align="center">	
		<input type="submit" name="btnSubmit" Value="Save" />&nbsp;&nbsp;
		<input type="submit" name="btnSubmit" Value="Save & Submit to Supervisor" />
		<input type="hidden" name="txtExtension" value='<%=Ext_ %>' />
		<input type="hidden" name="txtMonthP" value='<%=MonthP%>' />
		<input type="hidden" name="txtYearP" value='<%=YearP%>' />
		<input type="hidden" name="txtEmpName" value='<%=EmpName_%>' />
		<input type="hidden" name="txtPeriod" value='<%=Period_%>' />
		<input type="hidden" name="txtOffice" value='<%=Office_%>' />
		<input type="hidden" name="txtTotalCost" value='<%=TotalCost_%>' />
		<input type="hidden" name="txtLoginID" value='<%=user1_%>' />
	<td>
</tr>
<%		end if%>
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