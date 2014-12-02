<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<html>
<head>

<script type="text/javascript">
function checkall(obj)
{
	var c = document.frmMonthlyBilling.elements.length
	for (var x=0; x<frmMonthlyBilling.elements.length; x++)
	{
		cbElement = frmMonthlyBilling.elements[x]
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

	if (document.frmMonthlyBilling.cmbSupervisor.value == "" )
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
<TITLE>U.S. Embassy Zagreb - zBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">

</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">Monthly Billing</TD>
   </TR>
<tr>
        <td colspan="3" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
        <td colspan="1" align="right"><FONT color=#330099 size=2><A HREF="javascript:history.go(-1)">Back</A></font></TD>
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
user1_ = user_  'user1_ = right(user_,len(user_)-4)
'user1_ = "pranataw"
'response.write user1_ & "<br>"

strsql = "Select Max(YearP+MonthP) As Period From vwMonthlyBilling Where LoginID='" & user1_ & "'"
'response.write strsql & "<br>"
set rsPeriod = server.createobject("adodb.recordset")
set rsPeriod = BillingCon.execute(strsql)
if not rsPeriod.eof Then
	Period_ = rsPeriod("Period")
end if
'response.write Period_ & "<br>"

If Period_ <> "" Then
	curMonth_ = Right(Period_, 2)
	curYear_ = Left(Period_, 4)
Else
	curMonth_ = month(date())
	curYear_ = year(date())
End If


if len(curMonth_)= 1 then
	curMonth_ = "0" & curMonth_
end if

MonthP = Request("MonthP")
if MonthP ="" then
	MonthP = Request.Form("MonthList")
	if MonthP ="" then
		MonthP = curMonth_
	end if
end if

YearP = Request("YearP")
if YearP ="" then
	YearP = Request.Form("YearList")
	if YearP ="" then
		YearP = curYear_
	end if
end if

'MobilePhoneNumberP = Request("txtPhoneNumber")
MobilePhoneNumberP = Request("CellPhone")
if MobilePhoneNumberP ="" then
	MobilePhoneNumberP = Request.Form("NumberList")
end if

				strsql = "Select Distinct MobilePhone From vwMonthlyBilling Where LoginID='" & user1_ & "'"
				'response.write strsql & "<br>"
				set rsNumber = server.createobject("adodb.recordset")
				set rsNumber = BillingCon.execute(strsql)
				if not rsNumber.eof then
					MobilePhoneNumberFirst_ = rsNumber("MobilePhone")
				Else
					MobilePhoneNumberFirst_ = "No number"
				End If

If MobilePhoneNumberP = "" Then MobilePhoneNumberP = MobilePhoneNumberFirst_



%>
<table cellspadding="1" cellspacing="0" width="65%" border="1" align="center">
<tr align="Center">
	<td colspan="2" align="center">
		<form method="post" name="frmSearch" Action="MonthlyBilling.asp">
		<table  width="100%">
		<tr bgcolor="#000099">
			<td height="25" colspan="5"><strong>&nbsp;<span class="style5">Search &amp; Sort By </span></strong></td>
		</tr>
		<tr>
			<td>&nbsp;Period :</td>
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
					<Option value="09" <%if MonthP ="09" then %>Selected<%End If%> >September</Option>
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

			<td align="Right">Mobile Phone :</td>


			<td>
			<%	if not rsNumber.eof then  %>
				<select name="NumberList">
			<%
				do while not rsNumber.eof
					MobilePhoneNumber_ = rsNumber("MobilePhone")
			%>
					<OPTION value='<%=MobilePhoneNumber_%>'  <%if trim(MobilePhoneNumberP) = trim(MobilePhoneNumber_) then %>Selected<%End If%> >  <%= MobilePhoneNumber_  %></Option>
			<%
					rsNumber.MoveNext
				Loop%>
				</select>&nbsp;
			<% Else %>
				Number not assigned to you for this month
			<% End If %>

			</td>


			<td height="30" align="center">
			<!--	<input type="Button" name="btnBack" value="Back" onClick="Javascript:document.location.href('Default.asp');"> -->
				<input type="Button" name="btnBack" value="Back" onClick="javascript:history.go(-1);">
				<input type="submit" name="Submit" value="Search">
			</td>
		</tr>
		</table>
		</form>
	</td>
</tr>
</table>
<form method="post" action="MonthlyBillingUpdate.asp" name="frmMonthlyBilling" onSubmit="return validate_form()">
<table cellspadding="1" cellspacing="0" width="65%" bgColor="white">
<%
HomePhoneBillRp_ = 0
HomePhoneBillDlr_ = 0
HomePhonePrsBillRp_ = 0
HomePhonePrsBillDlr_ = 0
OfficePhonePrsBillRp_ = 0
OfficePhonePrsBillDlr_ = 0
OfficePhoneBillRp_ = 0
OfficePhoneBillDlr_ = 0
CellPhoneBillRp_ = 0
CellPhoneBillDlr_ = 0
CellPhonePrsBillRp_ = 0
CellPhonePrsBillDlr_ = 0
TotalShuttleBillRp_ = 0
TotalShuttleBillDlr_ = 0
TotalBillingRp_ = 0
TotalBillingDlr_ = 0



'strsql = "Exec spGetMonthlyBill '" & user1_ & "','" & MonthP & "','" & YearP & "'"
strsql = "Select * from vwMonthlyBilling Where LoginID='" & user1_ & "' And MonthP='" & MonthP & "' And YearP='" & YearP & "' And MobilePhone='" & MobilePhoneNumberP & "'"
'response.write strsql & "<br>"
set rsData = server.createobject("adodb.recordset")
set rsData = BillingCon.execute(strsql)
Period_ = MonthP & " - " & YearP
'response.write Period_  & "<br>"
if not rsData.eof then
	EmpID_ = rsData("EmpID")
	EmpName_ = rsData("EmpName")
	Office_ = rsData("Agency") & " - " & rsData("Office")
	Position_ = rsData("WorkingTitle")
	'OfficePhone_ = rsData("WorkPhone")
	'HomePhone_ = rsData("HomePhone")
	MobilePhone_ = rsData("MobilePhone")
	ExchangeRate_ = rsData("ExchangeRate")
	HomePhoneBillRp_ = rsData("HomePhoneBillRp")
	HomePhoneBillDlr_ = rsData("HomePhoneBillDlr")
	HomePhonePrsBillRp_ = rsData("HomePhonePrsBillRp")
	HomePhonePrsBillDlr_ = rsData("HomePhonePrsBillDlr")
	OfficePhonePrsBillRp_ = rsData("OfficePhonePrsBillRp")
	OfficePhonePrsBillDlr_ = rsData("OfficePhonePrsBillDlr")
	OfficePhoneBillRp_ = rsData("OfficePhoneBillRp")
	OfficePhoneBillDlr_ = rsData("OfficePhoneBillDlr")
	CellPhoneBillRp_ = rsData("CellPhoneBillRp")
	CellPhoneBillDlr_ = rsData("CellPhoneBillDlr")
	CellPhonePrsBillRp_ = rsData("CellPhonePrsBillRp")
	CellPhonePrsBillDlr_ = rsData("CellPhonePrsBillDlr")
	TotalShuttleBillRp_ = rsData("TotalShuttleBillRp")
	TotalShuttleBillDlr_ = rsData("TotalShuttleBillDlr")
	TotalBillingRp_ = rsData("TotalBillingRp")
	TotalBillingDlr_ = rsData("TotalBillingDlr")
	ProgressID_ = rsData("ProgressID")
	ProgressStatus_ = rsData("ProgressDesc")
	SupervisorEmail_ = rsData("SupervisorEmail")
	If SupervisorEmail_ = "" Then
		SupervisorEmail_ = rsData("EmailAddress")
	End If
	Notes_ = rsData("Notes")
	SpvRemark_ = rsData("SupervisorRemark")
	TotalBillingPrsAmount_ = rsData("TotalBillingAmountPrsRp")
	TotalBillingAmountPrsDlr_ = rsData("TotalBillingAmountPrsDlr")
'	response.write "TotalBilling: " & TotalBillingPrsAmount_ & "<br>"
'	response.write SupervisorEmail_
%>

<!-- <tr>
	<td colspan="6" align="center"><u>Billing Period (Month - Year) : <a class="FontContent"><%=Period_%></a></u></td>
</tr> -->
<tr>
          <td colspan="6" align="Left"><u><strong>Personal Info :<strong></u></TD>
</tr>
<!-- <tr>
	<td width="20%">Employee Name</td>
	<td width="1%">:</td>
	<td class="FontContent"><%=EmpName_%></td>
	<td>Agency / Office</td>
	<td width="1%">:</td>
	<td class="FontContent"><%=Office_%></td>
</tr> -->
<tr>
	<td colspan="6" align="Left">
	<table cellspadding="1" border="2" bordercolor="black" cellspacing="3" width="100%" bgColor="#999999" border="0">

		<tr BGCOLOR="#999999">
			<td colspan="3" style="border: none;"><FONT color=#FFFFFF><strong>Employee Name : <%=EmpName_%></strong></font></td>
			<td colspan="3" style="border: none;" align="right"><FONT color=#FFFFFF><strong>Phone Number : <%=MobilePhone_ %>&nbsp;</strong></font></td>
		</tr>
		<tr BGCOLOR="#999999">
			<td colspan="6" style="border: none;"><FONT color=#FFFFFF><strong>Position : <%=Position_%></strong></font></td>
		</tr>
		<tr BGCOLOR="#999999">
			<td colspan="6" style="border: none;"><FONT color=#FFFFFF><strong>Agency / Office : <%=Office_%></strong></font></td>
		</tr>
	</table>
	</td>
</tr>
<!-- <tr>
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
	<td>Mobile Phone</td>
	<td width="1%">:</td>
	<td class="FontContent" colspan="4"><%=MobilePhone_%></td>
</tr>
<tr>
	<td colspan="6"><hr></td>
</tr> -->

<tr>
	<td align="Left" colspan="5"><u><strong>Billing detail :<strong></u></TD>
</tr>
<tr>
	<td colspan="6">*Click on the bill for more detail</td>
</tr>
<tr>
	<td align="Left" colspan="6">
	<table cellspadding="1" border="1" bordercolor="black" cellspacing="0" width="100%" bgColor="white" border="0">
	<tr align="center" height=26>
		<td width=20%><strong>Action</strong></td>
		<td width=20%><strong>Billing Period</strong></td>
		<td width=20%><strong>Status</strong></td>
		<td width=20%><strong>Total Bill Amount (Kn.)</strong></td>
		<td width=20%><strong>Personal Amount Due (Kn.)</strong></td>
	</tr>
<!-- <%if cdbl(OfficePhoneBillRp_) > 0 Then %>
	<tr>
		<td><a href="OfficePhoneDetail.asp?Extension=<%=OfficePhone_ %>&MonthP=<%=MonthP%>&YearP=<%=YearP%>" target="_blank">Office Phone</a></td>
		<td width="20%" class="FontContent" align="right"><%=formatnumber(OfficePhoneBillRp_,-1) %>&nbsp;</td>
		<td width="20%" class="FontContent" align="right"><%=formatnumber(OfficePhonePrsBillRp_ ,-1) %>&nbsp;</td>
		<td width="20%" class="FontContent" align="right"><%=formatnumber(OfficePhonePrsBillDlr_,-1) %>&nbsp;</td>
	</tr>
<%else%>
	<tr>
		<td>Office Phone</td>
		<td class="FontContent" align="right">- &nbsp;</td>
		<td class="FontContent" align="right">- &nbsp;</td>
		<td class="FontContent" align="right">- &nbsp;</td>
	</tr>
<%end if%> -->
<!-- <%if cdbl(HomePhoneBillRp_) > 0 Then %>
	<tr>
		<td><a href="HomePhoneDetail.asp?HomePhone=<%=HomePhone_%>&MonthP=<%=MonthP%>&YearP=<%=YearP%>" target="_blank">Home Phone</a></td>
		<td class="FontContent" align="right"><%=formatnumber(HomePhoneBillRp_ ,-1) %>&nbsp;</td>
		<td class="FontContent" align="right"><%=formatnumber(HomePhonePrsBillRp_ ,-1) %>&nbsp;</td>
		<td class="FontContent" align="right"><%=formatnumber(HomePhonePrsBillDlr_ ,-1) %>&nbsp;</td>
	</tr>
<%else%>
	<tr>
		<td>Home Phone</td>
		<td class="FontContent" align="right">- &nbsp;</td>
		<td class="FontContent" align="right">- &nbsp;</td>
		<td class="FontContent" align="right">- &nbsp;</td>
	</tr>
<%end if%> -->
<%if cdbl(CellPhoneBillRp_ ) > 0 Then %>
	<tr height=26>
	<% if ProgressID_ <= 3 Then %>
		<td>&nbsp;<a href="CellPhoneDetail.asp?CellPhone=<%=MobilePhone_ %>&MonthP=<%=MonthP%>&YearP=<%=YearP%>">Tick your calls</a></td>
	<% else %>
		<td>&nbsp;<a href="CellPhoneDetail.asp?CellPhone=<%=MobilePhone_ %>&MonthP=<%=MonthP%>&YearP=<%=YearP%>">View Submitted Bill</a></td>
	<% end if %>
	        <TD align="right">&nbsp;<%= MonthP%>-<%= YearP%></font>&nbsp;</TD>
	        <TD align="right"><%=ProgressStatus_%>&nbsp;</font></TD>
		<td align="right"><%=formatnumber(CellPhoneBillRp_  ,-1) %>&nbsp;</td>
		<td align="right"><%=formatnumber(CellPhonePrsBillRp_ ,-1) %>&nbsp;</td>






	</tr>
<%else%>
	<tr>
		<td>Mobile Phone</td>
		<td class="FontContent" align="right">- &nbsp;</td>
		<td class="FontContent" align="right">- &nbsp;</td>
		<td class="FontContent" align="right">- &nbsp;</td>
	</tr>
<%end if%>
<!-- <%if cdbl(TotalShuttleBillRp_) > 0 Then %>
	<tr>
		<td><a href="ShuttleBusBillDetail.asp?Username=<%=user1_ %>&MonthP=<%=MonthP%>&YearP=<%=YearP%>" target="_blank">Shuttle Bus</a></td>
		<td class="FontContent" align="right"><%=formatnumber(TotalShuttleBillRp_ ,-1) %>&nbsp;</td>
		<td class="FontContent" align="right"><%=formatnumber(TotalShuttleBillRp_ ,-1) %>&nbsp;</td>
		<td class="FontContent" align="right"><%=formatnumber(TotalShuttleBillDlr_,-1) %>&nbsp;</td>
	</tr>
<%else%>
	<tr>
		<td>Shuttle Bus</td>
		<td class="FontContent" align="right">- &nbsp;</td>
		<td class="FontContent" align="right">- &nbsp;</td>
		<td class="FontContent" align="right">- &nbsp;</td>
	</tr>
<%end if%> -->
	</table>
	</TD>
</tr>
<!-- <tr>
	<td colspan="6">
	<table cellspadding="1" cellspacing="0" width="100%" bgColor="white" border="0">
	<tr>
		<td width="200px" align="center"><strong>Total</strong></td>
		<td width="190px" class="FontContent" align="right"><strong><u><%=formatnumber(TotalBillingRp_ , -1) %></u></strong>&nbsp;</td>
		<td width="240px" class="FontContent" align="right"><strong><u><%=formatnumber(TotalBillingPrsAmount_ , -1) %></u></strong>&nbsp;</td>
		<td width="250px" class="FontContent" align="right"><strong><u><%=formatnumber(TotalBillingAmountPrsDlr_ ,-1) %></u></strong>&nbsp;</td>
	</tr>
	</table>
	</td>
</tr> -->
<tr>
	<td colspan="6">
		<table cellspadding="1" cellspacing="0" bgColor="white" width="100%">
		<tr>
			<td width="12%"><strong>Supervisor</strong></td>
			<td width="1%">:</td>
			<td>
				<select name="cmbSupervisor" <% If ((ProgressID_ <> 1) and (ProgressID_ <> 3)) then %>Disabled<%End If%> >
					<option value="">--Select--</option>
			<%
				'strsql = "Select EmailAddress, LastName, FirstName, Office, WorkingTitle From vwPhoneCustomerList Where Type='Amer' and EmailAddress<>'' Order by LastName"
				'strsql = "Select EmailAddress, EmpName, Office, WorkingTitle From vwPhoneCustomerList Where len(EmailAddress)>5 and EmpType<>'Dummy' Order by EmpName"
				'strsql = "Select EmailAddress, EmpName, Office, WorkingTitle From vwPhoneCustomerList Where len(EmailAddress)>5 and EmpType = 'AMER' Order by EmpName"
				strsql = "Select EmailAddress, EmpName, Office, WorkingTitle From vwDirectReport Where len(EmailAddress)>5 and Type = 'AMER' Order by EmpName"
				'response.write strsql & "<br>"
				set rsSPV = server.createobject("adodb.recordset")
				set rsSPV = BillingCon.execute(strsql)
				do while not rsSPV.eof
					Ename_ = rsSPV("EmpName") & "(" & rsSPV("Office") & "-" & rsSPV("WorkingTitle") & ")"
			%>
					<OPTION value=<%=rsSPV("EmailAddress")%> <%if trim(SupervisorEmail_) = trim(rsSPV("EmailAddress")) then %> Selected<%End If%>  >  <%= EName_  %>
			<%
					rsSPV.MoveNext
				Loop%>
				</select>
			</td>
		</tr>
		<tr>
			<td valign="top"><strong>Note</strong></td>
			<td valign="top" width="1%">:</td>
			<td>
				<TextArea name="txtNotes" Rows="5" Cols="70" Wrap <% if (ProgressID_  <> 1) and (ProgressID_ <> 3) then%>ReadOnly<%End If%> ><%=Notes_%></textarea>
			</td>
		</tr>
<%		if (ProgressID_ <> 1)then%>
		<tr>
			<td colspan="3"><strong>Remarks/Correction(s) :</strong></td>
		</tr>
		<tr>
			<td valign="top">&nbsp;</td>
			<td valign="top" width="1%">&nbsp;</td>
			<td>
				<TextArea name="txtRemark" Rows="5" Cols="70" Wrap <% if (ProgressID_  <> 1) or (ProgressID_ <> 3) then%>ReadOnly<%End If%>><%=SpvRemark_ %></textarea>
			</td>
		</tr>
<%		end if%>
		</table>
	</td>
</tr>
<%

		'response.write TotalBillingRp_ & "<br>"
		'response.write TotalBillingPrsAmount_
%>
<tr>
	<td colspan="6" align="center">&nbsp;</td>
</tr>
<%		if (ProgressID_ = 1) or (ProgressID_ = 3) then%>
<tr>
	<td colspan="6" align="center">
<!--
		<input type="submit" name="btnSubmit" Value="Save" />&nbsp;&nbsp;
		<input type="submit" name="btnSubmit" Value="Save & Submit to Supervisor" />
-->
		<input type="submit" name="btnSubmit" Value="Submit to Supervisor" />
		<input type="hidden" name="txtMobilePhone" value='<%=MobilePhone_%>' />
		<input type="hidden" name="txtMonthP" value='<%=MonthP%>' />
		<input type="hidden" name="txtYearP" value='<%=YearP%>' />
		<input type="hidden" name="txtEmpID" value='<%=EmpID_ %>' />
		<input type="hidden" name="txtEmpName" value='<%=EmpName_%>' />
		<input type="hidden" name="txtPosition" value='<%=Position_%>' />
		<input type="hidden" name="txtPeriod" value='<%=Period_%>' />
		<input type="hidden" name="txtOffice" value='<%=Office_%>' />
		<input type="hidden" name="txtTotalCost" value='<%=TotalBillingRp_ %>' />
		<input type="hidden" name="txtTotalBillingPrsAmount" value='<%=TotalBillingPrsAmount_ %>' />
	<td>
</tr>
<%
		end if


Else%>
<tr>
	<td colspan="6" align="center">there is not data.</td>
</tr>
<% end if %>
<tr>
	<td colspan="6"><hr></td>
</tr>
</table>

<%

strsql = "Select * From vwMonthlyBilling Where LoginID ='" & user1_ & "' and (ProgressId='4' or ProgressId='5')"
set AwaitingRS = BillingCon.execute(strsql)

%>
<table cellspadding="1" cellspacing="0" width="65%" bgColor="white">
<tr>
	<td align="Left"><u><strong>Accumulated Debt :</strong></u></TD>
</tr>
<%
if not AwaitingRS.eof Then
%>
<tr>
	<td>*Click on each bill for more detail</td>
</tr>
<tr>
	<td align="Left">
	<table cellspadding="1" border="1" bordercolor="black" cellspacing="0" width="100%" bgColor="white" border="0">
	<tr align="center" height=26>
		<td width=20%><strong>Action</strong></td>
		<td width=20%><strong>Billing Period</strong></td>
		<td width=20%><strong>Status</strong></td>
		<td width=20%><strong>Total Bill Amount (Kn.)</strong></td>
		<td width=20%><strong>Personal Amount Due (Kn.)</strong></td>
	</tr>

<%
   PrevEmpName_ =""
   TotalBill_ = 0
   TotalPrs_ = 0
   do while not AwaitingRS.eof
	   if bg="#dddddd" then bg="ffffff" else bg="#dddddd"
	    if PrevEmpName_ <> AwaitingRS("EmpName") Then
		SubTotalBill_ = 0
		SubTotalPrs_ = 0
%>
		<tr BGCOLOR="#999999" height=26>
			<td colspan="2" style="border: none;"><FONT color=#FFFFFF><strong>&nbsp;Employee Name : <%=AwaitingRS("EmpName") %></strong></font></td>
			<td colspan="3" style="border: none;" align="right"><FONT color=#FFFFFF><strong>Phone Number : <%=AwaitingRS("MobilePhone") %>&nbsp;</strong></font></td>

		</tr>
<%
	    end if
%>

	   <TR bgcolor="<%=bg%>">
	        <TD>&nbsp;<a href="CellPhoneDetail.asp?CellPhone=<%=AwaitingRS("MobilePhone") %>&MonthP=<%= AwaitingRS("MonthP")%>&YearP=<%= AwaitingRS("YearP")%>" target="_blank">View Submitted Bill</a></TD>
	        <TD align="right">&nbsp;<%= AwaitingRS("MonthP")%>-<%= AwaitingRS("YearP")%></font>&nbsp;</TD>
	        <TD align="right"><%= AwaitingRS("ProgressDesc") %>&nbsp;</font></TD>
	        <TD align="right"><%= formatnumber(AwaitingRS("TotalBillingRp"),-1) %>&nbsp;</font></TD>
		<TD align="right">&nbsp;<%= formatnumber(AwaitingRS("TotalBillingAmountPrsRp"),-1) %>&nbsp;</font></TD>
	   </TR>

<%
		SubTotalBill_ = cdbl(SubTotalBill_) + cdbl(AwaitingRS("TotalBillingRp"))
		SubTotalPrs_ = cdbl(SubTotalPrs_) + cdbl(AwaitingRS("TotalBillingAmountPrsRp"))
		TotalBill_ = cdbl(TotalBill_) + cdbl(AwaitingRS("TotalBillingRp"))
		TotalPrs_ = cdbl(TotalPrs_) + cdbl(AwaitingRS("TotalBillingAmountPrsRp"))
		PrevEmpName_ = AwaitingRS("EmpName")
	   AwaitingRS.movenext
		if AwaitingRS.eof Then
%>
		<tr>
			<td colspan="3"><strong>&nbsp;SubTotal</strong></td>
			<td align="right"><strong>&nbsp;<%=formatnumber(SubTotalBill_,-1) %>&nbsp;</font></strong></td>
			<td align="right"><strong>&nbsp;<%=formatnumber(SubTotalPrs_,-1) %>&nbsp;</font></strong></td>
		</tr>
<%

		elseif (PrevEmpName_ <> AwaitingRS("EmpName")) Then

%>
		<tr>
			<td colspan="3"><strong>&nbsp;SubTotal</strong></td>
			<td align="right"><strong>&nbsp;<%=formatnumber(SubTotalBill_,-1) %>&nbsp;</font></strong></td>
			<td align="right"><strong>&nbsp;<%=formatnumber(SubTotalPrs_,-1) %>&nbsp;</font></strong></td>
		</tr>
<%
		end if
   loop
%>

		<tr  BGCOLOR="#999999" height=26>
			<td colspan="3"><FONT color=#FFFFFF><strong>&nbsp;Total</strong></font></td>
			<td align="right"><FONT color=#FFFFFF><strong>&nbsp;<%=formatnumber(TotalBill_,-1) %>&nbsp;</font></strong></td>
			<td align="right"><FONT color=#FFFFFF><strong>&nbsp;<%=formatnumber(TotalPrs_,-1) %>&nbsp;</font></strong></td>
		</tr>

<%
	        strsql = " select CashierMinimumAmount from PaymentDueDate"
       		set rst1 = server.createobject("adodb.recordset")
        	set rst1 = BillingCon.execute(strsql)
	       	if not rst1.eof then
		   CashierMinimumAmount_ = rst1("CashierMinimumAmount")
        	end if
		if cdbl(TotalPrs_) > cdbl(CashierMinimumAmount_) Then
%>
			<tr BGCOLOR="#990000" height=36>
				<td  colspan="5" align="center"><FONT color=#FFFFFF><strong>Your total accumulated debt is greater than <%=formatnumber(CashierMinimumAmount_,-1) %> Kuna. Please make the payment at cashier office.<br><%=CashierInfo%></font></strong></td>
			</tr>
<% 		else %>
			<tr BGCOLOR="#999999" height=26>
				<td colspan="5" align="center"><FONT color=#FFFFFF><strong>Your total accumulated debt is less than <%=formatnumber(CashierMinimumAmount_,-1) %> Kuna. No payment is necessary at this point.</font></strong></td>
			</tr>
<% 		end if %>
	</table>
	</td>
</tr>
<% else %>
<tr>
	<td align="Left">
	<table cellspadding="1" border="1" bordercolor="black" cellspacing="0" width="100%" bgColor="white" border="0">
	<tr align="center" BGCOLOR="#999999" height=26>
		<td><FONT color=#FFFFFF><strong>There is no accumulated debt for your cell phone(s).</strong></font></td>
	</tr>

<% end if %>


</table>



</form>
</BODY>
</html>