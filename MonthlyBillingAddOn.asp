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
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY BGCOLOR="#FFFFFF" TEXT="#000000" LINK="#0080FF" ALINK="0080FF" VLINK="#0055AA" MARGINWIDTH="8" MARGINHEIGHT="0" LEFTMARGIN="8" TOPMARGIN="0">
  <Center>
<% 
 dim user_ 
 dim user1_  
 dim rst 
 dim strsql
 
user_ = request.servervariables("remote_user") 
user1_ = right(user_,len(user_)-4)
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

MonthP = Request("Month")
if MonthP ="" then
	MonthP = Request.Form("MonthList")
	if MonthP ="" then
		MonthP = curMonth_ 
	end if
end if

YearP = Request("Year")
if YearP ="" then
	YearP = Request.Form("YearList")
	if YearP ="" then
		YearP = curYear_ 
	end if
end if

%>  
<table cellspadding="1" cellspacing="0" width="60%" border="1" align="center">
<tr align="Center">
	<td colspan="2" align="center">
		<form method="post" name="frmSearch" Action="MonthlyBillingAddOn.asp">
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
			<td height="30" align="center">
				<input type="Button" name="btnBack" value="Back" onClick="Javascript:document.location.href('DefaultAddOn.asp');">
				<input type="submit" name="Submit" value="Search">
			</td>
		</tr>			
		</table>
		</form>
	</td>
</tr>	
</table>
<form method="post" action="MonthlyBillingUpdateAddOn.asp" name="frmMonthlyBilling" onSubmit="return validate_form()"> 
<table cellspadding="1" cellspacing="0" width="60%" bgColor="white">  
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
strsql = "Select * from vwMonthlyBilling Where LoginID='" & user1_ & "' And MonthP='" & MonthP & "' And YearP='" & YearP & "'"
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
	OfficePhone_ = rsData("WorkPhone")
	HomePhone_ = rsData("HomePhone")
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
	<td class="FontContent">Kn. <%= FormatNumber(ExchangeRate_,0) %> / Dollar</td>

</tr>
<tr>
	<td>Payment Status</td>
	<td width="1%">:</td>
	<td class="FontContent" colspan="4"><%=ProgressStatus_%></td>
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
<%if cdbl(OfficePhoneBillRp_) > 0 Then %>
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
<%end if%>
<%if cdbl(HomePhoneBillRp_) > 0 Then %>
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
<%end if%>
<%if cdbl(CellPhoneBillRp_ ) > 0 Then %>
	<tr>
		<td><a href="CellPhoneDetail.asp?CellPhone=<%=MobilePhone_ %>&MonthP=<%=MonthP%>&YearP=<%=YearP%>" target="_blank">CellPhone</a></td>
		<td class="FontContent" align="right"><%=formatnumber(CellPhoneBillRp_  ,-1) %>&nbsp;</td>
		<td class="FontContent" align="right"><%=formatnumber(CellPhonePrsBillRp_ ,-1) %>&nbsp;</td>
		<td class="FontContent" align="right"><%=formatnumber(CellPhonePrsBillDlr_ ,-1) %>&nbsp;</td>
	</tr>
<%else%>
	<tr>
		<td>Mobile Phone</td>
		<td class="FontContent" align="right">- &nbsp;</td>
		<td class="FontContent" align="right">- &nbsp;</td>
		<td class="FontContent" align="right">- &nbsp;</td>
	</tr>
<%end if%>
<%if cdbl(TotalShuttleBillRp_) > 0 Then %>
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
<%end if%>
	</table>
	</TD>
</tr>
<tr>
	<td colspan="6">
	<table cellspadding="1" cellspacing="0" width="100%" bgColor="white" border="0">
	<tr>
		<td width="200px" align="center"><b>Total</b></td>
		<td width="190px" class="FontContent" align="right"><b><u><%=formatnumber(TotalBillingRp_ , -1) %></u></b>&nbsp;</td>
		<td width="240px" class="FontContent" align="right"><b><u><%=formatnumber(TotalBillingPrsAmount_ , -1) %></u></b>&nbsp;</td>
		<td width="250px" class="FontContent" align="right"><b><u><%=formatnumber(TotalBillingAmountPrsDlr_ ,-1) %></u></b>&nbsp;</td>
	</tr>
	</table>	
	</td>
</tr>
<tr>
	<td colspan="6">
		<table cellspadding="1" cellspacing="0" bgColor="white" width="100%">  
		<tr>
			<td width="12%"><b>Supervisor</b></td>
			<td width="1%">:</td>
			<td>
				<select name="cmbSupervisor" <% If ((ProgressID_ <> 1) and (ProgressID_ <> 3)) then %>Disabled<%End If%> >
					<option value="">--Select--</option>
			<%
'				strsql = "Select EmailAddress, LastName, FirstName, Office, WorkingTitle From vwPhoneCustomerList Where Type='Amer' and EmailAddress<>'' Order by LastName"
				strsql = "Select EmailAddress, EmpName, Office, WorkingTitle From vwPhoneCustomerList Where len(EmailAddress)>5 and type<>'Dummy' Order by EmpName"
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
			<td valign="top"><b>Note</b></td>
			<td valign="top" width="1%">:</td>
			<td>
				<TextArea name="txtNotes" Rows="5" Cols="70" Wrap <% if (ProgressID_  <> 1) and (ProgressID_ <> 3) then%>ReadOnly<%End If%> ><%=Notes_%></textarea>
			</td>
		</tr>
<%		if (ProgressID_ <> 1)then%>
		<tr>
			<td colspan="3"><b>Remarks/Correction(s) :</b></td>
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
		<input type="hidden" name="txtExtension" value='<%=OfficePhone_%>' />
		<input type="hidden" name="txtMonthP" value='<%=MonthP%>' />
		<input type="hidden" name="txtYearP" value='<%=YearP%>' />
		<input type="hidden" name="txtEmpID" value='<%=EmpID_ %>' />
		<input type="hidden" name="txtEmpName" value='<%=EmpName_%>' />
		<input type="hidden" name="txtPeriod" value='<%=Period_%>' />
		<input type="hidden" name="txtOffice" value='<%=Office_%>' />
		<input type="hidden" name="txtTotalCost" value='<%=TotalBillingRp_ %>' />
		<input type="hidden" name="txtTotalBillingPrsAmount" value='<%=TotalBillingPrsAmount_ %>' />
	<td>
</tr>
<%
		end if%>
<%Else%>
<tr>
	<td colspan="6" align="center">there is not data.</td>	
</tr>
<% end if %>
</table>
</form>
</BODY>
</html>