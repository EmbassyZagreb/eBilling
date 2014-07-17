<%@ Language=VBScript%>
<% ' VI 6.0 Scripting Object Model Enabled %>
 
 
<html>
<head>

<META name=VI60_defaultClientScript content=VBScript>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<!--#include file="connect.inc" -->


<script type="text/javascript">

function ClearFilter()
{
	document.forms['frmSearch'].elements['cmbAgency'].value="X";
	document.forms['frmSearch'].elements['cmbSection'].value="X";
	document.forms['frmSearch'].elements['cmbEmp'].value="X";
	document.forms['frmSearch'].elements['cmbStatus'].value=0;
}

function ValidateForm()
{
	valid = true;
	nRec = 0;
	for (var x=0; x<frmBillingList.elements.length; x++)
	{	
		cbElement = frmBillingList.elements[x]
		if ((cbElement.checked) && (cbElement.name=="cbApproval"))
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

function checkall(obj)
{
	for (var x=0; x<frmBillingList.elements.length; x++)
	{
		cbElement = frmBillingList.elements[x]
		if (cbElement.type == "checkbox")
		{
			cbElement.checked= obj.checked?true:false
		}
	}
}
</script>
<%
Dim user_ , user1_

user_ = request.servervariables("remote_user")
user1_ = right(user_,len(user_)-4)
'response.write user1_ & "<br>"

strsql = "select RoleID from Users where loginId ='" & user1_ & "'"
set UserRS = server.createobject("adodb.recordset")
'response.write strsql & "<br>"
set UserRS = BillingCon.execute(strsql)
if not UserRS.eof then
	UserRole_ = UserRS("RoleID")
Else
	UserRole_ = ""
end if


	curMonth_ = month(date())
	curYear_ = year(date())
	if len(curMonth_)= 1 then
		curMonth_ = "0" & curMonth_
	end if

	sMonthP = request("sMonthP")

	if sMonthP = "" Then sMonthP = Request.Form("sMonthList")
	if sMonthP = "" then
		sMonthP = curMonth_ 
	end if
	'response.write sMonthP

	sYearP = request("sYearP")
	if sYearP ="" Then sYearP = Request.Form("sYearList")
	if sYearP ="" then
		sYearP = curYear_ 
	end if

	eMonthP = request("eMonthP")
	if eMonthP = "" Then eMonthP = Request.Form("eMonthList")
	if eMonthP = "" then
		eMonthP = curMonth_ 
	end if

	eYearP = request("eYearP")
	if eYearP = "" Then eYearP = Request.Form("eYearList")
	if eYearP = "" then
		eYearP = curYear_ 
	end if

	Section_ = request("Section")
	if Section_ = "" Then Section_ = Request.Form("cmbSection")
	if Section_ = "" then
		Section_ = "X"
	end if

	EmpID_ = request("EmpID")
	if EmpID_ = "" Then EmpID_ = Request.Form("cmbEmp")
	if EmpID_ = "" then
		EmpID_ = "X"
	end if

	Status_ = request("Status")
	if Status_ = "" Then Status_ = Request.Form("cmbStatus")
	if Status_ = "" then
		Status_ = 0
	end if

%>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">BILLING SETTLEMENT</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>

<%

dim rs 
dim strsql
dim tombol
dim hlm
If (UserRole_ = "Admin") or (UserRole_ = "Voucher") or (UserRole_ = "FMC") or (UserRole_ = "Cashier")  Then
%>
<form method="post" name="frmSearch" Action="BillingSettlement.asp">
<table cellspadding="1" cellspacing="0" width="60%" border="1" align="center">
<tr align="Center">
	<td colspan="2" align="center">
		<table  width="100%">
		<tr bgcolor="#000099">
			<td height="25" colspan="6"><strong>&nbsp;<span class="style5">Criteria(s): </span></strong></td>
		</tr>
<tr>
			<td width="15%">&nbsp;Period&nbsp;</td>				
			<td>:</td>
			<td colspan="4">
				<Select name="sMonthList">
					<Option value="01" <%if sMonthP ="01" then %>Selected<%End If%>>January</Option>
					<Option value="02" <%if sMonthP ="02" then %>Selected<%End If%>>February</Option>
					<Option value="03" <%if sMonthP ="03" then %>Selected<%End If%>>March</Option>
					<Option value="04" <%if sMonthP ="04" then %>Selected<%End If%>>April</Option>
					<Option value="05" <%if sMonthP ="05" then %>Selected<%End If%>>May</Option>
					<Option value="06" <%if sMonthP ="06" then %>Selected<%End If%>>June</Option>
					<Option value="07" <%if sMonthP ="07" then %>Selected<%End If%>>July</Option>
					<Option value="08" <%if sMonthP ="08" then %>Selected<%End If%>>August</Option>
					<Option value="09" <%if sMonthP ="09" then %>Selected<%End If%>>September</Option>
					<Option value="10" <%if sMonthP ="10" then %>Selected<%End If%>>October</Option>
					<Option value="11" <%if sMonthP ="11" then %>Selected<%End If%>>November</Option>
					<Option value="12" <%if sMonthP ="12" then %>Selected<%End If%>>December</Option>
				</Select>&nbsp;
<%
				Year_ = Year(Date()) - 1
%>

				<Select name="sYearList">
<% 				Do While Year_ <= Year(Date()) %>
				<Option value='<%=Year_%>' <%if trim(Year_) = trim(sYearP) then %>Selected<%End If%> ><%=Year_%></Option>
<% 
			Year_ = Year_ + 1
			Loop %>	
				</Select>&nbsp;to&nbsp;
				<Select name="eMonthList">
					<Option value="01" <%if eMonthP ="01" then %>Selected<%End If%>>January</Option>
					<Option value="02" <%if eMonthP ="02" then %>Selected<%End If%>>February</Option>
					<Option value="03" <%if eMonthP ="03" then %>Selected<%End If%>>March</Option>
					<Option value="04" <%if eMonthP ="04" then %>Selected<%End If%>>April</Option>
					<Option value="05" <%if eMonthP ="05" then %>Selected<%End If%>>May</Option>
					<Option value="06" <%if eMonthP ="06" then %>Selected<%End If%>>June</Option>
					<Option value="07" <%if eMonthP ="07" then %>Selected<%End If%>>July</Option>
					<Option value="08" <%if eMonthP ="08" then %>Selected<%End If%>>August</Option>
					<Option value="09" <%if eMonthP ="09" then %>Selected<%End If%>>September</Option>
					<Option value="10" <%if eMonthP ="10" then %>Selected<%End If%>>October</Option>
					<Option value="11" <%if eMonthP ="11" then %>Selected<%End If%>>November</Option>
					<Option value="12" <%if eMonthP ="12" then %>Selected<%End If%>>December</Option>
				</Select>&nbsp;
<%
				Year_ = Year(Date()) - 1
%>

				<Select name="eYearList">
<% 				Do While Year_ <= Year(Date()) %>
				<Option value='<%=Year_%>' <%if trim(Year_) = trim(eYearP) then %>Selected<%End If%> ><%=Year_%></Option>
<% 
			Year_ = Year_ + 1
			Loop %>	
				</Select>
			</td>
		</tr>
		<tr>
			<td>&nbsp;Section&nbsp;</td>
			<td>:</td>
			<td>
<%
 				strsql ="select distinct Office from vwPhoneCustomerList Where Office<>'' order by Office"
				set SectionRS = server.createobject("adodb.recordset")
				set SectionRS = BillingCon.execute(strsql)
'				response.write strStr 
%>	
				<Select name="cmbSection">
					<Option value='X'>--All--</Option>
<%				Do While not SectionRS.eof %>
					<Option value='<%=SectionRS("Office")%>' <%if trim(Office_) = trim(SectionRS("Office")) then %>Selected<%End If%> ><%=SectionRS("Office")%></Option>
					
<%					SectionRS.MoveNext
				Loop%>
				</select>

			</td>
			<td>&nbsp;Status&nbsp;</td>
			<td>:</td>
			<td>
<!--
				<Select name="cmbStatus">
					<Option value="X" <%if Status ="X" then %>Selected<%End If%>>-All-</Option>
					<Option value="Pending" <%if Status = "Pending" then %>Selected<%End If%>>Pending</Option>
					<Option value="Completed" <%if Status = "Completed" then %>Selected<%End If%>>Completed</Option>
				</Select>&nbsp;
-->
<%
 				strsql ="select ProgressID, ProgressDesc from ProgressStatus Where ProgressID <10 Order By OrderNo"
				set StatusRS = server.createobject("adodb.recordset")
				set StatusRS = BillingCon.execute(strsql)
'				response.write strStr 
%>	
				<Select name="cmbStatus">
					<Option value='0'>--All--</Option>
<%				Do While not StatusRS.eof %>
					<Option value='<%=StatusRS("ProgressID")%>' <%if trim(Status_) = trim(StatusRS("ProgressID")) then %>Selected<%End If%> ><%=StatusRS("ProgressDesc")%></Option>
					
<%					StatusRS.MoveNext
				Loop%>
				</select>
			</td>
		</tr>
		<tr>
			<td>&nbsp;Employee&nbsp;</td>
			<td>:</td>
			<td colspan="2">
<%
 				strsql ="select EmpID, EmpName from vwPhoneCustomerList order by EmpName"
				set EmpRS = server.createobject("adodb.recordset")
				set EmpRS = BillingCon.execute(strsql)
'				response.write strStr 
%>	
				<Select name="cmbEmp">
					<Option value='X'>--All--</Option>
<%				Do While not EmpRS.eof 
%>
					<Option value='<%=EmpRS("EmpID")%>' <%if trim(EmpID_) = trim(EmpRS("EmpID")) then %>Selected<%End If%> ><%=EmpRS("EmpName") %></Option>
					
<%					EmpRS.MoveNext
				Loop%>
				</select>

			</td>	
		</tr>
		<tr>
			<td colspan="2">&nbsp;</td>
			<td align="left">
				<input type="Button" name="btnReset" value="Reset" onClick="Javascript:ClearFilter();">	
			</td>
			<td align="Left" colspan="3">
				<input type="submit" name="Submit" value="Search">
			</td>
		</tr>
		</table>
	</td>
</tr>	
</table>
</form>
<%
sPeriod = sYearP&sMonthP
ePeriod = eYearP&eMonthP

'response.write sPeriod & ePeriod 
strsql = "Select * From vwMonthlyBilling Where YearP+MonthP>='" & sPeriod & "' and YearP+MonthP<='" & ePeriod & "'"
strFilter=""
If Section_ <> "X" then
	strFilter=strFilter & " and Office='" & Section_ & "'"
End If

If EmpID_ <> "X" then
	strFilter =strFilter & " and EmpID='" & EmpID_ & "'"
End If

If Status_ <> "0" then
	strFilter =strFilter & " and ProgressID=" & Status_
End If


strsql = strsql  & strFilter
'response.write strsql & "<br>"

set DataRS = server.createobject("adodb.recordset")
DataRS.CursorLocation = 3
DataRS.Open strsql,BillingCon
'set DataRS=BillingCon.execute(strsql)

dim intPageSize,PageIndex,TotalPages 
dim RecordCount,RecordNumber,Count 
intpageSize=20 
PageIndex=request("PageIndex")

if PageIndex ="" then PageIndex=1 

if not DataRS.eof then
	RecordCount = DataRS.RecordCount   
	'response.write RecordCount & "<br>"
	RecordNumber=(intPageSize * PageIndex) - intPageSize 
	'response.write RecordNumber
	DataRS.PageSize =intPageSize 
	DataRS.AbsolutePage = PageIndex
	TotalPages=DataRS.PageCount 
	'response.write TotalPages & "<br>"
End If
'response.write strsql

dim intPrev,intNext 	
intPrev=PageIndex - 1 
intNext=PageIndex +1 


if not DataRS.eof Then

   dim no_  
   no_ = 1 + ((PageIndex*intPageSize)-intPageSize) 
   Count=1 
   do while not DataRS.eof  and Count<=intPageSize
	Period_ = DataRS("MonthP") & "-" & DataRS("YearP")
	MonthP_ = DataRS("MonthP")
	YearP_ = DataRS("YearP")
	EmpName_ = DataRS("EmpName")
	Office_ = DataRS("Agency") & " - " & DataRS("Office")
	Position_ = DataRS("WorkingTitle")
	OfficePhone_ = DataRS("WorkPhone")
	HomePhone_ = DataRS("HomePhone")
	MobilePhone_ = DataRS("MobilePhone")
	ExchangeRate_ = DataRS("ExchangeRate")
	LoginID_ = DataRS("LoginID")
	HomePhoneBillRp_ = DataRS("HomePhoneBillRp")
	HomePhoneBillDlr_ = DataRS("HomePhoneBillDlr")
	HomePhonePrsBillRp_ = DataRS("HomePhonePrsBillRp")
	HomePhonePrsBillDlr_ = DataRS("HomePhonePrsBillDlr")
	OfficePhonePrsBillRp_ = DataRS("OfficePhonePrsBillRp")
	OfficePhonePrsBillDlr_ = DataRS("OfficePhonePrsBillDlr")
	OfficePhoneBillRp_ = DataRS("OfficePhoneBillRp")
	OfficePhoneBillDlr_ = DataRS("OfficePhoneBillDlr")
	CellPhoneBillRp_ = DataRS("CellPhoneBillRp")
	CellPhoneBillDlr_ = DataRS("CellPhoneBillDlr")
	CellPhonePrsBillRp_ = DataRS("CellPhonePrsBillRp")
	CellPhonePrsBillDlr_ = DataRS("CellPhonePrsBillDlr")

	TotalShuttleBillRp_ = DataRS("TotalShuttleBillRp")
	TotalShuttleBillDlr_ = DataRS("TotalShuttleBillDlr")
	TotalBillingRp_ = DataRS("TotalBillingRp")
	TotalBillingDlr_ = DataRS("TotalBillingDlr")
	ProgressStatus_ = DataRS("ProgressDesc")
	TotalBillingPrsAmount_ = DataRS("TotalBillingAmountPrsRp")
	TotalBillingAmountPrsDlr_ = DataRS("TotalBillingAmountPrsDlr")
%>
	<table cellspadding="1" cellspacing="0" width="70%" bgColor="white">  
      	<tr>
		<td colspan="6" align="center"><u>Billing Period (Month - Year) : <a class="FontContent"><%=Period_%></a></u></td>
	</tr>
	<tr>
        	  <td align="Left"><u><b>Personal Info<b></u></TD>
	</tr>  
	<tr>
		<td width="20%">Employee Name</td>
		<td width="1%">:</td>
		<td class="FontContent" width="40%"><%=EmpName_%></td>
		<td width="20%">Agency / Office</td>
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
		<td>Exchange Rate</td>
		<td width="1%">:</td>
		<td class="FontContent">Kn. <%= FormatNumber(ExchangeRate_,-1) %> / Dollar</td>	
	</tr>
	<tr>
		<td>Mobile Phone</td>
		<td width="1%">:</td>
		<td class="FontContent"><%=MobilePhone_ %></td>
		<td>Total Billing</td>
		<td width="1%">:</td>
		<td class="FontContent">Kn. <%= FormatNumber(TotalBillingRp_ ,-1) %></td>	
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
		<td align="Left" colspan="6"><u><b>Billing Summary :<b></u></TD>
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
		<td><a href="OfficePhoneDetail.asp?Extension=<%=OfficePhone_ %>&MonthP=<%=DataRS("MonthP")%>&YearP=<%=DataRS("YearP")%>" target="_blank">Office Phone</a></td>
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
		<td><a href="HomePhoneDetail.asp?HomePhone=<%=HomePhone_%>&MonthP=<%=DataRS("MonthP")%>&YearP=<%=DataRS("YearP")%>" target="_blank">Home Phone</a></td>
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
		<td><a href="CellPhoneDetail.asp?CellPhone=<%=MobilePhone_ %>&MonthP=<%=DataRS("MonthP")%>&YearP=<%=DataRS("YearP")%>" target="_blank">CellPhone</a></td>
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
		<td><a href="ShuttleBusBillDetail.asp?Username=<%=user1_ %>&MonthP=<%=DataRS("MonthP")%>&YearP=<%=DataRS("YearP")%>" target="_blank">Shuttle Bus</a></td>
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
			<td align="center"><b>Total</b></td>
			<td width="240px" class="FontContent" align="right"><b><u><%=formatnumber(TotalBillingRp_ , -1) %></u></b>&nbsp;</td>
			<td width="240px" class="FontContent" align="right"><b><u><%=formatnumber(TotalBillingPrsAmount_ , -1) %></u></b>&nbsp;</td>
			<td width="240px" class="FontContent" align="right"><b><u><%=formatnumber(TotalBillingAmountPrsDlr_ ,-1) %></u></b>&nbsp;</td>
		</tr>
		</table>	
		</td>
	</tr>
	<tr>
		<td align="Left" colspan="5"><u><b>Billing detail :<b></u></TD>
	</tr>	
	
<%
strsql = "Exec spGetBilling 'Detail','" & OfficePhone_ & "','" & MonthP_ & "','" & YearP_ & "'"
'response.write strsql & "<br>"
set rsOfficePhone = BillingCon.execute(strsql) 
no_ = 1 
if not rsOfficePhone.eof then
%>
	<tr>
		<td colspan="6">
		<table cellspadding="1" cellspacing="1" width="95%" align="center">
		<tr>
			<td align="Left" colspan="5"><u><b>Office Phone detail :<b></u></TD>
		</tr>
		<tr align="center" cellpadding="0" cellspacing="0" >
			<TD width="5%"><strong>No.</strong></TD>
		       	<TD width="30%"><strong>Dialed Date/time</strong></TD>
			<TD><strong>Dialed Number</strong></TD>
			<TD width="15%"><strong>Amount (Kn.)</strong></TD>
			<TD width="15%"><strong>Personal used</strong></TD>
		</tr>
<%

		do while not rsOfficePhone.eof
   			if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
%>
		<tr bgcolor="<%=bg%>">
			<td align="right"><%=No_%>&nbsp;</td>
		        <td><FONT color=#330099 size=2>&nbsp;<%=rsOfficePhone("DialedDatetime")%></font></td> 
		       	<td><FONT color=#330099 size=2>&nbsp;<%=rsOfficePhone("DialedNumber")%></font></td> 
		        <td align="right"><FONT color=#330099 size=2><%=formatnumber(rsOfficePhone("Cost"),-1)%>&nbsp;</font></td> 
<%			if (ProgressID_ = 1) or (ProgressID_ = 3) then %>
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
		</tr>
		<tr>
			<td align="right" colspan="3"><b>Sub Total (Kn.) </b>&nbsp;</td>
			<td width="15%" class="FontContent" align="right"><b><%=formatnumber(OfficePhoneBillRp_,-1)%></b>&nbsp;</td>
			<td width="15%" class="FontContent" align="right"><b><u><%=formatnumber(OfficePhonePrsBillRp_ ,-1)%></u></b>&nbsp;</td>
		</tr>
		</table>
		</td>
	</tr>
	<%end if%>
<%
strsql = "Exec spGetHomephone '2','" & HomePhone_ & "','" & MonthP_ & "','" & YearP_ & "'"
'response.write strsql & "<br>"
set rsHomePhone = BillingCon.execute(strsql) 
if not rsHomePhone.eof then
%>
	<tr>
	<td colspan="6">
		<table cellspadding="1" cellspacing="1" width="95%" align="center">
		<tr>
			<td align="Left" colspan="5"><u><b>Home Phone detail :<b></u></TD>
		</tr>
		<tr align="center" cellpadding="0" cellspacing="0" >
			<TD width="5%"><strong>No.</strong></TD>
		       	<TD width="30%"><strong>Dialed Date/time</strong></TD>
			<TD><strong>Dialed Number</strong></TD>
			<TD width="15%"><strong>Amount (Kn.)</strong></TD>
			<TD width="15%"><strong>Personal used</strong></TD>
		</tr>

<%
		no_ = 1 
		do while not rsHomePhone.eof
   			if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
%>
		<tr bgcolor="<%=bg%>">
			<td align="right"><%=No_%>&nbsp;</td>
		        <td><FONT color=#330099 size=2>&nbsp;<%=rsHomePhone("DialedDatetime")%></font></td> 
		       	<td><FONT color=#330099 size=2>&nbsp;<%=rsHomePhone("DialedNumber")%></font></td> 
		        <td align="right"><FONT color=#330099 size=2><%=formatnumber(rsHomePhone("Cost"),-1)%>&nbsp;</font></td> 
			<td align="center">
				<Input type="Checkbox" name="cbPersonal" <%if rsHomePhone("isPersonal") = "Y" then%> Checked<%end if%> Disabled>
			</td>
		</tr>
		<%   
			rsHomePhone.movenext
			no_ = no_ + 1
		loop
		%>
		<tr>
			<td align="right" colspan="3"><b>Sub Total (Kn.) </b>&nbsp;</td>
			<td width="15%" class="FontContent" align="right"><b><%=formatnumber(HomePhoneBillRp_ ,-1)%></b>&nbsp;</td>
			<td width="15%" class="FontContent" align="right"><b><u><%=formatnumber(HomePhonePrsBillRp_ ,-1)%></u></b>&nbsp;</td>
		</tr>
		</table>
	</td>
  	</tr>
<%End if%>

<%
strsql = "Select * from CellPhoneDt Where PhoneNumber='" & MobilePhone_ & "' and MonthP='" & MonthP_ & "' and YearP='" & YearP_ & "' order by DialedDatetime"
'response.write strsql & "<br>"
set rsCellPhone = BillingCon.execute(strsql) 
if not rsCellPhone.eof then
%>
	<tr>
	<td colspan="6">
		<table cellspadding="1" cellspacing="1" width="95%" align="center">
		<tr>
			<td align="Left" colspan="5"><u><b>Cell Phone detail :<b></u></TD>
		</tr>
		<tr align="center" cellpadding="0" cellspacing="0" >
			<TD width="5%"><strong>No.</strong></TD>
		       	<TD width="30%"><strong>Dialed Date/time</strong></TD>
			<TD><strong>Dialed Number</strong></TD>
			<TD width="15%"><strong>Amount (Kn.)</strong></TD>
			<TD width="15%"><strong>Personal used</strong></TD>
		</tr>

<%
		no_ = 1 
		do while not rsCellPhone.eof
   			if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
%>
		<tr bgcolor="<%=bg%>">
			<td align="right"><%=No_%>&nbsp;</td>
		        <td><FONT color=#330099 size=2>&nbsp;<%=rsCellPhone("DialedDatetime")%></font></td> 
		       	<td><FONT color=#330099 size=2>&nbsp;<%=rsCellPhone("DialedNumber")%></font></td> 
		        <td align="right"><FONT color=#330099 size=2><%=formatnumber(rsCellPhone("Cost"),-1)%>&nbsp;</font></td> 
		        <td align="center"><Input type="Checkbox" name="cbPersonal" <%if rsCellPhone("isPersonal") = "Y" then%> Checked<%end if%> disabled></td>
		<%   
			rsCellPhone.movenext
			no_ = no_ + 1
		loop
		%>
		</tr>
		<tr>
			<td align="right" colspan="3"><b>Sub Total (Kn.) </b>&nbsp;</td>
			<td width="15%" class="FontContent" align="right"><b><%=formatnumber(CellPhoneBillRp_ ,-1)%></b>&nbsp;</td>
			<td width="15%" class="FontContent" align="right"><b><u><%=formatnumber(CellPhonePrsBillRp_ ,-1)%></u></b>&nbsp;</td>
		</tr>
		</table>
	</td>
  	</tr>
<%End if%>
<%
strsql = "Exec spGetShuttleBusList '" & LoginID_ & "','" & MonthP_ & "','" & YearP_ & "'"
'response.write strsql & "<br>"
set rsShuttle = server.createobject("adodb.recordset") 
set rsShuttle = BillingCon.execute(strsql) 
If not rsShuttle.eof then
%>
	<tr>
		<td colspan="6">
		<table cellspadding="1" cellspacing="1" width="95%" align="center">
		<tr>
			<td align="Left" colspan="8"><u><b>Shuttle Bus detail :<b></u></TD>
		</tr>
		<TR align="center" cellpadding="0" cellspacing="0" >
			<TD width="6%"><strong>No.</strong></TD>
			<TD><strong>Date</strong></TD>
			<TD width="10%"><strong>AM</strong></TD>
			<TD width="10%"><strong>PM</strong></TD>
			<TD width="20%"><strong>Tot. Shuttle Qty</strong></TD>
			<TD width="20%"><strong>Tot. Shuttle Bill($)</strong></TD>
		</TR>    
<% 
		dim TotalQty_ , TotalAmount_ 
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
		TotalAmount_ = CDbl(TotalAmount_) + formatnumber(rsShuttle("TotalAmountPerDay"),-1)
		'response.write rsShuttle("TotalAmountPerDay")
		'response.write TotalAmount_ 
 		rsShuttle.movenext
   		no_ = no_ + 1
		loop
%>	
		<tr>
			<td align="right" colspan="4"><b>Sub Total&nbsp;</b></td>
			<td width="10%" class="FontContent" align="right"><b><%=formatnumber(TotalQty_ ,-1)%></b></td>
			<td width="10%" class="FontContent" align="right"><b>$&nbsp;&nbsp;&nbsp;<%=formatnumber(TotalAmount_ ,-1)%></b></td>
		</tr>
		<tr>
			<td align="right" colspan="5"><b>&nbsp;</b></td>
			<td width="10%" class="FontContent" align="right"><b>Kn.&nbsp;&nbsp;&nbsp;<u><%=formatnumber(TotalShuttleBillRp_ ,-1)%></u></b></td>
		</tr>
		</table>
		</td>
	</tr>
<%End if%>
	<tr>
		<td colspan="6" align="center">~End of Settlement~</td>
	</tr>
	<tr>
		<td colspan="6" align="center"><br><br></td>
	</tr>
	</table>
<%   
		Count=Count +1
	   DataRS.movenext
	   no_ = no_ + 1 
   loop 
	PageNo=1
%>
</table>


<table width="60%">
	<tr>
		<td align="right">
<%
		Do while PageNo<=TotalPages 
			if trim(pageNo) = trim(PageIndex) Then
%>		
				<label class="ActivePage"><%=PageNo%></label>&nbsp;
			<%Else%>
				<a href="BillingSettlement.asp?PageIndex=<%=PageNo%>&sMonthP=<%=sMonthP%>&sYearP=<%=sYearP%>&eMonthP=<%=eMonthP%>&eYearP=<%=eYearP%>&Section=<%=Section_%>&EmpID=<%=EmpID_%>&Status=<%=Status%>"><%=PageNo%></a>&nbsp;
<%	
			End If						
			PageNo=PageNo+1
		Loop
%>
		</td>
	</tr>
</table>
<%
else 
%>
	<table cellspadding="1" cellspacing="0" width="100%">  
	<tr>
        	<td><br></TD>
	</tr>
	<tr>
		<td align="center">not data.</td>
	</tr>
	<tr>
        	<td><br></TD>
	</tr>
	<tr>
		<td align="center"><a href="Default.asp"><img src="images/Back.gif" border="0" alt="Go..Back" WIDTH="83" HEIGHT="25"></a></td>
	</tr>	
	</table>
<% end if %>
<%Else%>
	<table>
		<tr>
			<td>You do not have permission to access this site.</td>
		</tr>
		<tr>
			<td>Please <a href="http://zagrebws03.eur.state.sbu/WebPASS/eservices/MainPage.asp">Submit Request </a> or contact Zagreb ISC Helpdesk at ext.3333.</td>
		</tr>
	</table>
<% end if %>
</body> 

</html>


