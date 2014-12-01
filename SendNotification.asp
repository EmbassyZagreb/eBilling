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
	//document.forms['frmSearch'].elements['cmbBillType'].value="C";
	document.forms['frmSearch'].elements['cmbSection'].value="X";
	document.forms['frmSearch'].elements['cmbSectionGroup'].value="X";
	document.forms['frmSearch'].elements['cmbEmp'].value="X";
	document.forms['frmSearch'].elements['cmbStatus'].value="0";
	document.forms['frmSearch'].elements['txtEmailAddress'].value="";
	document.forms['frmSearch'].elements['cmbSentStatus'].value="0";
}

function ValidateForm()
{
	valid = true;
	nRec = 0;
	for (var x=0; x<frmARReminder.elements.length; x++)
	{	
		cbElement = frmARReminder.elements[x]
		if ((cbElement.checked) && (cbElement.name=="cbApproval"))
		{
			nRec++;
		}
	}
	if (nRec == 0)
	{
		alert("Please select data that you want to send !!!");
		valid = false;
	}
	return valid;
}

function checkall(obj)
{
	for (var x=0; x<frmARReminder.elements.length; x++)
	{
		cbElement = frmARReminder.elements[x]
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
user1_ = user_  'user1_ = right(user_,len(user_)-4)
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


strsql = "Select Max(YearP+MonthP) As Period From vwMonthlyBilling"
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




'curMonth_ = month(date())
'curYear_ = year(date())
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

Agency_ = request("Agency")
if Agency_ = "" Then Agency_ = Request.Form("cmbAgency")
if Agency_ = "" then
	Agency_ = "X"
end if

Section_ = request("Section")
if Section_ = "" Then Section_ = Request.Form("cmbSection")
if Section_ = "" then
	Section_ = "X"
end if

SectionGroup_ = request("SectionGroup")
if SectionGroup_ = "" Then SectionGroup_ = Request.Form("cmbSectionGroup")
if SectionGroup_ = "" then
	SectionGroup_ = "X"
end if

EmpID_ = request("EmpID")
if EmpID_ = "" Then EmpID_ = Request.Form("cmbEmp")
if EmpID_ = "" then
	EmpID_ = "X"
end if

EmailAddress_ = request("EmailAddress")
if EmailAddress_ = "" Then EmailAddress_ = Request.Form("txtEmailAddress")


BillType = request("BillType")
if BillType = "" Then BillType = Request.Form("cmbBillType")
if BillType = "" then
	BillType = "C"
end if

Status_ = request("Status")
if Status_ = "" Then Status_ = Request.Form("cmbStatus")
if Status_ = "" then
	Status_ = 0
end if

SentStatus_ = request("SentStatus")
if SentStatus_ = "" Then SentStatus_ = Request.Form("cmbSentStatus")
if SentStatus_ = "" then
	SentStatus_ = 1
end if

SortBy_ = Request.Form("SortList")
'response.write "SortBy" & SortBy_
if (SortBy_ ="") then
	if Request("SortBy")<>"" then
		SortBy_ = Request("SortBy")	
	Else
		SortBy_ = "EmpName"
	end if
end if

Order_ = Request.Form("OrderList")
if (Order_ ="") then
	if Request("Order")<>"" then
		Order_ = Request("OrderList")	
	Else
		Order_ = "Asc"
	end if
end if


NoRecord_ = request("NoRecord")
if NoRecord_ = "" Then NoRecord_ = Request.Form("cmbNoRecord")
if NoRecord_ = "" then
	NoRecord_ = 1
end if

%>
<TITLE>U.S. Embassy Zagreb - zBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">BILLING NOTIFICATION</TD>
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
If (UserRole_ = "Admin") or (UserRole_ = "Voucher") or (UserRole_ = "FMC") Then
%>
<form method="post" name="frmSearch" Action="SendNotification.asp">
<table cellspadding="1" cellspacing="0" width="70%" border="1" align="center">
<tr align="Center">
	<td colspan="2" align="center">
		<table  width="100%">
		<tr bgcolor="#000099">
			<td height="25" colspan=6"><strong>&nbsp;<span class="style5">Criteria(s): </span></strong></td>
		</tr>
		<tr>
			<td width="15%" align="right">&nbsp;Period&nbsp;</td>				
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
				Year_ = Year(Date()) - 2
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
				Year_ = Year(Date()) - 2
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
			<td align="right">&nbsp;Agency&nbsp;</td>
			<td>:</td>
			<td>
<%
 				strsql ="select distinct Agency from vwPhoneCustomerList Where Agency<>'' order by Agency"
				set AgencyRS = server.createobject("adodb.recordset")
				set AgencyRS = BillingCon.execute(strsql)
'				response.write strStr 
%>	
				<Select name="cmbAgency">
					<Option value='X'>--All--</Option>
<%				Do While not AgencyRS.eof %>
					<Option value='<%=AgencyRS("Agency")%>' <%if trim(Agency_) = trim(AgencyRS("Agency")) then %>Selected<%End If%> ><%=AgencyRS("Agency")%></Option>
					
<%					AgencyRS.MoveNext
				Loop%>
				</select>

			</td>
<!--			<td align="right">Bill Type&nbsp;</td>
			<td>:</td>
			<td>
				<Select name="cmbBillType">
					<Option value="X" <%if BillType ="X" then %>Selected<%End If%>>-All-</Option>
					<Option value="O" <%if BillType = "O" then %>Selected<%End If%>>Office Phone</Option>
					<Option value="C" <%if BillType = "C" then %>Selected<%End If%>>Cell Phone</Option>
				<Option value="H" <%if BillType = "H" then %>Selected<%End If%>>Home Phone</Option>
					<Option value="S" <%if BillType = "S" then %>Selected<%End If%>>Shuttle Bus</Option>
				</Select>&nbsp;
			</td> -->
		</tr>
		<tr>
			<td align="right">Section&nbsp;</td>
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
					<Option value='<%=SectionRS("Office")%>' <%if trim(Section_) = trim(SectionRS("Office")) then %>Selected<%End If%> ><%=SectionRS("Office")%></Option>
					
<%					SectionRS.MoveNext
				Loop%>
				</select>
			</td>
			<td align="right">Section by Group&nbsp;</td>
			<td>:</td>
			<td>
<%
 				strsql ="select distinct SectionGroup from vwPhoneCustomerList Where SectionGroup<>'' order by SectionGroup"
				set SectionGroupRS = server.createobject("adodb.recordset")
				set SectionGroupRS = BillingCon.execute(strsql)
'				response.write strStr 
%>	
				<Select name="cmbSectionGroup">
					<Option value='X'>--All--</Option>
<%				Do While not SectionGroupRS.eof %>
					<Option value='<%=SectionGroupRS("SectionGroup")%>' <%if trim(SectionGroup_) = trim(SectionGroupRS("SectionGroup")) then %>Selected<%End If%> ><%=SectionGroupRS("SectionGroup")%></Option>
					
<%					SectionGroupRS.MoveNext
				Loop%>
				</select>
			</td>
		</tr>
		<tr>	
			<td align="right">&nbsp;Employee&nbsp;</td>
			<td>:</td>
			<td>
<%
 				strsql ="select EmpID, EmpName from vwPhoneCustomerList order by EmpName"
				set EmpRS = server.createobject("adodb.recordset")
				set EmpRS = BillingCon.execute(strsql)
'				response.write strStr 
%>	
				<Select name="cmbEmp">
					<Option value='X'>--All--</Option>
<%				Do While not EmpRS.eof 
					EmpName_ = trim(EmpRS("EmpName"))
%>
					<Option value='<%=EmpRS("EmpID")%>' <%if trim(EmpID_) = trim(EmpRS("EmpID")) then %>Selected<%End If%> ><%=EmpName_ %></Option>
					
<%					EmpRS.MoveNext
				Loop%>
				</select>

			</td>
			<td align="right">Invoice Status&nbsp;</td>
			<td>:</td>
			<td>
<%
 				strsql ="select ProgressID, ProgressDesc from ProgressStatus Order By OrderNo"
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
<!--					<Option value='99' <%if trim(Status_) = 99 then %>Selected<%End If%>>Completed - Threshold</Option> -->
				</select>

			</td>	
		<tr>
		<tr>
			<td align="right">&nbsp;Email address&nbsp;</td>
			<td>:</td>
			<td>
				<input type="input" name="txtEmailAddress" value="<%=EmailAddress_%>" size=30 />(put "-" to find empty email)
			</td>	
			<td align="right">&nbsp;Sent Status&nbsp;</td>
			<td>:</td>
			<td>
<%
 				strsql ="select SendMailStatusID, SendMailStatusDesc from SendMailStatus"
				set SentStatusRS = server.createobject("adodb.recordset")
				set SentStatusRS = BillingCon.execute(strsql)
'				response.write strStr 
%>	
				<Select name="cmbSentStatus">
					<Option value='0'>--All--</Option>
<%				Do While not SentStatusRS.eof %>
					<Option value='<%=SentStatusRS("SendMailStatusID")%>' <%if trim(SentStatus_) = trim(SentStatusRS("SendMailStatusID")) then %>Selected<%End If%> ><%=SentStatusRS("SendMailStatusDesc")%></Option>
					
<%					SentStatusRS.MoveNext
				Loop%>
				</select>

			</td>		
		</tr>
		<tr>
			<td align="right">Show record&nbsp;</td>
			<td>:</td>
			<td>
				<Select name="cmbNoRecord">
					<Option value="1" <%if NoRecord_ = "1" then %>Selected<%End If%> >All</Option>
					<Option value="25" <%if NoRecord_ = "25" then %>Selected<%End If%> >25</Option>
					<Option value="50" <%if NoRecord_ = "50" then %>Selected<%End If%> >50</Option>
					<Option value="100" <%if NoRecord_ = "100" then %>Selected<%End If%> >100</Option>
				</Select>&nbsp;per Page
			</td>			
			<td align="right">Sort By&nbsp;</td>
			<td>:</td>
			<td>
				<Select name="SortList">
					<Option value="EmpName" <%if SortBy_ ="EmpName" then %>Selected<%End If%> >Employee Name</Option>
					<Option value="Office" <%if SortBy_ ="Office" then %>Selected<%End If%> >Section</Option>
					<Option value="SendMailStatusDesc" <%if SortBy_ ="SendMailStatusDesc" then %>Selected<%End If%> >Send Mail Status</Option>
					<Option value="ProgressDesc" <%if SortBy_ ="ProgressDesc" then %>Selected<%End If%> >Status</Option>
				</Select>&nbsp;
				<Select name="OrderList">
					<Option value="Asc" <%if Order_ ="Asc" then %>Selected<%End If%> >Asc</Option>
					<Option value="Desc" <%if Order_ ="Desc" then %>Selected<%End If%> >Desc</Option>
				</Select>
			</td
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
strsql = "Select CeilingAmount From PaymentDueDate"
set DataRS=BillingCon.execute(strsql)
if not DataRS.eof Then
	CeilingAmount_ = DataRS("CeilingAmount")
Else
	CeilingAmount_ = 0
End If

sPeriod = sYearP&sMonthP
ePeriod = eYearP&eMonthP
'response.write sPeriod & ePeriod 
'strsql = "Select * From vwMonthlyBilling Where ProgressID=5 and YearP+MonthP>='" & sPeriod & "' and YearP+MonthP<='" & ePeriod & "'"
'strsql = "Select * From vwMonthlyBilling Where ProgressID=1 and YearP+MonthP>='" & sPeriod & "' and YearP+MonthP<='" & ePeriod & "'"
strsql = "Select * From vwMonthlyBilling Where YearP+MonthP>='" & sPeriod & "' and YearP+MonthP<='" & ePeriod & "'"
strFilter=""
If Agency_ <> "X" then
	strFilter=strFilter & " and Agency='" & Agency_ & "'"
End If

If Section_ <> "X" then
	strFilter=strFilter & " and Office='" & Section_ & "'"
End If

If SectionGroup_ <> "X" then
	strFilter=strFilter & " and SectionGroup='" & SectionGroup_ & "'"
End If

If EmpID_ <> "X" then
	strFilter =strFilter & " and EmpID='" & EmpID_ & "'"
End If
'response.write "EmailAddress_ :" & EmailAddress_ 
If EmailAddress_ <> "" then
	If EmailAddress_ = "-" then
		strFilter =strFilter & " and EmailAddress=''"
	Else
		strFilter =strFilter & " and EmailAddress like '%" & EmailAddress_  & "%'"
	End If
End If

'response.write EmpID_ 
If Status_ <> "0" then
'	if Status_ ="99" Then
'		strFilter =strFilter & " and ProgressID=6 and CellPhonePrsBillRp<=" & CeilingAmount_ 
'	else
		strFilter =strFilter & " and ProgressID=" & Status_
'	end if
End If

If SentStatus_ <> "0" then
	strFilter =strFilter & " and SendMailStatusID=" & SentStatus_
End If
'response.write SentStatus_ 

If BillType = "O" then
	strFilter =strFilter & " and OfficePhoneBillRp>0 "
ElseIf BillType = "H" then
	strFilter =strFilter & " and HomePhoneBillRp>0 "
ElseIf BillType = "C" then
	strFilter =strFilter & " and CellPhoneBillRp>0 "
ElseIf BillType = "S" then
	strFilter =strFilter & " and TotalShuttleBillRp >0 "
End If

'strsql = strsql  & strFilter & " Order By EmpName,YearP,MonthP"
strsql = strsql  & strFilter & " Order By " & SortBy_ & " "& Order_
'response.write strsql & "<br>"

set DataRS = server.createobject("adodb.recordset")
DataRS.CursorLocation = 3
DataRS.Open strsql,BillingCon
'set DataRS=BillingCon.execute(strsql)

'response.write "NoRecord_" & NoRecord_

dim intPageSize,PageIndex,TotalPages 
dim RecordCount,RecordNumber,Count 
intpageSize = NoRecord_ 
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

%>
<form method="post" name="frmARReminder" Action="SendNotificationSendMail.asp" onSubmit="return ValidateForm()">
<table width="100%">
<tr>
	<td align="Left"><input type="button" name="btnExport" value="Export to Excel" onClick="javascript:document.location.href('SendNotificationPrint.asp?sMonthP=<%=sMonthP%>&sYearP=<%=sYearP%>&eMonthP=<%=eMonthP%>&eYearP=<%=eYearP%>&Agency=<%=Agency_%>&Section=<%=Section_%>&SectionGroup=<%=SectionGroup_%>&SortBy=<%=SortBy_%>&Order=<%=Order_%>&EmpID=<%=EmpID_%>&BillType=<%=BillType%>&Status=<%=Status_%>&SentStatus=<%=SentStatus_%>');"/></td>
	<td align="Right"><input type="submit" name="btnSendNotice" value="Send Notice(s)" /></td>
</tr>
</table>
<%If BillType = "X" then %>
	<table border="1" bordercolor="#EEEEEE" cellpadding="0" cellspacing="0" width="100%"  class="FontText">
	    <TR BGCOLOR="#330099" align="center">
        	 <TD rowspan="2" width="35px" align="right"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
	         <TD rowspan="2" Width="15%"><strong><label STYLE=color:#FFFFFF>Employee Name</label></strong></TD>
        	 <TD rowspan="2"><strong><label STYLE=color:#FFFFFF>Billing<br>Period</label></strong></TD>
	         <TD rowspan="2"><strong><label STYLE=color:#FFFFFF>Section</label></strong></TD>
	         <TD rowspan="2" Width="7%"><strong><label STYLE=color:#FFFFFF>Phone<br>Number</label></strong></TD>
		 <TD colspan="3"><strong><label STYLE=color:#FFFFFF>Billing Amount (Kn.)</label></strong></TD>
	         <TD rowspan="2" Width="30px">
			<input type="checkbox" name="cbAll" value="true" onclick="checkall(this)" />
		 </TD>
	         <TD rowspan="2" align="center"><strong><label STYLE=color:#FFFFFF>Sent<br>Status</label></strong></TD>
	         <TD rowspan="2"><strong><label STYLE=color:#FFFFFF>Sent Date</label></strong></TD>
	         <TD rowspan="2"><strong><label STYLE=color:#FFFFFF>Email</label></strong></TD>
	         <TD rowspan="2"><strong><label STYLE=color:#FFFFFF>Status</label></strong></TD>
	    </TR>
	    <tr BGCOLOR="#330099" align="center">
        	<!--  <TD width="7%"><strong><label STYLE=color:#FFFFFF>Home Phone</label></strong></TD>  -->
	         <TD width="7%"><strong><label STYLE=color:#FFFFFF>Office Phone</label></strong></TD>
        	 <TD width="7%"><strong><label STYLE=color:#FFFFFF>Mobile Phone</label></strong></TD>
	        <!--  <TD width="7%"><strong><label STYLE=color:#FFFFFF>Shuttle Bus</label></strong></TD>  -->
        	 <TD width="7%"><strong><label STYLE=color:#FFFFFF>Total</label></strong></TD>
	    </tr>
<%else%>
	<table border="1" bordercolor="#EEEEEE" cellpadding="0" cellspacing="0" width="100%"  class="FontText">
	    <TR BGCOLOR="#330099" align="center">
        	 <TD width="35px" align="right"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
	         <TD Width="15%"><strong><label STYLE=color:#FFFFFF>Employee Name</label></strong></TD>
        	 <TD><strong><label STYLE=color:#FFFFFF>Billing<br>Period</label></strong></TD>
	         <TD><strong><label STYLE=color:#FFFFFF>Section</label></strong></TD>
	         <TD Width="7%"><strong><label STYLE=color:#FFFFFF>Phone<br>Number</label></strong></TD>
	    <!--     <TD rowspan="2" width="100px"><strong><label STYLE=color:#FFFFFF>Bill Type</label></strong></TD>  -->
		 <TD Width="9%"><strong><label STYLE=color:#FFFFFF>Billing Amount (Kn.)</label></strong></TD>
		 <TD Width="9%"><strong><label STYLE=color:#FFFFFF>Personal Usage (Kn.)</label></strong></TD>
	         <TD Width="3%">
			<input type="checkbox" name="cbAll" value="true" onclick="checkall(this)" />
		 </TD>
	         <TD align="center"><strong><label STYLE=color:#FFFFFF>Sent<br>Status</label></strong></TD>
	         <TD Width="9%"><strong><label STYLE=color:#FFFFFF>Sent Date</label></strong></TD>
	         <TD><strong><label STYLE=color:#FFFFFF>Email</label></strong></TD>
	         <TD><strong><label STYLE=color:#FFFFFF>Status</label></strong></TD>
	    </TR>
	 <!--    <tr BGCOLOR="#330099" align="center">
        	 <TD width="6%"><strong><label STYLE=color:#FFFFFF>In Kuna (Kn.)</label></strong></TD>
	         <TD width="6%"><strong><label STYLE=color:#FFFFFF>in US Dollar ($)</label></strong></TD>
	    </tr>  -->
<%end if%>
<%   
 dim no_  
  
if (NoRecord_ = 1) then
 no_ = 1
 do while not DataRS.eof

	   TotalBillingRp_ = 0
	   if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
%>
      
	   <TR bgcolor="<%=bg%>">
	        <TD align="right"><%=no_ %>&nbsp;</font></TD>
	        <TD>&nbsp;<%=DataRS("EmpName") %></TD>
	        <TD align="right">&nbsp;<%= DataRS("MonthP")%>-<%= DataRS("YearP")%></font>&nbsp;</TD>
	        <TD>&nbsp;<%=DataRS("Office") %> </font></TD>
	        <TD>&nbsp;<%=DataRS("MobilePhone") %> </font></TD>
<%If BillType <> "X" then 
		If BillType ="O" then
			BillTypeDesc ="Office Phone"
		ElseIf BillType ="C" then
			BillTypeDesc ="Cell Phone"
		ElseIf BillType ="H" then
			BillTypeDesc ="Home Phone"
		ElseIf BillType ="S" then
			BillTypeDesc ="Shuttle Bus"
		End If
%>
	      <!--  <TD>&nbsp;<%=BillTypeDesc %> </font></TD>  -->
<%end if%>
<!-- <%		If (BillType = "H") or (BillType = "X") Then 
			If CDbl(DataRS("HomePhonePrsBillRp")) > 0 then %>
				<td align="right">
					<a href="HomePhoneDetail.asp?HomePhone=<%=DataRS("HomePhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("HomePhonePrsBillRp"),-1) %></a>
				</td>
<%				TotalBillingRp_ = TotalBillingRp_ + cdbl(DataRS("HomePhonePrsBillRp")) %>
<%			Else %>
				<td align="right">-&nbsp;</td>
<%			End If%>
<%		End If %>   -->
<%		If (BillType = "O")  or (BillType = "X") Then 
			If CDbl(DataRS("OfficePhonePrsBillRp")) > 0 then %>
				<td align="right">
					<a href="OfficePhoneDetail.asp?Extension=<%=DataRS("WorkPhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("OfficePhonePrsBillRp"),-1) %></a>
				</td>
<%			TotalBillingRp_ = cdbl(TotalBillingRp_ )+cdbl(DataRS("OfficePhonePrsBillRp"))%>
<%			Else %>
				<td align="right">-&nbsp;</td>
<%			End If%>
<%		End If %>
<%		If (BillType = "C")  or (BillType = "X") Then 
			If CDbl(DataRS("CellPhoneBillRp")) > 0 then %>
				<td align="right">
					<a href="CellPhoneDetail.asp?CellPhone=<%=DataRS("MobilePhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("CellPhoneBillRp"),-1) %></a>
				</td>
<%			TotalBillingRp_ = cdbl(TotalBillingRp_ )+cdbl(DataRS("CellPhoneBillRp"))%>
<%			Else %>
				<td align="right">-&nbsp;</td>
<%			End If%>
<%		End If %>
<!-- <%		If (BillType = "S")  or (BillType = "X") Then 
			If CDbl(DataRS("TotalShuttleBillRp")) > 0 then %>
				<td align="right">
					<a href="ShuttleBusBillDetail.asp?Username=<%=DataRS("LoginID") %>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("TotalShuttleBillRp"),-1) %></a>
				</td>
<%			TotalBillingRp_ = cdbl(TotalBillingRp_ )+cdbl(DataRS("TotalShuttleBillRp"))%>
<%			Else %>
				<td align="right">-&nbsp;</td>
<%			End If%>
<%		End If %>   -->


<%		If (BillType = "H") Then 
			If CDbl(DataRS("HomePhonePrsBillRp")) > 0 Then %>
				<td align="right">
					<a href="HomePhoneDetail.asp?HomePhone=<%=DataRS("HomePhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("HomePhonePrsBillRp"),-1) %></a>
				</td>
<%			Else %>
				<td align="right">-&nbsp;</td>
<%			End If %>
<%		End If %>

<%		If (BillType = "O") Then
			If CDbl(DataRS("OfficePhonePrsBillRp")) > 0 Then %>
			<td align="right">
				<a href="OfficePhoneDetail.asp?Extension=<%=DataRS("WorkPhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("OfficePhonePrsBillRp"),-1) %></a>
			</td>
<%			Else %>
				<td align="right">-&nbsp;</td>
<%			End If %>
<%		End If %>
<%		If (BillType = "C") Then 
			If CDbl(DataRS("CellPhonePrsBillRp")) > 0 Then %>
				<td align="right">
					<a href="CellPhoneDetail.asp?CellPhone=<%=DataRS("MobilePhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("CellPhonePrsBillRp"),-1) %></a>
				</td>
<%			Else%>
				<td align="right">-&nbsp;</td>
<%			End If %>
<%		End If %>
<%		If (BillType = "S") Then
			If CDbl(DataRS("TotalShuttleBillRp")) > 0 Then %>
				<td align="right">
					<a href="ShuttleBusBillDetail.asp?Username=<%=DataRS("LoginID") %>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("TotalShuttleBillRp"),-1) %></a>
				</td>
<%			Else%>
				<td align="right">-&nbsp;</td>
<%			End If %>
<%		End If %>
<%		If CDbl(DataRS("HomePhonePrsBillDlr")) > 0 Then
			If (BillType = "H") Then %>
				<td align="right">
					<a href="HomePhoneDetail.asp?HomePhone=<%=DataRS("HomePhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("HomePhonePrsBillDlr"),-1) %></a>
				</td>
<%			End If %>
<%		ElseIf (CDbl(DataRS("HomePhonePrsBillDlr")) = 0) and (BillType = "H") Then %>
			<td align="right">-&nbsp;</td>
<%		End If %>
<%		If CDbl(DataRS("OfficePhonePrsBillDlr")) > 0 Then 
			If (BillType = "O") Then %>
			<td align="right">
				<a href="OfficePhoneDetail.asp?Extension=<%=DataRS("WorkPhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("OfficePhonePrsBillDlr"),-1) %></a>
			</td>
<%			End If %>
<%		ElseIf (CDbl(DataRS("OfficePhonePrsBillDlr")) = 0) and (BillType = "O") Then %>
			<td align="right">-&nbsp;</td>
<%		End If %>
<!-- <%		If CDbl(DataRS("CellPhonePrsBillDlr")) > 0 Then
			If (BillType = "C") Then %>
			<td align="right">
				<a href="CellPhoneDetail.asp?CellPhone=<%=DataRS("MobilePhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("CellPhonePrsBillDlr"),-1) %></a>
			</td>
<%			End If %>
<%		ElseIf (CDbl(DataRS("CellPhonePrsBillDlr")) = 0) and (BillType = "C")  Then %>
			<td align="right">-&nbsp;</td>
<%		End If %>   -->
<%		If CDbl(DataRS("TotalShuttleBillDlr")) > 0 Then
			If (BillType = "S") Then %>
				<td align="right">
					<a href="ShuttleBusBillDetail.asp?Username=<%=DataRS("LoginID") %>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("TotalShuttleBillDlr"),-1) %></a>
				</td>
<%			End If %>
<%		ElseIf (CDbl(DataRS("TotalShuttleBillDlr")) = 0) and (BillType = "S") Then %>
			<td align="right">-&nbsp;</td>
<%		End If %>

<%If BillType = "X" then 
%>
	        <TD align="right"><%= formatnumber(TotalBillingRp_,-1) %>&nbsp;</font></TD>
<%End if%>
		<td align="center">
<%		If len(DataRS("EmailAddress"))>5 then %>
			<Input type="Checkbox" name="cbApproval" Value='<%=DataRS("EmpID")%><%=DataRS("MonthP")%><%=DataRS("YearP")%><%=BillType%>'>
<%		Else%>
			&nbsp;
<%		End If%>
		</td>
	        <TD align="center">&nbsp;<%=DataRS("SendMailStatusDesc") %> </font></TD>
	        <TD>&nbsp;<%=DataRS("SendMailDate") %> </font></TD>
	        <TD>&nbsp;<%=DataRS("EmailAddress") %> </font></TD>
	        <TD><a title="click to update the status" href="ChangeProgressStatus.asp?EmpID=<%=DataRS("EmpID")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%=DataRS("ProgressDesc") %> </font></a></TD>
	    </TR>

<%   
	   DataRS.movenext
	   no_ = no_ + 1 
   loop 
	PageNo=1
%>
</table>
<%
Else
   no_ = 1 + ((PageIndex*intPageSize)-intPageSize)
   Count=1 
   'response.write "intPageSize :" & intPageSize
   do while not DataRS.eof and cdbl(Count)<= cdbl(intPageSize)

	   TotalBillingRp_ = 0
	   if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
%>
      
	   <TR bgcolor="<%=bg%>">
	        <TD align="right"><%=no_ %>&nbsp;</font></TD>
	        <TD>&nbsp;<%=DataRS("EmpName") %></TD>
	        <TD align="right">&nbsp;<%= DataRS("MonthP")%>-<%= DataRS("YearP")%></font>&nbsp;</TD>
	        <TD>&nbsp;<%=DataRS("Office") %> </font></TD>
<%If BillType <> "X" then 
		If BillType ="O" then
			BillTypeDesc ="Office Phone"
		ElseIf BillType ="C" then
			BillTypeDesc ="Cell Phone"
		ElseIf BillType ="H" then
			BillTypeDesc ="Home Phone"
		ElseIf BillType ="S" then
			BillTypeDesc ="Shuttle Bus"
		End If
%>
	        <TD>&nbsp;<%=BillTypeDesc %> </font></TD>
<%end if%>
<%		If (BillType = "H") or (BillType = "X") Then 
			If CDbl(DataRS("HomePhonePrsBillRp")) > 0 then %>
				<td align="right">
					<a href="HomePhoneDetail.asp?HomePhone=<%=DataRS("HomePhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("HomePhonePrsBillRp"),-1) %></a>
				</td>
<%				TotalBillingRp_ = TotalBillingRp_ + cdbl(DataRS("HomePhonePrsBillRp")) %>
<%			Else %>
				<td align="right">-&nbsp;</td>
<%			End If%>
<%		End If %>
<%		If (BillType = "O")  or (BillType = "X") Then 
			If CDbl(DataRS("OfficePhonePrsBillRp")) > 0 then %>
				<td align="right">
					<a href="OfficePhoneDetail.asp?Extension=<%=DataRS("WorkPhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("OfficePhonePrsBillRp"),-1) %></a>
				</td>
<%			TotalBillingRp_ = cdbl(TotalBillingRp_ )+cdbl(DataRS("OfficePhonePrsBillRp"))%>
<%			Else %>
				<td align="right">-&nbsp;</td>
<%			End If%>
<%		End If %>
<%		If (BillType = "C")  or (BillType = "X") Then 
			If CDbl(DataRS("CellPhonePrsBillRp")) > 0 then %>
				<td align="right">
					<a href="CellPhoneDetail.asp?CellPhone=<%=DataRS("MobilePhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("CellPhonePrsBillRp"),-1) %></a>
				</td>
<%			TotalBillingRp_ = cdbl(TotalBillingRp_ )+cdbl(DataRS("CellPhonePrsBillRp"))%>
<%			Else %>
				<td align="right">-&nbsp;</td>
<%			End If%>
<%		End If %>
<%		If (BillType = "S")  or (BillType = "X") Then 
			If CDbl(DataRS("TotalShuttleBillRp")) > 0 then %>
				<td align="right">
					<a href="ShuttleBusBillDetail.asp?Username=<%=DataRS("LoginID") %>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("TotalShuttleBillRp"),-1) %></a>
				</td>
<%			TotalBillingRp_ = cdbl(TotalBillingRp_ )+cdbl(DataRS("TotalShuttleBillRp"))%>
<%			Else %>
				<td align="right">-&nbsp;</td>
<%			End If%>
<%		End If %>


<%		If (BillType = "H") Then 
			If CDbl(DataRS("HomePhonePrsBillRp")) > 0 Then %>
				<td align="right">
					<a href="HomePhoneDetail.asp?HomePhone=<%=DataRS("HomePhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("HomePhonePrsBillRp"),-1) %></a>
				</td>
<%			Else %>
				<td align="right">-&nbsp;</td>
<%			End If %>
<%		End If %>

<%		If (BillType = "O") Then
			If CDbl(DataRS("OfficePhonePrsBillRp")) > 0 Then %>
			<td align="right">
				<a href="OfficePhoneDetail.asp?Extension=<%=DataRS("WorkPhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("OfficePhonePrsBillRp"),-1) %></a>
			</td>
<%			Else %>
				<td align="right">-&nbsp;</td>
<%			End If %>
<%		End If %>
<%		If (BillType = "C") Then 
			If CDbl(DataRS("CellPhonePrsBillRp")) > 0 Then %>
				<td align="right">
					<a href="CellPhoneDetail.asp?CellPhone=<%=DataRS("MobilePhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("CellPhonePrsBillRp"),-1) %></a>
				</td>
<%			Else%>
				<td align="right">-&nbsp;</td>
<%			End If %>
<%		End If %>
<%		If (BillType = "S") Then
			If CDbl(DataRS("TotalShuttleBillRp")) > 0 Then %>
				<td align="right">
					<a href="ShuttleBusBillDetail.asp?Username=<%=DataRS("LoginID") %>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("TotalShuttleBillRp"),-1) %></a>
				</td>
<%			Else%>
				<td align="right">-&nbsp;</td>
<%			End If %>
<%		End If %>
<%		If CDbl(DataRS("HomePhonePrsBillDlr")) > 0 Then
			If (BillType = "H") Then %>
				<td align="right">
					<a href="HomePhoneDetail.asp?HomePhone=<%=DataRS("HomePhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("HomePhonePrsBillDlr"),-1) %></a>
				</td>
<%			End If %>
<%		ElseIf (CDbl(DataRS("HomePhonePrsBillDlr")) = 0) and (BillType = "H") Then %>
			<td align="right">-&nbsp;</td>
<%		End If %>
<%		If CDbl(DataRS("OfficePhonePrsBillDlr")) > 0 Then 
			If (BillType = "O") Then %>
			<td align="right">
				<a href="OfficePhoneDetail.asp?Extension=<%=DataRS("WorkPhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("OfficePhonePrsBillDlr"),-1) %></a>
			</td>
<%			End If %>
<%		ElseIf (CDbl(DataRS("OfficePhonePrsBillDlr")) = 0) and (BillType = "O") Then %>
			<td align="right">-&nbsp;</td>
<%		End If %>
<%		If CDbl(DataRS("CellPhonePrsBillDlr")) > 0 Then
			If (BillType = "C") Then %>
			<td align="right">
				<a href="CellPhoneDetail.asp?CellPhone=<%=DataRS("MobilePhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("CellPhonePrsBillDlr"),-1) %></a>
			</td>
<%			End If %>
<%		ElseIf (CDbl(DataRS("CellPhonePrsBillDlr")) = 0) and (BillType = "C")  Then %>
			<td align="right">-&nbsp;</td>
<%		End If %>
<%		If CDbl(DataRS("TotalShuttleBillDlr")) > 0 Then
			If (BillType = "S") Then %>
				<td align="right">
					<a href="ShuttleBusBillDetail.asp?Username=<%=DataRS("LoginID") %>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("TotalShuttleBillDlr"),-1) %></a>
				</td>
<%			End If %>
<%		ElseIf (CDbl(DataRS("TotalShuttleBillDlr")) = 0) and (BillType = "S") Then %>
			<td align="right">-&nbsp;</td>
<%		End If %>

<%If BillType = "X" then 
%>
	        <TD align="right"><%= formatnumber(TotalBillingRp_,-1) %>&nbsp;</font></TD>
<%End if%>
		<td align="center">
<%		If len(DataRS("EmailAddress"))>5 then %>
			<Input type="Checkbox" name="cbApproval" Value='<%=DataRS("EmpID")%><%=DataRS("MonthP")%><%=DataRS("YearP")%><%=BillType%>'>
<%		Else%>
			&nbsp;
<%		End If%>
		</td>
	        <TD>&nbsp;<%=DataRS("SendMailStatusDesc") %> </font></TD>
	        <TD>&nbsp;<%=DataRS("SendMailDate") %> </font></TD>
	        <TD>&nbsp;<%=DataRS("EmailAddress") %> </font></TD>
	        <TD>&nbsp;<%=DataRS("ProgressDesc") %> </font></TD>
	    </TR>

<%   
		Count=Count +1
		'response.write Count
	   DataRS.movenext
	   no_ = no_ + 1 
   loop 
	PageNo=1
%>
</table>

<table width="100%">
	<tr>
		<td align="right">
<%
		Do while PageNo<=TotalPages 
			if trim(pageNo) = trim(PageIndex) Then
%>		
				<label class="ActivePage"><%=PageNo%></label>&nbsp;
<%			Else%>
				<a href="SendNotification.asp?PageIndex=<%=PageNo%>&NoRecord=<%=NoRecord_%>&sMonthP=<%=sMonthP%>&sYearP=<%=sYearP%>&eMonthP=<%=eMonthP%>&eYearP=<%=eYearP%>&Agency=<%=Agency_%>&Section=<%=Section_%>&SectionGroup=<%=SectionGroup_%>&EmpID=<%=EmpID_%>&BillType=<%=BillType%>&Status=<%=Status_%>&SentStatus=<%=SentStatus_%>&SortBy=<%=SortBy_%>&Order=<%=Order_%>"><%=PageNo%></a>&nbsp;
<%	
			End If						
			PageNo=PageNo+1
		Loop
%>
		</td>
	</tr>
</table>
<%end if%>
</form>
<%
else 
%>
	<table cellspadding="1" cellspacing="0" width="100%">  
	<tr>
        	<td><br></TD>
	</tr>
	<tr>
		<td align="center">There is no data.</td>
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


