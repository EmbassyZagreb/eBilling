<%@ Language=VBScript%>
<% ' VI 6.0 Scripting Object Model Enabled %>
 
 
<html>
<head>

<META name=VI60_defaultClientScript content=VBScript>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<!--#include file="connect.inc" -->

<style type="text/css">
<!--

.FontText {
	font-size: small;
}


.Hint{
	color: gray;
	font-size: x-small;
}

-->
</style>
<script type="text/javascript">

function ClearFilter()
{
	document.forms['frmSearch'].elements['cmbAgency'].value="X";
	document.forms['frmSearch'].elements['cmbSection'].value="X";
	document.forms['frmSearch'].elements['cmbEmp'].value="X";
	document.forms['frmSearch'].elements['cmbStatus'].value="1";
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

EmpID_ = request("EmpID")
if EmpID_ = "" Then EmpID_ = Request.Form("cmbEmp")
if EmpID_ = "" then
	EmpID_ = "X"
end if

BillType = request("BillType")
if BillType = "" Then BillType = Request.Form("cmbBillType")
if BillType = "" then
	BillType = "X"
end if

Status_ = request("Status")
if Status_ = "" Then Status_ = Request.Form("cmbStatus")
if Status_ = "" then
	Status_ = 1
end if
%>
<TITLE>U.S. Mission Jakarta e-Billing</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">

<STYLE TYPE="text/css"><!--
  A:ACTIVE { color:#003399; font-size:8pt; font-family:Verdana; }
  A:HOVER { color:#003399; font-size:8pt; font-family:Verdana; }
  A:LINK { color:#003399; font-size:8pt; font-family:Verdana; }
  A:VISITED { color:#003399; font-size:8pt; font-family:Verdana; }
  body {scrollbar-3dlight-color:#FFFFFF; scrollbar-arrow-color:#E3DCD5; scrollbar-base-color:#FFFFFF; scrollbar-darkshadow-color:#FFFFFF;	scrollbar-face-color:#FFFFFF; scrollbar-highlight-color:#E3DCD5; scrollbar-shadow-color:#E3DCD5; }
  p { font-family: verdana; font-size: 12px; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; color: #003399; text-decoration: none}
  h3 { font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 16px; font-style: normal; line-height: normal; font-weight: bold; color: #003399; letter-spacing: normal; word-spacing: normal; font-variant: small-caps}
  td { font-family: verdana; font-size: 10px; font-style: normal; font-weight: normal; color: #000000}
  .title { font-size:14px; font-weight:bold; color:#000080; }
  .SubTitle { font-size:16px; font-weight:bold; color:#000080;  }
  A.menu { text-decoration:none; font-weight:bold; }
  A.mmenu { text-decoration:none; color:#FFFFFF; font-weight:bold; }
  .normal { font-family:Verdana,Arial; color:black}
  .style5 {color: #FFFFFF;}
  .ActivePage {color: red; font-weight:bold; }
--></STYLE>
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
<form method="post" name="frmSearch" Action="zSendNotification.asp">
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
			<td align="right">Bill Type&nbsp;</td>
			<td>:</td>
			<td>
				<Select name="cmbBillType">
					<Option value="X" <%if BillType ="X" then %>Selected<%End If%>>-All-</Option>
					<Option value="O" <%if BillType = "O" then %>Selected<%End If%>>Office Phone</Option>
					<Option value="C" <%if BillType = "C" then %>Selected<%End If%>>Cell Phone</Option>
					<Option value="H" <%if BillType = "H" then %>Selected<%End If%>>Home Phone</Option>
					<Option value="S" <%if BillType = "S" then %>Selected<%End If%>>Shuttle Bus</Option>
				</Select>&nbsp;
			</td>
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
			<td align="right">Status&nbsp;</td>
			<td>:</td>
			<td>
<%
 				strsql ="select ProgressID, ProgressDesc from ProgressStatus"
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
			<td align="right">&nbsp;Employee&nbsp;</td>
			<td>:</td>
			<td colspan="4">
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
		<tr>
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

If EmpID_ <> "X" then
	strFilter =strFilter & " and EmpID='" & EmpID_ & "'"
End If

If Status_ <> "0" then
	strFilter =strFilter & " and ProgressID=" & Status_
End If

If BillType = "O" then
	strFilter =strFilter & " and OfficePhoneBillRp>0 "
ElseIf BillType = "H" then
	strFilter =strFilter & " and HomePhoneBillRp>0 "
ElseIf BillType = "C" then
	strFilter =strFilter & " and CellPhoneBillRp>0 "
ElseIf BillType = "S" then
	strFilter =strFilter & " and TotalShuttleBillRp >0 "
End If

strsql = strsql  & strFilter & " Order By EmpName,YearP,MonthP"
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

%>
<form method="post" name="frmARReminder" Action="zSendNotificationSendMail.asp" onSubmit="return ValidateForm()">
<table width="100%">
<tr>
	<td align="Left"><input type="button" name="btnExport" value="Export to Excel" onClick="javascript:document.location.href('SendNotificationPrint.asp?sMonthP=<%=sMonthP%>&sYearP=<%=sYearP%>&eMonthP=<%=eMonthP%>&eYearP=<%=eYearP%>&Agency=<%=Agency_%>&Section=<%=Section_%>&EmpID=<%=EmpID_%>&BillType=<%=BillType%>');"/></td>
	<td align="Right"><input type="submit" name="btnSendNotice" value="Send Notice(s)" /></td>
</tr>
</table>
<%If BillType = "X" then %>
	<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" width="120%"  class="FontText">
	    <TR BGCOLOR="#330099" align="center">
        	 <TD rowspan="2" width="35px" align="right"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
	         <TD rowspan="2"><strong><label STYLE=color:#FFFFFF>Employee Name</label></strong></TD>
        	 <TD rowspan="2" width="100px"><strong><label STYLE=color:#FFFFFF>Billing Period</label></strong></TD>
	         <TD rowspan="2"><strong><label STYLE=color:#FFFFFF>Section</label></strong></TD>
		 <TD colspan="5"><strong><label STYLE=color:#FFFFFF>Billing Amount (Rp.)</label></strong></TD>
	         <TD rowspan="2" Width="30px">
			<input type="checkbox" name="cbAll" value="true" onclick="checkall(this)" />
		 </TD>
	         <TD rowspan="2" width="50px"><strong><label STYLE=color:#FFFFFF>Sent Status</label></strong></TD>
	         <TD rowspan="2" width="80px"><strong><label STYLE=color:#FFFFFF>Sent Date</label></strong></TD>
	         <TD rowspan="2"><strong><label STYLE=color:#FFFFFF>Email</label></strong></TD>
	         <TD rowspan="2"><strong><label STYLE=color:#FFFFFF>Status</label></strong></TD>
	    </TR>
	    <tr BGCOLOR="#330099" align="center">
        	 <TD width="100px"><strong><label STYLE=color:#FFFFFF>Home Phone</label></strong></TD>
	         <TD width="100px"><strong><label STYLE=color:#FFFFFF>Office Phone</label></strong></TD>
        	 <TD width="100px"><strong><label STYLE=color:#FFFFFF>Mobile Phone</label></strong></TD>
	         <TD width="100px"><strong><label STYLE=color:#FFFFFF>Shuttle Bus</label></strong></TD>
        	 <TD width="100px"><strong><label STYLE=color:#FFFFFF>Total</label></strong></TD>
	    </tr>
<%else%>
	<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" width="120%"  class="FontText">
	    <TR BGCOLOR="#330099" align="center">
        	 <TD rowspan="2" width="35px" align="right"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
	         <TD rowspan="2"><strong><label STYLE=color:#FFFFFF>Employee Name</label></strong></TD>
        	 <TD rowspan="2" width="100px"><strong><label STYLE=color:#FFFFFF>Billing Period</label></strong></TD>
	         <TD rowspan="2"><strong><label STYLE=color:#FFFFFF>Section</label></strong></TD>
	         <TD rowspan="2" width="100px"><strong><label STYLE=color:#FFFFFF>Bill Type</label></strong></TD>
		 <TD rowspan="2" ><strong><label STYLE=color:#FFFFFF>Billing Amount (Rp.)</label></strong></TD>
		 <TD colspan="2"><strong><label STYLE=color:#FFFFFF>Personal Usage (Rp.)</label></strong></TD>
	         <TD rowspan="2" Width="3%">
			<input type="checkbox" name="cbAll" value="true" onclick="checkall(this)" />
		 </TD>
	         <TD rowspan="2"><strong><label STYLE=color:#FFFFFF>Sent Status</label></strong></TD>
	         <TD rowspan="2"><strong><label STYLE=color:#FFFFFF>Sent Date</label></strong></TD>
	         <TD rowspan="2"><strong><label STYLE=color:#FFFFFF>Email</label></strong></TD>
	         <TD rowspan="2"><strong><label STYLE=color:#FFFFFF>Status</label></strong></TD>
	    </TR>
	    <tr BGCOLOR="#330099" align="center">
        	 <TD width="10%"><strong><label STYLE=color:#FFFFFF>In Rupiah (Rp.)</label></strong></TD>
	         <TD width="10%"><strong><label STYLE=color:#FFFFFF>in US Dollar ($)</label></strong></TD>
	    </tr>
<%end if%>
<% 
   dim no_  
   no_ = 1 + ((PageIndex*intPageSize)-intPageSize)
   Count=1 
   do while not DataRS.eof   and Count<=intPageSize
	   TotalBillingRp_ = 0
	   if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
%>
      
	   <TR bgcolor="<%=bg%>">
	        <TD align="right"><FONT color=#330099 size=2><%=no_ %>&nbsp;</font></TD>
	        <TD>&nbsp;<%=DataRS("EmpName") %></TD>
	        <TD align="right"><FONT color=#330099 size=2>&nbsp;<%= DataRS("MonthP")%>-<%= DataRS("YearP")%></font>&nbsp;</TD>
	        <TD><FONT color=#330099 size=2>&nbsp;<%=DataRS("Office") %> </font></TD>
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
	        <TD><FONT color=#330099 size=2>&nbsp;<%=BillTypeDesc %> </font></TD>
<%end if%>
<%		If (BillType = "H") or (BillType = "X") Then 
			If CDbl(DataRS("HomePhonePrsBillRp")) > 0 then %>
				<td align="right">
					<a href="HomePhoneDetail.asp?HomePhone=<%=DataRS("HomePhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("HomePhonePrsBillRp"),0) %></a>
				</td>
<%				TotalBillingRp_ = TotalBillingRp_ + cdbl(DataRS("HomePhonePrsBillRp")) %>
<%			Else %>
				<td align="right">-&nbsp;</td>
<%			End If%>
<%		End If %>
<%		If (BillType = "O")  or (BillType = "X") Then 
			If CDbl(DataRS("OfficePhonePrsBillRp")) > 0 then %>
				<td align="right">
					<a href="OfficePhoneDetail.asp?Extension=<%=DataRS("WorkPhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("OfficePhonePrsBillRp"),0) %></a>
				</td>
<%			TotalBillingRp_ = cdbl(TotalBillingRp_ )+cdbl(DataRS("OfficePhonePrsBillRp"))%>
<%			Else %>
				<td align="right">-&nbsp;</td>
<%			End If%>
<%		End If %>
<%		If (BillType = "C")  or (BillType = "X") Then 
			If CDbl(DataRS("CellPhonePrsBillRp")) > 0 then %>
				<td align="right">
					<a href="CellPhoneDetail.asp?CellPhone=<%=DataRS("MobilePhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("CellPhonePrsBillRp"),0) %></a>
				</td>
<%			TotalBillingRp_ = cdbl(TotalBillingRp_ )+cdbl(DataRS("CellPhonePrsBillRp"))%>
<%			Else %>
				<td align="right">-&nbsp;</td>
<%			End If%>
<%		End If %>
<%		If (BillType = "S")  or (BillType = "X") Then 
			If CDbl(DataRS("TotalShuttleBillRp")) > 0 then %>
				<td align="right">
					<a href="ShuttleBusBillDetail.asp?Username=<%=DataRS("LoginID") %>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("TotalShuttleBillRp"),0) %></a>
				</td>
<%			TotalBillingRp_ = cdbl(TotalBillingRp_ )+cdbl(DataRS("TotalShuttleBillRp"))%>
<%			Else %>
				<td align="right">-&nbsp;</td>
<%			End If%>
<%		End If %>


<%		If (BillType = "H") Then 
			If CDbl(DataRS("HomePhonePrsBillRp")) > 0 Then %>
				<td align="right">
					<a href="HomePhoneDetail.asp?HomePhone=<%=DataRS("HomePhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("HomePhonePrsBillRp"),0) %></a>
				</td>
<%			Else %>
				<td align="right">-&nbsp;</td>
<%			End If %>
<%		End If %>

<%		If (BillType = "O") Then
			If CDbl(DataRS("OfficePhonePrsBillRp")) > 0 Then %>
			<td align="right">
				<a href="OfficePhoneDetail.asp?Extension=<%=DataRS("WorkPhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("OfficePhonePrsBillRp"),0) %></a>
			</td>
<%			Else %>
				<td align="right">-&nbsp;</td>
<%			End If %>
<%		End If %>
<%		If (BillType = "C") Then 
			If CDbl(DataRS("CellPhonePrsBillRp")) > 0 Then %>
				<td align="right">
					<a href="CellPhoneDetail.asp?CellPhone=<%=DataRS("MobilePhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("CellPhonePrsBillRp"),0) %></a>
				</td>
<%			Else%>
				<td align="right">-&nbsp;</td>
<%			End If %>
<%		End If %>
<%		If (BillType = "S") Then
			If CDbl(DataRS("TotalShuttleBillRp")) > 0 Then %>
				<td align="right">
					<a href="ShuttleBusBillDetail.asp?Username=<%=DataRS("LoginID") %>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("TotalShuttleBillRp"),0) %></a>
				</td>
<%			Else%>
				<td align="right">-&nbsp;</td>
<%			End If %>
<%		End If %>
<%		If CDbl(DataRS("HomePhonePrsBillDlr")) > 0 Then
			If (BillType = "H") Then %>
				<td align="right">
					<a href="HomePhoneDetail.asp?HomePhone=<%=DataRS("HomePhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("HomePhonePrsBillDlr"),0) %></a>
				</td>
<%			End If %>
<%		ElseIf (CDbl(DataRS("HomePhonePrsBillDlr")) = 0) and (BillType = "H") Then %>
			<td align="right">-&nbsp;</td>
<%		End If %>
<%		If CDbl(DataRS("OfficePhonePrsBillDlr")) > 0 Then 
			If (BillType = "O") Then %>
			<td align="right">
				<a href="OfficePhoneDetail.asp?Extension=<%=DataRS("WorkPhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("OfficePhonePrsBillDlr"),0) %></a>
			</td>
<%			End If %>
<%		ElseIf (CDbl(DataRS("OfficePhonePrsBillDlr")) = 0) and (BillType = "O") Then %>
			<td align="right">-&nbsp;</td>
<%		End If %>
<%		If CDbl(DataRS("CellPhonePrsBillDlr")) > 0 Then
			If (BillType = "C") Then %>
			<td align="right">
				<a href="CellPhoneDetail.asp?CellPhone=<%=DataRS("MobilePhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("CellPhonePrsBillDlr"),0) %></a>
			</td>
<%			End If %>
<%		ElseIf (CDbl(DataRS("CellPhonePrsBillDlr")) = 0) and (BillType = "C")  Then %>
			<td align="right">-&nbsp;</td>
<%		End If %>
<%		If CDbl(DataRS("TotalShuttleBillDlr")) > 0 Then
			If (BillType = "S") Then %>
				<td align="right">
					<a href="ShuttleBusBillDetail.asp?Username=<%=DataRS("LoginID") %>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("TotalShuttleBillDlr"),0) %></a>
				</td>
<%			End If %>
<%		ElseIf (CDbl(DataRS("TotalShuttleBillDlr")) = 0) and (BillType = "S") Then %>
			<td align="right">-&nbsp;</td>
<%		End If %>

<%If BillType = "X" then 
%>
	        <TD align="right"><FONT color=#330099 size=2><%= formatnumber(TotalBillingRp_,0) %>&nbsp;</font></TD>
<%End if%>
		<td align="center">
<%		If len(DataRS("EmailAddress"))>5 then %>
			<Input type="Checkbox" name="cbApproval" Value='<%=DataRS("EmpID")%><%=DataRS("MonthP")%><%=DataRS("YearP")%><%=BillType%>'>
<%		Else%>
			&nbsp;
<%		End If%>
		</td>
	        <TD><FONT color=#330099 size=2>&nbsp;<%=DataRS("SendMailStatusDesc") %> </font></TD>
	        <TD><FONT color=#330099 size=2>&nbsp;<%=DataRS("SendMailDate") %> </font></TD>
	        <TD><FONT color=#330099 size=2>&nbsp;<%=DataRS("EmailAddress") %> </font></TD>
	        <TD><FONT color=#330099 size=2>&nbsp;<%=DataRS("ProgressDesc") %> </font></TD>
	    </TR>

<%   
		Count=Count +1
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
			<%Else%>
				<a href="SendNotification.asp?PageIndex=<%=PageNo%>&sMonthP=<%=sMonthP%>&sYearP=<%=sYearP%>&eMonthP=<%=eMonthP%>&eYearP=<%=eYearP%>&Agency=<%=Agency_%>&Section=<%=Section_%>&EmpID=<%=EmpID_%>&BillType=<%=BillType%>"><%=PageNo%></a>&nbsp;
<%	
			End If						
			PageNo=PageNo+1
		Loop
%>
		</td>
	</tr>
</table>
</form>
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
			<td>Please <a href="http://jakartaws01.eap.state.sbu/CSC">Submit Request </a> or contact Jakarta CSC Helpdesk at ext.9111.</td>
		</tr>
	</table>
<% end if %>
</body> 

</html>

