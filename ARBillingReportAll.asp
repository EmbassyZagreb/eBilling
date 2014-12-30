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

EmpID_ = request("EmpID")
if EmpID_ = "" Then EmpID_ = Request.Form("cmbEmp")
if EmpID_ = "" then
	EmpID_ = "X"
end if

Status = request("Status")
if Status = "" Then Status = Request.Form("cmbStatus")
if Status = "" then
	Status = "Pending"
end if
%>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">A/R BILLING REPORT</TD>
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
If (UserRole_ = "Admin") or (UserRole_ = "Voucher") or (UserRole_ = "FMC") or (UserRole_ = "Cashier")  Then
%>
<form method="post" name="frmSearch" Action="ARBillingReportAll.asp">
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
					<Option value='<%=SectionRS("Office")%>' <%if trim(Office_) = trim(SectionRS("Office")) then %>Selected<%End If%> ><%=SectionRS("Office")%></Option>
					
<%					SectionRS.MoveNext
				Loop%>
				</select>

			</td>	
		</tr>
		<tr>
			<td align="right">&nbsp;Employee&nbsp;</td>
			<td>:</td>
			<td>
<%
 				'strsql ="select EmpID, EmpName, MobilePhone from vwPhoneCustomerList Where MobilePhone <> '' order by EmpName"
				strsql ="select distinct EmpID, EmpName from vwPhoneCustomerList Where MobilePhone <> '' order by EmpName"
				set EmpRS = server.createobject("adodb.recordset")
				set EmpRS = BillingCon.execute(strsql)
'				response.write strStr 
%>	
				<Select name="cmbEmp">
					<Option value='X'>--All--</Option>
<%				Do While not EmpRS.eof 
%>
					<Option value='<%=EmpRS("EmpID")%>' <%if trim(EmpID_) = trim(EmpRS("EmpID")) then %>Selected<%End If%> ><%=EmpRS("EmpName")%></Option>
					
<%					EmpRS.MoveNext
				Loop%>
				</select>

			</td>	
			<td align="right">Status&nbsp;</td>
			<td>:</td>
			<td>
				<Select name="cmbStatus">
					<Option value="X" <%if Status ="X" then %>Selected<%End If%>>-All-</Option>
					<Option value="Pending" <%if Status = "Pending" then %>Selected<%End If%>>Pending</Option>
					<Option value="Completed" <%if Status = "Completed" then %>Selected<%End If%>>Completed</Option>
				</Select>&nbsp;
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
'strsql = "Select * From vwMonthlyBilling Where YearP+MonthP>='" & sPeriod & "' and YearP+MonthP<='" & ePeriod & "'"
strsql = "Select * From vwMonthlyBilling Where YearP+MonthP>='" & sPeriod & "' and YearP+MonthP<='" & ePeriod & "' and MobilePhone <> ''"
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

If Status <> "X" then
	strFilter =strFilter & " and Status='" & Status & "'"
End If

strsql = strsql  & strFilter & " Order by EmpName"
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
<!-- <div align="right"><input type="submit" name="btnApproval" value="Approve" /></div> -->
<div align="right"><input type="button" name="btnExport" value="Export to Excel" onClick="javascript:document.location.href('ARBillingReportAllPrint.asp?sMonthP=<%=sMonthP%>&sYearP=<%=sYearP%>&eMonthP=<%=eMonthP%>&eYearP=<%=eYearP%>&Agency=<%=Agency_%>&Section=<%=Section_%>&EmpID=<%=EmpID_%>&Status=<%=Status%>');"/></div>
<table border="1" bordercolor="#EEEEEE" cellpadding="2" cellspacing="0" width="100%"  class="FontText">
    <TR BGCOLOR="#330099" align="center">
         <TD width="3%" align="right"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
         <TD width="15%"><strong><label STYLE=color:#FFFFFF>Employee Name</label></strong></TD>
         <TD width="5%"><strong><label STYLE=color:#FFFFFF>Number</label></strong></TD>
         <TD width="7%"><strong><label STYLE=color:#FFFFFF>Billing Period</label></strong></TD>
         <TD width="8%"><strong><label STYLE=color:#FFFFFF>Section</label></strong></TD>
	 <TD width="8%" ><strong><label STYLE=color:#FFFFFF>Billing Amount (Kn.)</label></strong></TD>
	 <TD width="8%" ><strong><label STYLE=color:#FFFFFF>Personal Amount (Kn.)</label></strong></TD>
<!--         <TD><strong><label STYLE=color:#FFFFFF>Note</label></strong></TD> -->
         <TD width="15%"><strong><label STYLE=color:#FFFFFF>Approved By</label></strong></TD>
         <TD width="15%"><strong><label STYLE=color:#FFFFFF>Status</label></strong></TD>
<!--     <TD><strong><label STYLE=color:#FFFFFF>Note</label></strong></TD> -->
    </TR>
<!--     <tr BGCOLOR="#330099" align="center">
        <TD><strong><label STYLE=color:#FFFFFF>Home Phone</label></strong></TD> 
         <TD><strong><label STYLE=color:#FFFFFF>Office Phone</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Mobile Phone</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Shuttle Bus</label></strong></TD> 
        <TD><strong><label STYLE=color:#FFFFFF>Home Phone</label></strong></TD> 
         <TD><strong><label STYLE=color:#FFFFFF>Office Phone</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Mobile Phone</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Shuttle Bus</label></strong></TD> 
         <TD width="6%"><strong><label STYLE=color:#FFFFFF>Total</label></strong></TD>

    </tr> -->

<% 
   dim no_  
   no_ = 1 + ((PageIndex*intPageSize)-intPageSize)
   Count=1 
   do while not DataRS.eof   and Count<=intPageSize
	   if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
%>
      
	   <TR bgcolor="<%=bg%>">
	        <TD align="right"><%=no_ %>&nbsp;</font></TD>
	        <TD>&nbsp;<%=DataRS("EmpName") %></TD>
		<TD>&nbsp;<%=DataRS("MobilePhone") %></TD>
	        <TD align="right">&nbsp;<%= DataRS("MonthP")%>-<%= DataRS("YearP")%></font>&nbsp;</TD>
	        <TD>&nbsp;<%=DataRS("Office") %> </font></TD>
<!--	        <TD align="right"><%= formatnumber(DataRS("HomePhoneBillRp"),-1) %>&nbsp;</font></TD>-->
<!--		<td align="right">
<%		If CDbl(DataRS("HomePhonePrsBillRp")) > 0 Then %>
			<a href="HomePhoneDetail.asp?HomePhone=<%=DataRS("HomePhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("HomePhoneBillRp"),-1) %></a>
<%		Else %>
			-
<%		End If %>
		&nbsp;</td>
		<td align="right">
<%		If CDbl(DataRS("OfficePhoneBillRp")) > 0 Then %>
			<a href="OfficePhoneDetail.asp?Extension=<%=DataRS("WorkPhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("OfficePhoneBillRp"),-1) %></a>
<%		Else %>
			-
<%		End If %>
		&nbsp;</td> -->
		<td align="right">
<%		If CDbl(DataRS("CellPhoneBillRp")) > 0 Then %>
			<a href="CellPhoneDetail.asp?CellPhone=<%=DataRS("MobilePhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("CellPhoneBillRp"),-1) %></a>
<%		Else %>
			-
<%		End If %>
		&nbsp;</td>
<!--		<td align="right">
<%		If CDbl(DataRS("TotalShuttleBillRp")) > 0 Then %>
			<a href="ShuttleBusBillDetail.asp?Username=<%=DataRS("LoginID") %>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("TotalShuttleBillRp"),-1) %></a>
<%		Else %>
			-
<%		End If %>
		&nbsp;</td> 
		<td align="right">
<%		If CDbl(DataRS("HomePhonePrsBillRp")) > 0 Then %>
			<a href="HomePhoneDetail.asp?HomePhone=<%=DataRS("HomePhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("HomePhonePrsBillRp"),-1) %></a>
<%		Else %>
			-
<%		End If %>
		&nbsp;</td>
		<td align="right">
<%		If CDbl(DataRS("OfficePhonePrsBillRp")) > 0 Then %>
			<a href="OfficePhoneDetail.asp?Extension=<%=DataRS("WorkPhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("OfficePhonePRSBillRp"),-1) %></a>
<%		Else %>
			-
<%		End If %>
		&nbsp;</td> -->
		<td align="right">
<%		If CDbl(DataRS("CellPhonePrsBillRp")) > 0 Then %>
			<a href="CellPhoneDetail.asp?CellPhone=<%=DataRS("MobilePhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("CellPhonePrsBillRp"),-1) %></a>
<%		Else %>
			-
<%		End If %>
		&nbsp;</td>
<!-- 		<td align="right">
<%		If CDbl(DataRS("TotalShuttleBillRp")) > 0 Then %>
			<a href="ShuttleBusBillDetail.asp?Username=<%=DataRS("LoginID") %>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("TotalShuttleBillRp"),-1) %></a>
<%		Else %>
			-
<%		End If %>
		&nbsp;</td>
	        <TD align="right"><%= formatnumber(DataRS("TotalBillingRp"),-1) %>&nbsp;</font></TD> -->
		<TD>&nbsp;<%= DataRS("ApprovalSupervisor") %></font></TD>
		<TD>&nbsp;<%= DataRS("ProgressDesc") %></font></TD>
<!--		<TD>&nbsp;<%= DataRS("Note") %></font></TD> -->
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
				<a href="ARBillingReportAll.asp?PageIndex=<%=PageNo%>&sMonthP=<%=sMonthP%>&sYearP=<%=sYearP%>&eMonthP=<%=eMonthP%>&eYearP=<%=eYearP%>&Agency=<%=Agency_%>&Section=<%=Section_%>&EmpID=<%=EmpID_%>&Status=<%=Status%>"><%=PageNo%></a>&nbsp;
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


