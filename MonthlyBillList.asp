<%@ Language=VBScript%>
<% ' VI 6.0 Scripting Object Model Enabled %>
 
 
<html>
<head>

<META name=VI60_defaultClientScript content=VBScript>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<!--#include file="connect.inc" -->

<script type="text/javascript">
function ValidateForm()
{
	valid = true;
	nRec = 0;
	for (var x=0; x<frmData.elements.length; x++)
	{	
		cbElement = frmData.elements[x]
		if ((cbElement.checked) && (cbElement.name=="cbPrint"))
		{
			nRec++;
		}
	}
	if (nRec == 0)
	{
		alert("Please select data that you want to print !!!");
		valid = false;
	}
	return valid;
}


function checkall(obj)
{
	for (var x=0; x<frmData.elements.length; x++)
	{
		cbElement = frmData.elements[x]
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

'user1_ = "martinwc"
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

%>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">MONTHLY BILL PRINT</TD>
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
%>
<form method="post" name="frmSearch" Action="MonthlyBillList.asp">
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
			<td colspan="3">
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
			<td>
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
strsql = "Select * From vwMonthlyBilling Where YearP+MonthP>='" & sPeriod & "' and YearP+MonthP<='" & ePeriod & "' and LoginId='" & user1_ & "'"

'response.write strsql & "<br>"

set DataRS = server.createobject("adodb.recordset")
DataRS.CursorLocation = 3
DataRS.Open strsql,BillingCon
'set DataRS=BillingCon.execute(strsql)

dim intPageSize,PageIndex,TotalPages 
dim RecordCount,RecordNumber,Count 
intpageSize=50 
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
<form method="post" name="frmData" Action="MonthlyBillListPrint.asp" onsubmit="return ValidateForm()">
<table width="100%">
<tr>
<!--	<td align="left"><input type="button" name="btnExport" value="Export to Excel" onClick="javascript:document.location.href('MonthlyBillListPrint.asp?sMonthP=<%=sMonthP%>&sYearP=<%=sYearP%>&eMonthP=<%=eMonthP%>&eYearP=<%=eYearP%>&EmpID=<%=DataRS("EmpID")%>');"/></td> -->
<!--	<td align="right"><input type="submit" name="btnPrint" value="Print" /></td> -->
</tr>
</table>
<table border="1" bordercolor="#EEEEEE" cellpadding="2" cellspacing="0" width="100%"  class="FontText">
    <TR BGCOLOR="#330099" align="center">
         <TD rowspan="2" width="3%" align="right"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
         <TD rowspan="2" width="15%"><strong><label STYLE=color:#FFFFFF>Employee Name</label></strong></TD>
         <TD rowspan="2" width="10%"><strong><label STYLE=color:#FFFFFF>Billing Period</label></strong></TD>
         <TD rowspan="2" width="8%"><strong><label STYLE=color:#FFFFFF>Section</label></strong></TD>
	 <TD colspan="5"><strong><label STYLE=color:#FFFFFF>Billing Amount (Kn.)</label></strong></TD>
<!--
         <TD rowspan="2"><strong><label STYLE=color:#FFFFFF>Paid Date</label></strong></TD>
         <TD rowspan="2"><strong><label STYLE=color:#FFFFFF>Receipt No.</label></strong></TD>		
-->
<!--
	 <TD rowspan="2" Width="3%">
		<input type="checkbox" name="cbAll" value="true" onclick="checkall(this)" />
	 </TD>
-->
    </TR>
    <tr BGCOLOR="#330099" align="center">
         <TD width="10%"><strong><label STYLE=color:#FFFFFF>Home Phone</label></strong></TD>
         <TD width="10%"><strong><label STYLE=color:#FFFFFF>Office Phone</label></strong></TD>
         <TD width="10%"><strong><label STYLE=color:#FFFFFF>Mobile Phone</label></strong></TD>
         <TD width="10%"><strong><label STYLE=color:#FFFFFF>Shuttle Bus</label></strong></TD>
         <TD width="8%"><strong><label STYLE=color:#FFFFFF>Total</label></strong></TD>
    </tr>

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
	        <TD align="right">&nbsp;<%= DataRS("MonthP")%>-<%= DataRS("YearP")%></font>&nbsp;</TD>
	        <TD>&nbsp;<%=DataRS("Office") %> </font></TD>
<!--	        <TD align="right"><%= formatnumber(DataRS("HomePhoneBillRp"),-1) %>&nbsp;</font></TD>-->
		<td align="right">
<%		If CDbl(DataRS("HomePhonePrsBillRp")) > 0 Then %>
			<a href="HomePhoneDetail.asp?HomePhone=<%=DataRS("HomePhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("HomePhonePrsBillRp"),-1) %></a>
<%		Else %>
			-
<%		End If %>
		&nbsp;</td>
		<td align="right">
<%		If CDbl(DataRS("OfficePhonePrsBillRp")) > 0 Then %>
			<a href="OfficePhoneDetail.asp?Extension=<%=DataRS("WorkPhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("OfficePhonePrsBillRp"),-1) %></a>
<%		Else %>
			-
<%		End If %>
		&nbsp;</td>
		<td align="right">
<%		If CDbl(DataRS("CellPhonePrsBillRp")) > 0 Then %>
			<a href="CellPhoneDetail.asp?CellPhone=<%=DataRS("MobilePhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("CellPhonePrsBillRp"),-1) %></a>
<%		Else %>
			-
<%		End If %>
		&nbsp;</td>
		<td align="right">
<%		If CDbl(DataRS("TotalShuttleBillRp")) > 0 Then %>
			<a href="ShuttleBusBillDetail.asp?Username=<%=DataRS("LoginID") %>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("TotalShuttleBillRp"),-1) %></a>
<%		Else %>
			-
<%		End If %>
		&nbsp;</td>
	        <TD align="right"><%= formatnumber(DataRS("TotalBillingRp"),-1) %>&nbsp;</font></TD>
<!--		<TD><FONT color=#330099 size=2>&nbsp;<%= DataRS("PaidDate") %></font></TD>
		<TD><FONT color=#330099 size=2>&nbsp;<%= DataRS("ReceiptNo") %></font></TD>
-->
<!--
		<td align="center">
			<Input type="Checkbox" name="cbPrint" Value='<%=DataRS("EmpID")%><%=DataRS("MonthP")%><%=DataRS("YearP")%>'>
		</td>
-->
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
				<a href="ARPaymentReport.asp?PageIndex=<%=PageNo%>&sMonthP=<%=sMonthP%>&sYearP=<%=sYearP%>&eMonthP=<%=eMonthP%>&eYearP=<%=eYearP%>"><%=PageNo%></a>&nbsp;
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
</body> 

</html>


