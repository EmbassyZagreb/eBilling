<%@ Language=VBScript%>
<% ' VI 6.0 Scripting Object Model Enabled %> 
<html>
<head>

<META name=VI60_defaultClientScript content=VBScript>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<!--#include file="connect.inc" -->


<script language="JavaScript" src="calendar.js"></script>
<script language="JavaScript">
function ClearFilter()
{
	document.forms['frmSearch'].elements['cmbEmp'].value ="X";
	document.forms['frmSearch'].elements['cmbOfficeSection'].value ="X";
	document.forms['frmSearch'].elements['cmbStatus'].value =4;

}

</script>

<%
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


dim intPageSize,PageIndex,TotalPages 
dim RecordCount,RecordNumber,Count 
intpageSize=20 
PageIndex=request("PageIndex")


EmpID_ = request("EmpID")
if EmpID_ = "" Then EmpID_ = Request.Form("cmbEmp")
if EmpID_ = "" then
	EmpID_ = "X"
end if

'response.write EmpName_

OfficeSection_ = Trim(request("OfficeSection"))
if OfficeSection_ ="" then
	OfficeSection_ = Trim(Request.Form("cmbOfficeSection"))
End If
'response.write OfficeSection_ 

Outstanding_ = Trim(request("Outstanding"))
if Outstanding_ ="" then
	Outstanding_ = Trim(Request.Form("txtOutstanding"))
End If
'response.write Outstanding_
if Outstanding_ = "" then
	Outstanding_ = 0
end if

Status_ = Trim(Request.Form("cmbStatus"))
if Status_ ="" then
	Status_ = Trim(request("Status"))
End If
'response.write Status_

if Status_ = "" then
	Status_ = 4
end if
%>


<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
   <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">PAYMENT LIST</TD>
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
%>

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
 
If (UserRole_ = "Admin") or (UserRole_ = "Voucher") or (UserRole_ = "FMC") or (UserRole_ = "Cashier") Then
	sPeriod = sYearP&sMonthP
	ePeriod = eYearP&eMonthP
	strsql = "spGetPaymentReceipt '" & sPeriod & "','" & ePeriod & "','" & EmpID_ & "'," & Outstanding_ & ",'" & OfficeSection_  & "'," & Status_ 
	set DataRS = server.createobject("adodb.recordset")
	'response.write strsql & "<br>"
	DataRS.CursorLocation = 3
	DataRS.open strsql,BillingCon

	if ((PageIndex ="") or (request.form("btnSearch")="Search")) then PageIndex=1 
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

%>
	<form method="post" name="frmSearch" action="PaymentReceiptList.asp">
	<table align="center" cellpadding="1" cellspacing="0" width="70%">
	<tr bgcolor="#000099">
		<td height="25" colspan="4"><strong>&nbsp;<span class="style5">Search</span></strong></td>
	</tr>
	<tr>
		<td width="20%">Period :</td>				
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
	</tr>			
	<tr>
		<td>Employee Name :</td>
<!--		<td><input name="txtEmpName" type="Input" size="30" Value=<%=EmpName_%>></td> -->
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
%>
					<Option value='<%=EmpRS("EmpID")%>' <%if trim(EmpID_) = trim(EmpRS("EmpID")) then %>Selected<%End If%> ><%=EmpRS("EmpName") %></Option>
					
<%					EmpRS.MoveNext
				Loop%>
				</select>

		</td>
		<td>Office Section :</td>
		<td>
			<select name="cmbOfficeSection">
				<option value="X">--All--</option>
				<% Dim OfficeRS
				    strsql ="select distinct Office from vwPhoneCustomerList Where Office<>'' order by Office"
				    set OfficeRS = server.createobject("adodb.recordset")    
	   			    set OfficeRS = BillingCon.execute(strsql) 
				    
				    Do while not OfficeRS.eof
%>
					<option value="<%=OfficeRS("Office")%>" <%if OfficeSection_ = OfficeRS("Office") then%>Selected<%End If%>><%=OfficeRS("Office")%></option>
<%
					OfficeRS.MoveNext
				    loop
				%>
			</select>
		</td>
	</tr>
	<tr>
		<td>Outstanding Payment :</td>
		<td><input name="txtOutstanding" type="Input" size="3" Value='<%=Outstanding_%>'>&nbsp;day(s)</td>
		<td>Billing Status :</td>
<!--

		<td>
			<select name="cmbStatus">
				<option value="X" <%if Status_ ="X" then%>Selected<%End If%>>--All--</option>
				<option value="P" <%if Status_ ="P" then%>Selected<%End If%>>Pending</option>
				<option value="F" <%if Status_ ="F" then%>Selected<%End If%>>Paid</option>
			</select>			
		</td>
-->
		<td>
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
		<td colspan="2"></td>
		<td colspan="2">
			<input type="submit" name="btnSearch" value="Search">
			&nbsp;&nbsp;<input type="button" name="btnClear" value="Reset filter" onclick="javascript:ClearFilter();">
		</td>
	</tr>
	<tr>
		<td colspan="4"><hr></td>
	</tr>
</table>
</form>
<%
if not DataRS.eof Then
%>
<!-- <div align="right"><input type="submit" name="btnApproval" value="Approve" /></div> -->
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" width="100%"  class="FontText">
    <TR BGCOLOR="#330099" align="center">
         <TD rowspan="2" width="3%" align="right"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
         <TD rowspan="2" width="15%"><strong><label STYLE=color:#FFFFFF>Employee Name</label></strong></TD>
         <TD rowspan="2" width="10%"><strong><label STYLE=color:#FFFFFF>Billing Period</label></strong></TD>
         <TD rowspan="2" width="8%"><strong><label STYLE=color:#FFFFFF>Section</label></strong></TD>
	 <TD colspan="5"><strong><label STYLE=color:#FFFFFF>Billing Amount (Kn.)</label></strong></TD>
         <TD rowspan="2"><strong><label STYLE=color:#FFFFFF>Paid Amount</label></strong></TD>
         <TD rowspan="2">
		<strong><label STYLE=color:#FFFFFF>Aging</label></strong>
	 </TD>
         <TD rowspan="2">
		<strong><label STYLE=color:#FFFFFF>Billing Status</label></strong>
	 </TD>
         <TD rowspan="2"width="5%">
		<strong><label STYLE=color:#FFFFFF>Action</label></strong>
	 </TD>
    </TR>
    <tr BGCOLOR="#330099" align="center">
         <TD width="9%"><strong><label STYLE=color:#FFFFFF>Home Phone</label></strong></TD>
         <TD width="9%"><strong><label STYLE=color:#FFFFFF>Office Phone</label></strong></TD>
         <TD width="9%"><strong><label STYLE=color:#FFFFFF>Mobile Phone</label></strong></TD>
         <TD width="9%"><strong><label STYLE=color:#FFFFFF>Shuttle Bus</label></strong></TD>
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
	        <TD align="right"><FONT color=#330099 size=2><%=no_ %>&nbsp;</font></TD>
	        <TD>&nbsp;<%=DataRS("EmpName") %></TD>
	        <TD align="right"><FONT color=#330099 size=2>&nbsp;<%= DataRS("MonthP")%>-<%= DataRS("YearP")%></font>&nbsp;</TD>
	        <TD><FONT color=#330099 size=2>&nbsp;<%=DataRS("Office") %> </font></TD>
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
<!--        <TD align="right"><FONT color=#330099 size=2><%= formatnumber(DataRS("OfficePhonePrsBillRp"),-1) %>&nbsp;</font></TD> -->
<!--	        <TD align="right"><FONT color=#330099 size=2>-&nbsp;</font></TD> -->
<!--	        <TD align="right"><FONT color=#330099 size=2><%= formatnumber(DataRS("TotalShuttleBillRp"),-1) %>&nbsp;</font></TD> -->
		<td align="right">
<%		If CDbl(DataRS("CellPhonePrsBillRp")) > 0 Then %>
			<a href="CellPhoneDetail.asp?CellPhone=<%=DataRS("MobilePhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>&AlternateEmailFlag=<%=dataRS("AlternateEmailFlag")%>" target="_blank"><%= formatnumber(DataRS("CellPhonePrsBillRp"),-1) %></a>
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
	        <TD align="right"><FONT color=#330099 size=2><%= formatnumber(DataRS("TotalBillingRp"),-1) %>&nbsp;</font></TD>
		<TD>
<%	If cdbl(DataRS("PaidAmountRp"))>0 Then %>
			<A HREF="PaymentReceiptDetail.asp?EmpId=<%=DataRS("EmpId")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("PaidAmountRp"),-1) %></A>
<%	Else %>
			- &nbsp;
<%	End If%>
		</TD>
		<TD><FONT color=#330099 size=2>&nbsp;<%= DataRS("Aging") %></font></TD>
		<TD><FONT color=#330099 size=2>&nbsp;<%= DataRS("Status") %></font></TD>
		<TD align="left">&nbsp;
<%	If (DataRS("ProgressID")=4) or DataRS("ProgressID")=8 or (DataRS("ProgressID")=5) Then %>		
			<A HREF="PaymentRecieptEntry.asp?EmpId=<%=DataRS("EmpId")%>&PageIndex=<%=PageIndex%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>&EmpName=<%=EmpName_%>&OfficeSection=<%=OfficeSection_%>&Outstanding=<%=Outstanding_%>&Status=<%=Status_ %>&AlternateEmailFlag=<%=dataRS("AlternateEmailFlag")%>">Entry</A>
<%	ElseIf (DataRS("ProgressID")=6) Then %>
			<a href="PaymentReceiptDetail.asp?EmpId=<%=DataRS("EmpId")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank">View</a>
<%	End If%>
		</TD>
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
				<a href="PaymentReceiptList.asp?PageIndex=<%=PageNo%>&sMonthP=<%=sMonthP%>&sYearP=<%=sYearP%>&eMonthP=<%=eMonthP%>&eYearP=<%=eYearP%>&EmpID=<%=EmpID_%>&OfficeSection=<%=OfficeSection_%>&Outstanding=<%=Outstanding_%>&Status=<%=Status_ %>"><%=PageNo%></a>&nbsp;
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
			href="http://zagrebws03.eur.state.sbu/WebPASS/eservices/MainPage.asp">Submit Request </a> or contact Zagreb ISC Helpdesk at ext.3333.</td>
		</tr>
	</table>
<% end if %>
</body> 

</html>


