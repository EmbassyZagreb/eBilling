<%@ Language=VBScript%>
<% ' VI 6.0 Scripting Object Model Enabled %>
 
 
<html>
<head>

<META name=VI60_defaultClientScript content=VBScript>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<!--#include file="connect.inc" -->


<%
Dim user_ , user1_

user_ = request.servervariables("remote_user")
user1_ = right(user_,len(user_)-4)
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

'curMonth_ = month(date())
'curYear_ = year(date())
if len(curMonth_)= 1 then
	curMonth_ = "0" & curMonth_
end if

sMonthP = request("sMonthP")

if sMonthP = "" Then sMonthP = Request.Form("sMonthList")
if sMonthP = "" then
	sMonthP = "01"
end if
'response.write sMonthP

sYearP = request("sYearP")
if sYearP ="" Then sYearP = Request.Form("sYearList")
if sYearP ="" then
	sYearP = curYear_ -1
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
  	<TD COLSPAN="4" ALIGN="center" Class="title">BILLING TRACKING</TD>
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

<form method="post" name="frmSearch" Action="MonthlyBillingTracking.asp">
<table cellspadding="1" cellspacing="0" width="65%" border="1" align="center">
<tr align="Center">
	<td colspan="2" align="center">
		<table  width="100%">
		<tr bgcolor="#000099">
			<td height="25" colspan="4"><strong>&nbsp;<span class="style5">Search By</span></strong></td>
		</tr>
		<tr>
			<td width="20%">&nbsp;Period&nbsp;</td>				
			<td>:</td>
			<td colspan="2">
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
			<td>&nbsp;Approval Status&nbsp;</td>
			<td>:</td>
			<td colspan="2">
				<Select name="cmbStatus">
					<Option value="X" <%if Status ="X" then %>Selected<%End If%>>-All-</Option>
					<Option value="Pending" <%if Status = "Pending" then %>Selected<%End If%>>Pending</Option>
					<Option value="Completed" <%if Status = "Completed" then %>Selected<%End If%>>Completed</Option>
				</Select>&nbsp;
			</td>
		</tr>
		<tr>
			<td height="30" align="center" colspan="4">
				<input type="Button" name="btnBack" value="Back" onClick="Javascript:document.location.href('Default.asp');">
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

'strsql = "Exec spApprovalList '" & user1_ & "','" & MonthP & "','" & YearP & "'"
If MonthP="XX" then
	if Status ="X" then
		strsql = "Select * From vwMonthlyBilling Where LoginID ='" & user1_& "'"
	else
		strsql = "Select * From vwMonthlyBilling Where LoginID ='" & user1_ & "' and Status='" & Status & "'"
	end if
Else
	if Status ="X" then
		strsql = "Select * From vwMonthlyBilling Where LoginID ='" & user1_ & "' and YearP+MonthP>='" & sPeriod & "' and YearP+MonthP<='" & ePeriod & "'"
	else
		strsql = "Select * From vwMonthlyBilling Where LoginID='" & user1_ & "' and YearP+MonthP>='" & sPeriod & "' and YearP+MonthP<='" & ePeriod & "' and Status='" & Status & "'"
	end if
End If
'response.write strsql & "<br>"
set DataRS = BillingCon.execute(strsql)
if not DataRS.eof Then
%>
<table cellspadding="1" cellspacing="0" width="65%" bgColor="white">  
<tr>
	<td align="Left">
	<table cellspadding="1" border="1" bordercolor="black" cellspacing="0" width="100%" bgColor="white" border="0">  
	<tr align="center" height=26>
		<td width=20%><b>Action</b></td>
		<td width=20%><b>Billing Period</b></td>
		<td width=20%><b>Status</b></td>
		<td width=20%><b>Total Bill Amount (Kn.)</b></td>
		<td width=20%><b>Personal Amount Due (Kn.)</b></td>
	</tr>

<% 
   PrevEmpName_ =""
   do while not DataRS.eof
	   if bg="#dddddd" then bg="ffffff" else bg="#dddddd" 
	    if PrevEmpName_ <> DataRS("EmpName") Then
		SubTotalBill_ = 0
		SubTotalPrs_ = 0
%>
		<tr BGCOLOR="#999999" height=26>
			<td colspan="2" style="border: none;"><FONT color=#FFFFFF><b>&nbsp;Employee Name : <%=DataRS("EmpName") %></b></font></td>
			<td colspan="3" style="border: none;" align="right"><FONT color=#FFFFFF><b>Phone Number : <%=DataRS("MobilePhone") %>&nbsp;</b></font></td>

		</tr>		
<%
	    end if
%>
      
	   <TR bgcolor="<%=bg%>">
	        <TD>&nbsp;<a href="MonthlyBilling.asp?CellPhone=<%=DataRS("MobilePhone") %>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>">Review your bill</a></TD>
	        <TD align="right">&nbsp;<%= DataRS("MonthP")%>-<%= DataRS("YearP")%></font>&nbsp;</TD>
	        <TD align="right"><%= DataRS("ProgressDesc") %>&nbsp;</font></TD>
	        <TD align="right"><%= formatnumber(DataRS("TotalBillingRp"),-1) %>&nbsp;</font></TD>
		<TD align="right">&nbsp;<%= formatnumber(DataRS("TotalBillingAmountPrsRp"),-1) %>&nbsp;</font></TD>
	   </TR>

<%   
		PrevEmpName_ = DataRS("EmpName")
	   DataRS.movenext
   loop 
%>
			
	</table>
	</td>
</tr>
<% else %>
<tr>
	<td align="Left">
	<table cellspadding="1" border="1" bordercolor="black" cellspacing="0" width="65%" bgColor="white" border="0">  
	<tr align="center" BGCOLOR="#999999" height=26>
		<td><FONT color=#FFFFFF><b>There are no bills for your cell phone(s).</b></font></td>
	</tr>




</table>


<% end if %>
</body> 

</html>


