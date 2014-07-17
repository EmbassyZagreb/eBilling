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

curMonth_ = month(date())
curYear_ = year(date())
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
<BODY BGCOLOR="#FFFFFF" TEXT="#000000" LINK="#0080FF" ALINK="0080FF" VLINK="#0055AA" MARGINWIDTH="8" MARGINHEIGHT="0" LEFTMARGIN="8" TOPMARGIN="0">
<CENTER>
<%

dim rs 
dim strsql
dim tombol
dim hlm
%>
<form method="post" name="frmSearch" Action="MonthlyBillingTrackingAddOn.asp">
<table cellspadding="1" cellspacing="0" width="60%" border="1" align="center">
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
<form name="frmTracking">
<!-- <div align="right"><input type="submit" name="btnApproval" value="Approve" /></div> -->
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" width="100%"  class="FontText">
    <TR BGCOLOR="#330099" align="center">
         <TD width="15px" align="right"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Employee Name</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Billing Period</label></strong></TD>	
         <TD><strong><label STYLE=color:#FFFFFF>Agency / Office</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Phone Ext.</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Total Cost</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>Note</label></strong></TD>
         <TD>
		<strong><label STYLE=color:#FFFFFF>Status</label></strong>
	 </TD>
    </TR>    

<% 
   dim no_  
   no_ = 1 
   do while not DataRS.eof  
	   if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
%>
      
	   <TR bgcolor="<%=bg%>">
	        <TD align="right"><FONT color=#330099 size=2><%=no_ %>&nbsp;</font></TD>
	        <TD>&nbsp;
		<%if DataRS("ProgressID")=2 then %>
			<A HREF="MonthlyBillingView.asp?Month=<%=DataRS("MonthP")%>&Year=<%=DataRS("YearP")%>&LoginID=<%=DataRS("LoginID")%>" target="_blank"><%=DataRS("EmpName") %></A>
		<%elseif DataRS("ProgressID")=1 then %>
			<A HREF="MonthlyBilling.asp?Month=<%=DataRS("MonthP")%>&Year=<%=DataRS("YearP")%>"><%=DataRS("EmpName") %></A>
		<%else%>
			<%=DataRS("EmpName") %>
		<%end if%>
		</TD>
	        <TD align="right"><FONT color=#330099 size=2>&nbsp;<%= DataRS("MonthP")%>-<%= DataRS("YearP")%></font>&nbsp;</TD>
	        <TD><FONT color=#330099 size=2>&nbsp;<%= DataRS("Office") %></font>   </TD>
	        <TD><FONT color=#330099 size=2>&nbsp;<%= DataRS("WorkPhone") %></font></TD>
	        <TD align="right"><FONT color=#330099 size=2>Kn. &nbsp;<%= formatnumber(DataRS("TotalBillingRp"),-1) %>&nbsp;</font></TD>
	        <TD><FONT color=#330099 size=2>&nbsp;<%= DataRS("Notes") %></font></TD>
		<TD><FONT color=#330099 size=2>&nbsp;<%= DataRS("ProgressDesc") %></font></TD>
	   </TR>

<%   
	   DataRS.movenext
	   no_ = no_ + 1 
   loop %>
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
		<td align="center">there is no request that need you to approve.</td>
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


