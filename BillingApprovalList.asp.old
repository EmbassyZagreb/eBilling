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
user1_ = user_  'user1_ = right(user_,len(user_)-4)
'response.write user1_ & "<br>"

strSql = "select EmailAddress from vwPhoneCustomerList where LoginID='" & user1_ & "'"
set DataRS = server.createobject("adodb.recordset")
'response.write strsql & "<br>"
set DataRS = BillingCon.execute(strsql)
if not DataRS.eof Then
	SPVEmailAddress_ = DataRS("EmailAddress")
Else
	SPVEmailAddress_ = "X"
end If

curMonth_ = month(date())
curYear_ = year(date())
if len(curMonth_)= 1 then
	curMonth_ = "0" & curMonth_
end if

MonthP = Request.Form("MonthList")
if MonthP ="" then
	MonthP = "XX"
end if

YearP = Request.Form("YearList")
if YearP ="" then
	YearP = curYear_ 
end if

ProgressID = Request("ProgressID")
if ProgressID ="" then
	ProgressID = Request.Form("cmbProgress")
	if ProgressID ="" then
		ProgressID = "2"
	end if
end if
%> 
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">BILLING APPROVAL REQUEST LIST</TD>
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
<form method="post" name="frmSearch" Action="BillingApprovalList.asp">
<table cellspadding="1" cellspacing="0" width="50%" border="1" align="center">
<tr align="Center">
	<td colspan="2" align="center">
		<table  width="100%">
		<tr bgcolor="#000099">
			<td height="25" colspan="7"><strong>&nbsp;<span class="style5">Search &amp; Sort By </span></strong></td>
		</tr>
		<tr>
			<td width="20%">&nbsp;Period&nbsp;</td>				
			<td>:</td>
			<td colspan="2">
				<Select name="MonthList">
					<Option value="XX" <%if MonthP ="XX" then %>Selected<%End If%>>-All-</Option>
					<Option value="01" <%if MonthP ="01" then %>Selected<%End If%>>January</Option>
					<Option value="02" <%if MonthP ="02" then %>Selected<%End If%>>February</Option>
					<Option value="03" <%if MonthP ="03" then %>Selected<%End If%>>March</Option>
					<Option value="04" <%if MonthP ="04" then %>Selected<%End If%>>April</Option>
					<Option value="05" <%if MonthP ="05" then %>Selected<%End If%>>May</Option>
					<Option value="06" <%if MonthP ="06" then %>Selected<%End If%>>June</Option>
					<Option value="07" <%if MonthP ="07" then %>Selected<%End If%>>July</Option>
					<Option value="08" <%if MonthP ="08" then %>Selected<%End If%>>August</Option>
					<Option value="09" <%if MonthP ="09" then %>Selected<%End If%>>September</Option>
					<Option value="10" <%if MonthP ="10" then %>Selected<%End If%>>October</Option>
					<Option value="11" <%if MonthP ="11" then %>Selected<%End If%>>November</Option>
					<Option value="12" <%if MonthP ="12" then %>Selected<%End If%>>December</Option>
				</Select>&nbsp;
<%
				Year_ = Year(Date()) - 1
%>

				<Select name="YearList">
<% 				Do While Year_ <= Year(Date()) %>
				<Option value='<%=Year_%>' <%if trim(Year_) = trim(YearP) then %>Selected<%End If%> ><%=Year_%></Option>		
<% 
			Year_ = Year_ + 1
			Loop %>	
				</Select>										
			</td>
		</tr>
		<tr>
			<td>&nbsp;Approval Status&nbsp;</td>
			<td>:</td>
			<td>
				<Select name="cmbProgress">
					<Option value="X" <%if ProgressID ="X" then %>Selected<%End If%>>-All-</Option>
					<Option value="2" <%if ProgressID = "2" then %>Selected<%End If%>>Pending</Option>
					<Option value="4" <%if ProgressID = "4" then %>Selected<%End If%>>Approved</Option>
				</Select>&nbsp;
			</td>
			<td height="30" align="center">
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



If MonthP="XX" then
	if ProgressID ="X" then
		strsql = "Select * From vwMonthlyBilling Where SupervisorEmail='" & SPVEmailAddress_ & "'"
	else
		strsql = "Select * From vwMonthlyBilling Where SupervisorEmail='" & SPVEmailAddress_ & "' and ProgressID="& ProgressID 
	end if
Else
	if ProgressID ="X" then
		strsql = "Select * From vwMonthlyBilling Where SupervisorEmail='" & SPVEmailAddress_ & "' and MonthP='" & MonthP & "' and YearP='" & YearP & "'"
	else
		strsql = "Select * From vwMonthlyBilling Where SupervisorEmail='" & SPVEmailAddress_ & "' and MonthP='" & MonthP & "' and YearP='" & YearP & "' and ProgressID="& ProgressID 
	end if
End If
'response.write strsql & "<br>"
set DataRS = BillingCon.execute(strsql)
if not DataRS.eof Then
%>
<form name="frmBillingList" Action="BillingApprovalAll.asp" onSubmit="return ValidateForm()">
<table width="100%">
<tr>
	<td class="Hint" align="left">*Click on employee name for approve one by one or hit approve button for approve selected record(s)</td>
	<td align="right"><input type="submit" name="btnApproval" value="Approve" /></td>
</tr>
</table>
<!-- <div align="right"><input type="submit" name="btnApproval" value="Approve" /></div> -->
<table border="1" bordercolor="#EEEEEE" cellpadding="2" cellspacing="0" width="100%"  class="FontText">
    <TR BGCOLOR="#330099" align="center" height=26>
         <TD width="3%" align="right"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
         <TD width="22%"><strong><label STYLE=color:#FFFFFF>Employee Name</label></strong></TD>
         <TD width="10%"><strong><label STYLE=color:#FFFFFF>Number</label></strong></TD>
         <TD width="10%"><strong><label STYLE=color:#FFFFFF>Billing Period</label></strong></TD>
	 <TD width="15%"><strong><label STYLE=color:#FFFFFF>Billing Amount (Kn.)</label></strong></TD>
	 <TD width="15%"><strong><label STYLE=color:#FFFFFF>Personal Amount (Kn.)</label></strong></TD>
<!--     <TD><strong><label STYLE=color:#FFFFFF>Note</label></strong></TD> -->
         <TD Width="22%"><strong><label STYLE=color:#FFFFFF>Status</label></strong></TD>
         <TD Width="3%"><input type="checkbox" name="cbAll" value="true" onclick="checkall(this)" /></TD>
    </TR>
<!--    <tr BGCOLOR="#330099" align="center">
         <TD width="8%"><strong><label STYLE=color:#FFFFFF>Home Phone</label></strong></TD>
         <TD width="8%"><strong><label STYLE=color:#FFFFFF>Office Phone</label></strong></TD>
         <TD width="8%"><strong><label STYLE=color:#FFFFFF>Mobile Phone</label></strong></TD>
         <TD width="10%"><strong><label STYLE=color:#FFFFFF>Shuttle Bus</label></strong></TD>
         <TD width="10%"><strong><label STYLE=color:#FFFFFF>Home Phone</label></strong></TD>
         <TD width="8%"><strong><label STYLE=color:#FFFFFF>Office Phone</label></strong></TD>
         <TD width="8%"><strong><label STYLE=color:#FFFFFF>Mobile Phone</label></strong></TD>
         <TD width="10%"><strong><label STYLE=color:#FFFFFF>Shuttle Bus</label></strong></TD>
         <TD width="8%"><strong><label STYLE=color:#FFFFFF>Total</label></strong></TD>
    </tr> -->
<% 
   dim no_  
   no_ = 1 
   do while not DataRS.eof  
	   if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
%>
      
	   <TR bgcolor="<%=bg%>">
	        <TD align="right"><%=no_ %>&nbsp;</font></TD>
	        <TD>&nbsp;
		<%if DataRS("ProgressID")=2 then %>
			<A HREF="BillingApproval.asp?EmpID=<%= DataRS("EmpID")%>&Month=<%=DataRS("MonthP")%>&Year=<%=DataRS("YearP")%>&LoginID=<%=DataRS("LoginID")%>"><%=DataRS("EmpName") %></A>
		<%else%>
			<%=DataRS("EmpName") %>
		<%end if%>
		</TD>
	        <TD align="right">&nbsp;<%= DataRS("MobilePhone")%></font>&nbsp;</TD>
	        <TD align="right">&nbsp;<%= DataRS("MonthP")%>-<%= DataRS("YearP")%></font>&nbsp;</TD>
	        <TD align="right"><%= formatnumber(DataRS("CellPhoneBillRp"),-1) %>&nbsp;</font></TD>
	        <TD align="right"><%= formatnumber(DataRS("CellPhonePrsBillRp"),-1) %>&nbsp;</font></TD>
		<TD>&nbsp;<%= DataRS("ProgressDesc") %></font></TD>
		<td align="center">
			<%if DataRS("ProgressID")=2 then%>
				<Input type="Checkbox" name="cbApproval" Value='<%=DataRS("EmpID")%><%=DataRS("MonthP")%><%=DataRS("YearP")%>'>
			<%else%>
				&nbsp;
			<%end if%>
		</td>
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
		<td align="center"><b>There are no requests that require your approval</b></td>
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


