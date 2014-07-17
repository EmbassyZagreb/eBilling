<%@ Language=VBScript%>
<% ' VI 6.0 Scripting Object Model Enabled %> 
<html>
<head>

<!--#include file="connect.inc" -->


-->
</style>
<script language="JavaScript" src="calendar.js"></script>
<script language="JavaScript">
function validate_form()
{
	valid = true;
	msg = ""

	if (document.frmGenerate.cmbEmp.value == "")
	{
		msg = "Please select Employee !!!\n"
		valid = false;
	}


	if (valid == false)
	{
		alert(msg)
	}
	return valid;
}	
</script>
<%

StartDate_ = Request.Form("txtStartDate")
if StartDate_ ="" then
	StartDate_ = Date()
end if

EndDate_ = Request.Form("txtEndDate")
if EndDate_ ="" then
	EndDate_ = Date()+7
end if

EmpID_ = Trim(Request.Form("cmbEmp"))

Mon_ = Trim(Request.Form("cbMon"))
Tue_ = Trim(Request.Form("cbTue"))
Wed_ = Trim(Request.Form("cbWed"))
Thu_ = Trim(Request.Form("cbThu"))
Fri_ = Trim(Request.Form("cbFri"))
Sat_ = Trim(Request.Form("cbSat"))
Sun_ = Trim(Request.Form("cbSun"))
Time_ = Trim(Request.Form("cmbTime"))
%>

<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">Generate Shuttle Bus Schedule(s)</TD>
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
Dim user_ , user1_, ShowData_

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

If (UserRole_ <> "") Then
	if (request.form("btnSubmit")="Generate Schedule") then 
		strsql = "exec spGenerateShuttleSchedule '" & EmpID_ & "','" & StartDate_ & "','" & EndDate_ & "','" & Mon_ & "','" & Tue_ & "','" & Wed_ & "','" & Thu_ & "','" & Fri_ & "','" & Sat_ & "','" & Sun_ & "','" & Time_ & "'"
		'response.write strsql & "<br>"
		BillingCon.execute(strsql)
		'response.write strsql & "<br>"
	end if
	If (request.form("btnSubmit")="Generate Schedule") or (request.form("btnSubmit")="Show Data") then 
		strsql = "Select * from vwShuttleSchedule Where EmpId='" & EmpID_ & "' And ShuttleDate >='" & StartDate_ & "' And ShuttleDate <='" & EndDate_ & "'"
		'response.write strsql & "<br>"
		set DataRS = server.createobject("adodb.recordset")
		set DataRS = BillingCon.execute(strsql)
		ShowData_ = "Yes"
	end if
%>
	<form method="post" name="frmGenerate" onSubmit="return validate_form();">
	<table align="center" cellpadding="1" cellspacing="1" width="60%" border="1">
	<tr>
		<td>
		<table align="center" cellpadding="2" cellspacing="0" width="100%">
		<tr bgcolor="#000099">
			<td height="25" colspan="4"><strong>&nbsp;<span class="style5">Schedule Parameters</span></strong></td>
		</tr>
		<tr>
			<td width="20%">Employee</td>
			<td width="1%">:</td>
			<td colspan="2">
				<Select name="cmbEmp">
					<Option name="">-- Select --</option>
<%
'			strsql = "Select EmpID, Case When FirstName='' Then LastName Else LastName+', '+FirstName End As EmpName From vwPhoneCustomerList Where Type='AMER' order by EmpName "
			strsql = "Select EmpID, EmpName From vwPhoneCustomerList order by EmpName "
			set EmpRS = server.createobject("adodb.recordset")
			'response.write strsql & "<br>"
			set EmpRS = BillingCon.execute(strsql)
			Do While not EmpRS.eof 
%>
				<Option value='<%=EmpRS("EmpID")%>'  <%If trim(EmpID_) = trim(EmpRS("EmpID")) Then%>Selected<%End If%> ><%=EmpRS("EmpName")%> </option>
			
<%
				EmpRS.MoveNext
			Loop
		%>
			</td>
		</tr>
		<tr>
			<td width="15%">Date Range</td>
			<td width="1%">:</td>
			<td colspan="2">
				<input name="txtStartDate" type="Input" size="10" maxlength="10" value='<%=StartDate_ %>' >
			       	<a href="javascript:cal0.popup();"><img src="images/calendar.gif" width="34" height="18" border="0" alt="Calendar"></a>&nbsp;to&nbsp;
				<input name="txtEndDate" type="Input" size="10" maxlength="10" Value='<%=EndDate_ %>' >
			       	<a href="javascript:cal1.popup();"><img src="images/calendar.gif" width="34" height="18" border="0" alt="Calendar"></a>
			</td>
		</tr>
		<tr>
			<td width="15%">Selected Day</td>
			<td width="1%">:</td>
			<td colspan="2">
				<input type="checkbox" name="cbMon" value="Y" Checked >&nbsp;Mon&nbsp;</input>
				<input type="checkbox" name="cbTue" value="Y" Checked >&nbsp;Tue&nbsp;</input>
				<input type="checkbox" name="cbWed" value="Y" Checked >&nbsp;Wed&nbsp;</input>
				<input type="checkbox" name="cbThu" value="Y" Checked >&nbsp;Thu&nbsp;</input>
				<input type="checkbox" name="cbFri" value="Y" Checked>&nbsp;Fri&nbsp;</input>
				<input type="checkbox" name="cbSat" value="Y" >&nbsp;Sat&nbsp;</input>
				<input type="checkbox" name="cbSun" value="Y" >&nbsp;Sun&nbsp;</input>
			</td>
		<tr>
		<tr>
			<td width="15%">Selected Time</td>
			<td width="1%">:</td>
			<td colspan="2">
				<select name="cmbTime">
					<option value="Both">Both</option>
					<option value="AM">AM</option>
					<option value="PM">PM</option>
				</Select>
			</td>
		</tr>
		<tr>
			<td colspan="4"><br></td>
		</tr>
		<tr>
			<td colspan="2"></td>
			<td><input type="submit" name="btnSUbmit" value="Show Data"></td>
			<td><input type="submit" name="btnSUbmit" value="Generate Schedule" onClick="return confirm('Are you sure generate schedule for this employee? Existing data will be replaced.')" ></td>	
		</tr>		
		</table>
		</td>
	</tr>
	</table>
		<script language="JavaScript">
	    	    var cal0 = new calendar1(document.forms['frmGenerate'].elements['txtStartDate']);
			cal0.year_scroll = true;
			cal0.time_comp = false;
	    	    var cal1 = new calendar1(document.forms['frmGenerate'].elements['txtEndDate']);
			cal1.year_scroll = true;
			cal1.time_comp = false;
		</script>
	</form>
	<form method="post" name="frmSchedule">
	<table align="center" cellpadding="1" cellspacing="0" width="40%" border="1" bordercolor="black"> 
	<TR BGCOLOR="#000099" align="center">
		<TD width="3%"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
		<TD width="10%"><strong><label STYLE=color:#FFFFFF>Shuttle Date</label></strong></TD>
		<TD width="6%"><strong><label STYLE=color:#FFFFFF>AM</label></strong></TD>
		<TD width="6%"><strong><label STYLE=color:#FFFFFF>PM</label></strong></TD>
		<TD width="10%"><strong><label STYLE=color:#FFFFFF>Action</label></strong></TD>
	</TR>    
<% 
		
		dim no_  
		no_ = 1 
		if ShowData_ = "Yes"  then
		do while not DataRS.eof
	   	if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
%>
	 	<TR bgcolor="<%=bg%>">
			<td align="right"><%=No_%>&nbsp;</td>
	       		<td><FONT color=#330099 size=2>&nbsp;<%=DataRS("ShuttleDate")%></font></td> 
		       	<td align="right"><FONT color=#330099 size=2><%=DataRS("AM")%>&nbsp;</font></td> 
        		<td align="right"><FONT color=#330099 size=2><%=DataRS("PM")%>&nbsp;</font></td> 
	        	<td>&nbsp;
			<A HREF="ShuttleBusScheduleEdit.asp?State=E&EmpID=<%=DataRS("EmpID")%>&ShuttleDate=<%=DataRS("ShuttleDate")%>" target="_new">Edit</A>
			&nbsp;
			<A HREF="ShuttleBusScheduleDeleteConfirm.asp?State=D&EmpID=<%=DataRS("EmpID")%>&ShuttleDate=<%=DataRS("ShuttleDate")%>" >Delete</A>
			</td> 
	  	 </TR>
<%   
 		DataRS.movenext
   		no_ = no_ + 1
		loop
		end if
%>
	</table>

	</form>	
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


