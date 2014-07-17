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

.style1 {color: #990000}
.style4 {color: #FFFFFF; font-weight: bold;}
.style5 {color: #FFFFFF;}
.style6 { font-size: 10px;}
}
-->
</style>
<script language="JavaScript">
function validate_form()
{
	valid = true;
	msg = ""
/*
	if (document.frmGenerate.cmbEmp.value == "")
	{
		msg = "Please select Employee.\n"
		valid = false;
	}
*/
	if (document.frmGenerate.cbMonth.value == "")
	{
		msg = msg + "Please select month of Billing Period.\n"
		valid = false;
	}

	if (document.frmGenerate.cbYear.value == "")
	{
		msg = msg + "Please select year of Billing Period.\n"
		valid = false;
	}


	if (valid == false)
	{
		alert(msg)
	}
	return valid;
}


function checkdata(MonthP, YearP, EmpID, combobox, Search, PageIndex)
{
	var e = document.getElementById(combobox); 
	var Used = e.options[e.selectedIndex].value;
	window.location= "ShuttleBusUpdate.asp?MonthP="+MonthP+"&YearP="+YearP+"&EmpID="+EmpID+"&Used="+Used+"&Search="+Search+"&PageIndex="+PageIndex;
}
</script>

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
  	<TD COLSPAN="4" ALIGN="center" Class="title">Input Shuttle Bus Usage</TD>
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

dim intPageSize,PageIndex,TotalPages 
dim RecordCount,RecordNumber,Count 
intpageSize=20 
PageIndex=request("PageIndex")

EmpID_ = Request.Form("cbEmp")
If EmpID_ ="" Then
	EmpID_ = Request("EmpID")	
End If
MonthP_ = Request.Form("cbMonth")
If MonthP_ ="" Then
	MonthP_ = Request("MonthP")	
End If
YearP_ = Request.Form("cbYear")
If YearP_ ="" Then
	YearP_= Request("YearP")	
End If

If (UserRole_ = "Admin") or (UserRole_ = "TRS") or (UserRole_= "FMC") Then
%>
	<form method="post" name="frmGenerate" onSubmit="return validate_form();">
	<table align="center" cellpadding="1" cellspacing="1" width="60%" border="1">
	<tr>
		<td>
		<table align="center" cellpadding="2" cellspacing="0" width="100%">
		<tr bgcolor="#000099">
			<td height="25" colspan="6"><strong>&nbsp;<span class="style5">Schedule Parameters</span></strong></td>
		</tr>
		<tr>
			<td width="15%">Employee</td>
			<td width="1%">:</td>
			<td>
				<Select name="cbEmp">
					<Option value="">-- All --</option>
<%
'			strsql = "Select EmpID, EmpName From vwPhoneCustomerList Where Type='AMER' order by EmpName "
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
			<td>Billing Period</td>				
			<td>:</td>
			<td>
				<Select name="cbMonth">
					<Option value="">--Select--</Option>
					<Option value="01" <%if MonthP_ ="01" then %>Selected<%End If%>>January</Option>
					<Option value="02" <%if MonthP_ ="02" then %>Selected<%End If%>>February</Option>
					<Option value="03" <%if MonthP_ ="03" then %>Selected<%End If%>>March</Option>
					<Option value="04" <%if MonthP_ ="04" then %>Selected<%End If%>>April</Option>
					<Option value="05" <%if MonthP_ ="05" then %>Selected<%End If%>>May</Option>
					<Option value="06" <%if MonthP_ ="06" then %>Selected<%End If%>>June</Option>
					<Option value="07" <%if MonthP_ ="07" then %>Selected<%End If%>>July</Option>
					<Option value="08" <%if MonthP_ ="08" then %>Selected<%End If%>>August</Option>
					<Option value="09" <%if MonthP_ ="09" then %>Selected<%End If%>>September</Option>
					<Option value="10" <%if MonthP_ ="10" then %>Selected<%End If%>>October</Option>
					<Option value="11" <%if MonthP_ ="11" then %>Selected<%End If%>>November</Option>
					<Option value="12" <%if MonthP_ ="12" then %>Selected<%End If%>>December</Option>
				</Select>&nbsp;
<%
				Year_ = Year(Date()) - 1
%>

				<Select name="cbYear">
				<Option value="">--Select--</Option>
<% 				Do While Year_ <= Year(Date()) %>

				<Option value='<%=Year_%>' <%if trim(Year_) = trim(YearP_) then %>Selected<%End If%> ><%=Year_%></Option>
<% 
			Year_ = Year_ + 1
			Loop %>	
				</Select>
			</td>
		</tr>
		<tr>
			<td colspan="6"><br></td>
		</tr>
		<tr>
			<td colspan="2"></td>
			<td><input type="submit" name="btnSubmit" value="Show Data" ></td>
		</tr>		
		</table>
		</td>
	</tr>
	</table>
	<br>
	<table align="center" cellpadding="1" cellspacing="0" width="60%" border="1" bordercolor="black"> 
	<TR BGCOLOR="#000099" align="center">
		<TD width="3%"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
		<TD><strong><label STYLE=color:#FFFFFF>Employee</label></strong></TD>
		<TD><strong><label STYLE=color:#FFFFFF>Office</label></strong></TD>
		<TD width="10%"><strong><label STYLE=color:#FFFFFF>Used</label></strong></TD>
		<TD width="12%"><strong><label STYLE=color:#FFFFFF>Action</label></strong></TD>
	</TR>    
<% 
		
		dim no_  
		no_ = 1 
	if (request.form("btnSubmit")="Show Data") or (PageIndex<>"") then 
		
		strsql = "exec spGetShuttleBusUsed '" & EmpID_ & "','" & MonthP_ & "','" & YearP_ & "'"
		set DataRS = server.createobject("adodb.recordset")
		'response.write strsql & "<br>"
		DataRS.CursorLocation = 3
		DataRS.open strsql,BillingCon
'		response.write "test" & strsql & "<br>"

		if ((PageIndex ="") or (request.form("btnSubmit")="Show Data")) then PageIndex=1 
		if not DataRS.eof then
			RecordCount = DataRS.RecordCount   
			'response.write RecordCount & "<br>"
			RecordNumber=(intPageSize * PageIndex) - intPageSize 
			'response.write RecordNumber
			DataRS.PageSize =intPageSize 
			DataRS.AbsolutePage = PageIndex
			TotalPages=DataRS.PageCount 
			'response.write
		End If
		no_ = 1 + ((PageIndex*intPageSize)-intPageSize)
		Count = 1
		do while not DataRS.eof and Count<=intPageSize
	   	if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
%>
	 	<TR bgcolor="<%=bg%>">
			<td align="right"><%=No_%>&nbsp;</td>
	       		<td><FONT color=#330099 size=2>&nbsp;<%=DataRS("EmpName")%></font></td> 
		       	<td><FONT color=#330099 size=2><%=DataRS("Office")%>&nbsp;</font></td> 
        		<td align="right">
<%
				Used_ = 1
%>
				<Select name="cbUsed<%=DataRS("EmpID")%>">
				<Option value="0">0</Option>
<% 				Do While Used_ <=70 %>

				<Option value='<%=Used_%>' <%if trim(Used_) = trim(DataRS("Used")) then %>Selected<%End If%> ><%=Used_%></Option>
<% 
					Used_ = Used_ + 1
				Loop %>	
				</Select>

			</td> 
	        	<td>&nbsp;
				<input type="button" value="Update" class=buttontext title="Click this button to Update" onclick="javascript:checkdata('<%=MonthP_%>','<%=YearP_%>','<%=DataRS("EmpID")%>','cbUsed<%=DataRS("EmpID")%>','<%=EMPID_ %>',<%=PageIndex%>);">
			</td> 
	  	 </TR>
<%   
			Count = Count+1
	 		DataRS.movenext
   			no_ = no_ + 1
		loop
	end if
%>
	</table>
	<table align="center" cellpadding="1" cellspacing="0" width="60%">
	<tr>
		<td align="right">
<%
		Do while PageNo<=TotalPages 
			if trim(pageNo) = trim(PageIndex) Then
%>		
				<label class="ActivePage"><%=PageNo%></label>&nbsp;
			<%Else%>
				<a href="InputShuttleBusUsage.asp?PageIndex=<%=PageNo%>&MonthP=<%=MonthP_%>&YearP=<%=YearP_ %>&EmpID=<%=EmpID_ %>"><%=PageNo%></a>&nbsp;
<%	
			End If						
			PageNo=PageNo+1
		Loop
%>
		</td>
	</tr>
	</table>
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

	</form>
</body> 

</html>


