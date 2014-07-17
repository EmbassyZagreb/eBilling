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
	document.forms['frmSearch'].elements['txtName'].value ="";
	document.forms['frmSearch'].elements['cmbStatus'].value ="C";
}

</script>

<%

dim intPageSize,PageIndex,TotalPages 
dim RecordCount,RecordNumber,Count 
intpageSize=20 
PageIndex=request("PageIndex")

Name_ = Trim(Request.Form("txtName"))
if Name_ ="" then
	Name_ = Trim(request("Name"))
End If
'response.write Name_

Status_ = Trim(Request.Form("cmbStatus"))
if Status_ ="" then
	if Trim(request("Status")) <> "" then
		Status_ = Trim(request("Status"))
	else
		Status_ = "C"
	end if	
End If
%>

<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">Non Employee LIST</TD>
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

If (UserRole_ <> "") Then

	strsql = "Select * from vwNonEmployeeList Where NonEmpName like '%" & Name_ & "%' and (Status='" & Status_ & "' or '" & Status_ & "'='X') order by NonEmpName"
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
	<form method="post" name="frmSearch" id="frmSearch" onSubmit="return validate_form();">
	<table align="center" cellpadding="1" cellspacing="0" width="70%">
	<tr bgcolor="#000099">
		<td height="25" colspan="4"><strong>&nbsp;<span class="style5">Search</span></strong></td>
	<tr>
		<td>Name :</td>
		<td><input name="txtName" type="Input" size="50" Value='<%=Name_%>'></td>
		<td>Status :</td>
		<td>
			<Select name="cmbStatus">
				<Option value="X" <%if Status_ ="X" then %>Selected<%End If%> >All</Option>
				<Option value="C" <%if Status_ ="C" then %>Selected<%End If%> >Current</Option>
				<Option value="D" <%if Status_ ="D" then %>Selected<%End If%> >Departed</Option>
			</Select>&nbsp;
		</td>
	</tr>
	<tr>
		<td colspan="3">
			<input type="submit" name="btnSearch" value="Search">
			&nbsp;&nbsp;<input type="button" name="btnClear" value="Clear filter" onclick="javascript:ClearFilter();">
		</td>
		<td><input type="button" name="btnExport" value="Export to Excel" onClick="javascript:document.location.href('NonEmployeeListPrint.asp?EmpName=<%=EmpName_%>&Status=<%=Status_%>');"/>
		</td>

	</tr>
	<tr>
		<td colspan="4"><hr></td>
	</tr>
</table>
</form>
<form method="post" name="frmPaymentList" action="" onSubmit="return ValidateCheckBox();">
<table width="90%">
<tr>
	<td><a href="NonEmployeeEdit.asp?State=I">Add New non employee</a></td>
</tr>
</table>
<table align="center" cellpadding="1" cellspacing="0" width="90%" border="1" bordercolor="black"> 
<TR BGCOLOR="#000099" align="center" cellpadding="0" cellspacing="0" >
	<TD width="30px"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
	<TD><strong><label STYLE=color:#FFFFFF>Name</label></strong></TD>
        <TD><strong><label STYLE=color:#FFFFFF>Agency Funding</label></strong></TD>
	<TD><strong><label STYLE=color:#FFFFFF>Email</label></strong></TD>
	<TD><strong><label STYLE=color:#FFFFFF>Remark</label></strong></TD>
       	<TD width="80px"><strong><label STYLE=color:#FFFFFF>Status</label></strong></TD>
</TR>    
<% 
	dim no_  
'	no_ = 1 
	no_ = 1 + ((PageIndex*intPageSize)-intPageSize)
	do while not DataRS.eof and Count<intPageSize
   	if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
%>
   	<TR bgcolor="<%=bg%>">
		<td align="right"><%=No_%>&nbsp;</td>
	        <td>&nbsp;<FONT color=#330099 size=2><A HREF="NonEmployeeEdit.asp?NonEmpID=<%=DataRS("NonEmpID")%>&State=E"><%= DataRS("NonEmpName") %></A></font></td>
        	<td><FONT color=#330099 size=2>&nbsp;<%=DataRS("AgencyDesc")%></font></td> 
        	<td><FONT color=#330099 size=2>&nbsp;<%=DataRS("Email")%></font></td> 
        	<td><FONT color=#330099 size=2>&nbsp;<%=DataRS("Remark")%></font></td> 
        	<td><FONT color=#330099 size=2>&nbsp;<%=DataRS("StatusName")%></font></td>
		</td> 
  	 </TR>
<%   
		Count=Count +1
 		DataRS.movenext
   		no_ = no_ + 1
	loop
%>
</table>
<table align="center" cellpadding="1" cellspacing="0" width="90%">
	<tr>
		<td align="right">
<%
		Do while PageNo<=TotalPages 
			if trim(pageNo) = trim(PageIndex) Then
%>		
				<label class="ActivePage"><%=PageNo%></label>&nbsp;
			<%Else%>
				<a href="NonEmployeeList.asp?PageIndex=<%=PageNo%>&Name=<%=Name_%>"><%=PageNo%></a>&nbsp;
<%	
			End If						
			PageNo=PageNo+1
		Loop
%>
		</td>
	</tr>
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


