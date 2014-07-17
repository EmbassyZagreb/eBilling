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
	document.forms['frmSearch'].elements['txtPhoneNumber'].value ="";
}

</script>

<%

dim intPageSize,PageIndex,TotalPages 
dim RecordCount,RecordNumber,Count 
intpageSize=20 
PageIndex=request("PageIndex")

PhoneNumber_ = Trim(Request.Form("txtPhoneNumber"))
if PhoneNumber_ ="" then
	PhoneNumber_= Trim(request("PhoneNumber"))
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
<BODY BGCOLOR="#FFFFFF" TEXT="#000000" LINK="#0080FF" ALINK="0080FF" VLINK="#0055AA" MARGINWIDTH="8" MARGINHEIGHT="0" LEFTMARGIN="8" TOPMARGIN="0">
<CENTER>
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

	strsql = "Select * from MsPersonalPhone Where Owner='" & user1_ & "' and PhoneNumber like '%" & PhoneNumber_& "%' "
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
	<table align="center" cellpadding="1" cellspacing="0" width="500px">
	<tr bgcolor="#000099">
		<td height="25" colspan="4"><strong>&nbsp;<span class="style5">Search</span></strong></td>
	<tr>
		<td>Name :</td>
		<td><input name="txtPhoneNumber" type="Input" size="30" Value='<%=PhoneNumber_%>'></td>
		<td>
			<input type="submit" name="btnSearch" value="Search">
			&nbsp;&nbsp;<input type="button" name="btnClear" value="Clear filter" onclick="javascript:ClearFilter();">
		</td>
<!--
		<td><input type="button" name="btnExport" value="Export to Excel" onClick="javascript:document.location.href('PersonalPhoneListPrint.asp?Pho=<%=EmpName_%>&Status=<%=Status_%>');"/>
		</td>
-->

	</tr>
	</table>
</form>
<form method="post" name="frmPaymentList" action="" onSubmit="return ValidateCheckBox();">
<table width="500px">
<tr>
	<td><a href="PersonalPhoneEditAddOn.asp?State=I">Add New Personal Phone</a></td>
</tr>
</table>
<table align="center" cellpadding="1" cellspacing="0" width="500px" border="1" bordercolor="black"> 
<TR BGCOLOR="#000099" align="center" cellpadding="0" cellspacing="0" >
	<TD width="30px"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
	<TD><strong><label STYLE=color:#FFFFFF>Phone Number</label></strong></TD>
	<TD><strong><label STYLE=color:#FFFFFF>Remark</label></strong></TD>
	<TD width="50px"></TD>
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
	        <td>&nbsp;<FONT color=#330099 size=2><A HREF="PersonalPhoneEditAddOn.asp?ID=<%=DataRS("ID")%>&State=E"><%= DataRS("PhoneNumber") %></A></font></td>
        	<td><FONT color=#330099 size=2>&nbsp;<%=DataRS("Remark")%></font></td> 
		</td> 
		<TD><a href="PersonalPhoneDeleteConfirmAddOn.asp?ID=<%=DataRS("ID")%>&State=E">Delete</a></TD>
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
				<a href="PersonalPhoneListAddOn.asp?PageIndex=<%=PageNo%>&PhoneNumber=<%=PhoneNumber_%>"><%=PageNo%></a>&nbsp;
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


