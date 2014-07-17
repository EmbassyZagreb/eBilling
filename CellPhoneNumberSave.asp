<%@ Language=VBScript %>
<!--#include file="connect.inc" -->

<html>
   <head>
   <script language="vbscript">
       <!--
        Sub btnBack_onclick
           history.back
	End Sub
        Sub btnClose_onclick
		close
	End Sub
       --> 
   </script>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">



<link href="style.css" rel="stylesheet" type="text/css">
<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  NAME="CDO for Windows 2000 Library" --> 
<%
dim CTL , ID_ , UserRole_
ID_ =  trim(request.form("txtID"))
State_ =  trim(request.form("State"))
PhoneNumber_ =  trim(request.form("txtPhoneNumber"))
PhoneType_ =  trim(request.form("PhoneTypeList"))
EmpID_ =  trim(request.form("EmployeeList"))
AlternateEmail_ =  trim(request.form("txtAlternateEmail"))
OwnerID_ =  trim(request.form("OwnerList"))
Remark_ =  replace(trim(request.form("txtRemark")),"'","''")
BillFlag_ =  trim(request.form("BillFlagList"))
Discontinued_ = trim(request.form("cmbDiscontinued"))
DiscontinuedDate_ = trim(request.form("txtDiscontinuedDate"))

user_ = request.servervariables("remote_user") 
UserName_ = right(user_,len(user_)-4)

if State_ ="I" Then
	ID_ = 0
End If


		'strsql = "select ID from vwCellPhoneNumberList where PhoneNumber = " & PhoneNumber_ 
       		'set ValidateRS = server.createobject("adodb.recordset") 
        	'set ValidateRS  = BillingCon.execute(strsql)
  		'if ValidateRS.eof then
			
			%>
			<meta http-equiv="refresh" content="1;url=CellPhoneNumberList.asp">
			<%
       		'	strsql = "Exec spCellPhoneNumber_IUD '" & State_ & "'," & ID_ & ",'" & PhoneNumber_ & "','" & PhoneType_ & "','" & EmpID_ & "','" & AlternateEmail_ & "','" & Remark_ & "','" & BillFlag_ & "','" & Discontinued_ & "','" & DiscontinuedDate_ & "','" & OwnerID_ & "','" & UserName_ & "'"
       		'	BillingCon.execute strsql
		'	Notification_ = "Your data has been saved."
		'else
			
		'	IDValidate_ = ValidateRS("ID")
		'	If ID_ = IDValidate_ then
		'	%>
		'	<meta http-equiv="refresh" content="1;url=CellPhoneNumberList.asp">
		'	<%
       		'	strsql = "Exec spCellPhoneNumber_IUD '" & State_ & "'," & ID_ & ",'" & PhoneNumber_ & "','" & PhoneType_ & "','" & EmpID_ & "','" & AlternateEmail_ & "','" & Remark_ & "','" & BillFlag_ & "','" & Discontinued_ & "','" & DiscontinuedDate_ & "','" & OwnerID_ & "','" & UserName_ & "'"
       		'	BillingCon.execute strsql
			Notification_ = "Your data has been saved."
		'	else			
		'	Notification_ = "Phone number " & PhoneNumber_ & " already exist!"
		'	end if
		'end if
%>
   </head>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">CELL PHONE UPDATE</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
<tr>
	<td colspan="4"><HR style="LEFT: 10px; TOP: 59px" align=center></td>
</tr>
</tr>
<tr>
	<td colspan="4"><br></td>
</tr>
</table>
<table border=0 width=100%>
<%
       strsql = "Exec spCellPhoneNumber_IUD '" & State_ & "'," & ID_ & ",'" & PhoneNumber_ & "','" & PhoneType_ & "','" & EmpID_ & "','" & AlternateEmail_ & "','" & Remark_ & "','" & BillFlag_ & "','" & Discontinued_ & "','" & DiscontinuedDate_ & "','" & OwnerID_ & "','" & UserName_ & "'"
       response.write strsql 
       BillingCon.execute strsql
%>               
<tr><td align=center><%=Notification_%></td></tr>
<tr><td>&nbsp;</td>
<!--
<tr><td align=center> 
<input type="button" value="Close" id="btnclose">
</td></tr>
-->
<tr>
	<td align="center"><br><a href="CellPhoneNumberEdit.asp?ID=<%=ID_%>&State=E"><img src="images/Back.gif" border="0" alt="Go..Back" WIDTH="83" HEIGHT="25"></a></td>
</tr>
</table>

   </body>
</html>