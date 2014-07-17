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
<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  NAME="CDO for Windows 2000 Library" --> 
<%
dim CTL , ID_ , UserRole_
ID_ =  trim(request.form("txtID"))
State_ =  trim(request.form("State"))
PhoneNumber_ =  trim(request.form("txtPhoneNumber"))
PhoneType_ =  trim(request.form("PhoneTypeList"))
EmpID_ =  trim(request.form("EmployeeList"))
AlternateEmail_ =  trim(request.form("txtAlternateEmail"))
Remark_ =  replace(trim(request.form("txtRemark")),"'","''")
BillFlag_ =  trim(request.form("BillFlagList"))
ShareFlag_ =  trim(request.form("ShareFlagList"))

user_ = request.servervariables("remote_user") 
UserName_ = right(user_,len(user_)-4)

if State_ ="I" Then
	ID_ = 0
End If

%>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>

<meta http-equiv="refresh" content="1;url=OfficePhoneNumberList.asp">

<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">

</HEAD>
<!--#include file="Header.inc" -->
   <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">OFFICE PHONE LIST</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>

<table border=0 width=100%>
<%
       strsql = "Exec spOfficePhoneNumber_IUD '" & State_ & "'," & ID_ & ",'" & PhoneNumber_ & "','" & PhoneType_ & "','" & EmpID_ & "','" & AlternateEmail_ & "','" & Remark_ & "','" & BillFlag_ & "','" & ShareFlag_ & "','" & UserName_ & "'"
	'response.write strsql 
       BillingCon.execute strsql
%>               
<tr><td align=center>Your data has already been saved. Thank you.</td></tr>
<tr><td>&nbsp;</td>
<tr><td align=center> 
<!-- <input type="button" value="Close" id="btnclose"> -->
</td></tr>
<tr>
	<td align="center"><br><a href="OfficePhoneNumberList.asp"><img src="images/Back.gif" border="0" alt="Go..Back" WIDTH="83" HEIGHT="25"></a></td>
</tr>
</table>

   </body>
</html>