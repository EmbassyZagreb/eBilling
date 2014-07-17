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
dim NonEmpID_ , UserRole_
NonEmpID_ =  trim(request.form("txtNonEmpID"))
State_ =  trim(request.form("State"))
NonEmpName_ =  trim(request.form("txtName"))
FundingAgency_ = Request.form("cmbFundingAgency")
Email_ =  trim(request.form("txtEmail"))
Remark_ =  trim(request.form("txtRemark"))
Status_ =  trim(request.form("cmbStatus"))
'response.write NonEmpID_ & "<br>"
'response.write UserRole_ & "<br>"
user_ = request.servervariables("remote_user") 
UserName_ = right(user_,len(user_)-4)

%>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
<TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">Non Employee Update</TD>
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
       strsql = "Exec spNonEmployee_IUD '" & State_ & "','" & NonEmpID_ & "','" & NonEmpName_  & "'," & FundingAgency_ & ",'" & Email_ & "','" & Remark_ & "','" & Status_ & "','" & UserName_ & "'"
	'response.write strsql 
       BillingCon.execute strsql
%>               
<tr><td align=center>Your data has been saved. </td></tr>
<tr>
	<td>&nbsp;</td>
<tr>
	<td align="center"><br><a href="NonEmployeeList.asp"><img src="images/Back.gif" border="0" alt="Go..Back" WIDTH="83" HEIGHT="25"></a></td>
</tr>
</table>

   </body>
</html>