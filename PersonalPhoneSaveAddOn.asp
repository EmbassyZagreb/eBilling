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
ID_ =  trim(request.form("txtID"))
State_ =  trim(request.form("State"))
PhoneNumber_ =  trim(request.form("txtPhoneNumber"))
Remark_ =  trim(request.form("txtRemark"))
'response.write UserRole_ & "<br>"
user_ = request.servervariables("remote_user") 
UserName_ = right(user_,len(user_)-4)

if ID_ = "" then
	ID_ =0
end if
%>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY BGCOLOR="#FFFFFF" TEXT="#000000" LINK="#0080FF" ALINK="0080FF" VLINK="#0055AA" MARGINWIDTH="8" MARGINHEIGHT="0" LEFTMARGIN="8" TOPMARGIN="0">
<CENTER>
<table border=0 width=100%>
<%
       strsql = "Exec spPersonalPhone_IUD '" & State_ & "'," & ID_ & ",'" & PhoneNumber_ & "','" & UserName_  & "','" & Remark_ & "','" & UserName_ & "'"
       'response.write strsql 
       BillingCon.execute strsql
%>               
<tr><td align=center>Your data has been saved. </td></tr>
<tr>
	<td>&nbsp;</td>
<tr>
	<td align="center"><br><a href="PersonalPhoneListAddOn.asp"><img src="images/Back.gif" border="0" alt="Go..Back" WIDTH="83" HEIGHT="25"></a></td>
</tr>
</table>

   </body>
</html>