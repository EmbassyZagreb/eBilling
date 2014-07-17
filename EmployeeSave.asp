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
EmpID_ = Request.form("txtEmpID")
'EmpType_ = Request.form("txtEmpType")
EmpType_ = Request.form("cmbType")
ReportTo_ = Request.form("cmbReportTo")
EmailAddress_ = Request.form("txtEmailAddress")
LoginID_ = Request.form("txtLoginID")
AlternateEmail_ = Request.form("txtAlternateEmail")
FundingAgency_ = Request.form("cmbFundingAgency")

State_ =  trim(request.form("State"))

EmpName_ =  trim(request.form("txtName"))
Post_ =  trim(request.form("cmbPostList"))
OfficeSection_ = trim(request.form("cmbOfficeList")) 
WorkingTitle_ = trim(request.form("txtWorkingTitle")) 
Agency_ = Request.form("cmbAgencyList")

Remark_ =  trim(request.form("txtRemark"))
Status_ =  trim(request.form("cmbStatus"))
user_ = request.servervariables("remote_user") 
UserName_ = right(user_,len(user_)-4)

%>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">

<meta http-equiv="refresh" content="1;url=EmployeeList.asp">

<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate" />
<meta http-equiv="Pragma" content="no-cache" />
<meta http-equiv="Expires" content="0" />

<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">EMPLOYEE LIST</TD>
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
	strsql = "Exec spEmployee_IUD '" & State_ & "','" & EmpID_ & "','" & EmpName_  & "'," & FundingAgency_ & ",'" & Post_ & "','" & EmpType_ & "','" & Agency_ & "','" & OfficeSection_ & "','" & WorkingTitle_ & "','" & EmailAddress_ & "','" & AlternateEmail_ & "','" & ReportTo_ & "','" & LoginID_ & "','" & Remark_ & "','" & Status_ & "','" & UserName_ & "'"


	'strsql = "Update MsEmployee Set SupervisorId ='" & ReportTo_ & "', EmailAddress='" & Email_ & "', LoginID='" & LoginID_ & "', AlternateEmail='" & AlternateEmail_ & "', AgencyId=" & FundingAgency_ & " Where EmpId='" & EmpID_ & "'"
	'response.write strsql 
	BillingCon.execute strsql
%>
              
<tr><td align=center>Your data has been updated.</td></tr>
<tr><td>&nbsp;</td>

<tr><td align=center>  
<!-- <input type="button" value="Close" id="btnclose"> -->
</td></tr>
<tr>
	<td align="center"><br><a href="EmployeeList.asp"><img src="images/Back.gif" border="0" alt="Go..Back" WIDTH="83" HEIGHT="25"></a></td>
</tr>

</table>

   </body>
</html>