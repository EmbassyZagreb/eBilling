<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<html>
<head>
<script language="vbscript">
       <!--
        Sub btnCancel_onclick
           history.back
	End Sub

       --> 
   </script>


<% 
 dim user_ 
 dim user1_  

 
 user_ = request.servervariables("remote_user") 
 user1_ = user_  'user1_ = right(user_,len(user_)-4)
'user1_ = "pranataw"
'response.write user1_ & "<br>"

%> 
<TITLE>U.S. Embassy Zagreb - zBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">USER SETTING</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>

<form method="post" action="UserSettingSave.asp"> 
<%  
 dim rst 
 dim strsql

 State_ = request("State")
' strsql = "Select A.EmployeesID, A.Type, A.EmpName, A.Office, A.WorkPhone, A.MobilePhone, A.EmailAddress, A.ReportTo, ISNULL(B.EmpName,'') As ReportToName From vwPhoneCustomerList A Left Join vwPhoneCustomerList B on (A.ReportTo=B.EmpID) where A.loginId='" & user1_ & "'"
 strsql = "Select * From vwPhoneCustomerList where LoginID='" & user1_ & "'"
  'response.write strsql 
  set rsData = server.createobject("adodb.recordset") 
  set rsData = BillingCon.execute(strsql)

  if not rsData.eof then 
	EmpID_ = rsData("EmpID")
	Type_ = rsData("EmpType")
	EmpName_ = rsData("EmpName")
	Office_ = rsData("Office")
	'WorkPhone_ = rsData("WorkPhone")
	MobilePhone_ = rsData("MobilePhone")
	EmailAddress_ = rsData("EmailAddress")
	ReportTo_ = rsData("SupervisorId")
  end if
	EmployeesID_ = EmpID_
 'response.write ReportTo_ 
%>             
<table align="center">
<tr>
  <td>Employee Name :</td>
  <td><label class="FontContent"><%=EmpName_ %></label>
  </td>
</tr>
<tr>
  <td>Office :</td>
  <td><label class="FontContent"><%=Office_ %></label>
  </td>
</tr>
<tr>
  <td>WorkPhone:</td>
  <td><label class="FontContent"><%=WorkPhone_ %></label>
  </td>
</tr>
<tr>
  <td>MobilePhone :</td>
  <td><label class="FontContent"><%=MobilePhone_ %></label>
  </td>
</tr>
<tr>
  <td>Email Address :</td>
  <td><label class="FontContent"><%=EmailAddress_ %></label>
  </td>
</tr>
<tr>
  <td>My Supervisor :</td>
  <td>
	<select id="cmbReportTo" name="cmbReportTo">
		<option value="">-- Select --</option>
<%
		Dim UserRS
		'strsql = "Select EmpID, ISNULL(EmpName,'')+' - '+ISNULL(Office,'') As EmpName from vwPhoneCustomerList Where LEN(ISNULL(EmpName,''))<>'' AND EmpType = 'AMER' Order by EmpName"
		strsql = "Select EmpID, ISNULL(EmpName,'')+' - '+ISNULL(Office,'') As EmpName from vwDirectReport Where LEN(ISNULL(EmpName,''))<>'' AND Type = 'AMER' Order by EmpName"
		response.write strsql & "<br>"
		set UserRS = server.createobject("adodb.recordset")
		set UserRS =BillingCon.execute(strsql)				        
		do while not UserRS.eof
			EmpID_ = UserRS("EmpID")
			Ename_ = UserRS("EmpName")
%>
	        <OPTION value='<%=EmpID_%>' <%if (trim(EmpID_) = trim(ReportTo_)) then%>selected<%end if%>><%= EName_  %>
<%
                 UserRS.movenext
	        loop
%>  
	</select>	
  </td>
</tr>


<tr>
  <td></td>
  <td><input type="submit" name="btnSubmit" value="Update">
      <input type="hidden" name="txtEmpID" value=<%=EmployeesID_ %>>
      <input type="hidden" name="txtEmpType" value=<%=Type_ %>>
      &nbsp;<input type="button" value="Cancel" name="btnCancel">
 </td>
</tr>  
<tr><td colspan=2>&nbsp;</td></tr>
</table>
</form>
</BODY>
</html>