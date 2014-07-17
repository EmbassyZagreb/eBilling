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
 user1_ = right(user_,len(user_)-4)
'user1_ = "pranataw"
'response.write user1_ & "<br>"

%> 
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY BGCOLOR="#FFFFFF" TEXT="#000000" LINK="#0080FF" ALINK="0080FF" VLINK="#0055AA" MARGINWIDTH="8" MARGINHEIGHT="0" LEFTMARGIN="8" TOPMARGIN="0">
<CENTER>

<form method="post" action="UserSettingSaveAddOn.asp">
<%  
 dim rst 
 dim strsql

 State_ = request("State")
 strsql = "Select A.EmployeesID, A.Type, A.EmpName, A.Office, A.WorkPhone, A.MobilePhone, A.EmailAddress, A.ReportTo, ISNULL(B.EmpName,'') As ReportToName From vwPhoneCustomerList A Left Join vwPhoneCustomerList B on (A.ReportTo=B.EmpID) where A.loginId='" & user1_ & "'"
 'response.write strsql 
  set rsData = server.createobject("adodb.recordset") 
  set rsData = BillingCon.execute(strsql)
  if not rsData.eof then 
	EmployeesID_ = rsData("EmployeesID")
	Type_ = rsData("Type")
	EmpName_ = rsData("EmpName")
	Office_ = rsData("Office")
	WorkPhone_ = rsData("WorkPhone")
	MobilePhone_ = rsData("MobilePhone")
	EmailAddress_ = rsData("EmailAddress")
	ReportTo_ = rsData("ReportTo")
  end if
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
		strsql = "Select EmpID, ISNULL(EmpName,'')+' - '+ISNULL(Office,'') As EmpName from vwPhoneCustomerList Where LEN(ISNULL(EmpName,''))<>''  Order by EmpName"
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
      &nbsp;<input type="button" value="Cancel" name="btnCancel" onclick="javascript:document.location.href('DefaultAddOn.asp')">
 </td>
</tr>  
<tr><td colspan=2>&nbsp;</td></tr>
</table>
</form>
</BODY>
</html>