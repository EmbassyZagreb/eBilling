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
 Amount_ = request("Amount") 
%> 
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">Pay.gov Form</TD>
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
 strsql = "Select A.EmpID, A.EmpType As Type, A.EmpName, A.Office, A.WorkPhone, A.MobilePhone, A.EmailAddress, A.SupervisorId As ReportTo, ISNULL(B.EmpName,'') As ReportToName From vwPhoneCustomerList A Left Join vwPhoneCustomerList B on (A.SupervisorId=B.EmpID)  where A.loginId='" & user1_ & "'"
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
  <td>Name :</td>
  <td>
	<input type="input" name="txtName" size="50" value="<%=EmpName_ %>" />
  </td>
</tr>
<tr>
  <td>Street Address :</td>
  <td>
	<input type="input" name="txtAddress" size="50" value="Enter your address here" />
  </td>
</tr>
<tr>
  <td>Street Address 2:</td>
  <td>
	<input type="input" name="txtAddress2" size="50" />
  </td>
</tr>
<tr>
  <td>City :</td>
  <td>
	<input type="input" name="txtCity" size="30" value="Zagreb" />
  </td>
</tr>
<tr>
  <td>State :</td>
  <td>
	<input type="input" name="txtState" size="30" />
  </td>
</tr>
<tr>
  <td>Zip/Postal code :</td>
  <td>
	<input type="input" name="txtZip" size="20" />
  </td>
</tr>
<tr>
  <td>Country :</td>
  <td>
	<input type="input" name="txtName" size="30" value="Croatia" />
  </td>
</tr>
<tr>
  <td>Daytime Phone Number :</td>
  <td>
	<input type="input" name="txtPhone" size="30" value=<%=MobilePhone_ %> />
  </td>
</tr>
<tr>
  <td>Email Address :</td>
  <td>
	<input type="input" name="txtEmail" size="50" value=<%=EmailAddress_ %> />
  </td>
</tr>
<tr>
  <td>Bill of Collection Number :</td>
  <td>
	<input type="input" name="txtBillNo" size="50" value="" />
  </td>
</tr>
<tr>
  <td>Payment owe To :</td>
  <td>
	<input type="input" name="txtPaymentTo" size="50" value="Zagreb - US EMBASSY (IDK/IDS/IDP)" />
  </td>
</tr>
<tr>
  <td>Reason for Payment :</td>
  <td>
	<select name="cmbReason">
		<option>Telephone/Fax Bills</option>
	</select>
  </td>
</tr>
<tr>
  <td valign="top">Payment Description :</td>
  <td>
	<textarea name="txtPaymentDesc" cols="40" rows="5"></textarea>
  </td>
</tr>
<tr>
  <td>Amount :</td>
  <td>
	<input type="input" name="txtAmount" size="10" value="<%=Amount_ %>" />
  </td>
</tr>
<tr><td colspan=2>&nbsp;</td></tr>
</table>
</form>
</BODY>
</html>