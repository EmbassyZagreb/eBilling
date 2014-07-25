<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="connect.inc" -->
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<style type="text/css">
<!--
.style1 {
	font-size: large;
	font-weight: bold;
	}

.FontText {
	font-size: small;
}

.FontContent {
	font-size: small;
        color: blue;
}

.FontJudul {
	font-size: 24px;
	font-weight: bold;
}

.FontComment {
	font-size: 18px;
	font-weight: bold;
}

-->
</style>
<link href="style.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
function validate_form()
{
	valid = true;

	if (document.frmAgencyEdit.txtAgencyName.value == "" )
	{
		alert("Please fill the agency name !!!");
		valid = false;
	}

	if (document.frmAgencyEdit.txtAgencyStripe.value == "" )
	{
		alert("Please fill the agency stripe !!!");
		valid = false;
	}

	if (document.frmAgencyNew.txtAgencyStripe.value == "" )
	{
		alert("Please fill the Fiscal Strip VAT !!!");
		valid = false;
	}

	if (document.frmAgencyNew.txtAgencyStripeNonVAT.value == "" )
	{
		alert("Please fill the Fiscal Strip Non VAT !!!");
		valid = false;
	}

	return valid;
}
</script>
<%

Dim user_ , user1_, UserRole_

user_ = request.servervariables("remote_user")
user1_ = user_  'user1_ = right(user_,len(user_)-4)

ID_ = request.querystring("ID")
'response.write "ServiceRecordID : " & serviceRecordId & "<br>"

%>
</head>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="SubTitle">Agency Edit</TD>
  </TR>
  <tr>
	<td colspan="3" align="Left" width="20%"><A HREF="Default.asp">Home</A></td>
	<td align="Right" width="20%"><A HREF="AgencyList.asp">Back</A></td>
  </tr> 
  <tr>
  	<td COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></td>
   </tr>
  </TABLE>
<form method="post" name="frmAgencyEdit" id="frmAgencyEdit" action="AgencySave.asp?Mode=U" onsubmit="return validate_form()">
<%
strsql = "select * from Users where loginId='" & user1_ & "'"
set RS_Query = server.createobject("adodb.recordset")
'response.write strsql & "<br>"
set RS_Query = BillingCon.execute(strsql)
'response.write RS_Query("RoleID") & "<br>"


if (trim(RS_Query("RoleID")) = "Admin") or (trim(RS_Query("RoleID")) = "Voucher") or (trim(RS_Query("RoleID")) = "FMC") then    
%>
	<table align="center">
	<%
	   dim rsAgency
	   strsql = "Select * From AgencyFunding Where AgencyID =" & ID_ 
	   set rsAgency= server.createobject("adodb.recordset")    
	'response.write strsql & "<br>"
	   set rsAgency= BillingCon.execute(strsql) 	  
'	   Response.write rsAgency("Disabled")
	%>
	<tr>
		<td align="right">Agency Code:</td>
		<td align="left"><input name="txtAgencyCode" type="Input" size="8" value='<%=rsAgency("AgencyFundingCode")%>' ></td>
	</tr>	
	<tr>
		<td align="right">Agency Name :</td>
		<td align="left"><input name="txtAgencyName" type="Input" size="60" value='<%=rsAgency("AgencyDesc")%>'></td>
	</tr>
	<tr>
		<td align="right">Fiscal Strip  VAT :</td>
		<td align="left"><input name="txtAgencyStripe" type="Input" size="100" value='<%=rsAgency("FiscalStripVAT")%>'></td>
	</tr>
	<tr>
		<td align="right">Fiscal Strip Non VAT :</td>
		<td align="left"><input name="txtAgencyStripeNonVAT" type="Input" size="100" value='<%=rsAgency("FiscalStripNonVAT")%>'></td>
	</tr>
	<tr>
		<td align="right">Disabled :</td>
		<td align="left">
		     <select name="txtAgencyType">
			<option value="">--Select--</option>
			<% if rtrim(rsAgency("Disabled"))="Y" then%>
				<option value="Y" selected>Yes</option>
			<%Else%>
				<option value="Y">Yes</option>
			<%End If%>
			<% if rtrim(rsAgency("Disabled"))="N" then%>
				<option value="N" Selected>No</option>
			<%Else%>
				<option value="N">No</option>
			<%End If%>
		     </select>
	</tr>
  	 <tr>
		<td colspan=2 align="center">
        		<input type="submit" value="Submit">
			&nbsp;<input type="button" value="Cancel" onClick="javascript:location.href='AgencyList.asp'">
			<INPUT TYPE="HIDDEN" NAME="txtID" value=<%=ID_%>>
    		</td>
  	</tr>
	</table>
<%Else %>
	<table>
		<tr>
			<td>You do not have permission to access this site.</td>
		</tr>
		<tr>
			<td>Please <a href="http://zagrebws03.eur.state.sbu/WebPASS/eservices/MainPage.asp">Submit Request </a> or contact Zagreb ISC Helpdesk at ext.3333.</td>
		</tr>
	</table>

<% end if %>
</form>
</body>
</html>
