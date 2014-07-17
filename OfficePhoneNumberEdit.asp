<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<html>
<head>

<script type="text/javascript">
function validate_form()
{
	valid = true;
	msg = ""

	if (document.frmOfficePhone.txtPhoneNumber.value == "" )
	{
		msg = msg + "Please fill in Phone Number!!!\n"
		valid = false;
	}

	if (document.frmOfficePhone.txtAlternateEmail.value != "" )
	{
		var alnum="a-zA-Z0-9";
		exp="^[^@\\s]+@(["+alnum+"+\\-]+\\.)+["+alnum+"]["+alnum+"]["+alnum+"]?$";
		emailregexp = new RegExp(exp);

		result = document.frmOfficePhone.txtAlternateEmail.value.match(emailregexp);
		if (result == null)
		{
			msg = msg + "Invalid data type for alternative email address !!!\n"
			valid = false;
		}
	}

	if (valid == false)
	{
		alert(msg);
	}
	return valid;
}
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
<form method="post" name="frmOfficePhone" action="OfficePhoneNumberSave.asp" onSubmit="return validate_form()"> 
<%  
 dim rst 
 dim strsql
 dim rst1
 dim today_


 today_ = now()

 ID_ = request("ID")
 State_ = request("State")
' State_ = "I"
  strsql = "select * from Users where loginId='" & user1_ & "'"
'response.write user1_ & "<br>"
'response.write strsql 
  set rst = server.createobject("adodb.recordset") 
  set rst = BillingCon.execute(strsql)
  if not rst.eof then 
     if (trim(rst("RoleID")) = "Admin") or (trim(rst("RoleID")) = "IM") or (mid(rst("RoleID"),1,3) = "FMC")  then  
	if State_ = "E" then
	        strsql = " select * from vwOfficePhoneNumberList where ID = " & ID_ 
       		set rst1 = server.createobject("adodb.recordset") 
		'response.write strsql 
        	set rst1 = BillingCon.execute(strsql)
	       	if not rst1.eof then 
        	   PhoneNumber_ = rst1("PhoneNumber") 
		   PhoneType_ = rst1("PhoneType")  
		   EmpID_ = rst1("EmpID")
		   EmailAddress_ = rst1("EmailAddress")
		   AlternateEmail_ = rst1("AlternateEmail")
		   Remark_ = rst1("Remark")
		   BillFlag_ = rst1("BillFlag")
		   ShareFlag_ = rst1("ShareFlag")
        	end if
       	end if
	'response.write State_ & "<br>"
	%>             
	<table align=center>  
	<tr>
	  <td>Phone Extension</td>
	  <td width="1%">:</td>
	  <td><input type="input" name="txtPhoneNumber" value='<%=PhoneNumber_ %>' size="15" maxlength="30" /></td>
	</tr>
	<tr>
	  <td>Phone Type</td>
	  <td width="1%">:</td>
	  <td>
		  <select name="PhoneTypeList">
<%
			Dim PhoneRS
			strsql = "select PhoneType, PhoneTypeName from PhoneType order by PhoneTypeName Desc"
			'response.write strsql & "<br>"
			set PhoneRS = server.createobject("adodb.recordset")
			set PhoneRS =BillingCon.execute(strsql)	
			do while not PhoneRS.eof
%>
			<option value=<%=PhoneRS("PhoneType")%> <%if PhoneType_ = PhoneRS("PhoneType") Then %>Selected<%End If%> ><%=PhoneRS("PhoneTypeName")%></option>
<%	      	        	PhoneRS.movenext
	        	loop
%>  
		  </select>
	  </td>
	</tr>
	<tr>
	  <td>Employee :</td>
	  <td width="1%">:</td>
	  <td>
<%
			Dim EmpRS
			strsql = "select EmpID, EmpName, Office from vwPhoneCustomerList order by EmpName"
			'response.write strsql & "<br>"
			set EmpRS = server.createobject("adodb.recordset")
			set EmpRS =BillingCon.execute(strsql)	
%>
		<select name="EmployeeList">
			<option value="">-- Vacant --</option>
	<%
			        
			do while not EmpRS.eof
				Ename_ = EmpRS("EmpName") 
				Ename_ = EName_ & "(" & EmpRS("Office") & ")"
				if EmpRS("EmpID") = EmpID_  then		
	%>
				        <OPTION value='<%=EmpRS("EmpID")%>' Selected>  <%= EName_  %>
	<%			Else%>
			        	<OPTION value='<%=EmpRS("EmpID")%>'>  <%= EName_  %>
	<%			End If
        	         EmpRS.movenext
	        	loop
	%>  
		</select>
	  </td>
	</tr>
	<tr>
	  <td>Email Address</td>
	  <td width="1%">:</td>
	  <td class="FontContent"><%=EmailAddress_%></td>
	</tr>
<!--
	<tr>
	  <td>Alternate Email</td>
	  <td width="1%">:</td>
	  <td><input type="input" name="txtAlternateEmail" size="50" Value='<%=AlternateEmail_%>' /></td>
	</tr>
-->
	<tr>
	  <td valign="top">Remark</td>
	  <td width="1%" valign="top">:</td>
	  <td><textarea name="txtRemark" cols="60" rows="3"><%=Remark_ %></textarea></td>
	</tr>
	<tr>
	  <td>Shared Phone ?</td>
	  <td width="1%">:</td>
	  <td>
		  <select name="ShareFlagList">
			<option value="Y" <%if ShareFlag_ ="Y" Then %>Selected<%End If%>>Yes</option>
			<option value="N" <%if ShareFlag_ ="N" Then %>Selected<%End If%>>No</option>
		  </select>
	  </td>
	</tr>
	<tr>
	  <td>Bill Charged ?</td>
	  <td width="1%">:</td>
	  <td>
		  <select name="BillFlagList">
			<option value="Y" <%if BillFlag_ ="Y" Then %>Selected<%End If%>>Yes</option>
			<option value="N" <%if BillFlag_ ="N" Then %>Selected<%End If%>>No</option>
		  </select>
	  </td>
	</tr>
	<tr>
	  <td>&nbsp;</td>
	  <td colspan="2"><input type="submit" name="btnSubmit" value="Submit">
		<%if State_= "E" then %>
		      <input type="hidden" name="txtId" value=<%=Id_ %>>
		<%End If%>
	      <input type="hidden" name="State" value=<%=State_ %> >
	      &nbsp;<input type="button" value="Cancel" name="btnCancel" onClick="Javascript:history.go(-1)">
	 </td>
	</tr>  
	<tr>
		<td colspan=3>&nbsp;</td></tr>
	</table>
    <%else %>
	<table>
	<tr>
		<td>You do not have permission to access this site.</td>
	</tr>
	<tr>
		<td>Please <a href="http://zagrebws03.eur.state.sbu/WebPASS/eservices/MainPage.asp">Submit Request </a> or contact Zagreb ISC Helpdesk at ext.3333.</td>
	</tr>
	</table>
<%   end if 
else %>
	<table align="center">
		<tr>
			<td>You do not have permission to access this site.</td>
		</tr>
		<tr>
			<td>Please <a href="http://zagrebws03.eur.state.sbu/WebPASS/eservices/MainPage.asp">Submit Request </a> or contact Zagreb ISC Helpdesk at ext.3333.</td>
		</tr>
	</table>

<%end if %>
</form>
</BODY>
</html>