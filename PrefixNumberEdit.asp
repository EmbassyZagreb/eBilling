<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<html>
<head>



<script type="text/javascript">
function validate_form()
{
	valid = true;
	msg = ""

	if (document.frmPrefixNumber.txtPrefix.value == "" )
	{
		msg = msg + "Please fill in Prefix Number !!!\n"
		valid = false;
	}

	if (document.frmPrefixNumber.cmbType.value == "" )
	{
		msg = msg + "Please select Prefix Type !!!\n"
		valid = false;
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
  	<TD COLSPAN="4" ALIGN="center" Class="title">PREFIX NUMBER UPDATE</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
<form method="post" name="frmPrefixNumber" action="PrefixNumberSave.asp" onSubmit="return validate_form()"> 
<%  
 dim rst 
 dim strsql
 dim rst1
 dim today_


 today_ = now()

 PrefixID_ = request("PrefixID")
 State_ = request("State")
' State_ = "I"
  strsql = "select * from Users where loginId='" & user1_ & "'"
'response.write user1_ & "<br>"
'response.write strsql 
  set rst = server.createobject("adodb.recordset") 
  set rst = BillingCon.execute(strsql)
  if not rst.eof then 
     if trim(rst("RoleID")) = "Admin" or (mid(rst("RoleID"),1,3) = "FMC") then  
	if State_ = "U" then
	        strsql = " select * from MsPrefixNumber where PrefixID= " & PrefixID_ 
       		set rst1 = server.createobject("adodb.recordset") 
		'response.write strsql 
        	set rst1 = BillingCon.execute(strsql)
	       	if not rst1.eof then 			
        	   Code_ = rst1("Code") 
		   Prefix_ = rst1("Prefix")  
		   Type_ = rst1("Type")
		   Description_ = rst1("Description")
        	end if
	else
		   Code_ = "-"
       	end if
	'response.write State_ & "<br>"
	%>             
	<table align=center>  
	<tr>
	  <td align="right">Code :</td>
	  <td class="FontContent"><%=Code_ %></td>
	</tr>
	<tr>
	  <td align="right">Prefix :</td>
	  <td><input type="input" name="txtPrefix" value='<%=Prefix_ %>' size="8" maxlength="10" /></td>
	</tr>
	<tr>
	  <td align="right">Type :</td>
	  <td>
<!--		<input type="input" name="txtType" value='<%=Type_ %>' size="40" maxlength="50" /> -->
		<select name="cmbType">
			<option value="">--select--</option>
			<option value="Celluler"  <%if trim(Type_)="Celluler" Then%>Selected<%end if%> >Celluler</option>
			<option value="InterLocal/SLJJ"  <%if trim(Type_)="InterLocal/SLJJ" Then%>Selected<%end if%> >InterLocal/SLJJ</option>
			<option value="International/SLI"  <%if trim(Type_)="International/SLI" Then%>Selected<%end if%> >International/SLI</option>
			<option value="Local"  <%if trim(Type_)="Local" Then%>Selected<%end if%> >Local</option>
		</select>
	  </td>
	</tr>
	<tr>
	  <td align="right">Description :</td>
	  <td><input type="input" name="txtDescription" value='<%=Description_ %>' size="40" maxlength="100" /></td>
	</tr>
	<tr>
	  <td>&nbsp;</td>
	  <td><input type="submit" name="btnSubmit" value="Submit">
		<%if State_= "U" then %>
		      	<input type="hidden" name="txtPrefixID" value=<%=PrefixID_ %>>
			<input type="hidden" name="txtCode" value='<%=Code_ %>' size="8" maxlength="10" />
		<%End If%>
	      <input type="hidden" name="State" value=<%=State_ %> >
	      &nbsp;<input type="button" value="Cancel" name="btnCancel" onClick="Javascript:history.go(-1)">
	 </td>
	</tr>  
	<tr>
		<td colspan=2>&nbsp;</td></tr>
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