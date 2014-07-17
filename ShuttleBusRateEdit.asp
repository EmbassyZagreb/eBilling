
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
<script type="text/javascript">
function validate_form()
{
	valid = true;
	msg = ""

	if (document.frmShuttleBusRate.txtShuttleBusRate.value == "" )
	{
		msg = msg + "Please fill in AM, minimum value is 0  !!!\n"
		valid = false;
	}
	else
	{
		var myRegExp = new RegExp("^[/+|/-]?[0-9]*[/.]?[0-9]*$");
		if (myRegExp.test(document.frmShuttleBusRate.txtShuttleBusRate.value) == false)
		{
			msg = msg + "Invalid data type for Shuttle Bus Rate!!!\n"
			valid = false;
		}
	}


	if (valid == false)
	{
		alert(msg)
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
  	<TD COLSPAN="4" ALIGN="center" Class="title">SHUTTLE BUS RATE UPDATE</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
<form method="post" align="center" name="frmShuttleBusRate" action="ShuttleBusRateSave.asp" onSubmit="return validate_form()"> 
<%  
 dim rst 
 dim strsql
 dim rst1
 dim today_

 ShuttleBusRateID_ = request("ShuttleBusRateID")
 ShuttleBusRate_ = request("ShuttleBusRate")
  
 State_ = request("State")
' State_ = "I"
  strsql = "select * from Users where loginId='" & user1_ & "'"
'response.write user1_ & "<br>"
'response.write strsql 
  set rst = server.createobject("adodb.recordset") 
  set rst = BillingCon.execute(strsql)
  if not rst.eof then 
     if trim(rst("RoleID")) = "Admin" or trim(rst("RoleID")) = "Cashier" or trim(rst("RoleID")) = "FMC" then 
	if State_ = "U" then
	        strsql = " select * from ShuttleBusRate where ShuttleBusRateID = '" & ShuttleBusRateID_ & "'"
       		set rst1 = server.createobject("adodb.recordset") 
		'response.write strsql 
        	set rst1 = BillingCon.execute(strsql)
	       	if not rst1.eof then 
        	   ShuttleBusRateDate_ = rst1("ShuttleBusRateDate") 
		   ShuttleBusRate_ = rst1("ShuttleBusRate")  
        	end if
	Else
        	   ShuttleBusRateDate_ = Date()
       	end if
'response.write State_ & "<br>"
%>             
<table align="center" cellpadding="1" cellspacing="0" width="50%">
<tr>
	<td align="right">Rate Date</td>
	<td width="1%">:</td>
	<td class="FontContent"><%=ShuttleBusRateDate_ %></td>
</tr>
<tr>
	<td align="right">Shuttle Bus Rate</td>
	<td width="1%">:</td>
	<td><input name="txtShuttleBusRate" value='<%=ShuttleBusRate_ %>' size="5" />
  </td>
</tr>
<tr>
  <td colspan="3"><br></td>
<tr>

  <td colspan="3" align="center"><input type="submit" name="btnSubmit" value="Submit">
<%if State_= "U" then %>
      <input type="hidden" name="txtShuttleBusRateID" value=<%=ShuttleBusRateID_ %>>
<%End If%>
      <input type="hidden" name="txtState" value=<%=State_ %> >
      &nbsp;<input type="button" value="Cancel" name="btnCancel">
 </td>
</tr>  
<tr><td colspan="3">&nbsp;</td></tr>
</table>
<%
   else 
%>
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

<%
end if 
%>
</form>
</BODY>
</html>