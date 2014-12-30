
<!--#include file="connect.inc" -->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>

<script language="JavaScript" src="calendar.js"></script>

<% 
'response.buffer=false



 dim user_ 
 dim user1_  

 
 user_ = request.servervariables("remote_user") 
 user1_ = user_  'user1_ = right(user_,len(user_)-4)
'user1_ = "pranataw"
'response.write user1_ & "<br>"

	Set fso = CreateObject("Scripting.FileSystemObject")

    Func = Request("Func")
    if isempty(Func) Then
    	Func = 1
    End if
    Select Case Func
    Case 1
	


%> 

<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">Maintenance Mode</TD>
   </TR>
	<tr>
        	<td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
	</tr>
	<TR>
		<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
	</TR>
  </TABLE>
<form method="post" name="frmPaymentDueDate" id="frmPaymentDueDate" action="Maintenancemode.asp?Func=2">
<% 
 
 dim rst 
 dim strsql
 dim Status_
 
 strsql = "select * from Users where loginId='" & user1_ & "'" 
  set rst = server.createobject("adodb.recordset") 
  set rst = BillingCon.execute(strsql)
  if not rst.eof then 
	if (trim(rst("RoleID")) = "Admin")  then

		If (fso.FileExists(flagfile)) Then
			Status_ = "Down"
		Else
			Status_ = "Up"
		End If


%>             
		<table align=center>  

		<tr>
		  <td align=right>Turn zBilling System in Maintenance mode :</td>
		  <td>	
		  <select name="cmbStatus">
			<option value="Down" <%if Status_ ="Down" Then %>Selected<%End If%>>Yes</option>
			<option value="Up" <%if Status_ ="Up" Then %>Selected<%End If%>>No</option>
		  </select>
		  </td>
		</tr>
		<tr>
			<td colspan="2"><br></td>
		</tr>
		<tr>
		  <td></td>
		  <td><input type="submit" id="btnSubmit" name="btnSubmit" value="Submit">
		      &nbsp;<input type="button" value="Cancel" name="btnCancel">
		 </td>
		</tr>  
		<tr>
		   <td colspan=2>&nbsp;</td>
		</tr>
		<tr>
		   <td colspan=2>SQL Server requires at least 60 seconds to automatically close open connections.</td>
		</tr>
		</table>
     <%  else %>
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

<% end if %>
</form>
</BODY>
</html>

<%
    Case 2
    ForWriting = 2
	
	Status_ =  trim(request.form("cmbStatus"))

	if Status_ = "Down" Then
		If (fso.FileExists(flagfile)) = False Then
			Set aFile = fso.CreateTextFile(flagfile, True)	
			aFile.WriteLine(user_)
			aFile.Close
		End If
	Else
		If (fso.FileExists(flagfile)) = True Then
			fso.DeleteFile(flagfile)
		End If	
	End If
	
    		Set aFile = nothing
    		Set fso = nothing
	
	Response.AddHeader "REFRESH","0;URL=MaintenanceMode.asp"	

    End Select
 %>	