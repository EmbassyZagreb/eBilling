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

<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
<% 
 dim user_ 
 dim user1_  

 
 user_ = request.servervariables("remote_user") 
 user1_ = right(user_,len(user_)-4)
'user1_ = "pranataw"
'response.write user1_ & "<br>"

if (session("Month") = "") or (session("Year") = "") then
	strsql = "Select MonthP, YearP From Period"
	'response.write strsql & "<br>"
	set rsData = server.createobject("adodb.recordset") 
	set rsData = BillingCon.execute(strsql)
	if not rsData.eof then
		session("Month") = rsData("MonthP")
		session("Year") = rsData("YearP")
	end if
end if
%> 

</head>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">PHONE LIST</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
<form method="post" action="PhoneSave.asp"> 
<table align="center" cellspadding="1" cellspacing="0" width="100%">  
<tr>
	<td colspan="2" align="center">Billing Period : <Label style="color:blue"><%=session("Month")%> - <%=session("Year")%></lable></td>
</tr>
<tr>
	<td colspan="2"><br></td>
</tr>
</table>
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
     if trim(rst("RoleID")) = "Admin" then  
	if State_ = "E" then
	        strsql = " select * from MsPhoneList where ID = " & ID_ 
       		set rst1 = server.createobject("adodb.recordset") 
		'response.write strsql 
        	set rst1 = BillingCon.execute(strsql)
	       	if not rst1.eof then 
        	   PhoneNumber_ = rst1("PhoneNumber") 
		   Location_ = rst1("Location")  
        	end if
       	end if
	'response.write State_ & "<br>"
	%>             
	<table align=center>  
	<tr>
	  <td>Phone Number / Ext. :</td>
	  <td><input type="input" name="txtPhoneNumber" value='<%=PhoneNumber_ %>' size="25" maxlength="20" /></td>
	</tr>
	<tr>
	  <td>Employee Name / Location :</td>
	  <td><input type="input" name="txtLocation" value='<%=Location_ %>' size="50" maxlength="60" /></td>	
	</tr>
	<tr>
	  <td>Phone Type :</td>
	  <td>
		  <select name="PhoneTypeList">
			<option value="O">Office Phone</option>
			<option value="C">Cell Phone</option>
			<option value="H">Home Phone</option>
		  </select>
	  </td>
	</tr>
	<tr>
	  <td>Post :</td>
	  <td>
		  <select name="PostList">
			<option value="ZAGREB">ZAGREB</option>
			<option value="PODGORICA">PODGORICA</option>
		  </select>
	  </td>
	</tr>
	<tr>
	  <td>&nbsp;</td>
	  <td><input type="submit" name="btnSubmit" value="Submit">
		<%if State_= "E" then %>
		      <input type="hidden" name="txtId" value=<%=Id_ %>>
		<%End If%>
	      <input type="hidden" name="State" value=<%=State_ %> >
	      &nbsp;<input type="button" value="Cancel" name="btnCancel">
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