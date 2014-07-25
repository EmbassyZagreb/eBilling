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
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">USER UPDATE</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>

<form method="post" action="UserSave.asp"> 
<%  
 dim rst 
 dim strsql
 dim rst1
 dim today_


 today_ = now()

 LoginID_ = request("LoginID")
 State_ = request("State")
' State_ = "I"
  strsql = "select * from Users where loginId='" & user1_ & "'"
'response.write user1_ & "<br>"
'response.write strsql 
  set rst = server.createobject("adodb.recordset") 
  set rst = BillingCon.execute(strsql)
  if not rst.eof then 
     if trim(rst("RoleID")) = "Admin" then  
        strsql = " select * from vwUserList where LoginID = '" & LoginID_ & "'"
       	set rst1 = server.createobject("adodb.recordset") 
	'response.write strsql 
        set rst1 = BillingCon.execute(strsql)
       	if not rst1.eof then 
           UserRole_ = rst1("RoleID") 
	   EmployeeName_ = rst1("EmployeeName")  
        end if
'response.write State_ & "<br>"
%>             
<table align=center>  
<%if State_ ="I" then %>
<tr>
  <td>User Name :</td>
  <td>
	<select id="LoginID" name="LoginID">
		<option value="">-- Select --</option>
<%
		Dim UserRS
		'strsql = "select * from [JakartaAP02\SQL05].Jakemployee.dbo.Employees where len(loginId)>2 And LoginID not in (Select LoginID From Users) order by LastName"
		strsql = "select * from MSEmployee where len(LoginID)>2 order by EmpName"
		response.write strsql & "<br>"
		set UserRS = server.createobject("adodb.recordset")
		set UserRS =BillingCon.execute(strsql)				        
		do while not UserRS.eof
			Ename_ = "" 
EName_ = UserRS("EmpName") 
			'if ltrim(UserRS("FirstName")) = "" then
			'	EName_ = UserRS("EmpName") 
			'else
   			'        EName_ = UserRS("EmpName") & ", " & UserRS("FirstName") 	
			'end if 
			'	Ename_ = EName_ & "(" & UserRS("OfficeLocation") & ")"
%>
	        <OPTION value='<%=UserRS("LoginID")%>'>  <%= EName_  %>
<%
                 UserRS.movenext
	        loop
%>  
	</select>
  </td>
</tr>
<%Else%>
<tr>
  <td>Login ID :</td>
  <td><label class="FontContent"><%=loginId_  %></label>
  </td>
</tr>
<tr>
  <td>Employee Name :</td>
  <td><label class="FontContent"><%=EmployeeName_ %></label>
  </td>
</tr>
<%End If%>
<tr>
  <td>User Role :</td>
  <td>
  <select name="userRole">
<%
		Dim RoleRS
		strsql = "select RoleID, RoleName from UserRoles Order By RoleName"
		'response.write strsql & "<br>"
		set RoleRS= server.createobject("adodb.recordset")
		set RoleRS= BillingCon.execute(strsql)				        
		do while not RoleRS.eof
%>
		<% if UserRole_ =RoleRS("RoleID") then%>
	              <option value=<%=RoleRS("RoleID")%> Selected><%=RoleRS("RoleName")%></option>
		<%else%>
	              <option value='<%=RoleRS("RoleID")%>'><%=RoleRS("RoleName")%></option>
		<%end if%>
<%		
		RoleRS.movenext
		loop
%>
  </select>
  </td>
</tr>  
<tr>
  <td></td>
  <td><input type="submit" name="btnSubmit" value="Submit">
<%if State_= "E" then %>
      <input type="hidden" name="LoginID" value=<%=loginId_ %>>
<%End If%>
      <input type="hidden" name="State" value=<%=State_ %> >
      &nbsp;<input type="button" value="Cancel" name="btnCancel">
 </td>
</tr>  
<tr><td colspan=2>&nbsp;</td></tr>
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

<%
rst.close
rst1.close
%>

</form>
</BODY>
</html>