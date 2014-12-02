<%@ Language=VBScript %>
<!--#include file="connect.inc" -->

<% dim LoginId_
ID_ = Request.QueryString("ID")
Name_ = Request.QueryString("Name")
%>

<html>
   <head>
<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  NAME="CDO for Windows 2000 Library" --> 
<link href="style.css" rel="stylesheet" type="text/css">
   </head>

<!--#include file="Header.inc" -->
  <tr>
	<TD COLSPAN="4" ALIGN="center" Class="title">EMPLOYEE DELETE</TD>
  </tr> 
  <tr>
	<td colspan="3" align="Left" width="20%"><A HREF="Default.asp">Home</A></td>
	<td align="Right" width="20%"><A HREF="AgencyList.asp">Back</A></td>
  </tr>  
  <tr>
  	<td COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></td>
   </tr>
   </table>
   
   
<%
strsql = "SELECT PhoneNumber, COUNT(*) As Total FROM MsCellPhoneNumber where EmpID='" & ID_ & "' Group By PhoneNumber"
set RS_Query = server.createobject("adodb.recordset")
'response.write strsql & "<br>"
set RS_Query = BillingCon.execute(strsql)

	if not RS_Query.eof then
	PhoneNumber_ = RS_Query("PhoneNumber")
	Total_ = RS_Query("Total")
	End If
%>


<form method="post" id="frmAgencyDelete" name="frmAgencyDelete" action="EmployeeSave.asp"> 
   <table>
   <tr>
   
<%If Total_>0 Then%>
	<td colspan="4" align=center>Cellphone <%=PhoneNumber_%> has been assigned to <font color=blue><strong><%=Name_ %></strong></font>. User cannot be deleted. </td>
<%Else%>
	<td colspan="4" align=center>Employee : <font color=blue><strong><%=Name_ %></strong></font> will be deleted, Continue ?</td>
<%End If%> 
 	
   </tr>
   <tr>
	<td colspan="4"><br></td>
   </tr>
   <tr>
	<td colspan="4" align=center>
	<%If Total_=0 Then%>
		<input type="Submit" value="Yes" id="btnDelete"> 	
		<INPUT TYPE="HIDDEN" NAME="txtEmpID" value=<%=ID_%>>
		<INPUT TYPE="HIDDEN" NAME="State" value="X">
	<%End If%> 	
		<input type="button" value="Cancel" id="btnCancel" onclick="self.history.back()"> 
	</td>
   </tr>
    <tr>

    </tr>
</table>
</form>
</body>
</html>