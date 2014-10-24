<%@ Language=VBScript %>
<!--#include file="connect.inc" -->

<% dim LoginId_
ID_ = Request.QueryString("ID")
AgencyDesc = Request.QueryString("AgencyDesc")
%>

<html>
   <head>
<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  NAME="CDO for Windows 2000 Library" --> 
<link href="style.css" rel="stylesheet" type="text/css">
   </head>

<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="SubTitle">Delete Agency</TD>
  </TR>
  <tr>
	<td colspan="3" align="Left" width="20%"><A HREF="Default.asp">Home</A></td>
	<td align="Right" width="20%"><A HREF="AgencyList.asp">Back</A></td>
  </tr>  
  <tr>
  	<td COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></td>
   </tr>

   </table>
   
   
   
   <%
strsql = "SELECT EmpName, COUNT(*) As Total FROM vwDirectReport where AgencyID='" & ID_ & "' Group By EmpName"
set RS_Query = server.createobject("adodb.recordset")
'response.write strsql & "<br>"
set RS_Query = BillingCon.execute(strsql)

	if not RS_Query.eof then
	EmpName_ = RS_Query("EmpName")
	Total_ = RS_Query("Total")
	End If
%>
   
   
   
<form method="post" id="frmAgencyDelete" name="frmAgencyDelete" action="AgencySave.asp?Mode=D"> 
   <table>
   <tr>
   
<%If Total_>0 Then%>
	<td colspan="4" align=center>Employee <%=EmpName_%> has been assigned to <font color=blue><strong><%=AgencyDesc %></strong></font>. Agency cannot be deleted. </td>
<%Else%>
	<td colspan="4" align=center>Agency : <font color=blue><strong><%=AgencyDesc %></strong></font> will be deleted, Continue ?</td>
<%End If%> 
   
   

   </tr>
   <tr>
	<td colspan="4"><br></td>
   </tr>

   <tr>
	<td colspan="4" align=center>
		<%If Total_=0 Then%>
		<input type="Submit" value="Yes" id="btnDelete"> 
		<INPUT TYPE="HIDDEN" NAME="txtID" value=<%=ID_%>>
			<%End If%> 
		<input type="button" value="Cancel" id="btnCancel" onclick="self.history.back()"> 
	</td>
   </tr>
    <tr>
	<td colspan="4">
		<INPUT TYPE="HIDDEN" NAME=LoginID value='<%=LoginID_%>'>
	</td>
    </tr>
</table>
</form>
</body>
</html>