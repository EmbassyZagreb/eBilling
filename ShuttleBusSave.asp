<%@ Language=VBScript %>
<!--#include file="connect.inc" -->

<html>
   <head>
   <script language="vbscript">
       <!--
        Sub btnBack_onclick
           history.back
	End Sub
        Sub btnClose_onclick
		close
	End Sub
       --> 
   </script>
<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  NAME="CDO for Windows 2000 Library" --> 
<%
ShuttleID_ = trim(request.form("ShuttleID"))
if ShuttleID_ = "" then
	ShuttleID_ = 0
end if
State_ = trim(request.form("State"))
EmpID_ = trim(request.form("cmbEmpID"))
TransportDate_ = trim(request.form("txtTransportDate"))
EventType_ = trim(request.form("cmbEventType"))
Qty_ = trim(request.form("txtQty"))
if Qty_ = "" then
	Qty_ = 0
end if

'response.write LoginID_ & "<br>"
'response.write UserRole_ & "<br>"

%>
   </head>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">Shuttle Bus Payment</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>

<table border=0 width=100%>
<%
       strsql = "Exec spShuttleBill_IUD '" & State_  & "'," & ShuttleID_ & ",'" & EmpID_ & "','" & TransportDate_ & "','" & EventType_ & "'," & Qty_  
      'response.write strsql 
       BillingCon.execute strsql
%>               
<tr><td align=center>Your data has already been saved. Thank you.</td></tr>
<tr><td>&nbsp;</td>
<tr><td align=center> 
<input type="button" value="Close" id="btnclose">
</td></tr>
<tr>
	<td align="center"><br><a href="ShuttleBusList.asp"><img src="images/Back.gif" border="0" alt="Go..Back" WIDTH="83" HEIGHT="25"></a></td>
</tr>
</table>

   </body>
</html>