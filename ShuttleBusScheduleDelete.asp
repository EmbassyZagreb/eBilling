<%@ Language=VBScript %>
<!--#include file="connect.inc" -->

<% 
EmpID_ = trim(request.form("txtEmpID"))
ShuttleDate_ = trim(request.form("txtShuttleDate"))
'   response.write LoginID_
%>

<html>
   <head>
<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  NAME="CDO for Windows 2000 Library" --> 
   </head>

<BODY bgcolor=white background=images/bg07.gif alink=blue link=blue vlink=blue >
<%
       strsql = "Delete ShuttleUserSchedule Where EmpID= " & EmpID_ & " And ShuttleDate='" & ShuttleDate_ & "'"
'	response.write strsql 
       BillingCon.execute strsql
        response.redirect("GenerateShuttleBusSchedule.asp")		
%>
</body>
</html>