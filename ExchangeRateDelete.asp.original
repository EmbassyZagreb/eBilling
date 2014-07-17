<%@ Language=VBScript %>
<!--#include file="connect.inc" -->

<% dim LoginId_
   ExchangeID_ =  trim(request.form("txtExchangeID"))
   State_ =  trim(request.form("txtState"))
'   response.write LoginID_
%>

<html>
   <head>
<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  NAME="CDO for Windows 2000 Library" --> 
   </head>

<BODY bgcolor=white background=images/bg07.gif alink=blue link=blue vlink=blue >
<%
       strsql = "Exec spExchangeRate_IUD '" & State_ & "'," & ExchangeID_ 
'	response.write strsql 
       BillingCon.execute strsql
        response.redirect("ExchangeRateList.asp")		
%>
</body>
</html>