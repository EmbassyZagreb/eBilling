<%@ Language=VBScript%>
<!--#include file="connect.inc" -->
<html>
<%
dim transID, translation, sorting
transID = request.querystring("ID")
translation = request.querystring("translation")

dim rs
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "SELECT * from CallTypeTranslation  where TransID = " & transID & ";", BillingCon, 1,3
rs("English") = translation
rs.update
rs.close
Response.AddHeader "REFRESH","0;URL=TranslationTable.asp"

%>
</html>