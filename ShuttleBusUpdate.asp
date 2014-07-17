<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<% 
	MonthP = Request("MonthP") 
	YearP = Request("YearP")
	EmpID = Request("EmpID")
	Search = Request("Search")
	Used = Request("Used")
	PageIndex = Request("PageIndex")
	strSQL="Exec  spShuttleBusUsedUpdate '" & EmpID & "','" & MonthP & "','" & YearP & "'," & Used
'	response.write strsql 
       	BillingCon.execute strsql
	response.redirect("InputShuttleBusUsage.asp?PageIndex="& PageIndex & "&MonthP=" & MonthP & "&YearP=" & YearP & "&EmpID="& Search)
%>               