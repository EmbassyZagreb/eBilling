<%@ Language=VBScript%>
<% ' VI 6.0 Scripting Object Model Enabled %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"> 

<head>

<!--#include file="connect.inc" -->

<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</head>
<BODY>
<%	
	Extension_ = Request.Form("txtExtension")
	response.write Extension_ & "<br>"
	MonthP_ = Request.Form("txtMonthP")
	response.write MonthP_ & "<br>"
	YearP_ = Request.Form("txtYearP")
	response.write YearP_ & "<br>"

	strsql = "Update BillingDt Set isPersonal='' Where Extension='" & Extension_ & "' And MonthP='" & MonthP_ & "' And YearP='" & YearP_ & "'"
'	response.write strsql & "<Br>"  
	BillingCon.execute(strsql) 

	For Each loopIndex in Request.Form("cbPersonal")
'	response.write loopIndex
				
		strsql = "Update BillingDt Set isPersonal='Y' Where Extension='" & Extension_ & "' And MonthP='" & MonthP_ & "' And YearP='" & YearP_ &"' And CallRecordID= " & loopIndex 
		response.write strsql & "<Br>"  
		BillingCon.execute(strsql) 
	next

	response.redirect("OfficePhoneBilling.asp")
%>
</body> 

</html>
