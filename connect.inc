<%
	Dim BillingCon
	Set BillingCon= Server.CreateObject("ADODB.Connection")
	BillingCon.Open "Driver={SQL Server};Server=ZAGREBAP02\UTILITIES;Database=eBilling;uid=ZGBeBilling;pwd=Pass9$pass;Persist Security Info=False;Network=TCP/IP;connection timeout=300"

	Dim BillingDL
	BillingDL ="zgbphonebill@state.gov"

	Dim SMTPServer
	SMTPServer = "nccsmtprelay.irm.state.gov"

	WebSiteAddress = "http://zagrebws02.eur.state.sbu:8080/eBilling"

	Dim CashierInfo
	CashierInfo = "Embassy Compound TOB: Tuesday and Friday, 09:00-12:00"

	'Wording that must be added in CallType column (dbo.CellphoneDt table) during monthly import procedure if exemption is required
	Dim AlwaysExemptedCallType_ 	
	AlwaysExemptedCallType_ = "(Always Exempted)"

	Dim ExemptedIfOfficialCallType_ 	
	ExemptedIfOfficialCallType_ = "(Exempted If Official)"
%>