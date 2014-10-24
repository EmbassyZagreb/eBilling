<%
	Dim BillingCon
	Set BillingCon= Server.CreateObject("ADODB.Connection")
	BillingCon.Open "Driver={SQL Server};Server=ZAGREBDIN-DC11\UTILITIES;Database=dev-eBilling;uid=ZGBeBilling;pwd=Pass9$pass;Persist Security Info=False;Network=TCP/IP;connection timeout=120"

	Dim BillingDL
	BillingDL ="zgbphonebill@state.gov"

	Dim SMTPServer
	SMTPServer = "nccsmtprelay.irm.state.gov"

	WebSiteAddress = "http://192.168.2.8:8080/dev-eBilling"

	Dim CashierInfo
	CashierInfo = "Embassy Compound TOB: Tuesday and Friday, 09:00-12:00"

	'Wording that must be added in CallType column (dbo.CellphoneDt table) during monthly import procedure if exemption is required
	Dim AlwaysExemptedCallType_ 	
	AlwaysExemptedCallType_ = "(Always Exempted)"

	Dim ExemptedIfOfficialCallType_ 	
	ExemptedIfOfficialCallType_ = "(Exempted If Official)"
%>
