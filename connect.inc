<%
	Const FOR_READING = 1
	Dim flagfile, filecontent, maintenancemode
	flagfile = server.MapPath(".") & "\maintenance\delete2getalive.log"


	filecontent = ""
	maintenancemode = False
	Set fso = CreateObject("Scripting.FileSystemObject")	
	If (fso.FileExists(flagfile)) Then
		maintenancemode = True
		Set objInputFile = fso.OpenTextFile(flagfile, FOR_READING)
		filecontent	= objInputFile.readLine
		objInputFile.close 
	End If	
	Set fso = nothing

	If request.servervariables("remote_user") <> filecontent and maintenancemode = True Then
		If Request.ServerVariables("URL") <> "/test-eBilling/1SystemDown.asp" Then 
			Response.AddHeader "REFRESH","0;URL=1SystemDown.asp"
		Else
			Response.AddHeader "REFRESH","15;URL=1SystemDown.asp"
		End If		
	Else
		Dim BillingCon
		Set BillingCon= Server.CreateObject("ADODB.Connection")
		BillingCon.Open "Driver={SQL Server};Server=ZAGREBAP02\UTILITIES;Database=test-eBilling;uid=ZGBeBilling;pwd=Pass9$pass;Persist Security Info=False;Network=TCP/IP;connection timeout=120"		
	End If		

	Dim BillingDL
	BillingDL ="zBillingTest@state.gov"

	Dim SMTPServer
	SMTPServer = "nccsmtprelay.irm.state.gov"

	WebSiteAddress = "http://zagrebws02.eur.state.sbu:8080/test-eBilling"

	Dim CashierInfo
	CashierInfo = "Embassy Compound TOB: Tuesday and Friday, 09:00-12:00"

	'Wording that must be added in CallType column (dbo.CellphoneDt table) during monthly import procedure if exemption is required
	Dim AlwaysExemptedCallType_ 	
	AlwaysExemptedCallType_ = "(Always Exempted)"

	Dim ExemptedIfOfficialCallType_ 	
	ExemptedIfOfficialCallType_ = "(Exempted If Official)"

	Dim SiteHeader
	SiteHeader = "Test Site Version"
%>
