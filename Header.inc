<!-- <body onload=window.location='maintenance/systemdown.html'> -->
<BODY BGCOLOR="#FFFFFF" TEXT="#000000" LINK="#0080FF" ALINK="0080FF" VLINK="#0055AA" MARGINWIDTH="8" MARGINHEIGHT="0" LEFTMARGIN="8" TOPMARGIN="0">
  <Center><FONT COLOR=#009900><B>SENSITIVE BUT UNCLASSIFIED</Center></FONT></B>
  <BR>
<CENTER>
  <IMG SRC="images/embassytitle2.jpeg" WIDTH="661" HEIGHT="80" BORDER="0"> 
  <TABLE WIDTH="65%" BORDER="0" CELLPADDING="0" CELLSPACING="0">
  <CAPTION><H3 STYLE="font-size:17px;color:#000040">Mission Zagreb - zBilling Application</H3></CAPTION>

<% If maintenancemode Then %>
	<CAPTION id="maintenancemode"><h1><%=SiteHeader%> - MAINTENANCE MODE - Website not available for the public!</h1></CAPTION>
<% Else %>
	<CAPTION><H3 STYLE="font-size:19px;color:red"><%=SiteHeader%></H3></CAPTION>
<% End If %>