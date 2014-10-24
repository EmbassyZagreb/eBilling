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
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>

<meta http-equiv="refresh" content="1;url=AgencyList.asp">

<%
dim CTL , LoginID_ , UserRole_
Dim AgencyCode_, Appropriation_, Allotment_, ObligationNo_, OrgCode_, Function_, Object_, AwardAmount_

'Mode = Request.Form("Mode")
Mode = Request.QueryString("Mode")
if Mode="I" then
	ID_ = 0
Else
	ID_ = Request.Form("txtID") 
end if
'response.write Mode & "<br>"
AgencyCode_ = Request.Form("txtAgencyCode")
'response.write AgencyCode_ & "<br>"
AgencyDesc_ = Request.Form("txtAgencyName")
AgencyStripe_ = Request.Form("txtAgencyStripe")
AgencyStripeNonVAT_ = Request.Form("txtAgencyStripeNonVAT")
AgencyType_ = Request.Form("txtAgencyType")
%>
   </head>

<!--#include file="Header.inc" -->
<tr>
	<td colspan="2"><HR style="LEFT: 10px; TOP: 59px" align=center></td>
</tr>
</table>
<table border=0 width=100%>
<%
       strsql = "Exec spAgency_IUD '" & Mode & "'," & ID_  & ",'" & AgencyCode_ & "','" & AgencyDesc_ & "','" & AgencyStripe_ & "','" & AgencyStripeNonVAT_ & "','" & AgencyType_ & "'"
       'Response.Write 	strsql
       BillingCon.execute strsql
%>               
<%If Mode="D" Then%>
	<tr><td align=center>Your data has already Deleted. Thank you.</td></tr>
<%Else%>
	<tr><td align=center>Your data has already saved. Thank you.</td></tr>
<%End If%>
<tr><td align=center> 
<!-- <input type="button" value="Close" id="btnclose"> -->
</td></tr>
<tr>
	<td align="center"><br><a href="AgencyList.asp"><img src="images/Back.gif" border="0" alt="Go..Back" WIDTH="83" HEIGHT="25"></a></td>
</tr>
<%
'response.redirect("AgencyList.asp?msg=List of Agency updated !!!")
%>
</table>

   </body>
</html>