
<!--#include file="connect.inc" -->
<html>
<head>
<script language="vbscript">
       <!--
        Sub btnCancel_onclick
           history.back
	End Sub

       --> 
   </script>

<% 
 dim user_ 
 dim user1_  

 
 user_ = request.servervariables("remote_user") 
 user1_ = user_  'user1_ = right(user_,len(user_)-4)
'user1_ = "pranataw"
'response.write user1_ & "<br>"

%> 

<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">EXCHANGE RATE UPDATE</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
<form method="post" name="frmExchangeRate" action="ExchangeRateSave.asp"> 
<%  
 dim rst 
 dim strsql
 dim rst1
 dim today_


 curMonth_ = month(date())
 curYear_ = year(date())
 if len(curMonth_)= 1 then
	curMonth_ = "0" & curMonth_
 end if

' response.write curMonth_
' response.write curYear_

 ExchangeID_ = request("ExchangeID")

 if ExchangeID_ = "" then
	 ExchangeDate_ = date()
 else
	 ExchangeDate_ = request("ExchangeDate")
 end if

 ExchangeMonth_ = request.form("ExchangeMonth")

 if ExchangeMonth_ = "" then
	 ExchangeMonth_ = curMonth_
 end if

 ExchangeYear_ = request.form("ExchangeYear")
 if ExchangeYear_ = "" then
	 ExchangeYear_ = curYear_ 
 end if

 ExchangeRate_ = request("ExchangeRate")
  
 State_ = request("State")
' State_ = "I"
  strsql = "select * from Users where loginId='" & user1_ & "'"
'response.write user1_ & "<br>"
'response.write strsql 
  set rst = server.createobject("adodb.recordset") 
  set rst = BillingCon.execute(strsql)
  if not rst.eof then 
     if trim(rst("RoleID")) = "Admin" or trim(rst("RoleID")) = "FMC" or trim(rst("RoleID")) = "Voucher" then 
	if State_ = "E" then
	        strsql = " select * from ExchangeRate where ExchangeID = '" & ExchangeID_ & "'"
       		set rst1 = server.createobject("adodb.recordset") 
		'response.write strsql 
        	set rst1 = BillingCon.execute(strsql)
	       	if not rst1.eof then 
        	   ExchangeMonth_ = rst1("ExchangeMonth") 
        	   ExchangeYear_ = rst1("ExchangeYear") 
		   ExchangeRate_ = rst1("ExchangeRate")  
        	end if
       	end if
'response.write State_ & "<br>"
%>             
<table align=center>  
<tr>
	<td>Period :</td>				
	<td>
		<Select name="MonthList">
			<Option value="01" <%if ExchangeMonth_ ="01" then %>Selected<%End If%> >January</Option>
			<Option value="02" <%if ExchangeMonth_ ="02" then %>Selected<%End If%> >February</Option>
			<Option value="03" <%if ExchangeMonth_ ="03" then %>Selected<%End If%> >March</Option>
			<Option value="04" <%if ExchangeMonth_ ="04" then %>Selected<%End If%> >April</Option>
			<Option value="05" <%if ExchangeMonth_ ="05" then %>Selected<%End If%> >May</Option>
			<Option value="06" <%if ExchangeMonth_ ="06" then %>Selected<%End If%> >June</Option>
			<Option value="07" <%if ExchangeMonth_ ="07" then %>Selected<%End If%> >July</Option>
			<Option value="08" <%if ExchangeMonth_ ="08" then %>Selected<%End If%> >August</Option>
			<Option value="09" <%if ExchangeMonth_ ="09" then %>Selected<%End If%> >Sepetember</Option>
			<Option value="10" <%if ExchangeMonth_ ="10" then %>Selected<%End If%> >October</Option>
			<Option value="11" <%if ExchangeMonth_ ="11" then %>Selected<%End If%> >November</Option>
			<Option value="12" <%if ExchangeMonth_ ="12" then %>Selected<%End If%> >December</Option>
		</Select>&nbsp;
<%
		Year_ = Year(Date()) - 1
'					response.write YearP_
%>

		<Select name="YearList">
<% 				Do While Year_ <= Year(Date()) %>
					<Option value='<%=Year_%>' <%if trim(Year_) = trim(ExchangeYear_) then %>Selected<%End If%> ><%=Year_%></Option>		
<% 
				Year_ = Year_ + 1
				Loop %>	
		</Select>										
	</td>
</tr>
<tr>
  <td>Exchange Rate :</td>
  <td><input name="txtExchangeRate" value='<%=ExchangeRate_ %>' size="10" align="right" />
  </td>
</tr>

<tr>
  <td></td>
  <td><input type="submit" name="btnSubmit" value="Submit">
<%if State_= "E" then %>
      <input type="hidden" name="txtExchangeID" value=<%=ExchangeID_ %>>
<%End If%>
      <input type="hidden" name="txtState" value=<%=State_ %> >
      &nbsp;<input type="button" value="Cancel" name="btnCancel">
 </td>
</tr>  
<tr><td colspan=2>&nbsp;</td></tr>
</table>
<%
   else 
%>
	<table>
		<tr>
			<td>You do not have permission to access this site.</td>
		</tr>
		<tr>
			<td>Please <a href="http://zagrebws03.eur.state.sbu/WebPASS/eservices/MainPage.asp">Submit Request </a> or contact Zagreb ISC Helpdesk at ext.3333.</td>
		</tr>
	</table>
 
<%   end if 
else %>
	<table align="center">
		<tr>
			<td>You do not have permission to access this site.</td>
		</tr>
		<tr>
			<td>Please <a href="http://zagrebws03.eur.state.sbu/WebPASS/eservices/MainPage.asp">Submit Request </a> or contact Zagreb ISC Helpdesk at ext.3333.</td>
		</tr>
	</table>

<%
end if 
%>
</form>
</BODY>
</html>