<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<html>
<head>


<% 
 dim user_ 
 dim user1_  

 user_ = request.servervariables("remote_user") 
 user1_ = right(user_,len(user_)-4)
'user1_ = "pranataw"
'response.write user1_ & "<br>"

EmpID_ = trim(Request("EmpID"))
'response.write "HomePhone_  :" & HomePhone_ & "<br>"
MonthP_ = Request("MonthP")
'response.write MonthP_ & "<br>"
YearP_ = Request("YearP")
'response.write YearP_ & "<br>"

%> 
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">UPDATE BILLING STATUS</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>

<form method="post" action="ChangeProgressStatusSave.asp" name="frmStatus"> 
<%  
	 strsql = "Select * from vwMonthlyBilling Where EmpID='" & EmpID_ & "' And MonthP='" & MonthP_ & "' And YearP='" & YearP_ & "'"
	'response.write strsql & "<br>"
	set rsData = server.createobject("adodb.recordset") 
	set rsData = BillingCon.execute(strsql) 
	Period_ = MonthP & " - " & YearP
	'response.write Period_  & "<br>"
	if not rsData.eof then
		EmpID_ = rsData("EmpID")
		EmpName_ = rsData("EmpName")
		Office_ = rsData("Agency") & " - " & rsData("Office")
		Position_ = rsData("WorkingTitle")
		OfficePhone_ = rsData("WorkPhone")
		HomePhone_ = rsData("HomePhone")
		MobilePhone_ = rsData("MobilePhone")
		HomePhonePrsBillRp_ = rsData("HomePhonePrsBillRp")
		OfficePhonePrsBillRp_ = rsData("OfficePhonePrsBillRp")
		CellPhonePrsBillRp_ = rsData("CellPhonePrsBillRp")
		TotalShuttleBillRp_ = rsData("TotalShuttleBillRp")
		ProgressID_ = rsData("ProgressID")
		TotalBillingPrsAmount_ = rsData("TotalBillingAmountPrsRp")
	'	response.write "TotalBilling: " & TotalBillingPrsAmount_ & "<br>"
	end if
%>             
<table align="center">
<tr>
  <td>Employee Name</td>
  <td width="1px">:</td>
  <td><label class="FontContent"><%=EmpName_ %></label></td>
</tr>
<tr>
  <td>Office</td>
  <td>:</td>
  <td><label class="FontContent"><%=Office_ %></label>
  </td>
</tr>
<tr>
  <td>Position</td>
  <td>:</td>
  <td><label class="FontContent"><%=Position_ %></label>
  </td>
</tr>
<tr>
  <td>Office Phone</td>
  <td>:</td>
  <td><label class="FontContent"><%=OfficePhone_ %></label>
  </td>
</tr>
<tr>
  <td>Home Phone</td>
  <td>:</td>
  <td><label class="FontContent"><%=HomePhone_%></label>
  </td>
</tr>
<tr>
  <td>Cell Phone</td>
  <td>:</td>
  <td><label class="FontContent"><%=MobilePhone_%></label>
  </td>
</tr>
<tr>
  <td>Home Phone bill</td>
  <td>:</td>
  <td><label class="FontContent"><%= formatnumber(HomePhonePrsBillRp_ ,-1) %></label>
  </td>
</tr>
<tr>
  <td>Office Phone bill</td>
  <td>:</td>
  <td><label class="FontContent"><%= formatnumber(OfficePhonePrsBillRp_,-1) %></label>
  </td>
</tr>
<tr>
  <td>Cell Phone bill</td>
  <td>:</td>
  <td><label class="FontContent"><%= formatnumber(CellPhonePrsBillRp_,-1) %></label>
  </td>
</tr>
<tr>
  <td>Invoice Status</td>
  <td>:</td>
  <td>
<%
 	strsql ="select ProgressID, ProgressDesc from ProgressStatus"
	set StatusRS = server.createobject("adodb.recordset")
	set StatusRS = BillingCon.execute(strsql)
%>	
	<Select name="cmbStatus">
<%	Do While not StatusRS.eof %>
		<Option value='<%=StatusRS("ProgressID")%>' <%if trim(ProgressID_) = trim(StatusRS("ProgressID")) then %>Selected<%End If%> ><%=StatusRS("ProgressDesc")%></Option>				
<%		StatusRS.MoveNext
	Loop%>
	</select>
	</td>	
</tr>
<tr>
  <td colspan="3"><br></td>
</tr>
<tr>
  <td></td>
  <td></td>
  <td><input type="submit" name="btnSubmit" value="Update">
      <input type="hidden" name="txtEmpID" value='<%=EmpID_  %>'>
      <input type="hidden" name="txtMonthP" value='<%=MonthP_  %>'>
      <input type="hidden" name="txtYearP" value='<%=YearP_  %>'>
      &nbsp;<input type="button" value="Cancel" name="btnCancel" onClick="window.close()">
 </td>
</tr>  
<tr><td colspan=2>&nbsp;</td></tr>
</table>
</form>
</BODY>
</html>