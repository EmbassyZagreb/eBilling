<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<html>
<head>
<META name=VI60_defaultClientScript content=VBScript>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script language="JavaScript" src="calendar.js"></script>
<script language="vbscript">
<!--
Sub btnCancel_onclick
          history.back
End Sub
--> 
</script>

<script type="text/javascript">
function validate_form()
{
	valid = true;
	msg = ""

	if (document.frmData.txtQty.value == "" )
	{
		msg = msg + "Please fill in your quantity of person !!!\n"
		valid = false;
	}
	else
	{
		var myRegExp = new RegExp("^[/+|/-]?[0-9]*[/.]?[0-9]*$");
		if (myRegExp.test(document.frmData.txtQty.value) == false)
		{
			msg = msg + "Invalid data type for quantity of person !!!\n"
			valid = false;
		}
	}


	if (valid == false)
	{
		alert(msg)
	}
	return valid;
}
</script>
<% 
 dim user_ 
 dim user1_  

 
 user_ = request.servervariables("remote_user") 
 user1_ = right(user_,len(user_)-4)
'response.write user1_ & "<br>"

%> 
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
   <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">Shuttle Bus Payment</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
<form name="frmData" method="post" action="ShuttleBusSave.asp" onSubmit="return validate_form()"> 
<%  
 dim rsUser
 dim strsql
 dim rsData

 ShuttleID_ = request("ShuttleID")
 State_ = request("State")
 If State_ = "" then
	 State_ = "I"
 End If
  strsql = "select * from Users where LoginID='" & user1_ & "'"
'response.write user1_ & "<br>"
'response.write strsql 
  set UserRS = server.createobject("adodb.recordset") 
  set UserRS = BillingCon.execute(strsql)
  if not UserRS.eof then 
	if (trim(UserRS("RoleID")) = "Admin") or (trim(UserRS("RoleID")) = "Trs") then  
		'response.write State_ & "<br>"
		if State_ ="E" Then
	        	strsql = " select * from vwShuttleBus where ShuttleID = " & ShuttleID_ 
	     		set rsData = server.createobject("adodb.recordset") 
			'response.write strsql 
	        	set rsData = BillingCon.execute(strsql)
		       	if not rsData.eof then 
        		   EmpID_ = rsData("EmpID") 
			   EmpName_ = rsData("EmpName")
			   TransportDate_ = rsData("TransportDate")
			   EventType_ = rsData("EventType")
			   QtyPerson_ = rsData("QtyPerson")
        		end if
		end if
		'response.write EmpID_ 
%>             
	<table align=center>  
	<tr>
	  <td align="right">Employeee Name :</td>
  	  <td>
		<select name="cmbEmpID">
		<option value="">-- Select --</option>
<%
		Dim UserRS
		strsql = "select EmpID, LastName, FirstName, Office from vwPhoneCustomerList order by LastName"
		'response.write strsql & "<br>"
		set UserRS = server.createobject("adodb.recordset")
		set UserRS =BillingCon.execute(strsql)				        
		do while not UserRS.eof
			Ename_ = "" 
			if ltrim(UserRS("FirstName")) = "" then
				EName_ = UserRS("LastName") 
			else
   			        EName_ = UserRS("LastName") & ", " & UserRS("FirstName") 	
			end if 
				Ename_ = EName_ & "(" & UserRS("Office") & ")"

			if trim(UserRS("EmpID")) = trim(EmpID_) then
%>
		        	<OPTION value='<%=UserRS("EmpID")%>' Selected><%= Ename_ %>
<%			Else%>
			        <OPTION value='<%=UserRS("EmpID")%>'>  <%= Ename_ %>
<%			end if

        	        UserRS.movenext
	        loop
%>  
		</select>
	  </td>
	</tr>
	<tr>
		<td align="right">Date :</td>
	  	<td><input name="txtTransportDate" type="Input" size="10" value='<%=TransportDate_ %>' maxlength="10">
			<a href="javascript:cal0.popup();"><img src="images/calendar.gif" width="34" height="18" border="0" alt="Calendar"></a></td>
					
	</tr>
	<tr>
		<td align="right">Time :</td>
		<td>
		    <select name="cmbEventType">
			<Option value="">--Select--</Option>
<%			if EventType_ = "AM" then %>
				<Option value="AM" selected>AM</Option>
<%			else%>
				<Option value="AM">AM</Option>
<%			end if%>
<%			if EventType_ = "PM" then %>
				<Option value="PM" selected>PM</Option>
<%			else%>
				<Option value="PM">PM</Option>
<%			end if%>
		    </select>
		</td>
	</tr>
	<tr>
		<td align="right">Qty. of shuttle used :</td>
		<td><input name="txtQty" size="2" value='<%=QtyPerson_%>'/></td>
	</tr>
	<tr>
	  	<td colspan="2"><br></td>
	</tr>
	<tr>
	  	<td>&nbsp;</td>
	  	<td><input type="submit" name="btnSubmit" value="Submit">
			<%if State_= "E" then %>
			      <input type="hidden" name="ShuttleID" value=<%=ShuttleID_ %>>
			<%End If%>
			      <input type="hidden" name="State" value=<%=State_ %> >	
			      &nbsp;<input type="button" value="Cancel" name="btnCancel">
		 </td>
	</tr>  
	<tr>
		<td colspan=2>&nbsp;</td>
	</tr>
	</table>
	<%else %>
	<table>
		<tr>
			<td>You do not have permission to access this site.</td>
		</tr>
		<tr>
			<td>Please <a href="http://zagrebws03.eur.state.sbu/WebPASS/eservices/MainPage.asp">Submit Request </a> or contact Zagreb ISC Helpdesk at ext.3333.</td>
		</tr>
	</table>
<% end if
else %>
	<table align="center">
		<tr>
			<td>You do not have permission to access this site.</td>
		</tr>
		<tr>
			<td>Please <a href="http://zagrebws03.eur.state.sbu/WebPASS/eservices/MainPage.asp">Submit Request </a> or contact Zagreb ISC Helpdesk at ext.3333.</td>
		</tr>
	</table>
<% end if %>
	<script language="JavaScript">
	    var cal0 = new calendar1(document.forms['frmData'].elements['txtTransportDate']);
		cal0.year_scroll = true;
		cal0.time_comp = false;
	</script>
</form>

</BODY>
</html>