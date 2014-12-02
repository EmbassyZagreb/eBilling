<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<html>
<head>

<script language="JavaScript" src="calendar.js"></script>
<script type="text/javascript">
function validate_form()
{
	valid = true;
	msg = ""

	if (document.frmCellPhone.txtPhoneNumber.value == "" )
	{
		msg = msg + "Please fill in Phone Number!!!\n"
		valid = false;
	}


	if (document.frmCellPhone.txtAlternateEmail.value != "" )
	{
		var alnum="a-zA-Z0-9";
		exp="^[^@\\s]+@(["+alnum+"+\\-]+\\.)+["+alnum+"]["+alnum+"]["+alnum+"]?$";
		emailregexp = new RegExp(exp);

		result = document.frmCellPhone.txtAlternateEmail.value.match(emailregexp);
		if (result == null)
		{
			msg = msg + "Invalid data type for alternative email address !!!\n"
			valid = false;
		}
	}

	if (valid == false)
	{
		alert(msg);
	}
	return valid;
}

function DiscontinueOnChange(obj)
{
	if (obj.selectedIndex == 0)
	{
		document.frmCellPhone.txtDiscontinuedDate.style.visibility = "visible";

	}
	else
	{
		document.frmCellPhone.txtDiscontinuedDate.style.visibility = "hidden";
		document.frmCellPhone.txtDiscontinuedDate.value = '';

	}
}

window.onload = function() 
{ 	
	if (document.frmCellPhone.txtDiscontinued.value == 'N')
	{
		document.frmCellPhone.txtDiscontinuedDate.style.visibility = "hidden";
	}
	else
	{
		document.frmCellPhone.txtDiscontinuedDate.style.visibility = "visible";
	}
};
</script>
<% 
 dim user_ 
 dim user1_  

 
 user_ = request.servervariables("remote_user") 
 user1_ = user_  'user1_ = right(user_,len(user_)-4)
'user1_ = "pranataw"
'response.write user1_ & "<br>"

%> 

<TITLE>U.S. Embassy Zagreb - zBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">CELL PHONE UPDATE</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>

<form method="post" name="frmCellPhone" action="CellPhoneNumberSave.asp" onsubmit="return validate_form();"> 
<%  
 dim rst 
 dim strsql
 dim rst1
 dim today_


 today_ = now()

 ID_ = request("ID")
 State_ = request("State")
' State_ = "I"
  strsql = "select * from Users where loginId='" & user1_ & "'"
'response.write user1_ & "<br>"
'response.write strsql 
  set rst = server.createobject("adodb.recordset") 
  set rst = BillingCon.execute(strsql)
  if not rst.eof then 
     if (trim(rst("RoleID")) = "Admin") or (trim(rst("RoleID")) = "IM") or (mid(rst("RoleID"),1,3) = "FMC") then  
	if State_ = "E" then
	        strsql = " select * from vwCellPhoneNumberList where ID = " & ID_ 
       		set rst1 = server.createobject("adodb.recordset") 
		'response.write strsql 
        	set rst1 = BillingCon.execute(strsql)
	       	if not rst1.eof then 
        	   PhoneNumber_ = rst1("PhoneNumber") 
		   PhoneType_ = rst1("PhoneType")
		   EmpID_ = rst1("EmpID")
		   EmailAddress_ = rst1("EmailAddress")
		   OwnerID_ = rst1("OwnerID")
		   AlternateEmail_ = rst1("AlternateEmail")
		   Remark_ = rst1("Remark")
		   BillFlag_ = rst1("BillFlag")
		   Discontinued_ = rst1("Discontinued")
		   DiscontinuedDate_ = rst1("DiscontinuedDate")
        	end if
       	end if
	'response.write State_ & "<br>"
	%>             
	<input type="hidden" name="txtDiscontinued" value='<%=Discontinued_ %>' size="25" maxlength="20" />
	<table align=center>  
	<tr>
	  <td>Phone Number</td>
	  <td width="1%">:</td>
	  <td><input type="input" name="txtPhoneNumber" value='<%=PhoneNumber_ %>' size="25" maxlength="20" />&nbsp;(e.g.: 38591XXX)</td>
	</tr>
	<tr>
	  <td>Phone Type</td>
	  <td width="1%">:</td>
	  <td>
		  <select name="PhoneTypeList">
<%
			Dim PhoneRS
			strsql = "select PhoneType, PhoneTypeName from PhoneType order by PhoneTypeName Desc"
			'response.write strsql & "<br>"
			set PhoneRS = server.createobject("adodb.recordset")
			set PhoneRS =BillingCon.execute(strsql)	
			do while not PhoneRS.eof
%>
			<option value=<%=PhoneRS("PhoneType")%> <%if PhoneType_ = PhoneRS("PhoneType") Then %> Selected<%End If%> ><%=PhoneRS("PhoneTypeName")%></option>
<%	      	        	PhoneRS.movenext
	        	loop
%>  
		  </select>
	  </td>
	</tr>
	<tr>
	  <td>Owner</td>
	  <td width="1%">:</td>
	  <td>
<%
			Dim EmpRS
			'strsql = "select EmpID, EmpName, Office, EmpID, EmailAddress, Remark, Case When [Status]='C' Then 'Current' When [Status]='D' Then 'Departed' Else '' End As Status from vwPhoneCustomerList order by EmpName"
			strsql = "select EmpID, EmpName, Office, EmpID, EmailAddress, Remark, StatusName from vwDirectReport Where Status='C' order by EmpName"

			'response.write strsql & "<br>"
			set EmpRS = server.createobject("adodb.recordset")
			set EmpRS =BillingCon.execute(strsql)	
%>
		<select name="EmployeeList">
			<option value="">-- Vacant --</option>
	<%
			        
			do while not EmpRS.eof
				Ename_ = EmpRS("EmpName") 
				'Ename_ = EName_ & "(" & EmpRS("EmpID") & "-" & EmpRS("Office") & "-" & EmpRS("EmailAddress") & " - " & EmpRS("StatusName") &" - " & EmpRS("Remark") &")" 
				Ename_ = EName_ & "(" & EmpRS("Office") & "-" & EmpRS("EmailAddress") & " - " & EmpRS("Remark") &")" 
				if EmpRS("EmpID") = EmpID_  then		
	%>
				        <OPTION value='<%=EmpRS("EmpID")%>' Selected>  <%= EName_  %>
	<%			Else%>
			        	<OPTION value='<%=EmpRS("EmpID")%>'>  <%= EName_  %>
	<%			End If
        	         EmpRS.movenext
	        	loop
	%>  
		</select>
	  </td>
	</tr>
	<tr>
	  <td>Email Address</td>
	  <td width="1%">:</td>
	  <td class="FontContent"><%=EmailAddress_%></td>
	</tr>
<!--
	<tr>
	  <td>Alternate Email</td>
	  <td width="1%">:</td>
	  <td><input type="input" name="txtAlternateEmail" size="50" Value='<%=AlternateEmail_%>' />
	  </td>
	</tr>
-->
	<tr>
	  <td valign="top">Remark</td>
	  <td width="1%" valign="top">:</td>
	  <td><textarea name="txtRemark" cols="60" rows="3"><%=Remark_ %></textarea></td>
	</tr>
	<tr>
	  <td>Bill Charged ?</td>
	  <td width="1%">:</td>
	  <td>
		  <select name="BillFlagList">
			<option value="Y" <%if BillFlag_ ="Y" Then %>Selected<%End If%>>Yes - Approved by supervisor</option>
			<option value="P" <%if BillFlag_ ="P" Then %>Selected<%End If%>>Personal phone - Full payment</option>
			<option value="N" <%if BillFlag_ ="N" Then %>Selected<%End If%>>No</option>
		  </select>
	  </td>
	</tr>
	<tr>
	  <td>Discontinued ?</td>
	  <td width="1%">:</td>
	  <td>
		  <Select name="cmbDiscontinued" onChange="DiscontinueOnChange(this);">
			<Option value="N" <%if Discontinued_ ="N" then %>Selected<%End If%> >No</Option>
			<Option value="Y" <%if Discontinued_ ="Y" then %>Selected<%End If%> >Yes</Option>
 		 </Select>&nbsp;<input type="input" name="txtDiscontinuedDate" size="10" Value='<%=DiscontinuedDate_%>'" onclick="javascript:cal0.popup();" />
	  </td>
	</tr>
	<tr>
	  <td>&nbsp;</td>
	  <td width="1%">&nbsp;</td>
	  <td><input type="submit" name="btnSubmit" value="Submit">
		<%if State_= "E" then %>
		      <input type="hidden" name="txtId" value=<%=Id_ %>>
		<%End If%>
	      <input type="hidden" name="State" value=<%=State_ %> >
	      &nbsp;<input type="button" value="Cancel" name="btnCancel" onClick="Javascript:history.go(-1)">
	 </td>
	</tr>  
	<tr>
		<td colspan=3>&nbsp;</td></tr>
	</table>




<table border="0" bordercolor="#FFFFFF" cellpadding="2" cellspacing="0" width="65%"  class="FontText">
	<tr>
		<td><u><strong>Historical assignment of number <%= PhoneNumber_ %>:<strong></u></td>
	</tr>
	<tr>
		<td class="Hint" align="left">*To alter historical data 'Generate Monthly Billing' procedure must be executed. Procedure sets bill to <%if BillFlag_ ="P" then %>'Awaiting Payment'<%Else%>'Pending'<%End If%> status.</td>
	</tr>
</table>










<%

strsql = "Select * From vwMonthlyBilling Where MobilePhone = '" & PhoneNumber_ & "' Order by YearP+MonthP Desc"
set DataRS = server.createobject("adodb.recordset")
DataRS.CursorLocation = 3
DataRS.Open strsql,BillingCon
'set DataRS=BillingCon.execute(strsql)

dim intPageSize,PageIndex,TotalPages 
dim RecordCount,RecordNumber,Count 
intpageSize=10 
PageIndex=request("PageIndex")

if PageIndex ="" then PageIndex=1 

if not DataRS.eof then
	RecordCount = DataRS.RecordCount   
	'response.write RecordCount & "<br>"
	RecordNumber=(intPageSize * PageIndex) - intPageSize 
	'response.write RecordNumber
	DataRS.PageSize =intPageSize 
	DataRS.AbsolutePage = PageIndex
	TotalPages=DataRS.PageCount 
	'response.write TotalPages & "<br>"
End If
'response.write strsql

dim intPrev,intNext 	
intPrev=PageIndex - 1 
intNext=PageIndex +1 


if not DataRS.eof Then

%>
<!-- <div align="right"><input type="submit" name="btnApproval" value="Approve" /></div> -->
<!-- <div align="right"><input type="button" name="btnExport" value="Export to Excel" onClick="javascript:document.location.href('ARBillingReportAllPrint.asp?sMonthP=<%=sMonthP%>&sYearP=<%=sYearP%>&eMonthP=<%=eMonthP%>&eYearP=<%=eYearP%>&Agency=<%=Agency_%>&Section=<%=Section_%>&EmpID=<%=EmpID_%>&Status=<%=Status%>');"/></div> -->
<table border="1" bordercolor="#EEEEEE" cellpadding="2" cellspacing="0" width="65%"  class="FontText">
    <TR BGCOLOR="#330099" align="center">
         <TD width="7%"><strong><label STYLE=color:#FFFFFF>Billing Period</label></strong></TD>
         <TD width="15%"><strong><label STYLE=color:#FFFFFF>Employee Name</label></strong></TD>
         <TD width="8%"><strong><label STYLE=color:#FFFFFF>Section</label></strong></TD>
	 <TD width="8%" ><strong><label STYLE=color:#FFFFFF>Billing Amount (Kn.)</label></strong></TD>
	 <TD width="8%" ><strong><label STYLE=color:#FFFFFF>Personal Amount (Kn.)</label></strong></TD>
         <TD width="15%"><strong><label STYLE=color:#FFFFFF>Approved By</label></strong></TD>
         <TD width="15%"><strong><label STYLE=color:#FFFFFF>Status</label></strong></TD>
         <TD width="7%"><strong><label STYLE=color:#FFFFFF>Charged</label></strong></TD>

    </TR>
<% 
   dim no_  
   no_ = 1 + ((PageIndex*intPageSize)-intPageSize)
   Count=1 
   do while not DataRS.eof   and Count<=intPageSize
	   if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
%>
      
	   <TR bgcolor="<%=bg%>">
	        <TD align="right">&nbsp;<%= DataRS("MonthP")%>-<%= DataRS("YearP")%></font>&nbsp;</TD>
	        <TD>&nbsp;<%=DataRS("EmpName") %></TD>
	        <TD>&nbsp;<%=DataRS("Office") %> </font></TD>
		<td align="right">
<%		If CDbl(DataRS("CellPhoneBillRp")) > 0 Then %>
			<a href="CellPhoneDetail.asp?CellPhone=<%=DataRS("MobilePhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("CellPhoneBillRp"),-1) %></a>
<%		Else %>
			-
<%		End If %>
		&nbsp;</td>
		<td align="right">
<%		If CDbl(DataRS("CellPhonePrsBillRp")) > 0 Then %>
			<a href="CellPhoneDetail.asp?CellPhone=<%=DataRS("MobilePhone")%>&MonthP=<%= DataRS("MonthP")%>&YearP=<%= DataRS("YearP")%>" target="_blank"><%= formatnumber(DataRS("CellPhonePrsBillRp"),-1) %></a>
<%		Else %>
			-
<%		End If %>
		&nbsp;</td>
		<TD>&nbsp;<%= DataRS("ApprovalSupervisor") %></font></TD>
		<TD>&nbsp;<%= DataRS("ProgressDesc") %></font></TD>
	<td align="center">
		<%if DataRS("BillFlag")="Y" then%> Yes <%end if%>
		<%if DataRS("BillFlag")="P" then%> Personal <%end if%>
		<%if DataRS("BillFlag")="N" then%> No <%end if%>
	</td>

	   </TR>

<%   
		Count=Count +1
	   DataRS.movenext
	   no_ = no_ + 1 
   loop 
	PageNo=1
%>
</table>
<table width="65%">
	<tr>
		<td align="right">
<%
		Do while PageNo<=TotalPages 
			if trim(pageNo) = trim(PageIndex) Then
%>		
				<label class="ActivePage"><%=PageNo%></label>&nbsp;
			<%Else%>
				<a href="CellPhoneNumberEdit.asp?PageIndex=<%=PageNo%>&ID=<%=ID_%>&State=E"><%=PageNo%></a>&nbsp;
<%	
			End If						
			PageNo=PageNo+1
		Loop
%>
		</td>
	</tr>
</table>
<%
else 
%>
	<table cellspadding="1" cellspacing="0" width="100%">  
	<tr>
        	<td><br></TD>
	</tr>
	<tr>
		<td align="center">not data.</td>
	</tr>
	<tr>
        	<td><br></TD>
	</tr>
	<tr>
		<td align="center"><a href="Default.asp"><img src="images/Back.gif" border="0" alt="Go..Back" WIDTH="83" HEIGHT="25"></a></td>
	</tr>	
	</table>
<% end if %>

















    <%else %>
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

<%end if %>
<script language="JavaScript">
   	    var cal0 = new calendar1(document.forms['frmCellPhone'].elements['txtDiscontinuedDate']);
		cal0.year_scroll = true;
		cal0.time_comp = false;
</script>
</form>
</BODY>
</html>