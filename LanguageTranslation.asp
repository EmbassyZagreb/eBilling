
<!--#include file="connect.inc" -->
<html>
<head>

<script language="JavaScript" src="calendar.js"></script>
<script type="text/javascript">
function ClearFilter()
{
	document.forms['frmSearch'].elements['cmbProviderID'].value =0;
}
function validate_form()
{
	valid = true;
	msg = ""

	if (document.frmPaymentDueDate.txtCeilingAmount.value == "" )
	{
		msg = msg + "Please fill in Ceiling Amount !!!\n"
		valid = false;
	}
	else
	{
		var myRegExp = new RegExp("^[/+|/-]?[0-9]*[/.]?[0-9]*$");
		if (myRegExp.test(document.frmPaymentDueDate.txtCeilingAmount.value) == false)
		{
			msg = msg + "Invalid data type for Paid Amount !!!\n"
			valid = false;
		}
	}

	if (document.frmPaymentDueDate.txtDetailRecordAmount.value == "" )
	{
		msg = msg + "Please fill in hide detail record value !!!\n"
		valid = false;
	}
	else
	{
		var myRegExp = new RegExp("^[/+|/-]?[0-9]*[/.]?[0-9]*$");
		if (myRegExp.test(document.frmPaymentDueDate.txtDetailRecordAmount.value) == false)
		{
			msg = msg + "Invalid data type for hide detail record value !!!\n"
			valid = false;
		}
	}

	if (document.frmPaymentDueDate.txtCashierMinimumAmount.value == "" )
	{
		msg = msg + "Please fill in accumulated debt record value !!!\n"
		valid = false;
	}
	else
	{
		var myRegExp = new RegExp("^[/+|/-]?[0-9]*[/.]?[0-9]*$");
		if (myRegExp.test(document.frmPaymentDueDate.txtCashierMinimumAmount.value) == false)
		{
			msg = msg + "Invalid data type for accumulated debt record value !!!\n"
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
<META http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">LANGUAGE TRANSLATION</TD>
   </TR>
	<tr>
        	<td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
	</tr>
	<TR>
		<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
	</TR>
  </TABLE>

<%  
 dim rst 
 dim strsql
 dim objRS
 dim DescriptionBilled_
 dim Name_
 
 
 ProviderID_ = Request.Form("cmbProviderID")
'response.write ProviderID_
if ProviderID_ = "" Then ProviderID_ = request("ProviderID")
if ProviderID_ = "" then
	ProviderID_ = 0
end if
 
 
 strsql = "select * from Users where loginId='" & user1_ & "'"
'response.write user1_ & "<br>"
'response.write strsql 
  set rst = server.createobject("adodb.recordset") 
  set rst = BillingCon.execute(strsql)
  if not rst.eof then 
	if (trim(rst("RoleID")) = "Admin") then 
	        strsql = "select DescriptionID, ProviderID, NameProvider, DescriptionOriginal, DescriptionTranslated, DescriptionBilled from vwLanguageTranslation Where (ProviderID='" & ProviderID_ & "' or '" & ProviderID_ & "'='0') order by DescriptionOriginal"			
       		set objRS = server.createobject("adodb.recordset") 
		'response.write strsql 
        	set objRS = BillingCon.execute(strsql)
	%>	



	

	<table align="center" border="1" cellpadding="1" cellspacing="0" width="70%" bgcolor="white">
	<tr>
		<td>
		<form method="post" name="frmSearch">
		<table align="center" cellpadding="1" cellspacing="0" width="100%">		
		<tr bgcolor="#000099">
			<td height="25" colspan="4"><strong>&nbsp;<span class="style5">Search</span></strong></td>
		</tr>
		<tr>
			<td>&nbsp;Service Provider&nbsp;</td>
			<td>:</td>
			<td>
<%
 				strsql ="select ProviderID, ProviderName, ProviderDescription from Provider order by ProviderName"
				set ProviderRS = server.createobject("adodb.recordset")
				set ProviderRS = BillingCon.execute(strsql)
'				response.write strStr 
%>	
				<Select name="cmbProviderID">
					<Option value=0>--All--</Option>
<%				Do While not ProviderRS.eof %>
					<Option value='<%=ProviderRS("ProviderID")%>' <%if trim(ProviderID_) = trim(ProviderRS("ProviderID")) then %>Selected<%End If%> ><%=ProviderRS("ProviderName")%> - <%=ProviderRS("ProviderDescription")%></Option>
					
<%					ProviderRS.MoveNext
				Loop%>
				</select>
			</td>	
			<td align="right">
				<input type="submit" name="btnSearch" value="Search">&nbsp;<input type="button" name="btnClear" value="Reset filter" onclick="javascript:ClearFilter();">
			</td>
		</tr>
		</table></form>
		</td>
	</tr>
	</table>

	
	
	
	
<form method="post" name="frmLanguageTranslation" id="frmLanguageTranslation" action="LanguageTranslationSave.asp" onSubmit="return validate_form()">
	
<table border="1" bordercolor="#EEEEEE" cellpadding="2" cellspacing="0" width="70%">
     <TR BGCOLOR="#330099" align="center">
         <TD width="10%"><strong><label STYLE=color:#FFFFFF>Provider</label></strong></TD>
         <TD width="90%"><strong><label STYLE=color:#FFFFFF>Type Description</label></strong></TD>
		 <TD width="1"><strong><label STYLE=color:#FFFFFF>Type Translation</label></strong></TD>
		 <TD width="1"><strong><label STYLE=color:#FFFFFF>Billing</label></strong></TD>
     </TR>
		<br>
	 <%
	 
	 Do While Not objRS.EOF	
			 
	%>
			<tr>	
				<input type="hidden" name="ID" value="<%=objRS("DescriptionID")%>">
				<td align=left><%=objRS("NameProvider")%></td>
				<td align=left bgcolor="#DDDDDD"><%=objRS("DescriptionOriginal")%></td>
									
				<td align=left><input type="text" size="50" name="txtDescriptionTranslated_<%=objRS("DescriptionID")%>" value="<%=objRS("DescriptionTranslated")%>"></td>
				<td>
				<select name="cmbDescriptionBilledList_<%=objRS("DescriptionID")%>">
					<option value="">-</option> 

				<option value="O" <%if objRS("DescriptionBilled") ="O" Then %>Selected<%End If%>>Billed If Official</option>
				<option value="P" <%if objRS("DescriptionBilled") ="P" Then %>Selected<%End If%>>Billed If Private</option>
				<option value="A" <%if objRS("DescriptionBilled") ="A" Then %>Selected<%End If%>>Always Billed</option>
				<option value="N" <%if objRS("DescriptionBilled") ="N" Then %>Selected<%End If%>>Never Billed</option>
				</select>
			</td>
			
			</tr>

   <%
      objRS.MoveNext
   Loop
   %>
		</table>
		<table>
		<tr>
		  <td></td>
		  <td><input type="submit" id="btnSubmit" name="btnSubmit" value="Submit">
		      &nbsp;<input type="button" value="Cancel" name="btnCancel">
		 </td>
		</tr>  
		<tr>
		   <td colspan=2>&nbsp;</td>
		</tr>
		</table>
     <%  else %>
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

<% end if %>
</form>
</BODY>
</html>