<%@ Language=VBScript%>
<% ' VI 6.0 Scripting Object Model Enabled %>
 
 
<html>
<head>

<META name=VI60_defaultClientScript content=VBScript>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<!--#include file="connect.inc" -->

<link href="style.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
function ClearFilter()
{
	document.forms['frmSearch'].elements['txtFundingAgency'].value ='';
	document.forms['frmSearch'].elements['SortList'].value ='AgencyFundingCode';
}
</script>
</head>
<!--#include file="Header.inc" -->
   <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">Agency List</TD>
  </TR>
  <tr>
	<td colspan="4" align="Left" width="20%"><A HREF="Default.asp">Home</A></td>
  </tr>   
  <tr>
  	<td COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></td>
   </tr>
  </TABLE>
<%

dim rs 
dim strsql
dim tombol
dim hlm
%>

<%
Message = Request("msg")

Dim user_ , user1_, UserRole_

user_ = request.servervariables("remote_user")
user1_ = right(user_,len(user_)-4)


strsql = "select * from Users where loginId='" & user1_ & "'"
set RS_Query = server.createobject("adodb.recordset")
'response.write strsql & "<br>"
set RS_Query = BillingCon.execute(strsql)
'response.write RS_Query("RoleID") & "<br>"



FundingAgency_ = trim(request.form("txtFundingAgency"))
if FundingAgency_ ="" then
	FundingAgency_ = Request("FundingAgency")	
end if

SortBy_ = Request.Form("SortList")
'response.write "SortBy" & SortBy_
if (SortBy_ ="") then
	if Request("SortBy")<>"" then
		SortBy_ = Request("SortBy")	
	Else
		SortBy_ = "AgencyDesc"
	end if
end if

Order_ = Request.Form("OrderList")
if (Order_ ="") then
	if Request("Order")<>"" then
		Order_ = Request("OrderList")	
	Else
		Order_ = "Asc"
	end if
end if

     if (trim(RS_Query("RoleID")) = "Admin") or (trim(RS_Query("RoleID")) = "Voucher") or (trim(RS_Query("RoleID")) = "FMC") then    
'	strsql = "select * from AgencyFunding"
	strsql = "select * from AgencyFunding"_
		  &" Where (AgencyDesc like '%" & FundingAgency_  & "%' or '" & FundingAgency_ & "'='') "_
		  &"Order by " & SortBy_ & " " & Order_ 

	'response.write strsql & "<br>"
	set rsAgency = server.createobject("adodb.recordset")
	set rsAgency = BillingCon.execute(strsql)
%>	
	<table align="center" border="1" cellpadding="1" cellspacing="0" width="70%" bgcolor="white">
	<tr>
		<td>
		<form method="post" name="frmSearch">
		<table align="center" cellpadding="1" cellspacing="0" width="100%">		
		<tr bgcolor="#000099">
			<td height="25" colspan="6"><strong>&nbsp;<span class="style5">Search</span></strong></td>
		</tr>	
		<tr>
			<td width="20%">&nbsp;Funding Agency</td>
			<td width="1%">:</td>
			<td><input name="txtFundingAgency" type="Input" size="30" Value='<%=FundingAgency_%>'></td>	
			<td align="right">Sort By&nbsp;</td>
			<td>:</td>
			<td>
			<Select name="SortList">
				<Option value="AgencyFundingCode" <%if SortBy_ ="AgencyFundingCode" then %>Selected<%End If%> >Agency Code</Option>
				<Option value="AgencyDesc" <%if SortBy_ ="AgencyDesc" then %>Selected<%End If%> >Funding Agency</Option>
				<Option value="FiscalStripVAT" <%if SortBy_ ="FiscalStripVAT" then %>Selected<%End If%> >Fiscal Strip VAT</Option>
				<Option value="FiscalStripNonVAT" <%if SortBy_ ="FiscalStripNonVAT" then %>Selected<%End If%> >FiscalStripNonVAT</Option>
			</Select>&nbsp;
			<Select name="OrderList">
				<Option value="Asc" <%if Order_ ="Asc" then %>Selected<%End If%> >Asc</Option>
				<Option value="Desc" <%if Order_ ="Desc" then %>Selected<%End If%> >Desc</Option>
			</Select>
			</td
		</tr>
		<tr>
			<td colspan="3"></td>
			<td><input type="submit" name="btnSearch" value="Search"></td>
			<td>&nbsp;</td>
			<td>&nbsp;&nbsp;<input type="submit" name="btnClear" value="Reset filter" onclick="javascript:ClearFilter();">
			    &nbsp;&nbsp;&nbsp;&nbsp;<input type="button" name="btnExport" value="Export to Excel" onClick="javascript:document.location.href('AgencyListPrint.asp?FundingAgency=<%=FundingAgency_%>&SortBy=<%=SortBy_%>&Order=<%=Order_%>');"/>
			</td>
		</tr>
		</table>
		</form>
		</td>
	</tr>
	</table>
	<table width="70%">		
<%	if Message<>"" then %>
	<tr>
		<td align="center" class="FontHint"><%=Message%></td>		
	</tr>
<%	end if	%>
	</table>
     <table width="100%">
	<tr>
		<td align="left"><a href="AgencyNew.asp">Add New Agency</a></td>
	</tr>
     </table>
	<table border="1" bordercolor="#EEEEEE" cellpadding="2" cellspacing="0" width="100%"  class="FontText">
	    <TR BGCOLOR="#330099" align="center">
		 <TD width=3%><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
	         <TD width=8%><strong><label STYLE=color:#FFFFFF>Agency Code</label></strong></TD>
	         <TD width="15%"><strong><label STYLE=color:#FFFFFF>Funding Agency</label></strong></TD>
	         <TD><strong><label STYLE=color:#FFFFFF>Fiscal Strip VAT</label></strong></TD>
	         <TD><strong><label STYLE=color:#FFFFFF>Fiscal Strip Non VAT</label></strong></TD>
	         <TD width="5%"><strong><label STYLE=color:#FFFFFF>Disabled</label></strong></TD>
	         <TD colspan="2" align="Center">
			<strong><label STYLE=color:#FFFFFF>Action</label></strong>
		</TD>
	    </TR>    

<% 
	   dim no_  
	   no_ = 1 
	   do while not rsAgency.eof  
		   if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 

		If rsAgency("Disabled")="Y" Then
			AgencyType_ ="Yes"
		ElseIf rsAgency("Disabled")="N" Then
			AgencyType_ ="No"
		Else
			AgencyType_ =""
		End If
%>

	   <TR bgcolor="<%=bg%>">
       		<TD align="right"> <%= no_ %> </font>&nbsp;</TD>
        	<TD align="right">&nbsp;<%= rsAgency("AgencyFundingCode")%></font>&nbsp;</TD>
	        <TD>&nbsp;<%= rsAgency("AgencyDesc") %></font></TD>
	        <TD>&nbsp;<%= rsAgency("FiscalStripVAT") %></font></TD>
	        <TD>&nbsp;<%= rsAgency("FiscalStripNonVAT") %></font></TD>
	        <TD align="right"><%= AgencyType_ %></font>&nbsp;</TD>
		<TD>
			&nbsp;<A HREF="AgencyEdit.asp?ID=<%= rsAgency("AgencyID")%>&Mode=E" >Edit</A></font>
		</TD>
		<TD>
			&nbsp;<A HREF="AgencyDelete.asp?ID=<%= rsAgency("AgencyID")%>&AgencyDesc='<%= rsAgency("AgencyDesc")%>'">Delete</A></font>
		</TD>
	   </TR>

<%   
	   	rsAgency.movenext
   		no_ = no_ + 1 
	  loop
%>
	</TABLE>
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

<%
end if 
%>
</body> 

</html>


