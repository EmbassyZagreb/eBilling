<%@ Language=VBScript%>
<% ' VI 6.0 Scripting Object Model Enabled %>
 
 
<html>
<head>

<META name=VI60_defaultClientScript content=VBScript>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<!--#include file="connect.inc" -->

<script language="JavaScript">
function ClearFilter()
{
	document.forms['frmSearch'].elements['txtPrefix'].value ='';
	document.forms['frmSearch'].elements['SortList'].value ='Code';
}
</script>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<!--#include file="Header.inc" -->
   <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">PREFIX NUMBER LIST</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
<%

dim rs 
dim strsql
dim tombol
dim hlm
%>

<%

dim x , ticket_, user_ , user1_
dim intPageSize,PageIndex,TotalPages 
dim RecordCount,RecordNumber,Count 
intpageSize=50 
PageIndex=request("PageIndex")


user_ = request.servervariables("remote_user")
user1_ = right(user_,len(user_)-4)

Prefix_ = trim(request.form("txtPrefix"))
if Prefix_ ="" then
	Prefix_ = Request("Prefix")	
end if

SortBy_ = Request.Form("SortList")
'response.write "SortBy" & SortBy_
if (SortBy_ ="") then
	if Request("SortBy")<>"" then
		SortBy_ = Request("SortBy")	
	Else
		SortBy_ = "Code"
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

strsql = "select * from Users where loginId='" & user1_ & "'"
set RS_Query = server.createobject("adodb.recordset")
'response.write strsql & "<br>"
set RS_Query = BillingCon.execute(strsql)

  if not RS_Query.eof then 
     if trim(RS_Query("RoleID")) = "Admin" or (mid(RS_Query("RoleID"),1,3) = "FMC") then     
	strsql = "Select PrefixID, Code, Prefix, Type, Description "_
		  &"From MsPrefixNumber Where (Prefix like '" & Prefix_  & "%' or '" & Prefix_ & "'='') "_
		  &"Order by " & SortBy_ & " " & Order_ 
		
	'response.write strsql & "<br>"
	set rs = server.createobject("adodb.recordset") 
	rs.CursorLocation = 3
	rs.open strsql,BillingCon

	if ((PageIndex ="") or (request.form("btnSearch")="Search")) then PageIndex=1 
	if not rs.eof then
		RecordCount = rs.RecordCount   
		'response.write RecordCount & "<br>"
		RecordNumber=(intPageSize * PageIndex) - intPageSize 
		'response.write RecordNumber
		rs.PageSize =intPageSize 
		rs.AbsolutePage = PageIndex
		TotalPages=rs.PageCount 
		'response.write TotalPages & "<br>"
	End If
'	set rs = BillingCon.execute(strsql) 


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
			<td width="20%">&nbsp;Prefix Number</td>
			<td width="1%">:</td>
			<td><input name="txtPrefix" type="Input" size="20" Value='<%=Prefix_%>'></td>
			<td>Sort By&nbsp;</td>
			<td>:</td>
			<td>
				<Select name="SortList">
					<Option value="Prefix" <%if SortBy_ ="Prefix" then %>Selected<%End If%> >Prefix</Option>
					<Option value="Code" <%if SortBy_ ="Code" then %>Selected<%End If%> >Code</Option>
					<Option value="Type" <%if SortBy_ ="Type" then %>Selected<%End If%> >Type</Option>
					<Option value="Description" <%if SortBy_ ="Description" then %>Selected<%End If%> >Description</Option>
				</Select>&nbsp;
				<Select name="OrderList">
					<Option value="Asc" <%if Order_ ="Asc" then %>Selected<%End If%> >Asc</Option>
					<Option value="Desc" <%if Order_ ="Desc" then %>Selected<%End If%> >Desc</Option>
				</Select>
			</td
		</tr>
		<tr>
			<td align="left">
				&nbsp;&nbsp;<input type="Button" name="btnBack" value="Back" onClick="Javascript:document.location.href('Default.asp');">
			</td>
			<td colspan="5" align="center">&nbsp;&nbsp;
				
				&nbsp;&nbsp;<input type="submit" name="btnSearch" value="Search">
				&nbsp;&nbsp;<input type="button" name="btnClear" value="Reset filter" onclick="javascript:ClearFilter();">


			</td>
		</tr>
		</table></form>
		</td>
	</tr>
	</table>
	
     <table width="100%">
	<tr>
		<td width="50%"><a href="PrefixNumberEdit.asp?State=I">Add New Prefix Number</a></td>
	</tr>
     </table>
     <table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" width=100%>
     <TR BGCOLOR="#330099" align="center">
         <TD width="5%" align="center"><strong><label STYLE=color:#FFFFFF>NO</label></strong></TD>
         <TD width="15%"><strong><label STYLE=color:#FFFFFF>&nbsp;Code</label></strong></TD>
         <TD width="10%"><strong><label STYLE=color:#FFFFFF>&nbsp;Prefix</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>&nbsp;Type</label></strong></TD>
         <TD><strong><label STYLE=color:#FFFFFF>&nbsp;Description</label></strong></TD>
         <TD width="8%"><strong><label STYLE=color:#FFFFFF>&nbsp;Action</label></strong></TD>
     </TR>    
<% 
	   dim no_  
	   'no_ = 1 
	   no_ = 1 + ((PageIndex*intPageSize)-intPageSize)
	   do while not rs.eof and Count<intPageSize
	   if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
%>
      
     <TR bgcolor="<%=bg%>">
        <TD align=center ><FONT color=#330099 size=2> <%= no_ %> </font>   </TD>
        <TD><FONT color=#330099 size=2><A HREF="PrefixNumberEdit.asp?PrefixID=<%= rs("PrefixID")%>&State=U"> <%= rs("Code") %></A></font>   </TD>
        <TD><FONT color=#330099 size=2>&nbsp;<%= rs("Prefix") %></font></TD>
        <TD><FONT color=#330099 size=2>&nbsp;<%= rs("Type") %></font></TD>
        <TD><FONT color=#330099 size=2>&nbsp;<%= rs("Description") %></font>   </TD>
	<TD><FONT color=#330099 size=2><A HREF="PrefixNumberDeleteConfirm.asp?PrefixID=<%= rs("PrefixID")%>&State=D" >Delete</A></font>   </TD>
      </TR>

<%   
	   Count=Count +1
	   rs.movenext
	   no_ = no_ + 1 
	   loop
%>
     </TABLE>
     <table align="center" cellpadding="1" cellspacing="0" width="100%">
		<tr>
			<td align="right">
	<%
			Do while PageNo<=TotalPages 
				if trim(pageNo) = trim(PageIndex) Then
	%>		
					<label class="ActivePage"><%=PageNo%></label>&nbsp;
				<%Else%>
					<a href="PrefixNumberList.asp?PageIndex=<%=PageNo%>&Prefix=<%=Prefix_%>&SortBy=<%=SortBy_%>&Order=<%=Order_%>"><%=PageNo%></a>&nbsp;
	<%	
				End If						
				PageNo=PageNo+1
			Loop
	%>
			</td>
		</tr>
	</table>
     <table width="100%">
	<tr ><td align="Center">
	<a href="default.asp"><img src="images/Back.gif" border="0" alt="Go..Back" WIDTH="83" HEIGHT="25"></a>
	</td></tr>
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


