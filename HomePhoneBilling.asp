<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<html>
<head>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</head>
<!--#include file="Header.inc" -->
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">HOMEPHONE BILLING</TD>
   </TR>
<tr>
        <td colspan="4" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
</tr>
  <TR>
  	<TD COLSPAN="4"><HR style="LEFT: 10px; TOP: 59px" align=center></TD>
   </TR>
  </TABLE>
<% 
 dim user_ 
 dim user1_  
 dim rst 
 dim strsql
 
 user_ = request.servervariables("remote_user") 
  user1_ = right(user_,len(user_)-4)
user1_ = "TuchrelloWP"
'response.write user1_ & "<br>"

curMonth_ = month(date())
curYear_ = year(date())
if len(curMonth_)= 1 then
	curMonth_ = "0" & curMonth_
end if

MonthP = Request.Form("MonthList")
if MonthP ="" then
	MonthP = curMonth_ 
end if

YearP = Request.Form("YearList")
if YearP ="" then
	YearP = curYear_ 
end if

%>  
<%
strsql = "Exec spGetHomephone '" & user1_ & "','" & MonthP & "','" & YearP & "'"
'response.write strsql & "<br>"
set rsData = server.createobject("adodb.recordset") 
set rsData = BillingCon.execute(strsql) 
%>
<table cellspadding="1" cellspacing="0" width="60%" border="1" align="center">
<tr align="Center">
	<td colspan="2" align="center">
		<form method="post" name="frmSearch" Action="HomePhoneBilling.asp">
		<table  width="100%">
		<tr bgcolor="#000099">
			<td height="25" colspan="7"><strong>&nbsp;<span class="style5">Search &amp; Sort By </span></strong></td>
		</tr>
		<tr>
			<td>&nbsp;Period&nbsp;</td>				
			<td>:</td>
			<td>
				<Select name="MonthList">
					<Option value="01" <%if MonthP ="01" then %>Selected<%End If%> >January</Option>
					<Option value="02" <%if MonthP ="02" then %>Selected<%End If%> >February</Option>
					<Option value="03" <%if MonthP ="03" then %>Selected<%End If%> >March</Option>
					<Option value="04" <%if MonthP ="04" then %>Selected<%End If%> >April</Option>
					<Option value="05" <%if MonthP ="05" then %>Selected<%End If%> >May</Option>
					<Option value="06" <%if MonthP ="06" then %>Selected<%End If%> >June</Option>
					<Option value="07" <%if MonthP ="07" then %>Selected<%End If%> >July</Option>
					<Option value="08" <%if MonthP ="08" then %>Selected<%End If%> >August</Option>
					<Option value="09" <%if MonthP ="09" then %>Selected<%End If%> >Sepetember</Option>
					<Option value="10" <%if MonthP ="10" then %>Selected<%End If%> >October</Option>
					<Option value="11" <%if MonthP ="11" then %>Selected<%End If%> >November</Option>
					<Option value="12" <%if MonthP ="12" then %>Selected<%End If%> >December</Option>
				</Select>&nbsp;
<%
				Year_ = Year(Date()) - 1
%>

				<Select name="YearList">
<% 				Do While Year_ <= Year(Date()) %>
				<Option value='<%=Year_%>' <%if trim(Year_) = trim(YearP) then %>Selected<%End If%> ><%=Year_%></Option>		
<% 
			Year_ = Year_ + 1
			Loop %>	
				</Select>										
			</td>
			<td height="30" align="center">
				<input type="Button" name="btnBack" value="Back" onClick="Javascript:document.location.href('Default.asp');">
				<input type="submit" name="Submit" value="Search">
			</td>
		</tr>			
		</table>
		</form>
	</td>
</tr>	
</table><br>
<table cellspadding="1" cellspacing="0" width="80%" bgColor="white" align="center">  
<%if not rsData.eof then%>
<tr>
	<td colspan="3" align="center"><h3>Billing Period : <Label style="color:blue"><%=MonthP%> - <%=YearP%></label></h3></td>
</tr>
<tr>
        <td align="Left" colspan="3"><u><b>Employee Info<b></u></TD>
</tr>  
<tr>
	<td width="25%">&nbsp;Employee Name</td>
	<td width="1%">:</td>
	<td class="FontContent"><%=rsData("EmpName")%></td>
</tr>
<tr>
	<td>&nbsp;Agency / Office</td>
	<td width="1%">:</td>
	<td class="FontContent"><%=rsData("OfficeLocation")%></td>
</tr>
<tr>
	<td>&nbsp;Homephone</td>
	<td width="1%">:</td>
	<td class="FontContent"><%=rsData("Nomor")%></td>
</tr>
<tr>
	<td colspan="3"><hr></td>
</tr>
<tr>
	<td align="Left" colspan="3"><u><b>Billing Detail :<b></u></td>
</tr>
<tr>
	<td colspan="3">
	<table cellspadding="1" cellspacing="0" width="100%">
	<tr>
		<td width="20%">&nbsp;Monthly Fee</td>
		<td width="1%">:</td>
		<td width="1%">Kn.</td>
		<td width="10%" class="FontContent" align="right"><%=formatnumber(rsData("Abonemen"),-1)%></td>
		<td width="10%">&nbsp;</td>
		<td width="15%">17</td>
		<td width="1%">:</td>
		<td width="1%">Kn.</td>
		<td width="10%" class="FontContent" align="right"><%=formatnumber(rsData("17"),-1)%></td>
	</tr>
	<tr>
		<td>&nbsp;Local call</td>
		<td width="1%">:</td>
		<td width="1%">Kn.</td>
		<td class="FontContent" align="right"><%=formatnumber(rsData("Lokal"),-1)%></td>
		<td width="10%">&nbsp;</td>
		<td>Operator</td>
		<td width="1%">:</td>
		<td width="1%">Kn.</td>
		<td class="FontContent" align="right"><%=formatnumber(rsData("Operator"),-1)%></td>
	</tr>
	<tr>
		<td>&nbsp;SLJJ</td>
		<td width="1%">:</td>
		<td width="1%">Kn.</td>
		<td class="FontContent" align="right"><%=formatnumber(rsData("SLJJ"),-1)%></td>
		<td width="10%">&nbsp;</td>
		<td>Air Time</td>
		<td width="1%">:</td>
		<td width="1%">Kn.</td>
		<td class="FontContent" align="right"><%=formatnumber(rsData("AirTime"),-1)%></td>
	</tr>
	<tr>
		<td>&nbsp;STB</td>
		<td width="1%">:</td>
		<td width="1%">Kn.</td>
		<td class="FontContent" align="right"><%=formatnumber(rsData("STB"),-1)%></td>
		<td width="10%">&nbsp;</td>
		<td>Quota</td>
		<td width="1%">:</td>
		<td width="1%">Kn.</td>
		<td class="FontContent" align="right"><%=formatnumber(rsData("Quota"),-1)%></td>
	</tr>
	<tr>
		<td>&nbsp;JAPATI</td>
		<td width="1%">:</td>
		<td width="1%">Kn.</td>
		<td class="FontContent" align="right"><%=formatnumber(rsData("JAPATI"),-1)%></td>
		<td width="10%">&nbsp;</td>
		<td>Miscellaneous</td>
		<td width="1%">:</td>
		<td width="1%">Kn.</td>
		<td class="FontContent" align="right"><%=formatnumber(rsData("Lain2"),-1)%></td>
	</tr>
	<tr>
		<td>&nbsp;Int. call (SLI007)</td>
		<td width="1%">:</td>
		<td width="1%">Kn.</td>
		<td class="FontContent" align="right"><%=formatnumber(rsData("SLI007"),-1)%></td>
		<td width="10%">&nbsp;</td>
		<td>Tax</td>
		<td width="1%">:</td>
		<td width="1%">Kn.</td>
		<td class="FontContent" align="right"><%=formatnumber(rsData("PPN"),-1)%></td>
	</tr>
	<tr>
		<td>&nbsp;Int. call (001+008)</td>
		<td width="1%">:</td>
		<td width="1%">Kn.</td>
		<td class="FontContent" align="right"><%=formatnumber(rsData("001+008"),-1)%></td>
		<td width="10%">&nbsp;</td>
		<td>Stamp fee</td>
		<td width="1%">:</td>
		<td width="1%">Kn.</td>
		<td class="FontContent" align="right"><%=formatnumber(rsData("Meterai"),-1)%></td>
	</tr>
	<tr>
		<td colspan="9"><hr></td>
	</tr>
	<tr>
		<td colspan="5">&nbsp;</td>
		<td><b>Total</b></td>
		<td width="1%">:</td>
		<td width="1%">Kn.</td>
		<td class="FontContent" align="right"><b><%=formatnumber(rsData("Total"),-1)%></b></td>
		<td width="10%">&nbsp;</td>
	</tr>
	<tr>
		<td colspan="5">&nbsp;</td>
		<td><b>Payment Status</b></td>
		<td width="1%">:</td>
		<td width="1%">&nbsp;</td>
		<td class="FontContent" align="right"><b><%=rsData("Status")%></b></td>
	</tr>	
	</table>
	</td>
</tr>


<%else%>
<tr>
	<td colspan="6" align="center">&nbsp;</td>	
</tr>
<tr>
	<td colspan="6" align="center">there is no data.</td>	
</tr>

<%end if%>
</table>
</BODY>
</html>