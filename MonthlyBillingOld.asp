<%@ Language=VBScript %>
<!--#include file="connect.inc" -->
<html>
<head>
<style type="text/css">

<script type="text/javascript">
function checkall(obj)
{
	var c = document.frmOfficePhoneBilling.elements.length
	for (var x=0; x<frmOfficePhoneBilling.elements.length; x++)
	{
		cbElement = frmOfficePhoneBilling.elements[x]
		if (cbElement.type == "checkbox")
		{
			cbElement.checked= obj.checked?true:false
		}
	}
}

function validate_form()
{
	valid = true;
	msg = ""


	if (document.frmOfficePhoneBilling.txtSpvEmail.value != "" )
	{
		var alnum="a-zA-Z0-9";
		exp="^[^@\\s]+@(["+alnum+"+\\-]+\\.)+["+alnum+"]["+alnum+"]["+alnum+"]?$";
		emailregexp = new RegExp(exp);

		result = document.frmOfficePhoneBilling.txtSpvEmail.value.match(emailregexp);
		if (result == null)
		{
			msg = msg + "Invalid data type for supervisor email address !!!\n"
			valid = false;
		}
	}
	else
	{
		msg = "Please fill in your supervisor mail !!!\n"		
		valid = false;
	}

	if (valid == false)
	{
		alert(msg)
	}
	return valid;
}
</script>
<TITLE>U.S. Embassy Zagreb - eBilling Application</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=windows-1250">
<link href="style.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY BGCOLOR="#FFFFFF" TEXT="#000000" LINK="#0080FF" ALINK="0080FF" VLINK="#0055AA" MARGINWIDTH="8" MARGINHEIGHT="0" LEFTMARGIN="8" TOPMARGIN="0">
  <Center><FONT COLOR=#009900><B>SENSITIVE BUT UNCLASSIFIED</Center></FONT></B>
  <BR>
<CENTER>
  <IMG SRC="images/embassytitle2.jpeg" WIDTH="661" HEIGHT="80" BORDER="0"> 
  <TABLE WIDTH="65%" BORDER="0" CELLPADDING="0" CELLSPACING="0">
  <CAPTION><H3 STYLE="font-size:17px;color:#000040">Mission zagreb - Billing Application</H3></CAPTION>
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">Monthly Billing</TD>
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
user1_ = "MartinWC"
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
Dim TotalOffBill_, TotalHomeBill_, TotalBill_ 


strsql = "Exec spGetBilling 'Header','" & user1_ & "','" & MonthP & "','" & YearP & "'"
'response.write strsql & "<br>"
set rsOfficePhone = server.createobject("adodb.recordset") 
set rsOfficePhone = BillingCon.execute(strsql)
if not rsOfficePhone.eof then
	SpvEmail_ = rsOfficePhone("SpvEmail")
	Notes_ = rsOfficePhone("Notes")
	SpvRemark_ = rsOfficePhone("SpvRemark")
	TotalOffBill_ = rsOfficePhone("TotalCost")
	Ext_ = rsOfficePhone("Extension")
	Status_ = rsOfficePhone("Status")
else
	TotalOffBill_ = 0
end if

strsql = "Exec spGetHomephone '" & user1_ & "','" & MonthP & "','" & YearP & "'"
'response.write strsql & "<br>"
set rsHomePhone = server.createobject("adodb.recordset") 
set rsHomePhone = BillingCon.execute(strsql) 

if not rsHomePhone.eof then
	TotalHomeBill_ = rsHomePhone("Total")	
else
	TotalHomeBill_ = 0
end if
'response.write TotalOffBill_ & "<br>"
'response.write CDbl(TotalHomeBill_ )
TotalBill_ = CDbl(TotalOffBill_) + CDbl(TotalHomeBill_)

strsql = "Exec spGetShuttleBusList '" & user1_ & "','" & MonthP & "','" & YearP & "'"
'response.write strsql & "<br>"
set rsShuttle = server.createobject("adodb.recordset") 
set rsShuttle = BillingCon.execute(strsql) 


%>
<table cellspadding="1" cellspacing="0" width="60%" border="1" align="center">
<tr align="Center">
	<td colspan="2" align="center">
		<form method="post" name="frmSearch" Action="MonthlyBilling.asp">
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
'					response.write Year_
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
<form method="post" action="UpdateOfficePhoneBilling.asp" name="frmOfficePhoneBilling" onSubmit="return validate_form();"> 
<table cellspadding="1" cellspacing="0" width="100%" bgColor="white">  
<%
strsql = "Select Case When isNull(FirstName,'')='' Then LastName Else LastName+', '+FirstName End As EmpName, Agency, Office, WorkPhone, HomePhone from vwPhoneCustomerList Where LoginID='" & user1_ & "'"
'response.write strsql & "<br>"
set rsPersonalData = server.createobject("adodb.recordset") 
set rsPersonalData = BillingCon.execute(strsql) 
Period_ = MonthP & " - " & YearP
if not rsPersonalData.eof then
	EmpName_ = rsPersonalData("EmpName")
	Office_ = rsPersonalData("Agency") & " - " & rsPersonalData("Office")
	OfficePhone_ = rsPersonalData("WorkPhone")
	HomePhone_ = rsPersonalData("HomePhone")
end if
'response.write Period_  & "<br>"
%>
<tr>
          <td align="Left"><u><b>Personal Info<b></u></TD>
</tr>  
<tr>
	<td width="12%">Employee Name</td>
	<td width="1%">:</td>
	<td width="30%" class="FontContent"><%=EmpName_%></td>
	<td width="20%">Billing Period (Month - Year)</td>
	<td width="1%">:</td>
	<td class="FontContent"><%=Period_%></td>
</tr>
<tr>
	<td>Agency / Office</td>
	<td width="1%">:</td>
	<td class="FontContent"><%=Office_%></td>
	<td>Homephone</td>
	<td width="1%">:</td>
	<td class="FontContent"><%=HomePhone_ %></td>
</tr>
<tr>
	<td>Office Phone</td>
	<td width="1%">:</td>
	<td class="FontContent"><%=OfficePhone_ %></td>
	<td>Total Bill</td>
	<td width="1%">:</td>
	<td class="FontContent">Rp &nbsp;<%=formatnumber(TotalBill_,0) %></td>
</tr>
<tr>
	<td colspan="6"><hr></td>
</tr>
<%if not rsShuttle.eof then%>
<tr>
	<td align="Left" colspan="6"><u><b>Shuttle bus bill :<b></u></td>
</tr>
<tr>
	<td colspan="6">
	<table align="center" cellpadding="1" cellspacing="0" width="40%" border="1" bordercolor="black"> 
	<TR BGCOLOR="#000099" align="center" cellpadding="0" cellspacing="0" >
		<TD width="6%"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
		<TD><strong><label STYLE=color:#FFFFFF>Date</label></strong></TD>
		<TD width="10%"><strong><label STYLE=color:#FFFFFF>AM</label></strong></TD>
		<TD width="10%"><strong><label STYLE=color:#FFFFFF>PM</label></strong></TD>
		<TD width="20%"><strong><label STYLE=color:#FFFFFF>Tot. Shuttle Qty</label></strong></TD>
		<TD width="20%"><strong><label STYLE=color:#FFFFFF>Tot. Shuttle Bill($)</label></strong></TD>
	</TR>    
<% 
		dim no_ , TotalQty_ , TotalAmount_ 
		no_ = 1 
		TotalQty_ = 0
		TotalAmount_ = 0
		do while not rsShuttle.eof 
	   	if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
			TotalQty_ = TotalQty_ + rsShuttle("TotalPerDay")
%> 
 	<TR bgcolor="<%=bg%>">
		<td align="right"><%=No_%>&nbsp;</td>
        	<td><FONT color=#330099 size=2><%=rsShuttle("ShuttleDate")%>&nbsp;</font></td> 
        	<td align="right"><FONT color=#330099 size=2><%=rsShuttle("AM")%>&nbsp;</font></td> 
		<td align="right"><FONT color=#330099 size=2><%=rsShuttle("PM")%>&nbsp;</font></td> 
        	<td align="right"><FONT color=#330099 size=2><%=rsShuttle("TotalPerDay")%>&nbsp;</font></td> 
		<td align="right"><FONT color=#330099 size=2><%=rsShuttle("TotalAmountPerDay")%>&nbsp;</font></td>
  	 </TR>
<%   
		TotalAmount_ = TotalAmount_ + formatnumber(rsShuttle("TotalAmountPerDay"),-1)
		'response.write rsShuttle("TotalAmountPerDay")
		'response.write TotalAmount_ 
 		rsShuttle.movenext
   		no_ = no_ + 1
	loop
%>	
	<tr>
		<td align="right" colspan="4"><b>Total&nbsp;</b></td>
		<td width="10%" class="FontContent" align="right"><b><%=formatnumber(TotalQty_ ,-1)%></b></td>
		<td width="10%" class="FontContent" align="right"><b><%=formatnumber(TotalAmount_  ,-1)%></b></td>
	</tr>
	</table>
	</td>
</tr>
<tr>
	<td colspan="6"><hr></td>
</tr>
<!--
<tr>
	<td colspan="6">
	<table align="center" cellpadding="1" cellspacing="0" width="40%">
	<tr> -->
<!--		<td colspan="2">&nbsp;</td> -->
<!--		<td align="right" colspan="4"><b>Total&nbsp;</b></td>
		<td width="10%" class="FontContent" align="right"><b><%=formatnumber(TotalQty_ ,-1)%></b></td>
		<td width="10%" class="FontContent" align="right"><b><%=formatnumber(TotalAmount_  ,-1)%></b></td>
	</tr>
	</table>
	</td>
</tr>
-->
<%else%>
<tr>
	<td colspan="6" align="center">&nbsp;</td>	
</tr>
<tr>
	<td align="Left" colspan="6"><u><b>Shuttle bus bill :<b></u></td>
</tr>
<tr>
	<td colspan="6" align="center">there is not data.</td>	
</tr>
<%end if%>

<%if not rsHomePhone.eof then%>
<tr>
	<td colspan="6"><hr></td>
</tr>
<tr>
	<td align="Left" colspan="6"><u><b>Home phone bill :<b></u></td>
</tr>
<tr>
	<td colspan="6">
	<table cellspadding="1" cellspacing="0" width="100%">
	<tr>
		<td width="20%">&nbsp;Monthly Fee</td>
		<td width="1%">:</td>
		<td width="1%">Rp.</td>
		<td width="10%" class="FontContent" align="right"><%=formatnumber(rsHomePhone("Abonemen"),-1)%></td>
		<td width="10%">&nbsp;</td>
		<td width="15%">17</td>
		<td width="1%">:</td>
		<td width="1%">Rp.</td>
		<td width="10%" class="FontContent" align="right"><%=formatnumber(rsHomePhone("17"),-1)%></td>
	</tr>
	<tr>
		<td>&nbsp;Local call</td>
		<td width="1%">:</td>
		<td width="1%">Rp.</td>
		<td class="FontContent" align="right"><%=formatnumber(rsHomePhone("Lokal"),-1)%></td>
		<td width="10%">&nbsp;</td>
		<td>Operator</td>
		<td width="1%">:</td>
		<td width="1%">Rp.</td>
		<td class="FontContent" align="right"><%=formatnumber(rsHomePhone("Operator"),-1)%></td>
	</tr>
	<tr>
		<td>&nbsp;SLJJ</td>
		<td width="1%">:</td>
		<td width="1%">Rp.</td>
		<td class="FontContent" align="right"><%=formatnumber(rsHomePhone("SLJJ"),-1)%></td>
		<td width="10%">&nbsp;</td>
		<td>Air Time</td>
		<td width="1%">:</td>
		<td width="1%">Rp.</td>
		<td class="FontContent" align="right"><%=formatnumber(rsHomePhone("AirTime"),-1)%></td>
	</tr>
	<tr>
		<td>&nbsp;STB</td>
		<td width="1%">:</td>
		<td width="1%">Rp.</td>
		<td class="FontContent" align="right"><%=formatnumber(rsHomePhone("STB"),-1)%></td>
		<td width="10%">&nbsp;</td>
		<td>Quota</td>
		<td width="1%">:</td>
		<td width="1%">Rp.</td>
		<td class="FontContent" align="right"><%=formatnumber(rsHomePhone("Quota"),-1)%></td>
	</tr>
	<tr>
		<td>&nbsp;JAPATI</td>
		<td width="1%">:</td>
		<td width="1%">Rp.</td>
		<td class="FontContent" align="right"><%=formatnumber(rsHomePhone("JAPATI"),-1)%></td>
		<td width="10%">&nbsp;</td>
		<td>Miscellaneous</td>
		<td width="1%">:</td>
		<td width="1%">Rp.</td>
		<td class="FontContent" align="right"><%=formatnumber(rsHomePhone("Lain2"),-1)%></td>
	</tr>
	<tr>
		<td>&nbsp;Int. call (SLI007)</td>
		<td width="1%">:</td>
		<td width="1%">Rp.</td>
		<td class="FontContent" align="right"><%=formatnumber(rsHomePhone("SLI007"),-1)%></td>
		<td width="10%">&nbsp;</td>
		<td>Tax</td>
		<td width="1%">:</td>
		<td width="1%">Rp.</td>
		<td class="FontContent" align="right"><%=formatnumber(rsHomePhone("PPN"),-1)%></td>
	</tr>
	<tr>
		<td>&nbsp;Int. call (001+008)</td>
		<td width="1%">:</td>
		<td width="1%">Rp.</td>
		<td class="FontContent" align="right"><%=formatnumber(rsHomePhone("001+008"),-1)%></td>
		<td width="10%">&nbsp;</td>
		<td>Stamp fee</td>
		<td width="1%">:</td>
		<td width="1%">Rp.</td>
		<td class="FontContent" align="right"><%=formatnumber(rsHomePhone("Meterai"),-1)%></td>
	</tr>
	<tr>
		<td colspan="9"><hr></td>
	</tr>
	<tr>
		<td colspan="5">&nbsp;</td>
		<td align="right"><b>Sub Total&nbsp;</b></td>
		<td width="1%">:</td>
		<td width="1%">Rp.</td>
		<td class="FontContent" align="right"><b><%=formatnumber(TotalHomeBill_ ,-1)%></b></td>
		<td width="10%">&nbsp;</td>
	</tr>
<!--
	<tr>
		<td colspan="5">&nbsp;</td>
		<td><b>Payment Status</b></td>
		<td width="1%">:</td>
		<td width="1%">&nbsp;</td>
		<td class="FontContent" align="right"><b><%=rsHomePhone("Status")%></b></td>
	</tr>	
-->
	</table>
	</td>
</tr>
<%else%>
<tr>
	<td colspan="6" align="center">&nbsp;</td>	
</tr>
<tr>
	<td colspan="6"><hr></td>
</tr>
<tr>
	<td align="Left" colspan="6"><u><b>Home phone bill :<b></u></td>
</tr>
<tr>
	<td colspan="6" align="center">there is not data.</td>	
</tr>
<%end if%>
<%if not rsOfficePhone.eof then%>
<tr>
	<td colspan="6"><hr></td>
</tr>
<tr>
	<td align="Left" colspan="6"><u><b>Office phone bill :<b></u></TD>
</tr>
<tr>
	<td colspan="6" align="Center">
		<table cellspadding="0" cellspacing="0" bordercolor="black" border="1" width="80%" bgColor="white">  
		<tr bgcolor="#330099" align="center" cellpadding="0" cellspacing="0" >
			<TD width="5%"><strong><label STYLE=color:#FFFFFF>No.</label></strong></TD>
		       	<TD><strong><label STYLE=color:#FFFFFF>Dialed Date/time</label></strong></TD>
			<TD width="20%"><strong><label STYLE=color:#FFFFFF>Dialed Number</label></strong></TD>
			<TD width="15%"><strong><label STYLE=color:#FFFFFF>Amount (Rp.)</label></strong></TD>
			<TD width="15%"><strong><label STYLE=color:#FFFFFF>Personal used</label></strong><br>
<%			if (Status_ = "Pending") or (Status_ = "Correction") then %>
				<input type="checkbox" name="cbAll" value="true" onclick="checkall(this)" />
<%			end if %>
			</TD>
		</tr>
		<%
		strsql = "Exec spGetBilling 'Detail','" & user1_ & "','" & MonthP & "','" & YearP & "'"
		'response.write strsql & "<br>"
		set rsOfficePhone = BillingCon.execute(strsql) 
		no_ = 1 
		do while not rsOfficePhone.eof
   			if bg="#D7E3F4" then bg="ffffff" else bg="#D7E3F4" 
		%>
			<tr bgcolor="<%=bg%>">
				<td align="right"><%=No_%>&nbsp;</td>
			        <td><FONT color=#330099 size=2>&nbsp;<%=rsOfficePhone("DialedDatetime")%></font></td> 
		        	<td><FONT color=#330099 size=2>&nbsp;<%=rsOfficePhone("DialedNumber")%></font></td> 
			        <td align="right"><FONT color=#330099 size=2><%=formatnumber(rsOfficePhone("Cost"),-1)%>&nbsp;</font></td> 
<%			if (Status_ = "Pending") or (Status_ = "Correction") then %>
			        <td align="center">
				<%if rsOfficePhone("isPersonal") = "Y" then%>
					<Input type="Checkbox" Id="cbPersonal" name="cbPersonal" Value='<%=rsOfficePhone("CallRecordID")%>' Checked>				
				<%else%>
					<Input type="Checkbox" Id="cbPersonal" name="cbPersonal" Value='<%=rsOfficePhone("CallRecordID")%>' >
				<%end if%>
				</td>
<%			else%>
			        <td align="center">
				<%if rsOfficePhone("isPersonal") = "Y" then%>
					<Input type="Checkbox" Id="cbPersonal" name="cbPersonal" Value='<%=rsOfficePhone("CallRecordID")%>' Checked disabled>				
				<%else%>
					<Input type="Checkbox" Id="cbPersonal" name="cbPersonal" Value='<%=rsOfficePhone("CallRecordID")%>'  disabled>
				<%end if%>
				</td>
<%			end if %>
		<%   
			rsOfficePhone.movenext
			no_ = no_ + 1
		loop
		%>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td colspan="6" align="Right">
		<table cellspadding="0" cellspacing="0" bordercolor="black" width="80%" bgColor="white">  
		<tr>
			<td width="70%" align="right"><b>Sub Total </b>&nbsp;</td>
			<td width="1%">:</td>
			<td width="1%">&nbsp;Rp.</td>
			<td class="FontContent" align="right"><b><%=formatnumber(TotalOffBill_,-1)%></b></td>
			<td width="15%">&nbsp;</td>
		</tr>
<!--
		<tr>
			<td align="right"><b>Payment Status</b>&nbsp;</td>
			<td width="1%">:</td>
			<td width="1%">&nbsp;Rp.</td>
			<td class="FontContent" align="right"><b><%=Status_ %></b></td>
			<td width="15%">&nbsp;</td>
		</tr>
-->
		</table>
	</td>
</tr>
<tr>
	<td colspan="6">
		<table cellspadding="1" cellspacing="0" bgColor="white" width="100%">  
		<tr>
			<td width="12%">Supervisor Email</td>
			<td width="1%">:</td>
			<td> <input type="input" name="txtSpvEmail" size="50" value='<%=SpvEmail_%>' /></td>
		</tr>
		<tr>
			<td valign="top">Note</td>
			<td valign="top" width="1%">:</td>
			<td>
				<TextArea name="txtNotes" Rows="5" Cols="90" Wrap <% if (Status_ <> "Pending") and (Status_ <> "Correction") then%>ReadOnly<%End If%> ><%=Notes_%></textarea>
			</td>
		</tr>
<%		if (Status_ <> "Pending")then%>
		<tr>
			<td colspan="3"><b>Remarks/Correction(s) :</b></td>
		</tr>
		<tr>
			<td colspan="3">
				<TextArea name="txtRemark" Rows="5" Cols="90" Wrap <% if (Status_ <> "Pending") or (Status_ <> "Correction") then%>ReadOnly<%End If%>><%=SpvRemark_ %></textarea>
			</td>
		</tr>
<%		end if%>
		</table>
	</td>
</tr>
<tr>
	<td colspan="6" align="center">&nbsp;</td>
</tr>
<%		if (Status_ = "Pending") or (Status_ = "Correction") then%>
<tr>
	<td colspan="6" align="center">	
		<input type="submit" name="btnSubmit" Value="Save" />&nbsp;&nbsp;
		<input type="submit" name="btnSubmit" Value="Save & Submit to Supervisor" />
		<input type="hidden" name="txtExtension" value='<%=Ext_ %>' />
		<input type="hidden" name="txtMonthP" value='<%=MonthP%>' />
		<input type="hidden" name="txtYearP" value='<%=YearP%>' />
		<input type="hidden" name="txtEmpName" value='<%=EmpName_%>' />
		<input type="hidden" name="txtPeriod" value='<%=Period_%>' />
		<input type="hidden" name="txtOffice" value='<%=Office_%>' />
		<input type="hidden" name="txtTotalCost" value='<%=TotalCost_%>' />
		<input type="hidden" name="txtLoginID" value='<%=user1_%>' />
	<td>
</tr>
<%		end if%>
<%else%>
<tr>
	<td colspan="6" align="center">&nbsp;</td>	
</tr>
<tr>
	<td colspan="6"><hr></td>
</tr>
<tr>
	<td align="Left" colspan="6"><u><b>Office phone bill :<b></u></TD>
</tr>
<tr>
	<td colspan="6" align="center">there is not data.</td>	
</tr>
<%end if%>
</table>
</form>
</BODY>
</html>