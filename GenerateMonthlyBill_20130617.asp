
<!--#include file="connect.inc" -->
<html>
<head>
<script language="JavaScript" src="calendar.js"></script>
<STYLE TYPE="text/css"><!--
  A:ACTIVE { color:#003399; font-size:8pt; font-family:Verdana; }
  A:HOVER { color:#003399; font-size:8pt; font-family:Verdana; }
  A:LINK { color:#003399; font-size:8pt; font-family:Verdana; }
  A:VISITED { color:#003399; font-size:8pt; font-family:Verdana; }
  body {scrollbar-3dlight-color:#FFFFFF; scrollbar-arrow-color:#E3DCD5; scrollbar-base-color:#FFFFFF; scrollbar-darkshadow-color:#FFFFFF;	scrollbar-face-color:#FFFFFF; scrollbar-highlight-color:#E3DCD5; scrollbar-shadow-color:#E3DCD5; }
  p { font-family: verdana; font-size: 12px; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; color: #003399; text-decoration: none}
  h3 { font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 16px; font-style: normal; line-height: normal; font-weight: bold; color: #003399; letter-spacing: normal; word-spacing: normal; font-variant: small-caps}
  td { font-family: verdana; font-size: 12px; font-style: normal; font-weight: normal; color: #003399}
  .title { font-size:16px; font-weight:bold; color:#000080; }
  .SubTitle { font-size:16px; font-weight:bold; color:#000080;  }
  A.menu { text-decoration:none; font-weight:bold; }
  A.mmenu { text-decoration:none; color:#FFFFFF; font-weight:bold; }
  .normal { font-family:Verdana,Arial; color:black}
  .style5 {color: #FFFFFF;}
  .ActivePage {color: red; font-weight:bold; }
--></STYLE>
<script language="javascript">
		<!--

		var timerID = null;
		var timerRunning = false;
		var timeValue = 1000;  //the time increment in mS
		var count = 0;
		var finish = false;
		//load up the images for the progress bar
		image00 = new Image(); image00.src='images/image-00.gif';
		image01 = new Image(); image01.src='images/image-01.gif';
		image02 = new Image(); image02.src='images/image-02.gif';
		image03 = new Image(); image03.src='images/image-03.gif';
		image04 = new Image(); image04.src='images/image-04.gif';
		image05 = new Image(); image05.src='images/image-05.gif';
		image06 = new Image(); image06.src='images/image-06.gif';
		image07 = new Image(); image07.src='images/image-07.gif';
		image08 = new Image(); image08.src='images/image-08.gif';
		image09 = new Image(); image09.src='images/image-09.gif';
		image10 = new Image(); image10.src='images/image-10.gif';


		function increment() {
			count += 1;
			if (count == 0) {document.images.bar.src=image00.src;}
			if (count == 1) {document.images.bar.src=image01.src;}
			if (count == 2) {document.images.bar.src=image02.src;}
			if (count == 3) {document.images.bar.src=image03.src;}
			if (count == 4) {document.images.bar.src=image04.src;}
			if (count == 5) {document.images.bar.src=image05.src;}
			if (count == 6) {document.images.bar.src=image06.src;}
			if (count == 7) {document.images.bar.src=image07.src;}
			if (count == 8) {document.images.bar.src=image08.src;}
			if (count == 9) {document.images.bar.src=image09.src;}
			//If you want it to repeat the bar continuously then use this line:
			if (count == 10) {document.images.bar.src=image10.src; count=-1;}
			//If you want it to stop repeating the bar then use this line:
			//if (count == 10) {document.images.bar.src=image10.src; end();}
		}

		function stopclock() {
			if (timerRunning)
				clearInterval(timerID);
			timerRunning = false;	
		}

		function end() {
			if (finish == true) {
				stopclock();
				window.close();
			}
			else {
				finish = true; 
			}
		}

		function startclock() {
			stopclock();
			timerID = setInterval("increment()", timeValue);
			timerRunning = true;
			document.images.bar.src=image00.src;
		}

		function Send_onclick(frmSubmit) {
'			document.images.bar.style.display = 'block';
'			startclock();
'			frmSubmit.submit();			
		}

		//-->
		</script>
</head>
<BODY BGCOLOR="#FFFFFF" TEXT="#000000" LINK="#0080FF" ALINK="0080FF" VLINK="#0055AA" MARGINWIDTH="8" MARGINHEIGHT="0" LEFTMARGIN="8" TOPMARGIN="0">
  <Center><FONT COLOR=#009900><B>SENSITIVE BUT UNCLASSIFIED</Center></FONT></B>
  <BR>
<CENTER>
  <IMG SRC="images/embassytitle2.jpeg" WIDTH="661" HEIGHT="80" BORDER="0"> 
  <TABLE WIDTH="65%" BORDER="0" CELLPADDING="0" CELLSPACING="0">
  <CAPTION><H3 STYLE="font-size:17px;color:#000040">Mission Jakarta - Billing Application</H3></CAPTION>
  <TR>
  	<TD COLSPAN="4" ALIGN="center" Class="title">GENERATE MONTHLY BILL</TD>
   </TR>
<tr>
        <td colspan="3" align="left"><FONT color=#330099 size=2><A HREF="Default.asp">Main Menu</A></font></TD>
        <td align="right"><FONT color=#330099 size=2><A HREF="AdminPage.asp">Back</A></font></TD>
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
strsql = "select * from Users where loginId='" & user1_ & "'"
set RS_Query = server.createobject("adodb.recordset")
'response.write strsql & "<br>"
set RS_Query = BillingCon.execute(strsql)

if not RS_Query.eof then
	UserRole_ = RS_Query("RoleID")
end if

curMonth_ = month(date())
curYear_ = year(date())
if len(curMonth_)= 1 then
	curMonth_ = "0" & curMonth_
end if

MonthP = request("MonthP")

if MonthP = "" Then MonthP = Request.Form("MonthList")
if MonthP = "" then
	MonthP = curMonth_ 
end if
'response.write MonthP

YearP = request("YearP")
if YearP ="" Then YearP = Request.Form("YearList")
if YearP ="" then
	YearP = curYear_ 
end if


%>  
<form method="post" name="frmSearch" Action="GenerateMonthlyBillProcess.asp" onSubmit="return confirm('Are you sure ?');">
<table cellspacing="0" cellpadding="2">  
<%if (UserRole_= "Admin") or (UserRole_= "FMC") or (UserRole_= "Cashier") or (UserRole_= "Voucher") then %>
<tr align="Center">
	<td colspan="2" align="center">
		<table  width="100%">
		<tr bgcolor="#000099">
			<td height="25" colspan="3"><strong>&nbsp;<span class="style5">Criteria(s): </span></strong></td>
		</tr>
		<tr>
			<td width="15%" align="right">&nbsp;Period&nbsp;</td>				
			<td>:</td>
			<td>
				<Select name="MonthList">
					<Option value="01" <%if MonthP ="01" then %>Selected<%End If%>>January</Option>
					<Option value="02" <%if MonthP ="02" then %>Selected<%End If%>>February</Option>
					<Option value="03" <%if MonthP ="03" then %>Selected<%End If%>>March</Option>
					<Option value="04" <%if MonthP ="04" then %>Selected<%End If%>>April</Option>
					<Option value="05" <%if MonthP ="05" then %>Selected<%End If%>>May</Option>
					<Option value="06" <%if MonthP ="06" then %>Selected<%End If%>>June</Option>
					<Option value="07" <%if MonthP ="07" then %>Selected<%End If%>>July</Option>
					<Option value="08" <%if MonthP ="08" then %>Selected<%End If%>>August</Option>
					<Option value="09" <%if MonthP ="09" then %>Selected<%End If%>>September</Option>
					<Option value="10" <%if MonthP ="10" then %>Selected<%End If%>>October</Option>
					<Option value="11" <%if MonthP ="11" then %>Selected<%End If%>>November</Option>
					<Option value="12" <%if MonthP ="12" then %>Selected<%End If%>>December</Option>
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
		</tr>	
		<tr>
			<td align="right">&nbsp;Employee&nbsp;</td>
			<td>:</td>
			<td>
<%
 				strsql ="select EmpID, EmpName from vwPhoneCustomerList order by EmpName"
				set EmpRS = server.createobject("adodb.recordset")
				set EmpRS = BillingCon.execute(strsql)
'				response.write strStr 
%>	
				<Select name="cmbEmp">
					<Option value=''>--All--</Option>
<%				Do While not EmpRS.eof 
%>
					<Option value='<%=EmpRS("EmpID")%>' <%if trim(EmpID_) = trim(EmpRS("EmpID")) then %>Selected<%End If%> ><%=EmpRS("EmpName") %></Option>
					
<%					EmpRS.MoveNext
				Loop%>
				</select>

			</td>	
		</tr>	
		<tr>
			<td colspan="3"><br></td>
		</tr>
		<tr>
			<td colspan="2">&nbsp;</td>
			<td align="Left">
				<input type="submit" name="Submit" value="Process" LANGUAGE=javascript onclick="return Send_onclick(frmSearch)">
			</td>
		</tr>
		<tr>
			<td colspan="3" align="center">
				<img src="images/image-00.gif" name="bar" style="display: none;" align="middle" alt="Please wait.">
			</td>
		</tr>
		</table>
	</td>
</tr>	
<%else %>

<tr>
	<td>You do not have permission to access this site.</td>
</tr>
<tr>
	<td>Please <a href="/CSC">Submit Request </a> or contact Jakarta CSC Helpdesk at ext.9111.</td>
</tr>
<%end if %>
</table>
</form>
</BODY>
</html>