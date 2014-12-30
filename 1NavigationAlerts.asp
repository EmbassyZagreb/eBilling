
<div id="extra">
			<ul id="menu">
				<li>
					<a href="#">Cell Phones</a>
					<ul>
				

						<li class="current"><a href="CPSendNotification.asp">Send Notifications</a></li>	
						<li><a href="#">Send Reminders</a></li>	
						<li><a href="#">Remind Supervisors</a></li>							
					</ul>
				</li>
				<li>
					<a href="#">Residential Phones</a>
					<ul>
						<li><a href="#">Not available yet</a></li>
					</ul>
				</li>
			</ul>
</div>

</div>

<% If maintenancemode Then %>
	<div id="maintenancemode"><h1><%=SiteHeader%> - MAINTENANCE MODE - Website not available for the public!</h1> 
<% Else %>
	<div id="header"><h1><%=SiteHeader%></h1>
<% End If %>
<table widht=100% bolder = 0 cellpadding=0 cellspacing=0>
	<tr>
			
		<td colspan=2  style="background: url(images/top-navigation-slice.jpg) repeat left top;">
			<div id="navigation-left">
			<ul>
            	<li><a href="1MonthlyBilling.asp">Home</a></li>
                <li><a href="1BillingApproval.asp">Approve</a></li>
		<li><a href="default.asp">Cashier</a></li>
                <li><a href="default.asp">Manage</a></li>
		<li id="active"><a href="1SendNotification.asp">Alerts</a></li> 	            
		<li><a href="default.asp">Reports</a></li>
			</ul>
 		</div> 
		</td>
		
		<td  width=100% style="background: url(images/top-navigation-slice.jpg) repeat left top;">
		&nbsp;
		</td>
		<td widht=130px>
		<div id="navigation-right">
			<ul>
   	            <li><a href="default.asp">Admin</a></li>
				<li><a href="default.asp">Help</a></li>
			</ul>
		</div>  
		</td>		
	</tr>
</table> 
</div>