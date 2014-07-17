<%
Response.Buffer = False
server.scriptTimeout =25000
%>
<div id="ProgBar" style="font-family:Verdana; font-size=9pt;">Progress:<BR>
<TABLE style="color:red;" HEIGHT="16" Border=1><TR><TD BGCOLOR=RED ID=statuspic></TD></TR></TABLE><BR>
</div>
<script language="Javascript">var progBarWidth=250;</script>
<%
iProcessedSoFar = 0
iTotalRecords = 1000
strHTML=" <Table width=""80%""><TR><TD Width=""100%"" BGCOLOR=""gray"" align=""CENTER"">Results:<td></tr>"
for i = 0 to iTotalRecords
' next few lines are just a surrogate for whatever your processing function to be timed
	strHTML = strHTML & "<tr width=""100%""><td width=""100%"" BGCOLOR=""#FFCC66""> Your results</td></tr>"
	iProcessedSoFar = iProcessedSoFar + 1
	pctComplete = (iProcessedSoFar / iTotalRecords)
	if i mod 8 = 0 then
		ShowProgress pctComplete
	end if
next

FinishProgress
strHTML=strHTML & "</TABLE>"
Response.write strHTML

Sub ShowProgress(nPctComplete)
	Response.Write "<SCR" & "IPT LANGUAGE=""JavaScript"">" & vbCrlf
	Response.Write "statuspic.width = Math.ceil(" & nPctComplete & " * progBarWidth);" & vbCrlf
	Response.Write "</SCR" & "IPT>"
End Sub

Sub FinishProgress
	Response.Write "<SCR" & "IPT LANGUAGE=""JavaScript"">" & vbCrlf
	Response.Write "ProgBar.style.visibility ='hidden';" & vbCrLf
	Response.Write "</SCR" & "IPT>"
end sub
%>