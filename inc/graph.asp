<%
'2012-08-29 thermometer data

'high score = 207
'normal = 172
'low = 153

' add 6 to normal for high score
' subtract 3 from normal for low score

'Calculate the score of each factor
sqlFactor = "SELECT StatementID, Factor FROM SCUser.tblStatement"
Set rsFactor = dbCon.Execute(sqlFactor)
intA = 0
intB = 0
intC = 0
intD = 0
intE = 0
intF = 0
intTotal = 0
Do While Not rsFactor.EOF
	Select Case rsFactor("Factor")
		Case "A"
			intPos = rsFactor("StatementID")
			intScore = cInt(Right(Left(txtAnswers, intPos), 1))
			intA = intA + intScore
		Case "B"
			intPos = rsFactor("StatementID")
			intScore = cInt(Right(Left(txtAnswers, intPos), 1))
			intB = intB + intScore
		Case "C"
			intPos = rsFactor("StatementID")
			intScore = cInt(Right(Left(txtAnswers, intPos), 1))
			intC = intC + intScore
		Case "D"
			intPos = rsFactor("StatementID")
			intScore = cInt(Right(Left(txtAnswers, intPos), 1))
			intD = intD + intScore
		Case "E"
			intPos = rsFactor("StatementID")
			intScore = cInt(Right(Left(txtAnswers, intPos), 1))
			intE = intE + intScore
		Case "F"
			intPos = rsFactor("StatementID")
			intScore = cInt(Right(Left(txtAnswers, intPos), 1))
			intF = intF + intScore
	End Select
	intTotal = intTotal + intScore
	rsFactor.MoveNext
Loop
Set rsFactor = Nothing


sqlWatch = "SELECT * FROM SCUser.tblFactor ORDER BY Avarage"
Set rsWatch = dbCon.Execute(sqlWatch)
intWatchLow = cInt(rsWatch("Avarage"))
Do While Not rsWatch.EOF
	Select Case rsWatch("FactorID")
		Case "A"
			intAWatch = cInt(rsWatch("Avarage"))
			txtAFactor = rsWatch("Factor")
			intAWatchDif = intAwatch - intWatchLow
		Case "B"
			intBWatch = cInt(rsWatch("Avarage"))
			txtBFactor = rsWatch("Factor")
			intBWatchDif = intBwatch - intWatchLow
		Case "C"
			intCWatch = cInt(rsWatch("Avarage"))
			txtCFactor = rsWatch("Factor")
			intCWatchDif = intCwatch - intWatchLow
		Case "D"
			intDWatch = cInt(rsWatch("Avarage"))
			txtDFactor = rsWatch("Factor")
			intDWatchDif = intDwatch - intWatchLow
		Case "E"
			intEWatch = cInt(rsWatch("Avarage"))
			txtEFactor = rsWatch("Factor")
			intEWatchDif = intEwatch - intWatchLow
		Case "F"
			intFWatch = cInt(rsWatch("Avarage"))
			txtFFactor = rsWatch("Factor")
			intFWatchDif = intFwatch - intWatchLow
	End Select
	rsWatch.MoveNext
Loop
Set rsWatch = Nothing

'set the bar color red = below watch / blue = above watch
GrayBlock = "<img src=""../images/ch_gray.jpg"">"
If intA > intAWatch + 6 Then 
	ABlock = "<img src=""../images/ch_below.jpg"">"
Else
	ABlock = "<img src=""../images/ch_above.jpg"">"
End If
If intB > intBWatch + 6 Then 
	BBlock = "<img src=""../images/ch_below.jpg"">"
Else
	BBlock = "<img src=""../images/ch_above.jpg"">"
End If
If intC > intCWatch + 6 Then 
	CBlock = "<img src=""../images/ch_below.jpg"">"
Else
	CBlock = "<img src=""../images/ch_above.jpg"">"
End If
If intD > intDWatch + 6 Then 
	DBlock = "<img src=""../images/ch_below.jpg"">"
Else
	DBlock = "<img src=""../images/ch_above.jpg"">"
End If
If intE > intEWatch + 6 Then 
	EBlock = "<img src=""../images/ch_below.jpg"">"
Else
	EBlock = "<img src=""../images/ch_above.jpg"">"
End If
If intF > intFWatch + 6 Then 
	FBlock = "<img src=""../images/ch_below.jpg"">"
Else
	FBlock = "<img src=""../images/ch_above.jpg"">"
End If

'draw the chart
%>
Your total score was: <strong><%= intTotal %></strong>
<div align="center">
<table  border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td class="bodycopy" align="right" valign="top"><img src="../images/spacer.gif" width="1" height="185"><span style="color:red">High</span>&nbsp;<br><img src="../images/spacer.gif" width="1" height="48"><span style="color:#999900">Average</span>&nbsp;<br><img src="../images/spacer.gif" width="1" height="25"><span style="color:green">Low</span>&nbsp;<br></td>
		<td>			
<table border="0" cellspacing="1" cellpadding="0">
<tr>
	<td colspan="6"><img src="../images/ch_factors.jpg" width="424" height="50" border="0"></td>
</tr>
<% 
i = 50
Do While i > 10
	If i = intWatchLow - 3 Then
%>
<tr>
	<td colspan="6" align="center"><img src="../images/ch_green.jpg" width="424" height="2"></td>
</tr>
<%
	End If
	If i = intWatchLow  Then
%>
<tr>
	<td colspan="6" align="center"><img src="../images/ch_yellow.jpg" width="424" height="2"></td>
</tr>
<%
	End If
	If i = intWatchLow + 6 Then
%>
<tr>
	<td colspan="6" align="center"><img src="../images/ch_red.jpg" width="424" height="2"></td>
</tr>
<%
	End If
%>
<tr>
	<td align="center"><% If (intA - intAWatchDif) >= i Then Response.Write(ABlock) Else Response.Write(GrayBlock) %></td>
	<td align="center"><% If (intB - intBWatchDif) >= i Then Response.Write(BBlock) Else Response.Write(GrayBlock) %></td>
	<td align="center"><% If (intC - intCWatchDif) >= i Then Response.Write(CBlock) Else Response.Write(GrayBlock) %></td>
	<td align="center"><% If (intD - intDWatchDif) >= i Then Response.Write(DBlock) Else Response.Write(GrayBlock) %></td>
	<td align="center"><% If (intE - intEWatchDif) >= i Then Response.Write(EBlock) Else Response.Write(GrayBlock) %></td>
	<td align="center"><% If (intF - intFWatchDif) >= i Then Response.Write(FBlock) Else Response.Write(GrayBlock) %></td>
</tr>
<% 
	i = i - 1
Loop
%>
</table>
		</td>
	</tr>
	<tr>
		<td class="BodyCopy" align="right">Your Factor Scores</td>
		<td>			
<table border="0" cellspacing="1" cellpadding="0" width="100%">
<tr>
	<td class="BodyCopy" align="center">&nbsp;<br><%= intA %></td>
	<td class="BodyCopy" align="center">&nbsp;<br><%= intB %></td>
	<td class="BodyCopy" align="center">&nbsp;<br><%= intC %></td>
	<td class="BodyCopy" align="center">&nbsp;<br><%= intD %></td>
	<td class="BodyCopy" align="center">&nbsp;<br><%= intE %></td>
	<td class="BodyCopy" align="center">&nbsp;<br><%= intF %></td>
</tr>
</table>
		</td>
	</tr>
</table>
</div>
<%	
'calculate the score range for the score page
If intA > intAWatch + 6 Then 
	txtAFactorScore = "high"
ElseIf intA <= intAWatch + 6 AND intA >= intAWatch Then
	txtAFactorScore = "average"
Else
	txtAFactorScore = "low"
End If

If intB > intBWatch + 6 Then 
	txtBFactorScore = "high"
ElseIf intB <= intBWatch + 6 AND intB >= intBWatch Then
	txtBFactorScore = "average"
Else
	txtBFactorScore = "low"
End If

If intC > intCWatch + 6 Then 
	txtCFactorScore = "high"
ElseIf intC <= intCWatch + 6 AND intC >= intCWatch Then
	txtCFactorScore = "average"
Else
	txtCFactorScore = "low"
End If

If intD > intDWatch + 6 Then 
	txtDFactorScore = "high"
ElseIf intD <= intDWatch + 6 AND intD >= intDWatch Then
	txtDFactorScore = "average"
Else
	txtDFactorScore = "low"
End If

If intE > intEWatch + 6 Then 
	txtEFactorScore = "high"
ElseIf intE <= intEWatch + 6 AND intE >= intEWatch Then
	txtEFactorScore = "average"
Else
	txtEFactorScore = "low"
End If

If intF > intFWatch + 6 Then 
	txtFFactorScore = "high"
ElseIf intF <= intFWatch + 6 AND intF >= intFWatch Then
	txtFFactorScore = "average"
Else
	txtFFactorScore = "low"
End If

%>