<!-- Module 21 -->  <!-- practiceview.asp -->
<!-- #Include file="../inc/connect.inc" -->
<%
txtPageHeader = txtPageHeader & "Health Care Professionals"
%>
<!-- #Include file="../inc/mod-top.asp" -->
<%
Select Case Request.QueryString("act")
	Case Else
		intGroupID = Session("GID")
		sqlPractice = "SELECT tblPurchase.Quantity, tblCoupon.Code, tblStudent.LastName, tblStudent.FirstName, " &_
			"tblStressSchedule.TestDate, tblStressSchedule.Answers, tblStudent.StudentID, tblStressSchedule.StressScheduleID " &_
			"FROM SCUser.tblStressSchedule tblStressSchedule INNER JOIN SCUser.tblPurchase tblPurchase ON tblStressSchedule.PurchaseID = tblPurchase.PurchaseID " &_
			"INNER JOIN SCUser.tblStudent tblStudent ON tblStressSchedule.StudentID = tblStudent.StudentID " &_
			"INNER JOIN SCUser.tblCoupon tblCoupon ON tblPurchase.CouponID = tblCoupon.Code " &_
			"INNER JOIN SCUser.tblGroup tblGroup ON tblGroup.GroupID = tblCoupon.GroupID " &_
			"WHERE tblGroup.GroupID = " & intGroupID & " AND ((LEN(tblStressSchedule.Answers) = 60) OR " &_
			"(LEN(tblStressSchedule.Answers) IS NULL) OR (LEN(tblStressSchedule.Answers) = '')) " &_
			"ORDER BY tblCoupon.Code, tblStudent.LastName, tblStudent.FirstName, tblStressSchedule.StressScheduleID"
			
		'sqlPractice = "SELECT tblCoupon.Code FROM tblCoupon"
		
		'Response.Write(sqlPractice)
		
		Set rsPractice = dbCon.Execute(sqlPractice)
 		
		If rsPractice.BOF AND rsPractice.EOF Then
			Response.Write("<p style=""color:red; text-align: center;"">Sorry, no clients have taken Stresscom at this time.</p>")
		Else
%>
			<table>
				<tr>
					<td class="bodyhead">Patient</td>
					<td class="bodyhead">Test Date</td>
					<td class="bodyhead">Ctrl</td>
					<td class="bodyhead">Comp</td>
					<td class="bodyhead">Task</td>
					<td class="bodyhead">Change</td>
					<td class="bodyhead">Symp</td>
					<td class="bodyhead">Time</td>
					<td class="bodyhead">Total</td>
					<td class="bodyhead">&nbsp;</td>
				</tr>
<%
			Do While NOT rsPractice.EOF
			
			
			
			
'Calculate the score of each factor
txtAnswers = rsPractice("Answers")

If IsNull(txtAnswers) or Len(txtAnswers) = 0 or txtAnswers = "" Then
	intA = "n/a"
	intB = "n/a"
	intC = "n/a"
	intD = "n/a"
	intE = "n/a"
	intF = "n/a"
	intTotal = "n/a"
Else

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
End If		
			
			
			
			
			
%>
				<tr>
					<td class="bodycopy"><%= rsPractice("LastName") %>, <%= rsPractice("FirstName") %></td>
					<td class="bodycopy"><%= rsPractice("TestDate") %></td>
					<td class="bodycopy"><%= intA %></td>
					<td class="bodycopy"><%= intB %></td>
					<td class="bodycopy"><%= intC %></td>
					<td class="bodycopy"><%= intD %></td>
					<td class="bodycopy"><%= intE %></td>
					<td class="bodycopy"><%= intF %></td>
					<td class="bodycopy"><%= intTotal %></td>
					<td class="bodycopy">
<%
			If IsNull(rsPractice("TestDate")) OR rsPractice("TestDate") = "" Then
				Response.Write("N/A")
			Else
				Response.Write("<a href=""index.asp?mod=23&sid=" & rsPractice("StudentID") & "&tid=" & rsPractice("StressScheduleID") & """>[View/Print Results]</a>")
			End IF
%>
					</td>
				</tr>
<%
				rsPractice.MoveNext
			Loop
%>
			</table>
<%		
		End If
%>
			
<%
End Select
%>
<!-- #Include file="../inc/mod-bot.asp" -->