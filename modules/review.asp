<!-- Module 11 -->  <!-- review.asp -->
<!-- #Include file="../inc/connect.inc" -->
<%
txtPageHeader = "Select the Stress Schedule to review the outcome"
%>
<!-- #Include file="../inc/mod-top.asp" -->
<%
Select Case Request.QueryString("act")
	Case Else
		'check to see if ssfi has been completed
		sqlTest = "SELECT tblStressSchedule.StressScheduleID, tblStressSchedule.TestDate, tblStressSchedule.Answers, tblStudent.FirstName, tblStudent.LastName " &_
			"FROM SCUser.tblStressSchedule tblStressSchedule LEFT OUTER JOIN SCUser.tblStudent tblStudent ON tblStressSchedule.StudentID = tblStudent.StudentID " &_
			"WHERE tblStressSchedule.StudentID = " & Session("SID") & " AND Len(tblStressSchedule.Answers) = 60"
		Set rsTest = dbCon.Execute(sqlTest)
		
		If rsTest.BOF AND rsTest.EOF Then
			Response.Write("There are not any Stress Schedule scores for you to review.")
		Else
%>
<table>
	<tr>
		<td class="BodyCopy" align="center">Stress Schedule ID</td>
		<td width="10"></td>
		<td class="BodyCopy" align="center">Student</td>
		<td width="10"></td>
		<td class="BodyCopy" align="center">Test Taken</td>
		<td width="10"></td>
		<td></td>
	</tr>
<%
			Do While Not rsTest.EOF
%>
	<tr>
		<td class="BodyCopy" align="center"><%= rsTest("StressScheduleID") %></td>
		<td width="10"></td>
		<td class="BodyCopy" align="center"><%= rsTest("FirstName") %></td>
		<td width="10"></td>
		<td class="BodyCopy" align="center"><%= rsTest("TestDate") %></td>
		<td width="10"></td>
		<td><button onClick="window.location='index.asp?mod=12&tid=<%= rsTest("StressScheduleID") %>'">Review Scores</button></td>
	</tr>
<%				
				rsTest.MoveNext
			Loop
%>
	<tr>
		<td colspan="5"></td>
	</tr>
</table>
<%
		End If
End Select
%>		
<!-- #Include file="../inc/mod-bot.asp" -->