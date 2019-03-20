<!-- module 15 -->
<!-- #Include file="../inc/connect.inc" -->
<%
txtPageHeader = "Start Stresscom"
%>
<!-- #Include file="../inc/mod-top.asp" -->
<%
Select Case Request.QueryString("act")
	Case "tgo" 'update tblStressSchedule and start test
		'do this temporarily.....
		sqlTest = "SELECT tblStressSchedule.StressScheduleID, tblStudent.* " &_
			"FROM SCUser.tblStressSchedule tblStressSchedule JOIN SCUser.tblStudent tblStudent ON tblStressSchedule.StudentID = tblStudent.StudentID " &_
			"WHERE tblStressSchedule.StressScheduleID = " & Request.QueryString("tid")
		Set rsTest = dbCon.Execute(sqlTest)
		txtName = rsTest("FirstName") & " " & rsTest("LastName")
		intStressScheduleID = Request.QueryString("tid")
		
		'update tblStressSchedule
		'sqlUpdate = "UPDATE SCUser.tblStressSchedule SET Grade = " & intGrade & ", GPA = " & txtGPA & " " &_
		'	"WHERE StressScheduleID = " & intStressScheduleID
			
		'Set rsUpdate = dbcon.Execute(sqlUpdate)
		'Set rsUpdate = Nothing
		
		Response.Redirect("index.asp?mod=18&tid=" & intStressScheduleID)
		
		
		'this is the original code
		intStressScheduleID = Request.Form("fStressScheduleID")
		School = Request.Form("fSchool")
		Grade = Request.Form("fGrade")
		GPA = Request.Form("fGPA")
		
		'update tblStressSchedule
		sqlUpdate = "UPDATE SCUser.tblStressSchedule SET Grade = " & Grade & ", GPA = " & GPA & " " &_
			"WHERE StressScheduleID = " & intStressScheduleID
			
		Set rsUpdate = dbcon.Execute(sqlUpdate)
		Set rsUpdate = Nothing
		
		'update tblStudent
		sqlStUpdate = "UPDATE SCUser.tblStudent SET School = '" & School & "', Grade = " & Grade &_
			", GPA = " & GPA & " WHERE StudentID = " & Session("SID")
			
		Set rsStUpdate = dbcon.Execute(sqlStUpdate)
		Set rsStUpdate = Nothing
		
		'Response.Redirect("modules/ssfi.asp?tid=" & StressScheduleID)
		Response.Redirect("index.asp?mod=18&tid=" & intStressScheduleID)
		
	Case Else
		sqlTest = "SELECT tblStressSchedule.StressScheduleID, tblStudent.* " &_
			"FROM SCUser.tblStressSchedule tblStressSchedule JOIN SCUser.tblStudent tblStudent ON tblStressSchedule.StudentID = tblStudent.StudentID " &_
			"WHERE tblStressSchedule.StressScheduleID = " & Request.QueryString("tid")
		Response.Write(sqlTest)
		Set rsTest = dbCon.Execute(sqlTest)
		txtName = rsTest("FirstName") & " " & rsTest("LastName")
		intStressScheduleID = Request.QueryString("tid")
		
		Response.Redirect("index.asp?mod=15&act=tgo&tid=" & intStressScheduleID)
		
%>
			<form action="index.asp?mod=15&act=tgo" method="post">
				<input type="hidden" name="fStressScheduleID" value="<%= Request.QueryString("tid") %>">
				<table border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td class="BodyCopy" colspan="3"><%= txtName %>, please verify this information...</td>
				</tr>
				<tr>
					<td class="bodycopy" colspan="3" align="center"><hr></td>
				</tr>
				<tr>
					<td class="bodycopy" align="right">Name of your school:</td>
					<td width="10"></td>
					<td><input type="text" name="fSchool" size="50" value="<%= txtSchool %>"></td>
				</tr>
				<tr>
					<td class="bodycopy" align="right">Your Grade:</td>
					<td width="10"></td>
					<td><select name="fGrade">
							<option></option>
							<option value="9" <% If intGrade = 9 Then Response.Write("selected") End If %>>Freshman</option>
							<option value="10" <% If intGrade = 10 Then Response.Write("selected") End If %>>Sophmore</option>
							<option value="11" <% If intGrade = 11 Then Response.Write("selected") End If %>>Junior</option>
							<option value="12" <% If intGrade = 12 Then Response.Write("selected") End If %>>Senior</option>
						</select></td>
				</tr>
				<tr>
					<td class="bodycopy" align="right">Your current GPA:</td>
					<td width="10"></td>
					<td><input type="text" name="fGPA" size="12" value="<%= txtGPA %>"></td>
				</tr>
				<tr>
					<td class="BodyCopy" colspan="3" align="right">&nbsp;<br><input type="submit" value="Start the SSFI ->"></td>
				</tr>
				</table>
			</form>
			
<%
End Select
%>
<!-- #Include file="../inc/mod-bot.asp" -->