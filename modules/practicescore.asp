<!-- Module 12 -->  <!-- score.asp -->
<!-- #Include file="../inc/connect.inc" -->
<%
txtPageHeader = "Your Stresscom Results"
%>
<!-- #Include file="../inc/mod-top.asp" -->
<%
Select Case Request.QueryString("act")
	Case Else
		intStressScheduleID = Request.QueryString("tid")
		intStudentID = Request.QueryString("sid")
		
		'get the ssfi informantion
		sqlStressSch = "SELECT tblStressSchedule.*, tblStudent.FirstName, tblStudent.LastName, tblPurchase.Quantity " &_
			"FROM SCUser.tblStressSchedule tblStressSchedule JOIN SCUser.tblStudent tblStudent ON tblStressSchedule.StudentID = tblStudent.StudentID " &_
			"INNER JOIN SCUser.tblPurchase tblPurchase ON tblStressSchedule.PurchaseID = tblPurchase.PurchaseID " &_
			"WHERE StressScheduleID = " & intStressScheduleID & " AND tblStressSchedule.StudentID = " & intStudentID
		Set rsStressSch = dbCon.Execute(sqlStressSch)
		
		If rsStressSch.BOF AND rsStressSch.EOF Then
			Response.Write("You do not have the rights to view this test score.")
		Else
			txtName = rsStressSch("FirstName") & " " & rsStressSch("LastName")
			txtTestDate = rsStressSch("TestDate")
			txtAnswers = rsStressSch("Answers")
			intQuantity = rsStressSch("Quantity")
			intPurchaseID = rsStressSch("PurchaseID")
			Set rsStressSch = Nothing
			'calc the score
			intTotalScore = 0
			For s = 1 to 60
				intScore = Mid(txtAnswers, s, 1)
				intTotalScore = intTotalScore + Int(intScore)
			Next
			'make sure intTotalScore fits in the thermometer
			If intTotalScore < 130 Then
				txtTotalScore = "low"
			ElseIf intTotalScore > 230 Then
				txtTotalScore = "high"
			Else
				txtTotalScore = intTotalScore
			End If
			

%>
			<table>
				<tr>
					<td class="bodycopy" valign="top">
<p><span class="bodyhead">Stresscom Results for <%= txtName %></span><br><span class="bodycopy">Test date: <%= txtTestDate %></span></p>
<p><strong>Your overall score on the Stress Thermometer</strong></p>
<hr>
<%
			If cInt(intQuantity) = 1 Then
				txtThermoFile = "thermo-" & cStr(intTotalScore) & ".jpg"
				Response.Write("<div align=""center"">Your score is <strong>" & intTotalScore & "<br>&nbsp;<br></strong><img src=""/images/thermo/thermo-bars.jpg""><img src=""/images/thermo/" & txtThermoFile & """></div>")
			Else
				'Response.Write("<div align=""center""><img src=""/images/thermo/thermo-" & intTotalScore & ".gif""></div>")
				sqlSub = "SELECT * FROM SCUser.tblStressSchedule WHERE PurchaseID = " & cInt(intPurchaseID) & " ORDER BY StressScheduleID"
				Set rsSub = dbCon.Execute(sqlSub)
				
				'start table
				Response.Write("<div align=""center"">" & vbCrLf & "<table border=""0"" width=""470"">" & vbCrLf & "<tr>" & vbCrLf)
				Response.Write("<td class=""bodycopy"" colspan=""5"" align=""center""><strong>Your Subscription Scores</strong><hr></td>")
				Response.Write("</tr>" & vbCrLf & "<tr>" & vbCrLf)
				Response.Write("<td class=""bodycopy"" width=""50""></td>" & vbCrLf)
				Response.Write("<td class=""bodycopy"" width=""105""></td>" & vbCrLf)
				Response.Write("<td class=""bodycopy"" width=""105""></td>" & vbCrLf)
				Response.Write("<td class=""bodycopy"" width=""105""></td>" & vbCrLf)
				Response.Write("<td class=""bodycopy"" width=""105""></td>" & vbCrLf)
				Response.Write("</tr>" & vbCrLf & "<tr>" & vbCrLf)
				Response.Write("<td class=""bodycopy"" align=""center"">")
				Response.Write("<strong>&nbsp;<br>&nbsp;<hr>&nbsp;<br>&nbsp;<br>&nbsp;<br><img src=""/images/thermo/thermo-bars.jpg"" width=""50"" height=""320""></td>" & vbCrLf)
				
				'start table body with thermo graphics
				intSubCount = 0
				Do While NOT rsSub.EOF
					intSubCount = intSubCount + 1
					txtSubSSID = rsSub("StressScheduleID")
					txtSubTestDate = rsSub("TestDate")
					txtSubAnswers = rsSub("Answers")
					'get the first test date then set dates for the remaining administrations
					If intSubCount = 1 Then
						datFirstTest = rsSub("TestDate")
					End If
					
					If Len(txtSubAnswers) = 60 Then
						'calc the score
						intTotalScore = 0
						For s = 1 to 60
							intScore = Mid(txtSubAnswers, s, 1)
							intTotalScore = intTotalScore + Int(intScore)
						Next
						
						txtThermoFile = "thermo-" & cStr(intTotalScore) & ".jpg"
						
						'display the scores here for each subscription score
						Response.Write("<td class=""bodycopy"" align=""center"">")
						Response.Write("<strong>Test Date</strong><br>" & txtSubTestDate & "<hr>")
						Response.Write("Your score is<br><strong>" & intTotalScore & "<br>&nbsp;<br><img src=""/images/thermo/" & txtThermoFile & """ width=""105"" height=""320""></td>" & vbCrLf)
					Else
						
						intInterval = (intSubCount - 1) * 3
						txtSubTestDate = DateAdd("M", intInterval, datFirstTest)
						
						'display the scores here for each subscription score
						Response.Write("<td class=""bodycopy"" align=""center"" valign=""top"">")
						Response.Write("<strong>Take Test</strong><br>" & txtSubTestDate & "<hr>")
						Response.Write("&nbsp;<br><img src=""/images/spacer.gif"" width=""105"" height=""320""></td>" & vbCrLf)
					End If
					
					rsSub.MoveNext
				Loop
				Response.Write("</tr>" & vbCrLf & "</table>" & vbCrLf & "</div>" & vbCrLf)
			End If
%>
<hr>
<p style="text-align: right">Continued on the next page...</p>
<hr>
<p style="page-break-before: always"><strong>Your individual areas of stress.</strong></p>
<hr>
<!-- #Include file="../inc/graph.asp" -->
<hr width="50%">
<hr>
<hr width="50%">
					</td>
				</tr>
                <tr>
                	<td class="bodycopy" align="center">Review the answers to each statement <a href="/modules/ssanswers.asp?tid=<%= intStressScheduleID %>&sid=<%= intStudentID %>" target="_new">here</a>.</td>
                </tr>
			</table>
<%
		End If
End Select
%>
<!-- #Include file="../inc/mod-bot.asp" -->