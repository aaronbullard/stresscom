<!-- Module 13 -->  <!-- taketest.asp -->
<!-- #Include file="../inc/connect.inc" -->
<%
txtPageHeader = "Take Stresscom "
%>
<!-- #Include file="../inc/mod-top.asp" -->
<%
Select Case Request.QueryString("act")
'***************
	Case "gtt" 'assign student to Stress Schedule intStressScheduleID = Request.QueryString("tid")
'***************
		intStressScheduleID = Request.QueryString("tid")
		Response.Redirect("index.asp?mod=15&tid=" & intStressScheduleID)
'***************
	Case "fpr" 'from purchase
'***************
		'get Stress Schedules where student has not taken it yet
		sqlTestStd = "SELECT * FROM SCUser.tblStressSchedule WHERE StudentID = " & Session("SID") & " AND Answers IS NULL ORDER BY StressScheduleID"
		Response.Write(sqlteststd)
		Set rsTestStd = dbcon.Execute(sqlTestStd)
		If rsTestStd.BOF and rsTestStd.EOF Then
			Response.Redirect("index.asp?mod=13")
		Else
			intStressScheduleID = rsTestStd("StressScheduleID")
			Response.Redirect("index.asp?mod=13&act=gtt&tid=" & intStressScheduleID)
		End If
'***************
	Case Else
'***************
		intStudentID = Session("SID")

		'check to see if there is a purchased Stress Schedule 
		sqlPurchase = "SELECT * FROM SCUser.tblPurchase WHERE StudentID = " & intStudentID & " ORDER BY PurchaseID"
		Set rsPurchase = dbCon.Execute(sqlPurchase)
		
		
			
		
		If rsPurchase.BOF and rsPurchase.EOF Then
			Response.Write("<p>You need to purchare Stresscom before you can take it.</p><p>Click on the link in to the left to purchase Stresscom .<p>")
		Else
			Do While Not rsPurchase.EOF
				'Response.Write("<hr>")
				sqlTest = "SELECT * FROM SCUser.tblStressSchedule WHERE PurchaseID = " & rsPurchase("PurchaseID") & " ORDER BY StressScheduleID"
				Set rsTest = dbCon.Execute(sqlTest)
				
				If rsPurchase("Quantity") = 1 Then 'purchase was of a single administration
					'response.Write(sqltest & "<br>")
					Response.Write("<table width=""400""><tr><td class=""bodycopy"" colspan=""3"">")
					Response.Write("<strong>Single Administration</strong> - Purchased on " & rsPurchase("Date") & "<hr>")
					Response.Write("</td></tr>")
					Response.Write("<tr><td class=""bodycopy"" width=""25""></td><td class=""bodycopy"" width=""275""><strong>Test #" & intCount & "</strong> - ")

					If IsNull(rsTest("Answers")) Then
						Response.Write("Test not taken</td><td class=""bodycopy""><button onClick=""window.location='index.asp?mod=15&tid=" & rsTest("StressScheduleID") & "'"">Take Stresscom Now</button></td>")
					ElseIf Len(rsTest("Answers")) < 60 Then
						Response.Write("Test not finished</td><td class=""bodycopy""><button onClick=""window.location='index.asp?mod=15&tid=" & rsTest("StressScheduleID") & "'"">Finish Stresscom</button></td>")
					Else
						Response.Write("Stresscom has been completed</td><td class=""bodycopy""><button onClick=""window.location='index.asp?mod=12&tid=" & rsTest("StressScheduleID") & "'"">Review Scores</button></td>")
					End If
					Response.Write("</tr></table>")
				Else	'the purchase was a subscription
					Response.Write("<table><tr><td class=""bodycopy"" colspan=""3"">")
					Response.Write("<strong>Subscription</strong> - Purchased on " & rsPurchase("Date") & "<hr>")
					Response.Write("</td></tr>")
					datPurchase = rsPurchase("Date")
					intCount = 0
					txtToday = Date()
					Do While Not rsTest.EOF
						intCount = intCount + 1
						
						Select Case intCount
							Case 1
								If IsNull(rsTest("TestDate")) Then
									datFirstTest = Date()
								ElseIf Len(rsTest("Answers")) < 60 Then
									datFirstTest = Date()
								Else
									datFirstTest = rsTest("TestDate")
								End If
								txtTestDate = datFirstTest
							Case Else
								If IsNull(rsTest("TestDate")) Then
									intInterval = (intCount - 1) * 3
									txtTestDate = DateAdd("M", intInterval, datFirstTest)
								Else
									txtTestDate = rsTest("TestDate")
								End If
						End Select						
						
						Response.Write("<tr><td class=""bodycopy"" width=""25""></td><td class=""bodycopy"" width=""275""><strong>Test #" & intCount & "</strong> - ")

						If IsNull(rsTest("Answers")) Then
							'Response.Write("-- " & DateDiff("d", txtToday, txtTestDate) & " --<br>")
							If DateDiff("d", txtToday, txtTestDate) > 0 Then
								Response.Write("Test not taken</td><td class=""bodycopy"">Take Stresscom after " & txtTestDate)
							Else
								Response.Write("Test not taken</td><td class=""bodycopy""><button onClick=""window.location='index.asp?mod=15&tid=" & rsTest("StressScheduleID") & "'"">Take Stresscom " & txtTestDate & "</button></td>")
							End if
						ElseIf Len(rsTest("Answers")) < 60 Then
							Response.Write("Test not finished</td><td class=""bodycopy""><button onClick=""window.location='index.asp?mod=15&tid=" & rsTest("StressScheduleID") & "'"">Finish Stresscom " & txtTestDate & "</button></td>")
						Else
 							Response.Write("Stresscom&nbsp;completed&nbsp;" & cStr(rsTest("TestDate")) & "</td><td class=""bodycopy""><button onClick=""window.location='index.asp?mod=12&tid=" & rsTest("StressScheduleID") & "'"">Review Scores</button></td>")
						End If
						Response.Write("</tr>")
						rsTest.MoveNext
					Loop
					Response.Write("</table>")
				End If
				rsPurchase.MoveNext
			Response.Write("<hr>")
			Loop
%>
<%
				
		End If
End Select
%>			
<!-- #Include file="../inc/mod-bot.asp" -->