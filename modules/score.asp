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
		
		'get the ssfi informantion
		sqlStressSch = "SELECT tblStressSchedule.*, tblStudent.FirstName, tblStudent.LastName, tblPurchase.Quantity, tblPurchase.CouponID " &_
			"FROM SCUser.tblStressSchedule tblStressSchedule JOIN SCUser.tblStudent tblStudent ON tblStressSchedule.StudentID = tblStudent.StudentID " &_
			"INNER JOIN SCUser.tblPurchase tblPurchase ON tblStressSchedule.PurchaseID = tblPurchase.PurchaseID " &_
			"WHERE StressScheduleID = " & intStressScheduleID & " AND tblStressSchedule.StudentID = " & Session("SID")
		Set rsStressSch = dbCon.Execute(sqlStressSch)
		
		If rsStressSch.BOF AND rsStressSch.EOF Then
			Response.Write("You do not have the rights to view this test score.")
		Else
			txtName = rsStressSch("FirstName") & " " & rsStressSch("LastName")
			txtTestDate = rsStressSch("TestDate")
			txtAnswers = rsStressSch("Answers")
			intQuantity = rsStressSch("Quantity")
			intPurchaseID = rsStressSch("PurchaseID")
			txtCouponID = rsStressSch("CouponID")
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
<%
'show medical offer button if score is for freemedtest couponID
'If LCase(txtCouponID) = "freetestmed" Then
'	Response.Write("<a href=""index.asp?mod=24""><img src=""images/medoffer.jpg"" align=""right"" border=""0""></a>")
'End If
%>
<p>Thank you for taking the <strong>Stresscom</strong>. You have taken the first educational step in a process that can add years to your life, and make your days more productive.</p>
<p>Stress overload, or what we call exstress, increases the risk of heart disease, and lowers our body's natural immunity to other diseases. Stress also results in more absenteeism from work, lowers our productivity, sparks greater conflict with family and co-workers, and can even cause more accidents on the job.</p>
<p>Fortunately, <u>stress can be managed</u>. In the following pages, we'll talk about your stress levels and how you may start to manage them to your advantage.</p>
<p><strong>YOUR FIVE STRESS COMPARISONS</strong></p>
<p><strong>Comparison #1<br>Review your scores on the six Sources of Stress and compare them to definitions of the results that follow:</strong></p>
<ol type="A">
	<li><strong>Control/Responsibility</strong><br>If you do not have (or assert) an amount of control over a situation which is adequate for the amount of responsibility you have, or perceive you have, then stress will develop. For example, if your employer tells you it's your responsibility to produce a first-class sales brochure, but doesn't give you control over the budget or content of that brochure, you will probably experience stress.</li>
	<li><strong>Competition</strong><br>The need to compete is common in our society. We learn it in the classroom, and carry it into our work sites. It's natural. But no one can win all the time. That's why a person who strives to win at any cost may develop exstress.</li>
	<li><strong>Task Orientation With Perfection</strong><br>If you have a strong goal or task orientation and you need to complete the task in a perfect or near perfect manner, you may experience stress. No one is perfect all the time.</li>
	<li><strong>Change</strong><br>We live in a time when change is part of our daily lives. However, when life changes are occurring faster than we can adjust to them, exstress is generated. Some changes may force us to see ourselves in different ways and we need to be flexible enough to change with the flow of events. For example, in retirement, a person not only loses a job, he or she also loses the identity gained from the job. A person needs to be flexible to adapt to new identities, to avoid additional stress.</li>
	<li><strong>Symptomatology</strong><br>Many highly-stressed persons experience frequent sleeplessness, irritability, and depression. These symptoms are warning signals to check your level of stress.  Also, if you are suffereing from an illness, your concern, worry, and pain may, in itself, create more stress.</li>
	<li><strong>Time</strong><br>When a person has too many actual or perceived tasks to perform withing the time available, the person usually experiences heightened stress.</li>
</ol>
<!-- 
******************************
page break here !!!!!!
****************************** 
-->
<p style="page-break-before: always"><strong>Comparison #2<br>Look at your overall score on the Stress Thermometer.</strong></p>
<%
			If cInt(intQuantity) = 1 Then
				txtThermoFile = "thermo-" & cStr(intTotalScore) & ".png"
				Response.Write("<div align=""center"">Your score is <strong>" & intTotalScore & "<br>&nbsp;<br></strong><img src=""/images/thermo/201208/thermo-mean_bars.png""><img src=""/images/thermo/201208/" & txtThermoFile & """></div>")
			Else
				'Response.Write("<div align=""center""><img src=""/images/thermo/therm_" & intTotalScore & ".gif""></div>")
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
				Response.Write("<strong>&nbsp;<br>&nbsp;<hr>&nbsp;<br>&nbsp;<br>&nbsp;<br><img src=""/images/thermo/201208/thermo-mean_bars.png"" width=""50"" height=""325""></td>" & vbCrLf)
				
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
						
						txtThermoFile = "thermo-" & cStr(intTotalScore) & ".png"
						
						'display the scores here for each subscription score
						Response.Write("<td class=""bodycopy"" align=""center"">")
						Response.Write("<strong>Test Date</strong><br>" & txtSubTestDate & "<hr>")
						Response.Write("Your score is<br><strong>" & intTotalScore & "<br>&nbsp;<br><img src=""/images/thermo/201208/" & txtThermoFile & """ width=""105"" height=""325""></td>" & vbCrLf)
					Else
						
						intInterval = (intSubCount - 1) * 3
						txtSubTestDate = DateAdd("M", intInterval, datFirstTest)
						
						'display the scores here for each subscription score
						Response.Write("<td class=""bodycopy"" align=""center"" valign=""top"">")
						Response.Write("<strong>Take Test</strong><br>" & txtSubTestDate & "<hr>")
						Response.Write("&nbsp;<br><img src=""/images/spacer.gif"" width=""105"" height=""325""></td>" & vbCrLf)
					End If
					
					rsSub.MoveNext
				Loop
				Response.Write("</tr>" & vbCrLf & "</table>" & vbCrLf & "</div>" & vbCrLf)
			End If
%>
<hr>
<p><br>Along the right-hand side of the thermometer are the stress scores. Your stress "temperature" rises to a level that lines up with your stress score. Like reading your temperature, the higher the score, the more highly stressed you seem to be.</p>
<p>Along the left-hand side of the thermometer, also lined up with your stress "temperature," is the percent of people who have a lower stress score than you have. For example, a score of 185 lines up with 25%. That means 25 persons in 100 have a stress score of 185 or lower. This score of 185 defines the Exstress line in the next comparison.</p>
<p><strong>Comparison #3<br>Next, notice that the stress thermometer is divided into three categories – Exstress, Average, and Low Stress.</strong></p>
<p>Your Stresscom score places you in one of these categories. If you are in the average range, your stress results are comparable with the majority of people who have taken the test. If you are close to the Exstress line, you should consider taking steps to manage your stress. More on this later.</p>
<p>Even if your score is in the normal range one or more of the six areas of stress may be high (discussed under comparison #5), so remember we live in a stress-filled society. That means that even though your stress is average, it still can contribute to stress-related problems. Stress is something which may build over the years and its effect can be devastating.</p>
<!-- 
******************************
page break here !!!!!!
****************************** 
-->
<p style="page-break-before: always"><strong>Comparison #4<br>You may wish to compare your score with your peers or your work group.</strong></p>
<p>If your total score is 15 score points higher or lower than the average score of your peers, you can conclude your stress level is significantly higher or lower.</p>
<table width="500" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td>
			<table border="0" cellpadding="0" cellspacing="0">
				<tr class="bodycopy">
					<td></td>
					<td></td>
					<td>Average<br>Total Score</td>
				</tr>
				<tr class="bodycopy">
					<td><strong>1.</strong></td>
					<td><strong>Business</strong></td>
					<td></td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>Managers, Misc.</td>
					<td>171</td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>Managers, Auto, Japanese</td>
					<td>169</td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>Federal Bank Managers</td>
					<td>166</td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>Radio Marketing</td>
					<td>171</td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>Film Industry Employees</td>
					<td>172</td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>Retail Store Customers</td>
					<td>172</td>
				</tr>	
				<tr class="bodycopy">
					<td colspan="3">&nbsp;</td>
				</tr>
				<tr class="bodycopy">
					<td><strong>2.</strong></td>
					<td><strong>College/University</strong></td>
					<td></td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>University Graduates</td>
					<td>172</td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>Grad. Students – Counseling</td>
					<td>162</td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>University Students</td>
					<td>172</td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>University Faculty</td>
					<td>168</td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>Adult Ed. Participants</td>
					<td>169</td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>Student Affairs Professionals</td>
					<td>174</td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>Admission/Records Personnel</td>
					<td>171</td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>Comm. College Faculty</td>
					<td>168</td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>Foreign Students</td>
					<td>183</td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>College Health Fair Attendees</td>
					<td>174</td>
				</tr>	
				<tr class="bodycopy">
					<td colspan="3">&nbsp;</td>
				</tr>
				<tr class="bodycopy">
					<td><strong>3.</strong></td>
					<td><strong>Educational Professionals</strong></td>
					<td></td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>Teachers</td>
					<td>168</td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>School Counselors</td>
					<td>169</td>
				</tr>	
				<tr class="bodycopy">
					<td></td>
					<td>School Administrators</td>
					<td>172</td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>Special Educational Teachers</td>
					<td>172</td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>School Staff</td>
					<td>168</td>
				</tr>	
				<tr class="bodycopy">
					<td colspan="3">&nbsp;</td>
				</tr>
				<tr class="bodycopy">
					<td><strong>4.</strong></td>
					<td><strong>Technical</strong></td>
					<td></td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>Computer Science Managers</td>
					<td>175</td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>Computer Programmers</td>
					<td>168</td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>Computer Managers, Banking</td>
					<td>170</td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>Computer Design<br>/Manufacturing</td>
					<td>175</td>
				</tr>	
				<tr class="bodycopy">
					<td colspan="3">&nbsp;</td>
				</tr>
			</table>
		</td>
		<td>&nbsp;&nbsp;</td>
		<td valign="top">
			<table border="0" cellpadding="0" cellspacing="0">
				<tr class="bodycopy">
					<td></td>
					<td></td>
					<td>Average<br>Total Score</td>
				</tr>
				<tr class="bodycopy">
					<td><strong>5.</strong></td>
					<td><strong>Health Professionals</strong></td>
					<td></td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>Health Career Educators</td>
					<td>171</td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>Dentists</td>
					<td>175</td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>Endodontist</td>
					<td>175</td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>Dental Staff</td>
					<td>175</td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>Medical Professionals – Spouses</td>
					<td>183</td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>Hospital Managers</td>
					<td>183</td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>Hospital Risk Managers</td>
					<td>181</td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>Public Health, City</td>
					<td>184</td>
				</tr>	
				<tr class="bodycopy">
					<td colspan="3">&nbsp;</td>
				</tr>
				<tr class="bodycopy">
					<td><strong>6.</strong></td>
					<td><strong>City Workers</strong></td>
					<td></td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>Public Works</td>
					<td>178</td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>Clerical</td>
					<td>180</td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>Water and Power</td>
					<td>180</td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>Library</td>
					<td>178</td>
				</tr>	
				<tr class="bodycopy">
					<td colspan="3">&nbsp;</td>
				</tr>
				<tr class="bodycopy">
					<td><strong>7.</strong></td>
					<td><strong>Fire Fighters</strong></td>
					<td></td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>Fire Fighters, City</td>
					<td>177</td>
				</tr>	
				<tr class="bodycopy">
					<td colspan="3">&nbsp;</td>
				</tr>
				<tr class="bodycopy">
					<td><strong>8.</strong></td>
					<td><strong>Police</strong></td>
					<td></td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>Police, City</td>
					<td>178</td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>Police Academy Trainees</td>
					<td>184</td>
				</tr>	
				<tr class="bodycopy">
					<td colspan="3">&nbsp;</td>
				</tr>	
				<tr class="bodycopy">
					<td><strong>9.</strong></td>
					<td><strong>Miscellaneous</strong></td>
					<td></td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>Missionaries</td>
					<td>181</td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>U.S. Olympic Managers</td>
					<td>185</td>
				</tr>
				<tr class="bodycopy">
					<td></td>
					<td>Senior Citizens</td>
					<td>179</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<!-- 
******************************
page break here !!!!!!
****************************** 
-->
<p style="page-break-before: always"><strong>Comparison #5<br>Next, you can pinpoint your individual areas of stress.</strong></p>
<p>The following is a graph that shows your stress as it relates to issues of Control, Competition, Task Perfection, Change, Symptomatology, and Time Management. Your score in each area will be discussed as well as ways to reduce stress in each individual area.  Your high scores are represented in red.</p>
<hr>
<!-- #Include file="../inc/graph.asp" -->
<hr>
<p><strong>STRESS CHECKLIST</strong></p>
<p>Your stress profile tells you which areas in your life are producing stress. Each area is discussed below and techniques to manage stress are suggested.</p>
<p><strong>—YOUR CONTROL/RESPONSIBILITY STRESS IS:&nbsp;&nbsp;<%= UCase(txtAFactorScore) %></strong></p>
<p>If your score is high, this indicates that you may feel too many responsibilities that are not matched by a sense of control. For example, a mother who feels responsible for a child when the child lives in another state has inadequate control to match her feeling of responsibility.</p>
<p>A person may also feel a loss of control when he or she accepts too many life and work assignments.</p>
<p>You can lessen your stress in this area in two ways. First, you can reduce your responsibilities. Say "no" to being the leader of a work group or commmittee again. Tell yourself you don't always need to be the first at work in the morning. Check yourself if you find yourself trying to be responsible for what is someone elses task or area. Secondly, you can increase your control by setting realistic goals and assignments with equally realistic time and resource allocations.</p>
<p>We suggest the following for persons with high stress in this area:</p>
<ul>
	<li>If you can't say "no," take an assertiveness training course, or read a book on the topic.</li>
	<li>Set realistic goals at home and at work.</li>
	<li>Recognize that "The reed never tells the wind what to do."</li>
</ul>
<p><strong>Checklist for Responsibility/Control</strong></p>
<ul>
	<li>If you feel you have an "emergency within you," don't ignore it.</li>
	<li>Make realistic goals.</li>
	<li>If you can't say "no," learn how.</li>
	<li>Close the gaps between responsibilities and control.</li>
	<li>Quit thinking stress will pass— people often created much of their own stress.</li>
	<li>Listen to what your body is telling you.</li>
	<li>Don't wait until mid-life to manage your stress.</li>
	<li>Recognize that being fast is rarely better than taking time to do it right.  Slow yourself down.</li>
	<li>Practice positive rehearsals (see Stress Management Techniques below).</li>
	<li>It's okay to relax.  Acknowledge that you are still a good person and do not need to feel guilty, even when relaxed.</li>
</ul>
<p><strong>—YOUR COMPETITION STRESS IS:&nbsp;&nbsp;<%= UCase(txtBFactorScore) %></strong></p>
<p>If your score is high, you may want to re-think some of your attitudes about winning.</p>
<p>To compete is an American value, but none of us wins all the time. If you have an extra-competitive drive, you may discover this adds to your stress, especially if your goals and expectations greatly outpace your accomplishments.</p>
<p>Competition can be either external, with others, or internal with ourselves. If you have unrealistic goals for yourself, e.g., to become president of the company by age 35, to be a straight "A" student when your academic abilities are average, or to be the best mother on the block AND the best employee, you are creating an internal competition, that may be unrealistic.</p>
<p>To help reduce stress due to excessive competitiveness, we suggest the following:</p>
<ul>
	<li>Work at feeling good about yourself just because you're uniquely you. You don't have to be a super person in order to be acceptable to others and to yourself.</li>
	<li>Set goals that are more realistic given your abilities and resources.</li>
	<li>If you are competitive, try to use your extra competitive spirit in organized physical activities. This may help reduce stress.</li>
</ul>
<p><strong>Checklist for Competition</strong></p>
<ul>
	<li>Recognize that winning is not everything.</li>
	<li>Determin whether your enemy may be competing with your own high expectations.</li>
	<li>Stop and smell a flower or admire your natural surroundings.</li>
	<li>Understand that being a good person doesn't require competition or winning.</li>
	<li>Leave your competitiveness at work.</li>
	<li>Monitor your blood pressure.</li>
	<li>Recognize that being tough and fighting with everyone isn't winning.</li>
	<li>Don't allow yourself to be driven by "acute work advancement cravings."</li>
	<li>Allow your sports games to be merely fun.</li>
</ul>
<p><strong>—YOUR TASK ORIENTATION WITH PERFECTION STRESS IS:&nbsp;&nbsp;<%= UCase(txtCFactorScore) %></strong></p>
<p>Often exstress is developed when you seek perfection in a somewhat imperfect environment. This is particularly true when you need to perform perfect tasks at a faster and faster rate. To counter this type of stress, we suggest:</p>
<ul>
	<li>Realize that not all tasks require perfection. Demand perfection when the task truly requires it.</li>
	<li>Use time management techniques.</li>
	<li>Ask yourself if being perfect is important. As a rule, contributing your personal best is all that's needed.</li>
	<li>Learn to think of life in degrees (shades of gray) rather than in terms of perfection (black and white).</li>
	<li>Find a way to reduce your "anger in" (unexpressed anger).</li>
</ul>
<p><strong>Checklist for Perfection</strong></p>
<ul>
	<li>Be perfect only when required.</li>
	<li>Recongnize that there is a difference between being perfect and your personal best.</li>
	<li>Appreciate the gray areas of life.</li>
	<li>Acknowledge that sometimes being late is therapeutic.</li>
	<li>Laugh at your "anger in."</li>
	<li>Understand that a rapidly changing world has little room for the perfect.</li>
	<li>Recongnize that lashing out at others will not make them perfect.</li>
	<li>Do some deep breathing exercises when your "perfection" is frustrated.</li>
	<li>Admint that patience is really a virtue.</li>
	<li>Learn anew— your personality is not stuck with your desire for perfectibility.</li>
</ul>
<p><strong>—YOUR STRESS DUE TO LIFE CHANGES IS:&nbsp;&nbsp;<%= UCase(txtDFactorScore) %></strong></p>
<p>If your life is changing faster than you can adjust to it, slow it down. You don't need to do or be everything at once. For example, if you are going through a divorce, don't change jobs and move to another state all at once.</p>
<p>If you find your marketplace is changing rapidly, plan to increase your skills or reorganize slowly and deliberately. We suggest:</p>
<ul>
	<li>Think ahead; anticipate.</li>
	<li>Plan three years in advance.</li>
	<li>Look at how you are organizing yourself and your tasks.</li>
	<li>Budget the changes in your life.</li>
</ul>
<p><strong>Checklist for Change</strong></p>
<ul>
	<li>Don't resist change—contain it.</li>
	<li>Plan your future changes when possible.</li>
	<li>Spread out your major life events or transitions.</li>
	<li>Understand that change can be exciting and not just scary.</li>
	<li>Budget your major life events.</li>
	<li>Use "positive self talk" with your life changes.</li>
	<li>Practice "being in the moment."</li>
	<li>Understand that changing everything in your life at once can be disastrous.</li>
	<li>Recognize the differences between small hassles and major changes.</li>
	<li>Be aware that marital conflict often occurs during major life changes.</li>
</ul>
<p><strong>—YOUR STRESS SYMPTOMS ARE:&nbsp;&nbsp;<%= UCase(txtEFactorScore) %></strong></p>
<p>As you probably realized, the symptoms outlined in the Stresscom are those related to high stress. Such stress may lead to a series of health problems. Some of you have high stress because you have serious illness that you worry about.  High stress may depress your immune system, thus allowing health risks to increase. If your score on this area is high or you experience stress symptoms, we suggest:</p>
<ul>
	<li>Consult your physician.</li>
	<li>Develop a balanced diet.</li>
	<li>Incorporate a physical fitness program into your life. Remember to check with your doctor first.</li>
	<li>Get the proper sleep.</li>
	<li>Reduce or eliminate alcohol and drug consumption.</li>
	<li>Consider developing a lifestyle change toward improved wellness.</li>
</ul>
<p><strong>Checklist for Symptoms</strong></p>
<ul>
	<li>Develop a lifestyle of wellness.</li>
	<li>Enhance your own pharmacy—your immune system.</li>
	<li>Get your doctor to answer all of your important questions—write them down before your appointment.</li>
	<li>Discuss your health concerns with family members.</li>
	<li>Have a regular physical fitness program.</li>
	<li>Cut down on coffee or sources of caffeine.</li>
	<li>Recognize that stress and depression are related.</li>
	<li>Inform yourself about your symptoms over the internet.</li>
	<li>Balance food intake and metabolic burn.</li>
	<li>Don't get addicted to the adrenalin rush.</li>
	<li>Practice being in the moment.</li>
</ul>
<p><strong>—YOUR TIME MANAGEMENT STRESS IS:&nbsp;&nbsp;<%= UCase(txtFFactorScore) %></strong></p>
<p>You may feel you are unimportant, or unproductive when you are not "pressed for time," or on a "tight schedule." When you do not have enough time to do the job, your stress may be magnified. If your stress in this area is high, we suggest the following:</p>
<ul>
	<li>Employ time management techniques.</li>
	<li>Examine your time schedule and calendar. Put more time between appointments.</li>
	<li>Plan ahead.</li>
	<li>Try not to procrastinate.</li>
	<li>Recognize that "time is money," but you'd need to save a great deal of time to make up for the medical expense and care that can result from excess stress leading to hypertension.</li>
	<li>Watch that you don't become subject to the "hurry sickness" (discussed in the next section).</li>
</ul>
<p><strong>Checklist for Time</strong></p>
<ul>
	<li>Add time to your personal saving account.</li>
	<li>Understand that there is enough time in each day, if well planned.</li>
	<li>Don't procrastinate.</li>
	<li>Realize that kids are not more important than you.</li>
	<li>Budget your stress throughout the week.</li>
	<li>Don't feed your "hurry sickness."</li>
	<li>Avoid over-scheduling to make yourself seem important.</li>
	<li>Admit that not having enough time IS AN EXCUSE.</li>
	<li>Add "free time" to your daily schedule.</li>
	<li>Re-plan the 168 hours you are given each week.</li>
	<li>Use stress budgeting.</li>
</ul>
<hr width="50%">
<hr>
<hr width="50%">
<p><button onClick="window.location='?mod=19&tid=<%= intStressScheduleID %>'">Continue to Emotions of Stress</button></p>
					</td>
				</tr>
			</table>
<%
		End If
End Select
%>
<!-- #Include file="../inc/mod-bot.asp" -->