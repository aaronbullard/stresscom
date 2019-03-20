<!-- Module 10 -->
<!-- #Include file="../inc/connect.inc" -->
<%
txtPageHeader = "Register your student to use the SSFI"
%>
<!-- #Include file="../inc/mod-top.asp" -->
<%
Select Case Request.QueryString("act")
	Case "nst" 'registering a student, no assignment to ssfi
		txtName = Request.Form("fName")
		intPid = Session("pid")
		intBirthYear = Request.Form("fBirthYear")
		txtSchool = Request.Form("fSchool")
		
		sqlInsert = "INSERT INTO SCUser.tblStudent (Name, ParentID, BirthYear, School) " &_
			"VALUES ('" & txtName & "', " & intPid & ", " & intBirthYear & ", '" & txtSchool & "')"
		
		'Response.Write(Request.Form("cmdSubmit") & "<br>")
		Response.Write(sqlInsert)
		Set rsInsert = dbCon.Execute(sqlInsert)
		Set rsInsert = Nothing
		
		Response.Redirect("index.asp?mod=10")
		
	Case "frg" 'come here from parent registration or to register a new student
%>
			<table>
			<form action="index.asp?mod=10&act=nst" method="post">
				<tr>
					<td class="BodyCopy" colspan="5">Register a new student...</td>
				</tr>
				<tr>
					<td class="BodyCopy" colspan="5"><hr></td>
				</tr>
				<tr>
					<td></td>
					<td class="BodyCopy" align="right">Student's Name:</td>
					<td width="10"></td>
					<td><input type="text" name="fName"></td>
				</tr>
				<tr>
					<td></td>
					<td class="BodyCopy" align="right">Studenet's Birth Year:</td>
					<td width="10"></td>
					<td><select name="fBirthYear">
							<option>1994</option>
							<option>1993</option>
							<option>1992</option>
							<option>1991</option>
							<option>1990</option>
							<option>1989</option>
							<option>1988</option>
							<option>1987</option>
							<option>1986</option>
							<option>1985</option>
							<option>1984</option>
							<option></option>
						</select></td>
					<td></td>
				</tr>
				<tr>
					<td></td>
					<td class="BodyCopy" align="right">Student's School:</td>
					<td width="10"></td>
					<td><input type="text" name="fSchool"></td>
				</tr>
				<tr>
					<td class="BodyCopy" colspan="5"><hr></td>
				</tr>
				<tr>
					<td class="BodyCopy" colspan="5" align="center"><input type="submit" name="cmdSubmit" value="Register Your Student"></td>
				</tr>
			</form>
			</table>
<%
	Case Else
%>
			<table border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td class="BodyCopy" align="center">Registering your students will help you<br>to follow their progress to <strong>SMART GRADES™</strong>.</td>
				</tr>
				<tr>
					<td class="BodyCopy"><hr>
<%
		'check to see if parent has registered any students
		sqlStudent = "SELECT * FROM SCUser.tblStudent WHERE ParentID = " & Session("pid") & " ORDER BY Name"
		Set rsStudent = dbcon.Execute(sqlStudent)
		
		If Not rsStudent.BOF and Not rsStudent.EOF Then
			Response.Write("You have already registered...<br>")
			'show registered students
			Response.Write("<p>")
			Do While Not rsStudent.EOF
				response.Write("&nbsp;&nbsp;&nbsp;&nbsp;" & rsStudent("Name") & "<br>")
				rsStudent.MoveNext
			Loop
			Response.Write("</p>")
		Else
			Response.Write("You don't have any registerd Students yet.")
		End If
%>
					<hr></td>
				</tr>
			</table>

<%		
		Response.Write("<p><button OnClick=""location='index.asp?mod=10&act=frg'"">Register a New Student</button>")
		Response.Write("<p>--- OR ---</p>")
		Response.Write("<p><button OnClick=""location='index.asp?mod=9'"">Continue to Purchase the SSFI</button>")
		
		
End Select
%>
<!-- #Include file="../inc/mod-bot.asp" -->