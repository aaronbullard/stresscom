<%
Select Case Request.QueryString("sta")
	Case "snt" ' Message sent
		strStatus = "Thank you, your message has been sent"
	Case "fla" ' Failed Login Attemp
		strStatus = "Your Email Address and Password did not match"
	Case "noe" ' No email address found
		strStatus = "We are sorry, the email address you provide does not match an email on file."
	Case "ret" ' Password was retrieved
		strStatus = "Your password was emailed to you."
	Case "ngp" ' Group Password not found
		strStatus = "Your Administrative Password was not found."
End Select
%>