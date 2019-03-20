<!-- Module 8 -->  <!-- loggedin.asp -->
<!-- #Include file="../inc/connect.inc" -->
<%
If NOT IsEmpty(Session("Student")) Then 
	txtPageHeader = (Session("Student") & ", ")
End If 

txtPageHeader = txtPageHeader & "Thank you for using Stresscom"
%>
<!-- #Include file="../inc/mod-top.asp" -->
<%
Select Case Request.QueryString("act")
	Case "lin" 'logging in
		strLogin = Trim(Request.Form("fLogin"))
		strPass = Trim(Request.Form("fPass"))
		
		sqlSelect = "SELECT * FROM SCUser.tblStudent tblStudent WHERE Email = '" & strLogin & "' AND Password = '" & strPass & "'"
		Set rsLogin = dbCon.Execute(sqlSelect)
		
		If rsLogin.BOF AND rsLogin.EOF Then
			'Response.Write(sqlSelect)
			Response.Redirect("index.asp?mod=8&act=fla&sta=fla")
		Else
			Session("SID") = rsLogin("StudentID")
			Session("Student") = rsLogin("FirstName") & " " & rsLogin("LastName")
			
			Response.Redirect("index.asp?mod=8")
		End If
	Case "fla" 'failed login attempt
%>
		<table width="100%">
<%
If NOT Request.QueryString("sta") = "" Then
%>
			<tr>
				<td class="status" align="center"><%= strStatus %><br>&nbsp;</td>
			</tr>
<%
End If
%>		
			<tr>
				<td class="BodyCopy" align="center">If you have <a href="index.asp?mod=6">forgotten your password</a>, we can email it to you.</td>
			</tr>
		</table>
<%
	Case "lnp"
%>
<table border="0" cellpadding="0" cellspacing="0">
<form action="index.asp?mod=8&act=lin" method="post">
	<tr>
		<td class="Bodycopy">Log in to your existing account with your...<hr>Email Address:<br><input type="text" name="fLogin" width="117"></td>
	</tr>
	<tr>
		<td class="Bodycopy">Password:<br><input type="password" name="fPass" width="117"></td>
	</tr>
	<tr>
		<td class="Bodycopy" align="right"><hr><input type="submit" value="Login to Stresscom"></td>
	</tr>
	<tr>
		<td class="Bodycopy"><hr>Did you <a href="index.asp?mod=6">forget</a> your password?</td>
	</tr>
</form>
</table>	
<%
	Case Else
%>
			<p style="text-align: center;">Thank you for using <strong>Stresscom</strong>.</p>
			<p style="text-align: center;">You are on your way to better understanding and managing your stress with <strong>Stresscom</strong>.</p>
<hr width="50%">
			<p style="text-align: center;">To review your purchase history,<br>and/or purchase another administration of <strong>Stresscom</strong>,<br><a href="index.asp?mod=9" class="body">click here</a>.</p>
<hr width="50%">
			<p style="text-align: center;">If you would like to review your past <strong>Stresscom</strong> scores<br>or take the next text in your subscription,<br><a href="index.asp?mod=13" class="body">click here</a>.</p>
<%
End Select
%>
<!-- #Include file="../inc/mod-bot.asp" -->