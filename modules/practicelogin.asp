<!-- Module 20 -->  <!-- practicelogin.asp -->
<!-- #Include file="../inc/connect.inc" -->
<%
txtPageHeader = txtPageHeader & "Health Care Professionals"
%>
<!-- #Include file="../inc/mod-top.asp" -->
<%
Select Case Request.QueryString("act")
	Case "lin" 'logging in
		strPass = Trim(Request.Form("fGroupPass"))
		
		sqlSelect = "SELECT * FROM SCUser.tblGroup WHERE GroupPassword = '" & strPass & "'"
		Set rsLogin = dbCon.Execute(sqlSelect)
		
		If rsLogin.BOF AND rsLogin.EOF Then
			'Response.Write(sqlSelect)
			Response.Redirect("index.asp?mod=21&act=fla&sta=ngp")
		Else
			Session("GID") = rsLogin("GroupID")
			
			Response.Redirect("index.asp?mod=22")
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
				<td class="BodyCopy" align="center">Please try to <a href="index.asp?mod=21">log in again</a>.</td>
			</tr>
		</table>
<%
	Case Else
%>
			<p style="text-align: justify;"><strong>Stresscom</strong> can be a powerful tool to health care professionals.  In these troublesome times, stress can burden so many people.  Using <strong>Stresscom</strong> in your practice can help you to assist your patients in pin-pointing the areas of stress that weigh on them.</p>
			<p style="text-align: justify;">If you would like to know more about <strong>Stresscom</strong> and how it can help you and your practice, please contact us using our <a href="http://www.stresscom.net/index.asp?mod=2">Contact page</a>.</p>
			<hr width="50%">
			<p style="text-align: center;">If your practice is already taking advantage of <strong>Stresscom</strong>,<br>please login here to review your clients' results.</p>
			<form action="?mod=21&act=lin" method="post">
			<p style="text-align: center;">Administrative Password: <input type="password" name="fGroupPass"></p>
			<p style="text-align: center;"><input type="submit" value="Login"></p>
			</form>
<%
End Select
%>
<!-- #Include file="../inc/mod-bot.asp" -->