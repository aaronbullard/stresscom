<!-- #Include file="../inc/connect.inc" -->
<%
txtPageHeader = "Retrieve your forgotten password"
%>
<!-- #Include file="../inc/mod-top.asp" -->
<%
Select Case Request.QueryString("act")
	Case "get" 'retrieve the password
		sqlGet = "SELECT * FROM SCUser.tblStudent WHERE Email = '" & Request.Form("fEmail") & "'"
		Set rsGet = dbCon.Execute(sqlGet)
		
		If rsGet.BOF and rsGet.EOF Then
			Response.Redirect("?mod=6&sta=noe")
		Else
			txtFirstName = rsGet("FirstName")		
			txtEmail = rsGet("Email")
			txtPassword = rsGet("Password")
			
			txtMessage = txtFirstName & "," & vbCRLF & vbCRLF & "Accout: " & txtEmail & vbCRLF & "Password: " & txtPassword & vbCRLF & vbCRLF & "Thank You," & vbCRLF & vbCRLF & "StressCom Support"
					
			' Set the mail server configuration
			Set objConfig = CreateObject("CDO.Configuration")
			objConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
			objConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.stresscom.net"
			objConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10
			objConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
			objConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
			objConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "scadmin@stresscom.net"
			objConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "n0rcal"
			
			objConfig.Fields.Update		
			
			' Create and send the mail
			Set objMail = CreateObject("CDO.Message")
			
			' Use the config object created above
			Set objMail.Configuration = objConfig
			objMail.Sender = "StressCom Support"
			objMail.From = "StressCom Support <support@stresscom.com>"
			objMail.ReplyTo = "support@stresscom.com"
			objMail.To = txtEmail
			objMail.BCC = "scadmin@stresscom.net"
			objMail.Subject = "StressCom Password"
			objMail.TextBody = txtMessage
			
			objMail.Send
			set objMail = Nothing
			
			Set rsGet = Nothing
			
			Response.Redirect("index.asp?mod=6&sta=ret")
		End If
	Case Else
%>
			
			<form action="index.asp?mod=6&act=get" method="post">
				<table border="0" cellpadding="0" cellspacing="0">
<%
If NOT Request.QueryString("sta") = "" Then
%>
					<tr>
						<td class="status" colspan="3" align="center"><%= strStatus %><br>&nbsp;</td>
					</tr>
<%
End If
%>
				<tr>
					<td class="BodyCopy" colspan="3"><p align="center">Enter in the Email Address that you registered with.<br>Your password will be emailed to the same address.</p></td>
				</tr>
				<tr>
					<td class="bodycopy" colspan="3" align="center"><hr></td>
				</tr>
				<tr>
					<td class="bodycopy" align="right">Email Address:</td>
					<td width="10"></td>
					<td><input type="text" name="fEmail" size="50"></td>
				</tr>
				<tr>
					<td class="BodyCopy" colspan="3" align="center">&nbsp;<br><input type="submit" value="Retrieve Password"></td>
				</tr>
				</table>
			</form>
			
<%
End Select
%>
<!-- #Include file="../inc/mod-bot.asp" -->