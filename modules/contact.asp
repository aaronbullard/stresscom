<!-- Module 2 -->
<%
txtPageHeader = "Please feel free to contact us with questions or comments"
%>
<!-- #Include file="../inc/mod-top.asp" -->
<%
Select Case Request.QueryString("act")
	Case "snd"
		txtEmail = Request.Form("fEmail")
		txtName = Request.Form("fName")
		txtSubject = Request.Form("fSubject")
		txtMessage = Request.Form("fMessage")
				
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
		
		response.Write(txtEmail)
		response.Write(txtName)
		response.Write(txtSubject)
		response.Write(txtMessage)
		
		' Use the config object created above
		Set objMail.Configuration = objConfig
		objMail.Sender = txtName
		objMail.From = "scadmin@stresscom.net"
		objMail.ReplyTo = txtEmail
		objMail.To = "DrEHallberg33@aol.com, steffany@ombpress.com"
		objMail.Subject = txtSubject
		objMail.TextBody = txtMessage
		
		objMail.Send
		set objMail = Nothing
		
		Response.Redirect("index.asp?mod=2&sta=snt")
	Case Else
%>
			<form action="index.asp?mod=2&act=snd" method="post" >
				<table>
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
						<td class="BodyCopy">
							<p>Your Email Address:<br><input type="text" name="fEmail" size="50"></p>
							<p>Your Name:<br><input type="text" name="fName" size="50"></p>
							<p>Message Subject:<br><input type="text" name="fSubject" size="50"></p>
							<p>Your Message:<br><textarea cols="74" rows="10" name="fMessage"></textarea></p>
							<p><input type="submit" value="Send Message"></p>
						</td>
					</tr>
				</table>
			</form>
<%
End Select
%>
<!-- #Include file="../inc/mod-bot.asp" -->