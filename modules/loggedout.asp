<!-- Module 14 -->  <!-- loggedout.asp -->
<%
txtPageHeader = "Logged Out"
%>
<!-- #Include file="../inc/mod-top.asp" -->
<%

	Session.Abandon
	'response.Write("should have cleared sessions" & vbCRLF)
	'Response.Write("Parent " & session("Parent") & vbCRLF & "PID " & session("pid"))
	Response.Write("You are no longer logged in, please <a href=""index.asp?mod=8&act=lnp"" class=""body"">login</a> to take advantage of Stresscom")
%>
			
<!-- #Include file="../inc/mod-bot.asp" -->