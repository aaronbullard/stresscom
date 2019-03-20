<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="./inc/status.inc" -->
<HTML>
<HEAD>
	<title>Stresscom.net</title>
	<META name="robots" content="all">
	<META http-equiv="Pragma" Content="no-cache"> 
	<META name="description" content="SSFI - Smart Grades">
	<meta name="google-site-verification" content="ZPXG3qt-ntvP2hGMceDVLLqV_tT87J7sdEErzOr2YY4" />
	<link rel="stylesheet" type="text/css" href="./css/stresscom.css">

<script type="text/javascript">

  var _gaq = _gaq || [];
  _gaq.push(['_setAccount', 'UA-16681403-3']);
  _gaq.push(['_trackPageview']);

  (function() {
    var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
    ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
  })();

</script>
</HEAD>
	
	<body>
		<div align="center">
<!-- frame table -->
<body>
<div align="center">
<table border="2" bordercolor="#000000" background="images/Sky.gif" cellpadding="0" cellspacing="0" width="750">
	<tr>
		<td>
<!-- frame table -->
			<!--#include file="header.asp" -->
		</td>
	</tr>
	<tr>
		<td>
			<table border="2" bordercolor="#000000" cellpadding="0" cellspacing="0" width="100%">
				<tr>
					<td>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td><img src="./images/spacer.gif" width="25" height="25"></td>
								<td><img src="./images/spacer.gif" width="150" height="1"></td>
								<td><img src="./images/spacer.gif" width="25" height="8"></td>
								<td><img src="./images/spacer.gif" width="525" height="1"></td>
								<td><img src="./images/spacer.gif" width="25" height="1"></td>
							</tr>
							<tr>
								<td><img src="./images/spacer.gif" width="1" height="500"></td>
								<td valign="top">
<% 
'Determine the menus to show
Select Case request.QueryString("mod")
	Case 1,2,3,4,5,6,16,17,21,22,24,NULL
%>
<!--#include file="./navigators/nav-home.asp" -->
<!--#include file="./navigators/nav-tour.asp" -->
<!--#include file="./navigators/nav-login.asp" -->
<%
	Case 7,23
%>
<!--#include file="./navigators/nav-home.asp" -->
<!--#include file="./navigators/nav-tour.asp" -->
<%
	Case 8
		If Request.QueryString("act") = "fla" OR Request.QueryString("act") = "lnp" Then
%>
<!--#include file="./navigators/nav-home.asp" -->
<!--#include file="./navigators/nav-tour.asp" -->
<!--#include file="./navigators/nav-login.asp" -->
<%
		Else
%>
<!--#include file="./navigators/nav-home.asp" -->
<!--#include file="./navigators/nav-parent.asp" -->
<!--#include file="./navigators/nav-logout.asp" -->
<%
		End If
	Case 9,10,11,12,13,15,18,19,20
%>
<!--#include file="./navigators/nav-home.asp" -->
<!--#include file="./navigators/nav-parent.asp" -->
<!--#include file="./navigators/nav-logout.asp" -->
<%
	Case Else
%>
<!--#include file="./navigators/nav-home.asp" -->
<!--#include file="./navigators/nav-tour.asp" -->
<!--#include file="./navigators/nav-login.asp" -->
<%
End Select
%>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
	<tr>
		<td align="center"><button onClick="window.location='index.asp?mod=8&act=lnp'">Returning Users<br>Login</button></td>
	</tr>
</table>
<table border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td><img src="./images/spacer.gif" width="150" height="25"></td>
	</tr>
</table>
<table border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td><img src="./images/calmcool.jpg" width="150" height="120" border="0"></td>
	</tr>
</table>								

								</td>
								<td></td>
								<td valign="top">
<%
'Determin the module to show
Select Case Request.QueryString("mod")
	Case 1,NULL
%>
<!--#include file="./modules/home.asp" -->
<%
	Case 2
%>
<!--#include file="./modules/contact.asp" -->
<%
	Case 3
%>
<!--#include file="./modules/factors.asp" -->
<%
	Case 4
%>
<!--#include file="./modules/norm.asp" -->
<%
	Case 5
%>
<!--#include file="./modules/about.asp" -->
<%
	Case 6
%>
<!--#include file="./modules/forgot.asp" -->
<%
	Case 7
%>
<!--#include file="./modules/register.asp" -->
<%
	Case 8
%>
<!--#include file="./modules/loggedin.asp" -->
<%
	Case 9
%>
<!--#include file="./modules/purchase.asp" -->
<%
	Case 10
%>
<!--#include file="./modules/student.asp" -->
<%
	Case 11
%>
<!--#include file="./modules/review.asp" -->
<%
	Case 12
%>
<!--#include file="./modules/score.asp" -->
<%
	Case 13
%>
<!--#include file="./modules/taketest.asp" -->
<%
	Case 14
%>
<!--#include file="./modules/loggedout.asp" -->
<%
	Case 15
%>
<!--#include file="./modules/teststart.asp" -->
<%
	Case 16
%>
<!--#include file="./modules/samples.asp" -->
<%
	Case 17
%>
<!--#include file="./modules/using.asp" -->
<%
	Case 18
%>
<!--#include file="./modules/testhowto.asp" -->
<%
	Case 19
%>
<!--#include file="./modules/score-emotions.asp" -->
<%
	Case 20
%>
<!--#include file="./modules/score-technique.asp" -->
<%
	Case 21
%>
<!--#include file="./modules/practicelogin.asp" -->
<%
	Case 22
%>
<!--#include file="./modules/practiceview.asp" -->
<%
	Case 23
%>
<!--#include file="./modules/practicescore.asp" -->
<%
	Case 24
%>
<!--#include file="./modules/medlanding.asp" -->
<%
	Case Else
%>
<!--#include file="./modules/home.asp" -->
<%
End Select
%>
								</td>
								<td></td>
							</tr>
							<tr>
								<td><img src="./images/spacer.gif" width="25" height="25"></td>
								<td class="footer" align="center" colspan="3">Copyright &copy; 2008 Ombudsman Press, Inc.</td>
								<td><img src="./images/spacer.gif" width="25" height="25"></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</div>

<script type="text/javascript">
var gaJsHost = (("https:" == document.location.protocol) ? "https://ssl." : "http://www.");
document.write(unescape("%3Cscript src='" + gaJsHost + "google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E"));
</script>
<script type="text/javascript">
try {
var pageTracker = _gat._getTracker("UA-15261931-2");
pageTracker._trackPageview();
} catch(err) {}</script>

	</body>
</html>