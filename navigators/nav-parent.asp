<%
'only run if not attempting to login
If Request.QueryString("mod") = 8 AND Request.QueryString("act") = "lin" Then
	Response.Write("logging in")
Else
	'make sure you are logged in, if not, abandon all session veriables and redirect to logged out page
	If IsEmpty(Session("sid")) Then
		Response.Redirect("?mod=14&act=out")	
	End If
End If

txtNavTitle = "Account Menu"
%>
<!-- begin home nav table -->
<!--#include file="..\inc\nav-top.asp" -->
	<tr>
		<td class="MenuDataItem"><img src="./images/spacer.gif"></td>
		<td class="MenuDataItem">&#187; <a href="index.asp?mod=8" class="menu">Using your Stresscom Account</a><br><img src="../images/spacer.gif" width="1" height="5"></td>
		<td class="MenuDataItem"><img src="./images/spacer.gif"></td>
	</tr>
	<tr>
		<td class="MenuDataItem"><img src="./images/spacer.gif"></td>
		<td class="MenuDataItem">&#187; <a href="index.asp?mod=9" class="menu">Purchase Stresscom</a><br><img src="../images/spacer.gif" width="1" height="5"></td>
		<td class="MenuDataItem"><img src="./images/spacer.gif"></td>
	</tr>
	<tr>
		<td class="MenuDataItem"><img src="./images/spacer.gif"></td>
		<td class="MenuDataItem">&#187; <a href="index.asp?mod=13" class="menu">Your Stresscom Tests</a></td>
		<td class="MenuDataItem"><img src="./images/spacer.gif"></td>
	</tr>
<!-- 	<tr>
		<td class="MenuDataItem"><img src="./images/spacer.gif"></td>
		<td class="MenuDataItem">&#187; <a href="index.asp?mod=11" class="menu">Review your Stresscom Results</a></td>
		<td class="MenuDataItem"><img src="./images/spacer.gif"></td>
	</tr>
-->
<!--#include file="..\inc\nav-bot.asp" -->
<!-- end home nav table -->