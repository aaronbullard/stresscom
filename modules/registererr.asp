<!-- #Include file="../inc/connect.inc" -->
<html>
<Head>
<title>SmartGrades.net - SSFI Page One</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/ssfi-gold.css" rel="stylesheet" class="css/txt">
</Head>
<body>
<%
intRow = Request.Form("fStart")
'response.Write(intRow & "<br>")

'get and loop thru all the values of the form
txtAnswer = ""
For r=intRow to intRow + 19
	txtAnswer = txtAnswer + Request.Form("st" & cStr(r))
Next
'Response.Write(txtAnswer & "<br>")

If Len(txtAnswer) = 20 Then
	'update tblSSFI with the answers
	txtAnswers = cStr(Request.Form("fAnswers")) + cStr(txtAnswer)
	'Response.Write(txtAnswers & "<br>")
	
	sqlUpdate = "UPDATE SCUser.tblSSFI SET Answers = '" & cStr(txtAnswers) & "', TestDate = '" & cStr(Date) & "' " &_
		"WHERE SSFIID = " & Request.Form("fTID")
	'Response.Write(sqlUpdate & "<br>")
	
	Set rsUpdate = dbCon.Execute(sqlUpdate)
	
	Response.Redirect("ssfi.asp?tid=" & Request.Form("fTID"))
	'Response.Write("ssfi.asp?tid=" & Request.Form("fTID") & "<br>")
end if
%>
<div align="center">

<table cellpadding="3" cellspacing="0" bgcolor="#FFFFFF" width="325">
	<tr>
		<td class="TestTitle" align="right">&nbsp;<br>OOPS!!<br>&nbsp;</td>
	</tr>
	<tr>
		<td class="TestStatus">&nbsp;</td>
	</tr>
	<tr>
		<td class="TestStatus">You skipped over a statement on the previous page.</td>
	</tr>
	<tr>
		<td class="TestStatus">Please hit the Back button and make sure</td>
	</tr>
	<tr>
		<td class="TestStatus">to evaluate all the statements.</td>
	</tr>
	<tr>
		<td class="TestStatus">&nbsp;</td>
	</tr>
	<tr>
		<td class="TestHead">&nbsp;</td>
	</tr>
</table>
</div>
</body>
</head>