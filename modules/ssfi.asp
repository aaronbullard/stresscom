<!-- #Include file="../inc/connect.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<%
' get the page number of the ssfi by the charactor number of the results field in tblStressSchedule
sqlPage = "SELECT tblStressSchedule.*, tblStudent.FirstName, tblStudent.LastName " &_
	"FROM SCUser.tblStressSchedule tblStressSchedule JOIN SCUser.tblStudent tblStudent ON tblStressSchedule.StudentID = tblStudent.StudentID " &_
	"WHERE tblStressSchedule.StressScheduleID = " & Request.QueryString("tid")
Set rsPage = dbCon.Execute(sqlPage)

intAnswer = Len(rsPage("Answers"))
txtAnswer = rsPage("Answers")
txtName = rsPage("FirstName") & " " & rsPage("LastName")

txtDate = cStr(Date)

If IsNull(intAnswer) Then
	intAnswer = 0
End If

Select Case intAnswer
	Case 0 'page one
		ImageOne = "<img src=""../images/1.jpg"">"
		ImageTwo = "<img src=""../images/2.jpg"">"
		ImageThree = "<img src=""../images/3.jpg"">"
		ImageFour = "<img src=""../images/4.jpg"">"
		txtPageTitle = "StressCom.net - Stresscom Test Page One"
		intStart = 1
		txtNext = "Continue to the Stresscom Test Page 2 of 3"
	Case 20 'page two
		ImageOne = "<img src=""../images/1ch.jpg"">"
		ImageTwo = "<img src=""../images/2.jpg"">"
		ImageThree = "<img src=""../images/3.jpg"">"
		ImageFour = "<img src=""../images/4.jpg"">"
		txtPageTitle = "StressCom.net - Stresscom Test Page Two"
		intStart = 21
		txtNext = "Continue to the Stresscom Test Page 3 of 3"		
	Case 40 'page three
		ImageOne = "<img src=""../images/1ch.jpg"">"
		ImageTwo = "<img src=""../images/2ch.jpg"">"
		ImageThree = "<img src=""../images/3.jpg"">"
		ImageFour = "<img src=""../images/4.jpg"">"
		txtPageTitle = "StressCom.net - Stresscom Test Page Three"
		intStart = 41
		txtNext = "Review your test results"
	Case 60 'test has been completed
		'go to score page
		Response.Redirect("../index.asp?mod=12&act=scr&tid=" & Request.QueryString("tid"))
End Select

'get the statements for the appropriate page
sqlStatement = "SELECT * FROM SCUser.tblStatement " &_
	"WHERE StatementID >= " & intStart & " AND StatementID < " & (intStart + 20)

'Response.Write(sqlStatement)
'Response.Write(intStart)

Set rsStatement = dbCon.Execute(sqlStatement)

%>
<html>
<Head>
	<title><%= txtPageTitle %></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link href="../css/stresscom.css" rel="stylesheet" class="css/txt">
</Head>
<body>
<div align="center">

<table cellpadding="3" cellspacing="0" bgcolor="#E5EBDF" width="675">
	<tr>
		<td class="TestNumber"><img src="../images/spacer.gif" width="30" height="60"></td>
		<td colspan="6" class="TestName" align="right">Name:&nbsp;<strong><%= txtName %></strong><br>Date:&nbsp;<strong><%= txtDate %></strong></td>
		<td class="TestRtMargin"><img src="../images/spacer.gif" width="45" height="1"></td>
	</tr>
	<tr>
		<td class="TestNumber">&nbsp;</td>
		<td class="ContractHead" colspan="6"><strong>S T R E S S C O M</strong></td>
		<td class="TestRtMargin"><img src="../images/spacer.gif" width="1" height="25">&nbsp;</td>
	</tr>
	<tr>
		<td class="TestNumber">&nbsp;</td>
<!-- 	<td class="TestRight" colspan="6"><img src="../images/progress.jpg"><img src="../images/arrow.jpg"><%= ImageOne %><img src="../images/arrow.jpg"><%= ImageTwo %><img src="../images/arrow.jpg"><%= ImageThree %><img src="../images/arrow.jpg"><%= ImageFour %><img src="../images/arrow.jpg"><img src="../images/score.jpg"></td> -->
		<td class="TestRight" colspan="6">&nbsp;</td>
		<td class="TestRtMargin"><img src="../images/spacer.gif" width="1" height="25">&nbsp;</td>
	</tr>
	<tr>
		<td class="TestNumber">&nbsp;</td>
		<td class="TestStatement">&nbsp;</td>
		<td class="TestHead" width="60">Strongly<br>Disagree</td>
		<td class="TestHead" width="60">Disagree</td>
		<td class="TestHead" width="60">Undecided</td>
		<td class="TestHead" width="60">Agree</td>
		<td class="TestHead" width="60">Strongly<br>Agree</td>
		<td class="TestRtMargin"><img src="../images/spacer.gif" width="1" height="25">&nbsp;</td>
	</tr>
<form name="SSFI" action="ssfi2.asp" method="post">
<input type="hidden" name="fTID" value="<%= Request.QueryString("tid") %>">
<input type="hidden" name="fStart" value="<%= intStart %>">
<input type="hidden" name="fAnswers" value="<%= txtAnswer %>">
<%
Do While Not rsStatement.EOF
	intStatement = rsStatement("StatementID")
	txtStatement = rsStatement("Statement")
%>				
	<tr>
		<td class="TestNumber"><%= intStatement %>.</td>
		<td class="TestStatement"><%= txtStatement %></td>
		<td class="TestHead"><input class="TestRadio" type="radio" value="1" name="st<%= intStatement %>"></td>
		<td class="TestHead"><input class="TestRadio" type="radio" value="2" name="st<%= intStatement %>"></td>
		<td class="TestHead"><input class="TestRadio" type="radio" value="3" name="st<%= intStatement %>"></td>
		<td class="TestHead"><input class="TestRadio" type="radio" value="4" name="st<%= intStatement %>"></td>
		<td class="TestHead"><input class="TestRadio" type="radio" value="5" name="st<%= intStatement %>"></td>
		<td class="TestRtMargin"><img src="../images/spacer.gif" width="1" height="25">&nbsp;</td>
	</tr>
<%
	rsStatement.MoveNext
Loop
%>
	<tr>
		<td class="TestNumber">&nbsp;</td>
		<td class="TestStatement">&nbsp;</td>
		<td class="TestHead">&nbsp;</td>
		<td class="TestHead">&nbsp;</td>
		<td class="TestHead">&nbsp;</td>
		<td class="TestHead">&nbsp;</td>
		<td class="TestHead">&nbsp;</td>
		<td class="TestRtMargin"><img src="../images/spacer.gif" width="1" height="25">&nbsp;</td>
	</tr>
	<tr>
		<td class="TestNumber">&nbsp;</td>
		<td class="TestStatement">&nbsp;</td>
		<td class="TestRight" colspan="5"><input type="submit" value="<%= txtNext %>"</td>
		<td class="TestRtMargin"><img src="../images/spacer.gif" width="1" height="25">&nbsp;</td>
	</tr>
</form>
</table>
</div>
</body>
</html>