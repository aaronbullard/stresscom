<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<HTML>
<HEAD>
	<title>Stress Schedule Assessment</title>
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
<style>
tr, td {border-bottom: solid 1px black;}
</style>
</HEAD>
	
	<body>
	<!-- #Include file="../inc/connect.inc" -->
<%
intStressScheduleID = Request.QueryString("tid")
intStudentID = Request.QueryString("sid")

'get the ssfi informantion
sqlStressSch = "SELECT tblStressSchedule.*, tblStudent.FirstName, tblStudent.LastName, tblPurchase.Quantity " &_
	"FROM SCUser.tblStressSchedule tblStressSchedule JOIN SCUser.tblStudent tblStudent ON tblStressSchedule.StudentID = tblStudent.StudentID " &_
	"INNER JOIN SCUser.tblPurchase tblPurchase ON tblStressSchedule.PurchaseID = tblPurchase.PurchaseID " &_
	"WHERE StressScheduleID = " & intStressScheduleID & " AND tblStressSchedule.StudentID = " & intStudentID
Set rsStressSch = dbCon.Execute(sqlStressSch)

If rsStressSch.BOF AND rsStressSch.EOF Then
	Response.Write("You do not have the rights to view this test score.")
Else
	txtName = rsStressSch("FirstName") & " " & rsStressSch("LastName")
	txtTestDate = rsStressSch("TestDate")
	txtAnswers = rsStressSch("Answers")
	intQuantity = rsStressSch("Quantity")
	intPurchaseID = rsStressSch("PurchaseID")
	Set rsStressSch = Nothing
	
	'Get the test  and show the scores
	sqlAssessment = "SELECT tblStatement.*, tblFactor.* " &_
		"FROM SCUser.tblStatement tblStatement JOIN SCUser.tblFactor tblFactor on tblStatement.Factor = tblFactor.FactorID " &_
		"ORDER BY tblStatement.StatementID"
		
	Set rsAssessment = dbCon.Execute(sqlAssessment)

	'loop through the Assessment and display the answers
%>
<table cellpadding="1" cellspacing="1">
<tr>
	<td colspan="3"><h2>Stresscom Assessment</h3></td>
    <td><%= txtName %><br /><%= txtTestDate %></td>
</tr>
<%
	'loop through the Assessment and display the answers
	Do While Not rsAssessment.EOF
		intAID = rsAssessment("StatementID")
		txtFactor = Left(rsAssessment("Factor"), 4)
		txtStatement = rsAssessment("Statement")
		txtScore = Mid(txtAnswers, intAID, 1)
%>
<tr>
	<td><%= intAID %>. </td>
	<td><%= txtFactor %> - </td>
	<td><%= txtStatement %></td>
	<td align="center"><%= txtScore %></td>
</tr>
<%
		rsAssessment.MoveNext
	Loop
End If
%>
</table>
</body>
</HTML>