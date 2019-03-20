<!-- Module 23 -->  <!-- addingAIScodes.asp -->
<!-- #Include file="./inc/connect.inc" -->
<%
txtPageHeader = "Adding new codes, one time only"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<HTML>
	<HEAD>
		<title>Stresscom.net</title>
		<META name="robots" content="all">
		<META http-equiv="Pragma" Content="no-cache"> 
		<META name="description" content="SSFI - Smart Grades">
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
			<!-- header table -->
			<table border="2" bordercolor="#000000" cellpadding="0" cellspacing="0" width="750">
				<tr>
					<td><img src="./images/stresscom_header.jpg" width="750" height="160" border="0"></td>
				</tr>
			</table>
<!-- header table -->
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
								<td><img src="./images/spacer.gif" width="525" height="1"></td>
								<td><img src="./images/spacer.gif" width="25" height="1"></td>
							</tr>
							<tr>
								<td></td>
								<td valign="top">
<table bordercolor="#000000" cellSpacing="0" cellPadding="0" border="4" bgcolor="#FFFFFF" width="100%">
	<tr>
		<td>
		
<table border="0" cellSpacing="0" cellPadding="0" width="100%" bgcolor="E5E5E5">
	<tr>
		<td><img src="./images/spacer.gif" width="5" height="10"></td>
		<td><img src="./images/spacer.gif" width="100%" height="10"></td>
		<td><img src="./images/spacer.gif" width="5" height="10"></td>
	</tr>
	<tr>
		<td></td>
		<td class="BodyHead" align="center"><%= txtPageHeader %><hr></td>
		<td></td>
	</tr>
	<tr>
		<td><img src="./images/spacer.gif" width="5" height="8"></td>
		<td class="bodycopy"><img src="./images/spacer.gif" width="503" height="8"></td>
		<td><img src="./images/spacer.gif" width="5" height="8"></td>
	</tr>
	<tr>
		<td></td>
		<td class="bodycopy" valign="top">
			<div align="center">
			<table width="100%">
				<tr>
					<td class="bodycopy" valign="top">

<%
		Response.Write("Nothing will be added as I removed all code that would make any changes to the database!")
		
		intStart = 1
		intEnd = 1
		txtPrefix = "Hargrove"
		
		i = intStart		
		For  i =intStart To intEnd
			'set the leading zeros
			Select Case Len(CStr(i))
				Case 1
					txtCouponNumber = "0" + CStr(i)
				Case 2
					txtCouponNumber = "" + CStr(i)
				Case 3
					txtCouponNumber = CStr(i)
			End Select
			
			'create the new code
			txtCouponCode = txtPrefix + txtCouponNumber
			
			'add the new code to the database
			'sqlAddCode = "Insert Into SCUser.tblcoupon (GroupID, Code, AmountSgl, AmountSub, Valid, QuanLimit) " &_
			'	"VALUES (9, '" & txtCouponCode & "', 0, 0, 1, 1)"
	
			'Set rsAddCode = dbCon.Execute(sqlAddCode)
			
			'Set rsAddCode = Nothing
			
			Response.Write(txtCouponCode & vbCrLf)
		Next
					


%>
<hr>
<hr width="50%">
					</td>
				</tr>
			</table>

		</td>
		
    <td><img src="./images/spacer.gif" width="5" height="5"></td>
	</tr>
	<tr>
		<td><img src="./images/spacer.gif" width="5" height="5"></td>
		<td><img src="./images/spacer.gif" width="500" height="5"></td>
		<td><img src="./images/spacer.gif" width="5" height="5"></td>
	</tr>
	<tr>
		<td colspan="3"><img src="./images/spacer.gif" width="5" height="5"></td>
	</tr>
</table>
		
		</td>
	</tr>
</table>
								</td>
								<td></td>
							</tr>
							<tr>
								<td><img src="./images/spacer.gif" width="25" height="25"></td>
								<td class="footer" align="center">Copyright &copy; 2008 Ombudsman Press, Inc.</td>
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
	</body>
</html>