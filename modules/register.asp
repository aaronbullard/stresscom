<!-- Module 7 --> <!-- register.asp -->
<!-- #Include file="../inc/connect.inc" -->
<%
txtPageHeader = "Register and Purchase Stresscom&trade;"
%>
<!-- #Include file="../inc/mod-top.asp" -->
<%
Select Case Request.QueryString("act")
	Case "ins" 'INSERT
		txtFirstName = Request.Form("fFirstName")
		txtLastName = Request.Form("fLastName")
		txtAddress = Request.Form("fAddress")
		txtAddress2 = Request.Form("fAddress2")
		txtCity = Request.Form("fCity")
		txtState = Request.Form("fState")
		intZip = Request.Form("fZip")
		txtCountry = Request.Form("fCountry")
		txtEmail = Request.Form("fEmail")
		txtPassword = Request.Form("fPassword")
		txtCoupon = Request.Form("fCouponCode")
		intQuantity = Request.Form("fQuantity")
		
		'next make sure coupon code is valid for intErrorCode5
		If NOT txtCoupon = "" Then
			'a coupon was entered
			sqlCoupon = "SELECT * FROM SCUser.tblCoupon WHERE Code = '" & txtCoupon & "'"
			Set rsCoupon = dbCon.Execute(sqlCoupon)
			'response.Write(sqlCoupon)
			 
			If rsCoupon.BOF and rsCoupon.EOF Then
				'coupon code not found
				txtNote = "Sorry, No Discount Code found."
			ElseIf Now() < rscoupon("StartDate") and Now() > rsCoupon("StopDate") Then
				'coupon code found but expired
				txtNote = "Sorry, that Discount Code does not fall within the offer dates."
			ElseIF rsCoupon("Valid") = 0 Then
				'coupon code found but invalid offer	
				txtNote = "Sorry, this Discount Code is no longer offered."
			Else
				'coupon code found, within date, valid offer
				txtNote = "Discount Code accepted, discount noted below."
				txtDiscountCode = txtCoupon
			End If
		Else
			'no coupon code entered, standard pricing
			txtDiscountCode = "none"			
		End If	
		
		
		'next make sure a test quantity exists for intErrorCode 4
		If intQuantity = "" Then
			bitError = True
			intErrorCode = 4
		End If
		
		'next make sure email address is unique for intErrorCode 3
		sqlEmail = "SELECT * FROM SCUser.tblStudent WHERE Email = '" & txtEmail & "'"
		Set rsEmail = dbCon.Execute(sqlEmail)
		
		If rsEmail.BOF AND rsEmail.EOF Then
			'nothing
		Else
			intErrorCode = 3
		End If
		
		'Check to make sure the email address is valid for intErrorCode 2
		
		
		'first make sure all the fields are filled in for intErrorCode 1
		txtMissing = ""
		If txtFirstName = "" Then
			bitError = True
			txtMissing = "First Name <br>"
		End If
		
		If txtLastName = "" Then
			bitError = True	
			txtMissing = txtMissing & "Last Name <br>"
		End If

'removing required fields for address and city 10/3/2012
		
'		If txtAddress = "" Then
'			bitError = True	
'			txtMissing = txtMissing & "Address <br>"
'		End If
'		
'		If txtCity = "" Then
'			bitError = True	
'			txtMissing = txtMissing & "City <br>"
'		End If
'		
		If txtEmail = "" Then
			bitError = True	
			txtMissing = txtMissing & "Email Address <br>"
		End If
		
		If txtPassword = "" Then
			bitError = True	
			txtMissing = txtMissing & "Password <br>"
		End If
		
		'if there are blank fields, show the error
		If Len(txtMissing) > 1 Then
			intErrorCode = 1
		End If
		
		Select Case intErrorCode
			Case 1
%>
<table cellpadding="3" cellspacing="0" bgcolor="#FFFFFF" width="325">
	<tr>
		<td class="TestTitle" align="right">&nbsp;<br>OOPS!!<br>&nbsp;</td>
	</tr>
	<tr>
		<td class="TestStatus">&nbsp;</td>
	</tr>
	<tr>
		<td class="TestStatus">You missed a required field needed for registration.</td>
	</tr>
	<tr>
		<td class="TestStatus">Please hit the Back button and make sure</td>
	</tr>
	<tr>
		<td class="TestStatus">to fill in all required fields.</td>
	</tr>
	<tr>
		<td class="TestStatus"><%= txtMissing %></td>
	</tr>
	<tr>
		<td class="TestHead">&nbsp;</td>
	</tr>
</table>
<%
			Case 3
%>
<table cellpadding="3" cellspacing="0" bgcolor="#FFFFFF" width="325">
	<tr>
		<td class="TestTitle" align="right">&nbsp;<br>OOPS!!<br>&nbsp;</td>
	</tr>
	<tr>
		<td class="TestStatus">&nbsp;</td>
	</tr>
	<tr>
		<td class="TestStatus">The email you entered has already been registered.</td>
	</tr>
	<tr>
		<td class="TestStatus">If you would like to login to your existing account,</td>
	</tr>
	<tr>
		<td class="TestStatus">you can login in by <a href="index.asp?mod=8&act=lnp">clicking here</a>.</td>
	</tr>
	<tr>
		<td class="TestStatus">If you have forgotten your password,</td>
	</tr>
	<tr>
		<td class="TestStatus">you can retrieve it by <a href="index.asp?mod=6">clicking here</a>.</td>
	</tr>
	<tr>
		<td class="TestStatus">If you would like to use a new email address,</td>
	</tr>
	<tr>
		<td class="TestStatus">hit the back/previous button and change it.</td>
	</tr>
	<tr>
		<td class="TestHead">&nbsp;</td>
	</tr>
</table>
<%
			Case 4
%>
<table cellpadding="3" cellspacing="0" bgcolor="#FFFFFF" width="325">
	<tr>
		<td class="TestTitle" align="right">&nbsp;<br>OOPS!!<br>&nbsp;</td>
	</tr>
	<tr>
		<td class="TestStatus">&nbsp;</td>
	</tr>
	<tr>
		<td class="TestStatus">You didn't specify which test you wanted to purchase.</td>
	</tr>
	<tr>
		<td class="TestStatus">Please hit the Back button and make sure</td>
	</tr>
	<tr>
		<td class="TestStatus">to select your purchase choice.</td>
	</tr>
	<tr>
		<td class="TestHead">&nbsp;</td>
	</tr>
</table>
<%
			Case 5
%>
<table cellpadding="3" cellspacing="0" bgcolor="#FFFFFF" width="325">
	<tr>
		<td class="TestTitle" align="right">&nbsp;<br>OOPS!!<br>&nbsp;</td>
	</tr>
	<tr>
		<td class="TestStatus">&nbsp;</td>
	</tr>
	<tr>
		<td class="TestStatus"><%= txtNote %></td>
	</tr>
	<tr>
		<td class="TestStatus">Please hit the Back button and make sure</td>
	</tr>
	<tr>
		<td class="TestStatus">to select you entered the correct Discount Code.</td>
	</tr>
	<tr>
		<td class="TestHead">&nbsp;</td>
	</tr>
</table>
<%
			Case Else
				sqlInsert = "SET NOCOUNT ON; " &_
					"INSERT INTO SCUser.tblStudent (FirstName, LastName, Address, Address2, City, State, Zip, Country, Email, Password) " &_
					"VALUES ('" & txtFirstName & "', '" & txtLastName & "', '" &_
					txtAddress & "', '" & txtAddress2 & "', '" & txtCity & "', '" & txtState & "', '" &_
					intZip & "', '" &  txtCountry & "', '" & txtEmail & "', '" & txtPassword & "'); " &_
					"SELECT @@IDENTITY AS NewID;"
					
				Set rsInsert = dbCon.Execute(sqlInsert)
				Session("SID") = rsInsert("NewID")
				Session("Student") = txtFirstName & " " & txtLastName
				Session("DiscountCode") = txtDiscountCode
				Session("Quantity") = intQuantity
				Response.Redirect("index.asp?mod=9&act=buy&frm=7&dc=" & txtDiscountCode & "&qty=" & intQuantity)
			End Select
			
			'response.Write("txtDiscountCode: " & txtDiscountCode & "<br />txtCoupon: " & txtCoupon & "<br />txtNote: " & txtNote)
	Case Else
%>
			<form action="index.asp?mod=7&act=ins" method="post">
			<table border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td class="bodycopy" colspan="3" align="center">If you already have an account with Stresscom,<br>please <a href="index.asp?mod=8&act=lnp">login</a> here.<hr>					</td>
				</tr>
				<tr>
					<td class="bodycopy" align="right">First Name:</td>
					<td></td>
					<td><input type="text" name="fFirstName"></td>
				</tr>
				<tr>
					<td class="bodycopy" align="right">Last Name:</td>
					<td></td>
					<td><input type="text" name="fLastName"></td>
				</tr>
				<tr>
					<td class="bodycopy" align="right">Address:</td>
					<td></td>
					<td><input type="text" name="fAddress"></td>
				</tr>
				<tr>
					<td class="bodycopy" align="right">Address 2:</td>
					<td></td>
					<td><input type="text" name="fAddress2"></td>
				</tr>
				<tr>
					<td class="bodycopy" align="right">City:</td>
					<td></td>
					<td><input type="text" name="fCity"></td>
				</tr>
				<tr>
					<td class="bodycopy" align="right">State/Province/Region:</td>
					<td></td>
					<td><input type="text" name="fState"></td>
				</tr>
				<tr>
					<td class="bodycopy" align="right">ZIP/Postal Code:</td>
					<td></td>
					<td><input type="text" name="fZip"></td>
				</tr>
				<tr>
					<td class="bodycopy" align="right">Country:</td>
					<td></td>
					<td><input type="text" name="fCountry"></td>
				</tr>
				<tr>
					<td class="bodycopy" colspan="3" align="center"><hr></td>
				</tr>
				<tr>
					<td class="BodyCopy" align="right">Purchase Option:</td>
					<td width="10"></td>
					<td class="BodyCopy"><input type="radio" name="fQuantity" value="4"> $12.00 - Four Stresscom test administrations/interpretation<br>&nbsp;<br>
<!--						<input type="radio" name="fQuantity" value="4"> $19.95 - Stresscom subscription<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(4 administrations plus progress reports for one year)*</td>
				</tr>  -->
				<tr>
					<td class="BodyCopy" colspan="3" align="center"><hr>If you have a Discount Code from your physician,<br>school, or company, please enter it below.<br>&nbsp;<br>Leave Discount Code blank if you do not have one.<hr></td>
				</tr>
				<tr>
					<td class="BodyCopy" align="right">Company, Physician, or<br>School Discount Code:</td>
					<td width="10"></td>
					<td class="BodyCopy"><input type="text" name="fCouponCode"> (optional)</td>
				</tr>
				<tr>
					<td class="BodyCopy" colspan="3" align="center"><hr></td>
				</tr>
				<tr>
					<td class="bodycopy" colspan="3" align="center">
						The email address and password you enter<br>will be your log in when you return to our site<br>for future Stresscom tests and information.<hr>
					</td>
				</tr>
				<tr>
					<td class="bodycopy" align="right">Email address:</td>
					<td width="10"></td>
					<td><input type="text" name="fEmail"></td>
				</tr>
				<tr>
					
              <td class="bodycopy" align="right">Choose a Password:</td>
					<td></td>
					<td><input type="password" name="fPassword"></td>
				</tr>
				<tr>
					<td class="BodyCopy" colspan="3" align="center"><hr></td>
				</tr>
				<tr>
					<td class="BodyCopy" colspan="3" align="center">* The Stresscom year's subscription works best if you take Stresscom once each quarter of the year and apply the nine stress management techniques in between.<hr></td>
				</tr>
				<tr>
					<td colspan="3" align="right"><input type="submit" value="Continue to Purchase Stresscom"></td>
				</tr>
				<tr>
					<td class="BodyCopy" colspan="3" align="center"><hr></td>
				</tr>
			</table>
			</form>
<%
End Select
%>
<!-- #Include file="../inc/mod-bot.asp" -->