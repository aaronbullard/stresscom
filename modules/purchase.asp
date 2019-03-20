<!-- Module 9 -->  <!-- purchase.asp -->
<!-- #Include file="../inc/connect.inc" -->
<%
txtDC = Request.QueryString("dc")
If txtDC = "freetestmed" Then
	txtPageHeader = "Take the Stresscom Assessment"
Else
	txtPageHeader = "Purchase Stresscom"
End If
%>
<!-- #Include file="../inc/mod-top.asp" -->
<% 
Select Case Request.QueryString("act")
'***************
	Case "buy"
'***************
		
		'function to remove spaces from txtCoupon
		Function remSpace(strInp)
			If strInp <> "" Then
				remSpace = Replace(strInp, " ", "")
			End If
		End Function
		
		If Request.QueryString("frm") = 7 Then
			intQuantity = Request.QueryString("qty")
			txtCoupon = remSpace(Request.QueryString("dc"))
			If txtCoupon = "none" Then
				txtCoupon = ""
			End If
		Else
			intQuantity = Request.Form("fQuantity")
			txtCoupon = remSpace(Trim(UCase(Request.Form("fCouponCode"))))
		End If
		
		If cInt(intQuantity) = 1 Then
			txtItemName = "Stresscom.net Single Test Administration"
			txtFieldName = "AmountSgl"
			txtAmount = 12.00
		Else
			txtItemName = "Stresscom.net Subscription (4 Administrations)"
			txtFieldName = "AmountSub"
			txtAmount = 12.00
		End If
		
		If NOT txtCoupon = "" Then
			'a coupon was entered
			sqlCoupon = "SELECT * FROM SCUser.tblCoupon WHERE Code = '" & txtCoupon & "'"
			Set rsCoupon = dbCon.Execute(sqlCoupon)
			'response.Write(sqlCoupon)
			 
			If rsCoupon.BOF and rsCoupon.EOF Then
				'coupon code not found
				txtItemName = txtItemName & " (Invalid Discount Code - " & txtCoupon & ")"
				txtNote = "Sorry, the Discount Code '" & txtCoupon & "' could not be found."
				txtCoupon = ""
			ElseIf Now() < rscoupon("StartDate") and Now() > rsCoupon("StopDate") Then
				'coupon code found but expired
				txtItemName = txtItemName & " (Discount Code Expired - " & txtCoupon & ")"
				txtNote = "Sorry, the Discount Code '" & txtCoupon & "' does not fall within the offer dates."
				txtCoupon = ""
			ElseIF rsCoupon("Valid") = 0 Then
				'coupon code found but invalid offer			
				txtAmount = rsCoupon(txtFieldName)
				txtItemName = txtItemName & " (Discount Code Invalid - " & txtCoupon & ")"	
				txtNote = "Sorry, the Discount Code '" & txtCoupon & "' is no longer offered."
				txtCoupon = ""
			Else
				'coupon code found, within date, valid offer			
				txtAmount = rsCoupon(txtFieldName)
				txtItemName = txtItemName & " (Discount Code - " & txtCoupon & ")"
				txtNote = "The Discount Code '" & txtCoupon & "' has been accepted, discount noted below."
			End If
		Else
			'no coupon code entered, standard pricing
				txtItemName = txtItemName & " (No Discount Code)"			
		End If	
%>
			<table>
<%
If Not txtNote = "" Then
%>
				<tr>
					<td class="status" colspan="3" align="center"><%= txtNote %><br>&nbsp;</td>
				</tr>
				
<%
End If
%>
				<tr>
					<td class="BodyCopy" colspan="3" align="center"><strong>Review your Stresscom Purchase</strong><hr></td>
				</tr>
				<tr>
					<td class="BodyCopy">Purchasing:</td>
					<td width="10"></td>
					<td class="BodyCopy"><%= txtItemName %></td>
				</tr>
				<tr>
					<td class="BodyCopy">Price per item:</td>
					<td width="10"></td>
					<td class="BodyCopy"><%= FormatCurrency(txtAmount) %></td>
				</tr>
				<tr>
					<td class="BodyCopy">Quantity:</td>
					<td width="10"></td>
					<td class="BodyCopy"><%= intQuantity %> test administration(s)</td>
				</tr>
				<tr>
					<td class="BodyCopy">Discount Code:</td>
					<td width="10"></td>
					<td class="BodyCopy"><% If txtCoupon = "" Then Response.Write("N/A") Else Response.Write(txtCoupon) End If%> (optional)</td>
				</tr>
<%
'two form options, one for free ssfi and one for purchasing
If txtAmount = 0 Then
%>
			<form action="index.asp?mod=9&act=nch" method="post">
				<input type="hidden" name="fQuantity" value="1">
				<input type="hidden" name="fItemName" value="<%= txtItemName %>">
				<input type="hidden" name="fAmount" value="<%= txtAmount %>">
				<input type="hidden" name="fCoupon" value="<%= txtCoupon %>">
				<tr>
					<td class="BodyHead" colspan="3" align="center"><hr>You will <u>not</u> be charged for taking Stresscom.<hr></td>
				</tr>
				<tr>
					<td class="BodyCopy" colspan="3" align="center"><input type="submit" value="Continue to the Stresscom Test"></td>
				</tr>
			</form>
<%
Else
%>			
			<form action="https://www.paypal.com/cgi-bin/webscr" method="post">
			<!-- <form action="https://www.sandbox.paypal.com/cgi-bin/webscr" method="post"> -->
				<input type="hidden" name="cmd" value="_xclick">
				<input type="hidden" name="business" value="sales@ombudsmanpress.com">
				<input type="hidden" name="quantity" value="1">
				<input type="hidden" name="item_name" value="<%= txtItemName %>">
				<input type="hidden" name="amount" value="<%= txtAmount %>">
				<input type="hidden" name="page_style" value="OMB">
				<input type="hidden" name="no_shipping" value="1">
				<input type="hidden" name="return" value="http://www.stresscom.net/index.asp?mod=9&amp;act=bot">
				<input type="hidden" name="cancel_return" value="http://www.stresscom.net/index.asp?mod=9&amp;act=pcn">
				<input type="hidden" name="no_note" value="1">
				<input type="hidden" name="currency_code" value="USD">
				<input type="hidden" name="lc" value="US">
				<input type="hidden" name="bn" value="PP-BuyNowBF">
				<input type="hidden" name="first_name" value="<%= Request.Form("fFirst") %>">
				<input type="hidden" name="last_name" value="<%= Request.Form("fLast") %>">
				<input type="hidden" name="address1" value="<%= Request.Form("fAddress") %>">
				<input type="hidden" name="city" value="<%= Request.Form("fCity") %>">
				<input type="hidden" name="state" value="<%= Request.Form("fState") %>">
				<input type="hidden" name="zip" value="<%= Request.Form("fZip") %>">
				<input type="hidden" name="email" value="<%= Request.Form("fEmail") %>">
				<tr>
					<td class="BodyCopy" colspan="3" align="center"><hr>Clicking on the button below will take you to PayPal for your purchase.<br>&nbsp;<br>You do not need a PayPal account to pay via a credit card.<br>&nbsp;<br>Please note that it might take up to 30 seconds for the PayPal site to open.<hr></td>
				</tr>
				<tr>
					<td class="BodyCopy" colspan="3" align="center"><input type="submit" value="Continue to Purchase Stresscom"></td>
				</tr>
			</form>
<%
End If
%>
			</table>
			
<%
'***************
	Case "nch" ' no charge to take the ssfi
'***************
		response.Write("the section is under contruction<br>")
		
		txtAmount = Request.Form("fAmount")
		intQuantity = Request.Form("fQuantity")
		txtItemName = Request.Form("fItemName")
		txtCoupon = Request.Form("fCoupon")
		
		'get quantity based on ItemName
		If InStr(1, txtItemName, "Single", 1) > 0 Then
			intQuantity = 1
		Else
			intQuantity = 4
		End If
		response.Write("txtItemName = " & txtItemName & "<br>")
		response.Write("InStr = " & InStr(txtItemName, "Single") & "<br>")
		response.Write("intQuantity = " & intQuantity & "<br>")
		
		sqlGetLimit = "SELECT * FROM SCUser.tblCoupon WHERE Code = '" & txtCoupon & "'"
		response.Write(sqlGetLimit & "<br>")
		set rsGetLimit = dbcon.Execute(sqlGetLimit)
		
		intQuanLimit = rsGetLimit("QuanLimit")
		
		sqlCount = "SELECT * FROM SCUser.tblPurchase WHERE CouponID = '" & txtCoupon & "'"
		set rsCount = dbcon.Execute(sqlCount)
		
		intBought = 0
		If rsCount.BOF and rsCount.EOF Then
			intBought = 0
		Else
			Do While NOT rsCount.EOF
				intBought = intBought + 1
				rsCount.MoveNext
			Loop
		End If
		
		response.Write("Coupon Count = " & intBought & " of " & intQuanLimit & "<br>")
		
		Randomize
		upperlimit = 99000.0
		lowerlimit = 10000.0
		dblTran = Int((upperlimit - lowerlimit + 1)*Rnd() + lowerlimit)
		txtTran = Left(txtCoupon, 3) & CStr(dblTran)
		response.Write(cstr(txtTran) & "<br>")
		
		If intBought < intQuanLimit Then
			'add purchase to tblPurchase
			SID = Session("sid")
			
			sqlPurchase = "SET NOCOUNT ON; " &_
				"INSERT INTO SCUser.tblPurchase (StudentID, PayPalTransaction, CouponID, Quantity, Date, Total) " &_
				"VALUES(" & SID & ", '" & txtTran & "', '" & txtCoupon & "', " & intQuantity & ", '" & Date() & "', 0" &_
				"); SELECT @@IDENTITY AS NewID;"
			
			Set rsInsert = dbCon.Execute(sqlPurchase)
			
			intPurchaseID = rsInsert("NewID")
			
			Set rsInsert = Nothing
			
			'add ssfi record for how ever many tests were purchased
			For i = 1 to intQuantity
				sqlStressSchedule = "INSERT INTO SCUser.tblStressSchedule (StudentID, PurchaseID) VALUES (" & SID & ", " & intPurchaseID & ")"
				'response.Write(sqlssfi)
				Set rsInsert = dbCon.Execute(sqlStressSchedule)
				Set rsInsert = Nothing
			Next
			
			Response.Redirect("index.asp?mod=13&act=fpr")
		End If
	
'***************
	Case "bot" ' tests were bought successfully
'***************
Dim authToken, txToken
Dim query
Dim objHttp
Dim sQuerystring
Dim sParts, iParts, aParts
Dim sResults, sKey, sValue
Dim i, result
Dim firstName, lastName, itemName, mcGross, mcCurrency

'live token
authToken = "f67O9D5J68z2SH8xfKIa8sHA8UhRPtMe0PeLJSjOQFYFjOt_PPCDtXRJsiW"

'sandbox token
'authToken = "6CeqcMScTvZ1-KeNRurda00x61tSuSspB6Nd8TCWUwRvR63dZ5UUO2xHN1i"


txToken = Request.Querystring("tx")

query = "cmd=_notify-synch&tx=" & txToken & "&at=" & authToken

set objHttp = Server.CreateObject("Microsoft.XMLHTTP")
' set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")

'live server
objHttp.open  "POST", "http://www.paypal.com/cgi-bin/webscr", false

'sandbox server
'objHttp.open  "POST", "https://www.sandbox.paypal.com/cgi-bin/webscr", false


objHttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
objHttp.Send query

sQuerystring = objHttp.responseText

If Mid(sQuerystring,1,7) = "SUCCESS" Then
sQuerystring = Mid(sQuerystring,9)
sParts = Split(sQuerystring, vbLf)
iParts = UBound(sParts) - 1
ReDim sResults(iParts, 1)
For i = 0 To iParts
aParts = Split(sParts(i), "=")
sKey = aParts(0)
sValue = aParts(1)
sResults(i, 0) = sKey
sResults(i, 1) = sValue

Select Case sKey
Case "first_name"
firstName = unescape(replace(sValue, "+", " "))
Case "last_name"
lastName = unescape(replace(sValue, "+", " "))
Case "item_name"
txtItemName = unescape(replace(sValue, "+", " "))
Case "payment_date"
payDate = unescape(replace(sValue, "+", " "))
Case "mc_gross"
mcGross = unescape(replace(sValue, "+", " "))
Case "mc_currency"
mcCurrency = unescape(replace(sValue, "+", " "))
Case "txn_id"
txnID = unescape(replace(sValue, "+", " "))
Case "receipt_id"
receiptID = unescape(replace(sValue, "+", " "))
Case "quantity"
intQuantity = unescape(replace(sValue, "+", " "))
End Select
Next

'add purchase to tblPurchase
SID = Session("sid")
If InStr(txtItemName, "-") > 1 Then
	intTrim = InStr(txtItemName, "-")
	intILen = Len(txtItemName)
	intRight = (intILen - IntTrim) - 1
	txtCoupon = Right(txtItemName, intRight)
	txtCoupon = Left(txtCoupon, intRight - 1)
End If

'get quantity purchased based on ItemName
If InStr(1, txtItemName, "Single", 1) > 0 Then
	intQuantity = 1
Else
	intQuantity = 4
End If

sqlPurchase = "SET NOCOUNT ON; " &_
	"INSERT INTO SCUser.tblPurchase (StudentID, PayPalTransaction, CouponID, Quantity, Date, Total) " &_
	"VALUES(" & SID & ", '" & txnID & "', '" & txtCoupon & "', " & intQuantity & ", '" & Date() & "', " &_
	mcGross & "); SELECT @@IDENTITY AS NewID;"

Set rsInsert = dbCon.Execute(sqlPurchase)

intPurchaseID = rsInsert("NewID")

Set rsInsert = Nothing

'add tblStressSchedule record for how ever many tests were purchased
For i = 1 to intQuantity
	sqlStressSchedule = "INSERT INTO SCUser.tblStressSchedule (StudentID, PurchaseID) VALUES (" & SID & ", " & intPurchaseID & ")"
	'response.Write(sqlssfi)
	Set rsInsert = dbCon.Execute(sqlStressSchedule)
	Set rsInsert = Nothing
Next
%>
			<table>
				<tr>
					<td class="BodyCopy" colspan="3" align="center">Purchase Summary</td>
				</tr>
				<tr>
					<td class="BodyCopy" align="right"><strong>Buyer:</strong></td>
					<td width="10"></td>
					<td class="BodyCopy"><%= firstName & " " & lastName %></td>
				</tr>
				<tr>
					<td class="BodyCopy" align="right"><strong>Purchase Date:</strong></td>
					<td></td>
					<td class="BodyCopy"><%= payDate %></td>
				</tr>
				<tr>
					<td class="BodyCopy" align="right"><strong>Item:</strong></td>
					<td></td>
					<td class="BodyCopy"><%= txtItemName %></td>
				</tr>
				<tr>
					<td class="BodyCopy" align="right"><strong>Quantity:</strong></td>
					<td></td>
					<td class="BodyCopy"><%= intQuantity %></td>
				</tr>
				<tr>
					<td class="BodyCopy" align="right"><strong>Total:</strong></td>
					<td></td>
					<td class="BodyCopy"><%= FormatCurrency(mcGross) & " " & mcCurrency %></td>
				</tr>
				<tr>
					<td class="BodyCopy" align="right"><strong>Transaction ID:</strong></td>
					<td></td>
					<td class="BodyCopy"><%= txnID %></td>
				</tr>
				<tr>
					<td class="BodyCopy" colspan="3"><hr></td>
				</tr>
				<tr>
					<td class="BodyCopy" colspan="3">
						<p>Thank you for purchasing the <strong>Stresscom</strong> assessment.</p>
						<p align="right"><button onClick="window.location='index.asp?mod=13&act=fpr'">Continue to take Stresscom</button></td>
				</tr>
			</table>

<%

Else
'log for manual investigation
Response.Write("ERROR")
End If


'***************
	Case "pcn" 'purchase cancelled
'***************
%>
				Your Purchase has been cancelled.				
<%
'***************
	Case Else 'the default view
'***************
%>
			<table border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td class="BodyCopy" colspan="3">
<%
		'see if the parent has already purchased test.  If so, show the details
		sqlPurch = "SELECT * FROM SCUser.tblPurchase WHERE StudentID = " & Session("SID")
		Set rsPurch = dbcon.Execute(sqlPurch)
		If Not rsPurch.BOF and Not rsPurch.EOF Then
%>
				<table border="0" cellpadding="0" cellspacing="0">
					<tr>
						<td class="BodyCopy" colspan="9" align="center">Purchase History<hr></td>
					</tr>
					<tr>
						<td class="BodyCopy">Date</td>
						<td class="BodyCopy" width="10"></td>
						<td class="BodyCopy">Coupon</td>
						<td class="BodyCopy" width="10"></td>
						<td class="BodyCopy">Quantity</td>
						<td class="BodyCopy" width="10"></td>
						<td class="BodyCopy">Total</td>
						<td class="BodyCopy" width="10"></td>
						<td class="BodyCopy">Transaction ID</td>
					</tr>				
<%	
			Do While Not rsPurch.EOF
%>
					<tr>
						<td class="BodyCopy" align="center"><%= rsPurch("Date") %></td>
						<td class="BodyCopy"></td>
						<td class="BodyCopy" align="center"><%= rsPurch("CouponID") %></td>
						<td class="BodyCopy"></td>
						<td class="BodyCopy" align="center"><%= rsPurch("Quantity") %></td>
						<td class="BodyCopy"></td>
						<td class="BodyCopy" align="center"><%= FormatCurrency(rsPurch("Total")) %></td>
						<td class="BodyCopy"></td>
						<td class="BodyCopy" align="center"><%= rsPurch("PayPalTransaction") %></td>
					</tr>
<%	
				rsPurch.MoveNext
			Loop
%>
					<tr>
						<td class="BodyCopy" colspan="9" align="center"><hr></td>
					</tr>
				</table>				
<%	
		End If
		
		'Get the customer info to pass on to paypal
		sqlStudentInfo = "SELECT * FROM SCUser.tblStudent WHERE StudentID = " & Session("SID")
		Set rsStudentInfo = dbCon.Execute(sqlStudentInfo)
		Do While NOT rsStudentInfo.EOF
			txtFirst = rsStudentInfo("FirstName")
			txtLast = rsStudentInfo("LastName")
			txtAddress = rsStudentInfo("Address")
			txtCity = rsStudentInfo("City")
			txtState = rsStudentInfo("State")
			txtZip = rsStudentInfo("Zip")
			txtEmail = rsStudentInfo("Email")
			Exit Do
		Loop
%>
					</td>
				</tr>
			<form action="?mod=9&act=buy" method="post">
				<input type="hidden" name="fFirst" value="<%= txtFirst %>">
				<input type="hidden" name="fLast" value="<%= txtLast %>">
				<input type="hidden" name="fAddress" value="<%= txtAddress %>">
				<input type="hidden" name="fCity" value="<%= txtCity %>">
				<input type="hidden" name="fState" value="<%= txtState %>">
				<input type="hidden" name="fZip" value="<%= txtZip %>">
				<input type="hidden" name="fEmail" value="<%= txtEmail %>">
				<tr>
					<td class="BodyCopy" colspan="3" align="center">Purchase <strong>Stresscom</strong><hr></td>
				</tr>
				<tr>
					<td class="BodyCopy" align="right">Select:</td>
					<td width="10"></td>
					<td class="BodyCopy"><input type="radio" name="fQuantity" value="4"> $12.00 - 4 Stresscom test administration + interpretation<br>&nbsp;<br>
<!--						<input type="radio" name="fQuantity" value="4"> $19.95 - Stresscom subscription<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(4 administrations plus progress reports for one year)*</td>
				</tr>  -->
				<tr>
					<td class="BodyCopy" colspan="3" align="center"><hr>If you have a Discount Code from your physician,<br>school, or company agency, please enter it below.<br>&nbsp;<br>Leave Discount Code blank if you do not have one.<hr></td>
				</tr>
				<tr>
					<td class="BodyCopy" align="right">Company, Physician,<br>or School Discount Code:</td>
					<td width="10"></td>
					<td class="BodyCopy"><input type="text" name="fCouponCode"> (optional)</td>
				</tr>
				<tr>
					<td class="BodyCopy" colspan="3" align="center"><hr>* The Stresscom year's subscription works best if you take Stresscom once each quarter of the year and apply the nine stress management techniques in between.<hr></td>
				</tr>
				<tr>
					<td class="BodyCopy" colspan="3" align="center"><hr></td>
				</tr>
				<tr>
					<td class="BodyCopy" colspan="3" align="right">&nbsp;<br><input type="submit" value="Continue to Purchase Stresscom"></td>
				</tr>
			</form>
			</table>
<%
End Select
%>	
<!-- #Include file="../inc/mod-bot.asp" -->