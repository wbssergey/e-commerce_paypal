<!-- #include file ="CallerService.asp" -->
<%

	Response.Buffer = True
'-----------------------------------------------------------------------------
' DoDirectPaymentReceipt.asp

' Submits a credit card transaction to PayPal using a
' DoDirectPayment request.

' The code collects transaction parameters from the form
' displayed by DoDirectPayment.asp then constructs and sends
' the DoDirectPayment request string to the PayPal server.
' The paymentType variable becomes the PAYMENTACTION parameter
' of the request string.

' After the PayPal server returns the response, the code
' displays the API request and response in the browser.
' If the response from PayPal was a success, it displays the
' response parameters. If the response was an error, it
' displays the errors.

' Called by DoDirectPayment.asp.

' Calls CallerService.asp and APIError.asp.

'-----------------------------------------------------------------------------
	Dim firstName
	Dim lastName
	Dim creditCardType
	Dim creditCardNumber
	Dim expDateMonth
	Dim expDateYear
	Dim padDate
	Dim cvv2Number
	Dim address1
	Dim address2
	Dim city
	Dim state
	Dim zip
	Dim amount
	Dim currencyCode
	Dim paymentType
	Dim nvpstr
	Dim resArray
	Dim ack
	Dim message


	firstName			= Request.Form("primary_first_name") ' Request.Form("firstName")
	lastName			=  Request.Form("primary_last_name") ' Request.Form("lastName")


	creditCardType		= Request.Form("card_type") ' Request.Form("creditCardType")

    select case creditCardType
	case "1": creditCardType="MasterCard"  
	case "2": creditCardType="Visa"  
	Case "3": creditCardType="Amex"  
    Case "4": creditCardType="Discover"
	case else: creditCardType="unknown" 
    end select

	
	creditCardNumber	=  Trim(Request.Form("card_number")) ' Request.Form("creditCardNumber")

	
	expDateMonth		= Trim(Request.Form("card_expiry_month")) ' Request.Form("expDateMonth")
	expDateYear			= Trim(Request.Form("card_expiry_year")) ' Request.Form("expDateYear")
	padDate				= expDateMonth&expDateYear
	cvv2Number			= Trim(Request.Form("card_verif_number")) ' Request.Form("cvv2Number")
	address1			=  Trim(Request.Form("address")) ' Request.Form("address1")
	address2			= "" ' Request.Form("address2")
	city				= Trim(Request.Form("city")) ' Request.Form("city")
	state				= Trim(Request.Form("province")) ' Request.Form("state")
	zip					= Trim(Request.Form("postal_code")) ' Request.Form("zip")
	country             = Trim(Request.Form("cpayment_country"))
	amount				= Trim(Request.Form("total_amount")) ' Request.Form("amount")
	'currencyCode		=Request.Form("currency")
	currencyCode		= "USD"
	paymentType			= "Authorization"  'Request.Form("paymentType")
'-----------------------------------------------------------------------------
' Construct the request string that will be sent to PayPal.
' The variable $nvpstr contains all the variables and is a
' name value pair string with &as a delimiter
'-----------------------------------------------------------------------------


	nvpstr	=	"&PAYMENTACTION=" &paymentType & _
				"&AMT="&amount &_
				"&CREDITCARDTYPE="&creditCardType &_
				"&ACCT="&creditCardNumber & _
				"&EXPDATE=" & padDate &_
				"&CVV2=" & cvv2Number &_
				"&FIRSTNAME=" & firstName &_
				"&LASTNAME=" & lastName &_
				"&STREET=" & address1 &_
				"&CITY=" & city &_
				"&STATE=" & state &_
				"&ZIP=" &zip &_
				"&COUNTRYCODE=" & country  &_
				"&CURRENCYCODE=" & currencyCode
	nvpstr	=	URLEncode(nvpstr)


'-----------------------------------------------------------------------------
' Make the API call to PayPal,using API signature.
' The API response is stored in an associative array called gv_resArray
'-----------------------------------------------------------------------------

'response.write "stopcall" & nvpstr

'response.End


	Set resArray	= hash_call("doDirectPayment",nvpstr)

	ack = UCase(resArray("ACK"))
'----------------------------------------------------------------------------------
' Display the API request and API response back to the browser.
' If the response from PayPal was a success, display the response parameters
' If the response was an error, display the errors received
'----------------------------------------------------------------------------------
	If ack="SUCCESS" Then
		message="Thank you for your payment!"
	Else
		 Set SESSION("nvpErrorResArray") = resArray
		 Response.Redirect "APIError.asp"
	End If

%>
		<center>
			<font size="2" color="black" face="Verdana"><b>Pay Pal Response</b></font>
			<br>
			<br>
			<b>
				<%=message%>
			</b>
			<br>
			<br>
			<table width="400"">
				<tr>
					<td>
				    	<%
						transactionid=resArray("TRANSACTIONID")
						%>
						Transaction ID:</td><td><%=transactionid%></td>
				</tr>
				<tr>
					<td>
						Amount:</td>
					<td>USD <%=resArray("AMT")%></td>
				</tr>
				<tr>
					<td>
						AVS:</td>
					<td><%=resArray("AVSCODE")%></td>
				</tr>
				<tr>
					<td>
						CVV2:</td>
					<td><%=resArray("CVV2MATCH")%></td>
				</tr>
			</table>
		</center>
		<%
      If Err.Number <> 0 Then
	SESSION("ErrorMessage")	= ErrorFormatter(Err.Description,Err.Number,Err.Source,"DoDirectPaymentReceipt.asp")
	Response.Redirect "APIError.asp"
	Else
	SESSION("ErrorMessage")	= Null
	End If
    %>
