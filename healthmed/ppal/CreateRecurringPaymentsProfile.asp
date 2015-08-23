<!-- #include file ="CallerService.asp" -->
<%

	Response.Buffer = True
'-----------------------------------------------------------------------------
' Author : Sergey, Wiztel INC
'
' Calls CallerService.asp and APIError.asp.

'-----------------------------------------------------------------------------
	Dim firstName
	Dim lastName
	Dim creditCardType
	Dim creditCardNumber
	Dim expDateMonth
	Dim expDateYear
	Dim padDate
	'Dim cvv2Number
	'Dim address1
	'Dim address2
	'Dim city
	'Dim state
	'Dim zip
	Dim amount
	'Dim currencyCode
	'Dim paymentType
	Dim nvpstr
	Dim resArray
	Dim ack
	Dim message
    Dim startdate
	Dim billingperiod
    Dim billingfrequency
    'Dim dt

	'dt=Date()

	'startdate = Month(dt)&"/"&Day(dt)&"/"&Year(dt)

    startdate = Date()

	billingperiod ="Month"
    billingfrequency ="1"

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

	
	creditCardNumber	=  Request.Form("card_number") ' Request.Form("creditCardNumber") '+

	
	expDateMonth		= Request.Form("card_expiry_month") ' Request.Form("expDateMonth")
	expDateYear			= Request.Form("card_expiry_year") ' Request.Form("expDateYear")
	padDate				= expDateMonth&expDateYear
	'cvv2Number			= Request.Form("card_verif_number") ' Request.Form("cvv2Number")
	'address1			=  Request.Form("address") ' Request.Form("address1")
	'address2			= "" ' Request.Form("address2")
	'city				= Request.Form("city") ' Request.Form("city")
	'state				= Request.Form("province") ' Request.Form("state")
	'zip					= Request.Form("postal_code") ' Request.Form("zip")
	'country             = Request.Form("cpayment_country")
	amount				= Request.Form("total_amount_gross") ' Request.Form("amount")
	'currencyCode		=Request.Form("currency")
	'currencyCode		= "USD"
	'paymentType			= "Authorization"  'Request.Form("paymentType")
'-----------------------------------------------------------------------------
' Construct the request string that will be sent to PayPal.
' The variable $nvpstr contains all the variables and is a
' name value pair string with &as a delimiter
'-----------------------------------------------------------------------------
	nvpstr	=	"&AMT="&amount &_
				"&CREDITCARDTYPE="&creditCardType &_
				"&ACCT="&creditCardNumber & _
				"&EXPDATE=" & padDate &_
				"&FIRSTNAME=" & firstName &_
				"&LASTNAME=" & lastName &_
				"&PROFILESTARTDATE=" & startdate &_
                "&BILLINGPERIOD=" & billingperiod &_
                "&BILLINGFREQUENCY=" & billingfrequency 

'nvpstr	=	"&AMT="&amount &_
'				"&CREDITCARDTYPE="&creditCardType &_
'				"&ACCT="&creditCardNumber & _
'				"&EXPDATE=" & padDate &_
'				'"&CVV2=" & cvv2Number &_
'				"&FIRSTNAME=" & firstName &_
'				"&LASTNAME=" & lastName &_
'				'"&STREET=" & address1 &_
				'"&CITY=" & city &_
				'"&STATE=" & state &_
				'"&ZIP=" &zip &_
				'"&COUNTRYCODE=" & country  &_
				'"&CURRENCYCODE=" & currencyCode &_
'				"&PROFILESTARTDATE=" & startdate &_
 '               "&BILLINGPERIOD=" & billingperiod &_
  '              "&BILLINGFREQUENCY=" & billingfrequency 

	nvpstr	=	URLEncode(nvpstr)
'-----------------------------------------------------------------------------
' Make the API call to PayPal,using API signature.
' The API response is stored in an associative array called gv_resArray
'-----------------------------------------------------------------------------

'response.write "stopcall" & nvpstr

'response.End

	Set resArray	= hash_call("CreateRecurringPaymentsProfile",nvpstr)
	ack = UCase(resArray("ACK"))
'----------------------------------------------------------------------------------
' Display the API request and API response back to the browser.
' If the response from PayPal was a success, display the response parameters
' If the response was an error, display the errors received
'----------------------------------------------------------------------------------
	If ack="SUCCESS" Then
		message="Profile was created!"
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
						profileid=resArray("PROFILEID")
						%>
						Profile ID:</td><td><%=profileid%></td>
				</tr>
				<tr>
					<td>
						Status:</td>
					<td>USD <%=resArray("STATUS")%></td>
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
