<!-- #include file ="CallerService.asp" -->
<%

	Response.Buffer = True
'-----------------------------------------------------------------------------
' Author : Sergey, Wiztel INC
'
' Calls CallerService.asp and APIError.asp.

'-----------------------------------------------------------------------------
	Dim profileid


    profileid=Request.QueryString ("profileID")
	'-----------------------------------------------------------------------------
' Construct the request string that will be sent to PayPal.
' The variable $nvpstr contains all the variables and is a
' name value pair string with &as a delimiter
'-----------------------------------------------------------------------------
	nvpstr	=	"&PROFILEID=" &profileid
				
	nvpstr	=	URLEncode(nvpstr)
'-----------------------------------------------------------------------------
' Make the API call to PayPal,using API signature.
' The API response is stored in an associative array called gv_resArray
'-----------------------------------------------------------------------------

'response.write "stopcall" & nvpstr

'response.End

	Set resArray	= hash_call("GetRecurringPaymentsProfileDetails",nvpstr)
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
