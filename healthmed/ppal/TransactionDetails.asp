<!-- #include file ="CallerService.asp" -->
<%
'----------------------------------------------------------------------------------
' TransactionDetails.asp
' ======================
' Sends a GetTransactionDetails NVP API request to PayPal.

' The code retrieves the transaction ID and constructs the
' NVP API request string to send to the PayPal server. The
' request to PayPal uses an API Signature.

' After receiving the response from the PayPal server, the
' code displays the request and response in the browser. If
' the response was a success, it displays the response
' parameters. If the response was an error, it displays the
' errors received.

' Called by GetTransactionDetails.html.

' Calls CallerService.asp and APIError.asp.
'----------------------------------------------------------------------------------
	On Error Resume Next
	Dim transactionID 
	Dim nvpstr
	Dim resArray

	
	transactionID	= Request.QueryString("transactionID")
	doc=Request.QueryString("doc")
	If doc = "" Then doc = "0" End If
	
'	-------------------------------------------------------------------------------
' Construct the request string that will be sent to PayPal.
' The variable nvpStr contains all the variables and is a
' name value pair string with & as a delimiter 
'----------------------------------------------------------------------------------
	 nvpstr="&TRANSACTIONID="&transactionID
	 nvpstr=URLEncode(nvpstr)
'----------------------------------------------------------------------------------
' Make the API call to PayPal, using API signature.
' The API response is stored in an associative array called resArray 
'----------------------------------------------------------------------------------
	Set resArray=hash_call("gettransactionDetails",nvpstr)
	ack = UCase(resArray("ACK"))
'----------------------------------------------------------------------------------
' Display the API request and API response back to the browser.
' If the response from PayPal was a success, display the response parameters
' If the response was an error, display the errors received
'----------------------------------------------------------------------------------
	If ack="SUCCESS" Then
	  	message="Transaction Details"

		If doc = "1" Then
		 message="Authorization " & message
		End If
		
	     If doc = "2" Then
		 message="Capture " & message
		
		End If
		
	Else       
		 Set SESSION("nvpErrorResArray") = resArray
		 Response.Redirect "APIError.asp"
	End If	
		
dim dovoidurl
dim docaptureurl
dim dorefundurl
dim amount 

amount = resArray("AMT")

'--------------------------------------------------------------------------------------------
' If there is no Errors Construct the HTML page with a table of variables Loop through the associative array 
' for both the request and response and display the results.
'--------------------------------------------------------------------------------------------
%>

		<center>
		<font size="2" color="black" face="Verdana"><b><%=message%></b></font>
			<table width=600>
				<tr>
					<td >Account:<!--Payer:-->
					</td>
					<%
					v=resArray("RECEIVERBUSINESS")
					p=InStr(v,"paypal@wiztel.ca")
					If p > 0 then
					v="Wiztel Business Account"
                    Else
                    v="Fake Test Account"
                    End if
					%>
					<td><%=v%></td>
				</tr>
				<tr>
					<td >
						Payer ID:
					</td>
					<td><%=resArray("PAYERID")%></td>
				</tr>
				<tr>
					<td >
						First Name:
					</td>
					<td><%=resArray("FIRSTNAME")%></td>
				</tr>
				<tr>
					<td >
						Last Name:</td>
					<td><%=resArray("LASTNAME")%></td>
				</tr>
				<tr>
					<td >
						Parent Transaction ID (if any):
					</td>
					<td>
					</td>
				</tr>
				<tr>
					<td >
						Transaction ID:
					</td>
					<td><%=resArray("TRANSACTIONID")%></td>
				</tr>
				<tr>
					<td >
						Gross Amount:
					</td>
					<td><%=resArray("AMT")%></td>
				</tr>
				<tr>
					<td >
						Payment Status:
					</td>
					<td><%=resArray("PAYERSTATUS")%></td>
				</tr>
			</table>
		</center>

    <% 
    If Err.Number <> 0 Then 
	SESSION("ErrorMessage")	= ErrorFormatter(Err.Description,Err.Number,Err.Source,"TransactionDetails.asp")
	Response.Redirect "APIError.asp"
	Else
	SESSION("ErrorMessage")	= Null
	End If
    %>
