<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE>Payment Information</TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
<LINK href="i_style.css" rel="stylesheet" type="text/css">
</HEAD><BODY>
<%
Dim pagename
pagename="payment.asp"
session("pagename")="payment.asp"
%>
<FORM name="form1" METHOD=POST action="" >
<!--#include file=header.asp-->


<%Call menu()


Dim errmsg
Dim OrderID
OrderID=request.Form("OrderID")
If OrderID="" Then
OrderID=request.querystring("OrderID")

Else
End If
If request.Form("OrderID")<>"" And request.querystring("OrderID") ="" then
errmsg= "Order was inserted successfully."
End If

Dim Intension
Intension=request.Form("Intension")


Dim primary_first_name 
Dim primary_last_name
Dim card_type '1 master 2 visa 3 amex 4 discover
Dim card_number 
Dim card_expiry_month 
Dim card_expiry_year 
Dim card_verif_number 
Dim address 
Dim city 
Dim province 'short name
Dim HomePhone
Dim postal_code 
Dim cpayment_country 'short name
Dim total_amount 'total

'For Each Field In Request.Form
'response.write Field&":"
'Response.Write Request.Form(Field)&"<br>"
'Next
'For Each Field In Request.Form
'<INPUT TYPE="hidden" NAME="=Field" value="<%=Request.Form(Field) " ><%
	
'Next
paymentq="EXEC	sp_paymentinfo	@orderid = "&OrderID
Set paymentR=oConn.execute(paymentq)
If Not paymentR.eof Then

primary_first_name=paymentR("primary_first_name")
primary_last_name=paymentR("primary_last_name")
card_type=paymentR("card_type")
card_number=paymentR("card_number")
card_expiry_month=paymentR("card_expiry_month")
card_expiry_year=paymentR("card_expiry_year")
card_verif_number=paymentR("card_verif_number")
address=paymentR("address")
city=paymentR("city")
province=paymentR("province")
HomePhone=paymentR("phone")
postal_code=paymentR("postal_code")
cpayment_country=paymentR("cpayment_countryppal")

total_amount=paymentR("total_amount")


Else
Errmsg="Error: No such a order exists."
End if
paymentR.close
If card_type="1" Then
	card_type="2"
Else
	If card_type="2" Then
		card_type="1"
	End if
End if

msgpal=""
msgpaltrans=""
msgpalcapture=""
bisppal="y" '    bisiness ppal if not empty 
amount=0
transactionid=""
transactionidcap=""

'-----------------------------send payment--------------
'-------------------------------------------------------

If Intension="sendpaypal" Then

   amount=Trim(Request.Form("total_amount"))
   orderid=Trim(Request.Form("OrderID"))

   url="http://www.wiztel.ca/ppal/DoDirectPaymentReceipt.asp"
 
 
 'response.write "'"&Trim(Request.Form("total_amount")) & "' c2"
 'response.end
  msgpal=getWeb("post",url,Request.Form)

Call wx(msgpal)
If InStr(msgpal,"Invalid Data") Then

		sql="update tblorder set  transactiondate =dbo.GetLocalDateTimebyID(13) "
		sql=sql&" , transactionstatus =5 " ' 5 wrong card number

		sql=sql& " where orderid="&orderid&"  "

		msgpaltrans= msgpal
		oconn.execute(sql)
		msgpal=""

else
    p=InStr(msgpal,"ID:</td><td>")

    If p > 1 Then

	transactionid=Mid(msgpal,p+12,InStr(p+10,msgpal,"</td>")-p-12)
   	    
    url="http://www.wiztel.ca/ppal/transactiondetails.asp?transactionID="&transactionid&"&Submit=Submit&bisppal="&bisppal

    url1="http://www.wiztel.ca/ppal/DoCaptureReceipt.asp?authorization_id="&transactionid&"&amount="&amount
	url1=url1&"&CompleteCodeType=Complete&bisppal="&bisppal

    msgpaltrans=getWeb("get",url,"")
    	
    msgpalcapture=getWeb("get",url1,"")

    p=InStr(msgpalcapture,"ID:</td><td>")


    If p > 1 Then
	transactionidcap=Mid(msgpalcapture,p+12,InStr(p+10,msgpalcapture,"</td>")-p-12)
	End If
	

	End If ' p > 1
	
	If transactionidcap = "" Then 

		sql="update tblorder set transactionid='"&transactionid&"', transactionidcap='"&transactionidcap
		sql=sql&"' , transactiondate =dbo.GetLocalDateTimebyID(13) "
		sql=sql&" , transactionstatus =1 " ' 1 status pending

		sql=sql& " where orderid="&orderid&"  "

			errmsg= "Error. <br>"&msgpaltrans&"<br>"&msgpalcapture

	Else
	' to do : keep in database transactionid, transactionidcap
		sql="update tblorder set transactionid='"&transactionid&"', transactionidcap='"&transactionidcap
		sql=sql&"' , transactiondate =dbo.GetLocalDateTimebyID(13) "
		sql=sql&" , transactionstatus =2 " ' 2 status authorized

		sql=sql& " where orderid="&orderid&"  "

			oconn.execute(sql)

	End If
	
     
End if

If InStr(msgpal,"Processor Decline") Then

		sql="update tblorder set  transactiondate =dbo.GetLocalDateTimebyID(13) "
		sql=sql&" , transactionstatus =3 " ' 3 declined

		sql=sql& " where orderid="&orderid&"  "

		msgpaltrans= msgpal
		oconn.execute(sql)
		msgpal=""
End If

If InStr(msgpal,"Credit Card Verification Number") Then

		sql="update tblorder set  transactiondate =dbo.GetLocalDateTimebyID(13) "
		sql=sql&" , transactionstatus =4 " ' 4 wrong security code
		sql=sql& " where orderid="&orderid&"  "

		msgpaltrans= msgpal
		oconn.execute(sql)
		msgpal=""
End If

errmsg="Payment was successfully sent." 

errmsg=errmsg & "<br>"&msgpaltrans&"<br>"&msgpalcapture

Intension=""
End if



'----------------------------end of send---------------------
'------------------------------------------------------
HomePhone=Trim(HomePhone)
If validate(HomePhone,"^[1]\d{10}$") Then
	HomePhone=Mid(HomePhone,2,3)&"-"&Mid(HomePhone,5,3)&"-"&Mid(HomePhone,8,4)
End If


If validate(HomePhone,"^\d{10}$") Then
	HomePhone=Mid(HomePhone,1,3)&"-"&Mid(HomePhone,4,3)&"-"&Mid(HomePhone,7,4)
End If

%>
<table border="1" cellspacing="0" width="430px" cellpadding="10px">
<tr><td ><b>Payment Information</b><br><br>
<table width='100%'> 
<th align=left>First Name:</th><td><%=primary_first_name%></td></tr>
<th align=left>Last Name:</th><td><%=primary_last_name%></td></tr>
<th align=left>Card Type:</th><td><%
Select Case card_type
Case "1": response.write "Master Card"
Case "2": response.write "Visa"
Case "3": response.write "American Express"
Case "4": response.write "Discover"
End select

%></td></tr>
<th align=left>Card Number:</th><td><%=card_number%></td></tr>
<th align=left>Expiration Date:</th><td><%=card_expiry_month%>/<%=card_expiry_year%></td></tr>
<th align=left>CSC:</th><td><%=card_verif_number%></td></tr>
<th align=left>Billing Address:</th><td><%=address%></td></tr>
<th align=left>City:</th><td><%=city%></td></tr>
<th align=left>State:</th><td><%=province%></td></tr>
<th align=left>ZIP:</th><td><%=postal_code%></td></tr>
<th align=left>Phone :</th><td><%=HomePhone%></td></tr>
<th align=left>Country:</th><td><%=cpayment_country%></td></tr>
<th align=left>Total:</th><td>$<%=formatnumber(total_amount,2)%>
</td></tr>
</table><br><br>
<%
Set eeeR=oCOnn.execute("SELECT TransactionStatus,TransactionID FROM TblOrder WHERE     OrderID = "&OrderID)
If Not eeeR.eof  Then
		myTransactionStatus=eeeR("TransactionStatus")
		If IsNull(myTransactionStatus) Then
		myTransactionStatus="0"
		End If
		myTransactionID=eeeR("TransactionID")
		If IsNull(myTransactionID) Then
		myTransactionID=""
		End If
		If Trim(myTransactionID)="" And Cint(myTransactionStatus)<2  Then
		%>
		<INPUT TYPE="submit" name="MySubmit" value="Send Payment" onclick="doSubmit('1');return false;">
	
		<%

		Else
		
		End if
End If
eeeR.close
%>

</td></tr></table>
<INPUT TYPE="hidden" NAME="total_amount" value="<%=total_amount%>" >
<INPUT TYPE="hidden" NAME="Intension" value="" >

<INPUT TYPE="hidden" NAME="primary_first_name" value="<%=primary_first_name %>" >
<INPUT TYPE="hidden" NAME="primary_last_name" value="<%=primary_last_name %>" >
<INPUT TYPE="hidden" NAME="card_type" value="<%=card_type %>" >
<INPUT TYPE="hidden" NAME="card_number" value="<%=card_number %>" >
<INPUT TYPE="hidden" NAME="card_expiry_month" value="<%=card_expiry_month %>" >
<INPUT TYPE="hidden" NAME="card_expiry_year" value="<%=card_expiry_year %>" >
<INPUT TYPE="hidden" NAME="card_verif_number" value="<%=card_verif_number %>" >
<INPUT TYPE="hidden" NAME="address" value="<%=address %>" >
<INPUT TYPE="hidden" NAME="city" value="<%=city%>" >
<INPUT TYPE="hidden" NAME="province" value="<%=province%>" >
<INPUT TYPE="hidden" NAME="postal_code" value="<%=postal_code %>" >
<INPUT TYPE="hidden" NAME="cpayment_country" value="<%=cpayment_country %>" >

<INPUT TYPE="hidden" NAME="OrderID" value="<%=OrderID %>" >
<INPUT TYPE="hidden" NAME="bisppal" value="<%=bisppal%>" >

<%Call showerr(errmsg)%>
</form>
<!--#include file=footer.asp-->
</BODY>
</HTML>
<%oConn.close
set oConn=Nothing
%>
<script language=Javascript>

<!--
var fs=document.forms.form1; 
function doSubmit(i)
{
if (i==1)
{
if (confirm("Please be sure not send this transaction twice."))
{
	document.getElementById("Intension").value='sendpaypal';

	fs.action='';
	fs.submit();
	return false;
}

}

}

function doSubmittostep1(i)
{
fs.action='default.asp';
fs.submit();
return false;
}

function bill2(i)
{
if (i=="1")
{
if (confirm("Please be aware that shipping to a different address from billilng address creates a higher risk of Customer fraud.")==true){
	fs.submit();
	return false;
					}

}
else {
fs.submit();
return false;

}
}
function TrimString(sInString) {
  sInString = sInString.replace( /^\s+/g, "" );// strip leading
  return sInString.replace( /\s+$/g, "" );// strip trailing
}

//-->
</SCRIPT>


