<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE> Order Lists </TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
<LINK href="i_style.css" rel="stylesheet" type="text/css">
<%
siteRoot="/"
%>
	<%="<script language=javascript>dirty=false;siteRoot='"&siteRoot&"';window.status='Show Orders'</script>" &vbcrlf%>
	<%="<script language=javascript1.2 src=""" &siteRoot &"i_header.js""></script>"%>


</HEAD><BODY>

<FORM name=form1 method=post action="">
<%

bisppal="y" ' bisiness paypal account if not empty

session("OrderList")=""
Dim pagename
pagename="ShowOrders.asp"
session("pagename")="ShowOrders.asp"
%>
<!--#include file=header.asp-->
<%
Dim condelete
condelete=request.Form("condelete")

Dim search
search=request.Form("search")
If search="" Then
search=session("search")
Else
session("search")=search
End If

search=Trim (search)

Dim AffiliateID
AffiliateID=request.Form("AffiliateID")

If AffiliateID="" Then
AffiliateID=session("AffiliateID")
End If

If AffiliateID="" Then
AffiliateID="0"
Else
session("AffiliateID")=AffiliateID
End if


Dim nodelete
nodelete=request.Form("nodelete")

Call menu()
Dim RemoveID
RemoveID=request.Form("RemoveID")

Dim Remove
Remove=request.Form("Remove")

If nodelete="Cancel" then
Remove=""
RemoveID=""
End If



Dim UserName
UserName=request.Form("UserName")

If UserName="" Then
UserName=session("UserName")

End If
If UserName="" Then
UserName="0"
End If
session("UserName")=UserName

'-----------delete a record
If condelete="Confirm to delete" Then
	removeq="EXEC sp_DelterOrderbyOrderID @orderid = "&RemoveID
	oCOnn.execute(removeq)
	errmsg="Order "&Remove&" was successfully deleted."

	Remove=""
	RemoveID=""
End If
%>
<!--#include file=DateSelect.asp-->


<table>

<tr><td>&nbsp;</td></tr>
	<tr><td colspan=3><B>Entered By : </B>
	<select class=dataselect name="username" onchange="doSubmit(3);return false;">
	<option value="0" >----- Select All -----</option>
	<%
	Set usernameR=oConn.execute("SELECT  UserName FROM Tblusers ")
	While Not usernameR.eof
	%><option value="<%=usernameR("UserName")%>" <%=selected(usernameR("UserName"),username)%>><%=usernameR("UserName")%></option>
	<%
	
	usernameR.movenext
	Wend
	usernameR.close
	%>
	
	
	</select></td>
	<td colspan=8><B> Affiliate : </B> 
				<select class=dataselect name="AffiliateID" onchange="doSubmit(3);return false;">
				<option value="0" >----- Select All -----</option>
			<%
			affiq="SELECT AffiliateID,AffiliateName FROM TblAffiliates"
			Set addiqR=oConn.execute(affiq)

			Do Until  addiqR.eof
			%><option value="<%=addiqR("AffiliateID")%>"  
			<%=selected(addiqR("AffiliateID"),AffiliateID)%>><%=addiqR("AffiliateName")%></option><%
			addiqR.movenext
			Loop

			addiqR.close
			%>

			</select>
	</td>
	<td>&nbsp;<b>General Search :</b> <INPUT TYPE="text" NAME="search" value="<%=search%>" size=4 width=4></td>
	<td colspan=2><input type="submit" name="mySubmit" class=databutton value="Submit" Title="Get the order lists." onclick="doSubmit(4);return false;" /></td>
            </tr>
</table><br>
<table><tr><td style=letter-spacing:1;font-size:10pt;font-weight:bold;color:#646496  >Order List:</td></tr></table>

<table class=datatable>
<tr><th>OrderID</th><th>Name</th><th>BillingAddress</th> 
<th>City</th><th>Province</th><th>Country</th><th>Phone</th><th>Affiliate</th><th nowrap>Payment Info</th><th>OrderTime</th>
<th>Price</th><th>Shipping</th><th>Total</th><th>TranStatus</th><th>TranID</th><th>TranactionDate</th><th>ShippingTracking#</th><th>Shipper</th><th>User</th><th>Delete</th>
</tr>
<%
OderlistQ="EXEC	sp_showorder @OrderTimes = '"&Time1&"',@OrderTimee = '"&Time2&"'"
If Trim(UserName)="0" Then
Else
OderlistQ=OderlistQ&",@username='"&UserName&"'"
End If

If AffiliateID="0" Then
Else
OderlistQ=OderlistQ&",@AffiliateID='"&AffiliateID&"'"
End If

If search="" Or search="?" then 
Else
OderlistQ=OderlistQ&",@Search='"&search&"'"
End If

'debug here to write the SQL query
'response.write OderlistQ
Set OrderRe=OConn.Execute(OderlistQ)
rowstyle="class=datarow0"
If OrderRe.eof Then
%><tr><td align=center colspan=20><I>No Data were selected.</I></td></tr><% 
End if
Do Until OrderRe.eof
%>
<tr <%
If CStr(Trim(RemoveID))=trim(OrderRe("OrderID")) Then
response.write "class=dataactive"
else
response.write rowstyle
End if
%>  ><td nowrap><a href="orderdetail.asp?OrderID=<%=OrderRe("OrderID")%>" title="Edit Order Details." ><%=OrderRe("myOrderID")%></a></td>



<td nowrap><%=OrderRe("Fname")%>&nbsp;<%=OrderRe("Lname")%></td>
<td nowrap><%=OrderRe("BillingAddress")%></td>

<td nowrap><%=OrderRe("City")%></td>
<td nowrap><%=OrderRe("Province")%></td>
<td nowrap><%=OrderRe("country")%></td>
<%
Myhomephone=Trim(OrderRe("phone"))

Myhomephone=Replace(Myhomephone,"-","")
If validate(Myhomephone,"^[1]\d{10}$") Then
	Myhomephone=Mid(Myhomephone,2,3)&"-"&Mid(Myhomephone,5,3)&"-"&Mid(Myhomephone,8,4)
End If


If validate(Myhomephone,"^\d{10}$") Then
	Myhomephone=Mid(Myhomephone,1,3)&"-"&Mid(Myhomephone,4,3)&"-"&Mid(Myhomephone,7,4)
End If

%>
<td nowrap><%=Myhomephone%></td>

<td nowrap><%=OrderRe("Affiliate")%></td>
<td nowrap><a href="" title="Show Payment Information Details." onclick="window.open('payment.asp?OrderID=<%=OrderRe("OrderID")%>','mywindow','status=1,toolbar=0,width=620');return false;">Payment Info</a></td>

<td nowrap><%=OrderRe("OrderTime")%></td>

<%myycost= OrderRe("Cost")
If IsNull(myycost) Then
myycost="0"
End if%>
<td align=right nowrap>$<%=FormatNumber(myycost,2)%></td>
<td align=right nowrap>$<%=FormatNumber(OrderRe("ShippingCost"),2)%></td>
<td align=right nowrap><b>$<%=FormatNumber(OrderRe("Total"),2)%></b></td>

<%checkdate=OrderRe("TransactionDate")

If checkdate="" Or IsNull(checkdate) Then
	%><td colspan=5 align=left> 
	<a href="TransResult.asp?OrderID=<%=OrderRe("OrderID")%>" 
	title="This Transaction information is empty.">Add Transaction Result</a> </td><%

Else

	%>
	<td align=left nowrap <%
	
	If OrderRe("TransactionStatus")="Authorized" Then 
	%>class=dataactive<%
	End If

	If InStr(OrderRe("TransactionStatus"),"Declined") Then 
	%>class=dataalert<%
	End If

	If InStr(OrderRe("TransactionStatus"),"Code failed")  Or InStr(OrderRe("TransactionStatus"),"Invalid Card Number") Or InStr(OrderRe("TransactionStatus"),"Pending") Then 
	%>class=datayellow<%
	End If

	%>	><%=OrderRe("TransactionStatus")%></td>
	
	<%
	paypal=OrderRe("TransactionID")
    paypalc=OrderRe("TransactionIDcap")
	
	zlink= "http://healthmed.asp.ca/showpaltransid.asp?paypalid="&paypal&"&paypalidc="&paypalc&"&bisppal="&bisppal

	%>
	
	<td align=left nowrap><%If paypalc <> "" then%>
	<a href="/z.asp" OnClick="return acctWin(this,'<%=zLink%>')"><%=paypal%></a>
	<%else%>
	<%=paypal%>
	<%End if%>
	</td>
	

	<td align=left nowrap><%=checkdate%></td>
	<%
	myshippingnumber=OrderRe("ShippingTrackingNumber")
	myshippingnumber=Trim(myshippingnumber)
	If IsNull(myshippingnumber) Then
	myshippingnumber=""
	End If
	
	%>
	<td align=left nowrap>
	<%
	If myshippingnumber="" Then
	%><a href="Shippinginfo.asp?OrderID=<%=OrderRe("OrderID")%>">Tracking#</a><%
	Else
	response.write myshippingnumber
	End if
	%>
	</td>
	<td align=left><%
	ddshipper=OrderRe("Shipper")
	If IsNull(ddshipper) Then
	ddshipper=""
	End If
	If ddshipper<>"" then
		Set SR=oConn.execute("SELECT ShipperID, Shipper FROM  TblShipper where ShipperID="&ddshipper)
		If Not SR.eof Then
		response.write SR("Shipper")
		End If
		SR.close
	End if
	%></td>
		<%
End if%>

<td align=right nowrap><%=OrderRe("useby")%></td>
<td><a  href="" onclick=" doRemove('<%=OrderRe("OrderID")%>','<%=OrderRe("myOrderID")%>');return false;" title="Remove this order">Delete</a></td>
</tr><%
OrderRe.movenext

	If rowstyle="class=datarow0" Then
	rowstyle="class=datarow1"
	Else
	rowstyle="class=datarow0"
	End If
loop
OrderRe.close

%>
<input type=hidden name="Remove" id="Remove" value="<%=Remove%>">
<input type=hidden name="RemoveID" id="RemoveID" value="<%=RemoveID%>">
<input type=hidden name="bisppal"  value="<%=bisppal%>">

</table>
<br>
<%


Call showerr(errmsg)
If RemoveID=""   Then
Else
	%><font color="red"><B>Are you sure you want to delete Order '<%=Remove%>' ?</B></font><br><br> 
	<INPUT TYPE="submit" name="condelete" value="Confirm to delete" title="Confirm to delete this order">
	<INPUT TYPE="submit" name="nodelete"  value="Cancel"			title="Cancel this request"><br><br><%

End if
oConn.close
set oConn=Nothing
%>
</FORM>
</BODY>
</HTML>

<script language=Javascript>


function doSubmit(i)
{
var fs=document.forms.form1;
if (i=='1')
{

 
fs.action='Step2.asp';

fs.submit();
return false;
}


if (i=='2')
{


fs.action='';
document.getElementById("WAddtoOrder").value='AddtoOrder';

fs.submit();
return false;
}

if (i=='3')
{


fs.action='';
fs.submit();
return false;
}

if (i=='4')
{
if (document.getElementById("search").value==''
)
{
document.getElementById("search").value='?';

}

fs.action='';
fs.submit();
return false;
}
}


function doRemove(id,id2)
{

var fs=document.forms.form1; 
fs.action='';

document.getElementById("Remove").value=id2;
document.getElementById("RemoveID").value=id;

fs.submit();
return false;


}


</SCRIPT>


