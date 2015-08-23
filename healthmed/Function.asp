<%
'this file contaions alll the functions for this website
'-------------validate string by RegularExpression
Function Validate(str,pp)
	StringToSearch=str
	Set RegularExpressionObject = New RegExp
	With RegularExpressionObject
	.Pattern = pp
	.IgnoreCase = True
	.Global = True
	End With
	expressionmatch = RegularExpressionObject.Test(StringToSearch)
	set RegularExpressionObject = Nothing
	Validate=expressionmatch
End Function

'----------------------- is a string number?---------------
function isNumberm(vString)
    	
    	if not isnumeric(vString) then
    		isNumberm = false
    	else
    		
    	isNumberm = false
    	lCard=len(vString)
    	lC=right(vString,1)
    	cStat=0
    	for i=(lCard) to 1 step -1
    	tempChar= mid(vString,i,1)
    	tmp = InStr("0123456798",tempChar)
    		
    	if InStr("0123456798",tempChar) > 0 then
    		cStat = cStat + 1
    	end if
    	next
    	if lCard = cStat then isNumberm = true
    	
    	end if	

end Function
'-----------------------------show a list of orders---------------------------
Function Showsessionorder(removeornot,ShippingCost,edit,TaxRate,ApplyTax,width)
		'if removeornot=1 then show remove;if i=2 then no 'remove'
		'if edit=1 then show edit link
		Dim ProductCount
		ProductCount=1

		TotalPrice=0
		TaxRate=CDbl(TaxRate)
		%>

		<table><tr><td><b>Order Information</b><br>&nbsp;</td></tr>
		<%If removeornot=1 Or removeornot=2 Or removeornot=3 then%>
		<tr><td>&nbsp;Order Id:<INPUT TYPE="text" NAME="MyOrderID" value="<%=MyOrderID%>"><%Call star(ErrorID,"MyOrderID")%></td></tr>
		<%End if%></table>
		<table width=<%=width%>px>

		<tr ><th >#</th><th align=cente>Product</th><th align=center>Size</th><th align=center>Unit Price</th><th align=center>Quantity</th><th align=center>Extended</th>
		<%If removeornot=1 then%><th></th>
		<%End if%></tr>
		<%Orderitems=Split(OrderList,"&&")

		For Each orderitem In Orderitems

		order=Split(orderitem,"||")
		i=1
		for each ii In order

		If i=1 Then
		listProductid=ii
		End If
		If i=2 Then
		ListUnitQuantity=ii
		End If
		If i=3 Then
		Listunitprice=ii
		End If
		If i=4 Then
		ListQuantity=ii
		End If

		If i=5 Then
		ListProductName=ii
		End If
		If i=6 Then
		ListOrderID=ii
		End If
		If i=7 Then
		ListAlterUnitPrice=ii
		End If

		i=i+1
		next
		If Len(ListUnitQuantity)=1 Then
			If ListUnitQuantity="0" Then
				ListUnitQuantity=""
			End if

		End if
		%><tr><td><%=ProductCount%></td><td><%=ListProductName%></td><td align=center><%=ListUnitQuantity%></td><td align=right>$<%=FormatNumber(ListAlterUnitPrice,2)%></td><td align=center><%=ListQuantity%> </td>

		<td align=right>$<%=FormatNumber(ListAlterUnitPrice*ListQuantity,2)%></td>
		<%If removeornot=1 then%>
		<td align=center><a  href="" onclick=" doRemove('<%=ListOrderID%>');return false;" title="Remove this product">Remove</a></td>
		<%End if%>
		</tr><%
		If TotalPrice="" Or IsNull(TotalPrice) Then
		TotalPrice=ListAlterUnitPrice*ListQuantity*(1+TaxRate*0.01)
		Else
		TotalPrice=TotalPrice+(ListAlterUnitPrice*ListQuantity*(1+CDbl(TaxRate)*0.01))
		End if
		ProductCount=ProductCount+1
		Next
		If ApplyTax="ApplyTax" Then
		TotalPrice=TotalPrice+ShippingCost*(1+TaxRate*0.010)
		Else

		TotalPrice=TotalPrice+ShippingCost
		End if

		%>
		<tr><td>&nbsp;</td></tr>
		<tr><td colspan=5 align=right>Shipping Cost:</td><td align=right>$

		<%If ( removeornot=2 And edit=0 ) Or removeornot=3 Then
		response.write FormatNumber(ShippingCost,2)
		Else
		%>
		<INPUT TYPE="text" NAME="ShippingCost" id="ShippingCost" value="<%=FormatNumber(ShippingCost,2)%>" size=4 width=4>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<%End if%></td></tr>
		<tr><td colspan=3 ></td><td align=right> </td ><td colspan=2><%
		If edit=1 Then 
			%><INPUT TYPE="checkbox" NAME="ApplyTax" value="ApplyTax" <%If ApplyTax="ApplyTax"Then %>checked<%End if%> onclick="submit();"  >Apply Shipping Tax<%
			Else
			If ApplyTax="ApplyTax" Then %>
			Shipping Tax: Yes<%
			Else
			%>Shipping Tax: No<%
			End if
		End If
		%></td></tr>
		<tr><td colspan=3></td><td colspan=2 align=right>Tax Rate:</td><td align=right>
		<%If removeornot=1 Then%>
		<INPUT TYPE="text" NAME="TaxRate" value="<%=FormatNumber(TaxRate,2)%>" size=4 width=4>
		<%else%>
		<%=FormatNumber(TaxRate,2)%>
		<%End if%>
		%</td></tr>
		<tr><td colspan=3><%If removeornot=2 Or removeornot=3 then%>
		<INPUT TYPE="submit" value="Modify Orders" onclick="doSubmittostep1('1');return false;">
		<%End if%> </td><td colspan=2 align=right><h4>Total Price:</h4></td><td align=right ><h4>$<%=FormatNumber(TotalPrice,2)%></h4></td></tr>
		</table>
		<%
		Showsessionorder=TotalPrice
		End Function

		Function ShowAffiliate()


		%><table>
		<tr> <td style=letter-spacing:1;font-size:10pt;font-weight:bold;color:#646496>Affiliate:</td>
		<td>
		<select name="Afficiate" >
		<%
		affiq="SELECT AffiliateID,AffiliateName FROM TblAffiliates"
		Set addiqR=oConn.execute(affiq)

		Do Until  addiqR.eof
		%><option value="<%=addiqR("AffiliateID")%>" 
		<%If CInt(addiqR("AffiliateID"))= CInt(Afficiate) Then
		response.write "selected" 
		End if%>><%=addiqR("AffiliateName")%></option><%
		addiqR.movenext
		Loop

		addiqR.close
		%>

		</select>
		</td>
		<td><a href="step1.asp" title="Add a new Affiliate." onclick="wantwindows();return false;">Add...</a></td>
		</tr></table>
		<script language=Javascript>
		function wantwindows()
		{
		var winref=window.open('Addnewaffiliate.asp','mywindow','status=1,toolbar=0,width=620');
		while (winref.open)
		{

		}

		}

		</script>
		<%
End Function


'------------------------
Function showerr(Errmsg)
	If Errmsg="" Then
	Else
	%>
	<h4><font color="red"><%
	response.write Errmsg&"<br><br>"&"</font></h4>"
	End If
End Function


Function star(errorid,errorid1)
	If CStr(errorid)=CStr(errorid1) Then
	%><font style="color:red">*</font><%End if
End Function

function selected(str1,str2)
	if lcase(str1&"")=lcase(str2&"") and (len(str1)>0 or len(str2)>0) then selected=" selected " else selected=""
end Function

'------------------
Function menu()

	%>
	<table border="1" cellspacing="0"  cellpadding="10px" height=600px width=600px>
	<tr><td width=150px valign="top" nowrap>
	[<a href="<%=strm%>default.asp" >Home</a>]<br>
	[<a href="<%=strm%>step1.asp" onclick="cleansession();return false">Create New Order</a>]<br>
	[<a href="<%=strm%>showorders.asp" >Show Order List</a>]<br>
	[<a href="<%=strm%>transresult.asp" >Enter Transaction Result</a>]<br>
	[<a href="<%=strm%>Shippinginfo.asp" >Enter Shipping Info</a>]<br>
	


	<br>[<a href="" onclick="dologout('1');return false;">Log out</a>]
	<INPUT TYPE="hidden" name="mylogout" id="mylogout" value="">

	</td><td valign="top" width=auto>
	<script language=Javascript>

	function dologout(i)
	{
	var fs=document.forms.form1;
	document.getElementById("mylogout").value='mylogout';
	fs.action='';
	fs.submit();
	return false;

	}
	function cleansession()
	{
	var fs=document.forms.form1;
	document.getElementById("mylogout").value='cleansession';
	fs.action='step1.asp';
	fs.submit();
	return false;
	}

	</script>
<%End Function 




'-----------------------------show a list of orders---------------------------
Function GetExtended(OrderLista)
	'if removeornot=1 then show remove;if i=2 then no 'remove'
	'if edit=1 then show edit link
	Dim Orderitems
	Dim orderitem
	Dim order
	Dim i
	Dim ii
	Dim listProductid
	Dim ListUnitQuantity
	Dim Listunitprice
	Dim ListQuantity
	Dim ListProductName
	Dim ListOrderID
	Dim ListAlterUnitPrice
	Dim TotalPrice
	Orderitems=Split(OrderLista,"&&")

	For Each orderitem In Orderitems

	order=Split(orderitem,"||")
	i=1
	for each ii In order

	If i=1 Then
	listProductid=ii
	End If
	If i=2 Then
	ListUnitQuantity=ii
	End If
	If i=3 Then
	Listunitprice=ii
	End If
	If i=4 Then
	ListQuantity=ii
	End If

	If i=5 Then
	ListProductName=ii
	End If
	If i=6 Then
	ListOrderID=ii
	End If
	If i=7 Then
	ListAlterUnitPrice=ii
	End If

	i=i+1
	next
	If Len(ListUnitQuantity)=1 Then
		If ListUnitQuantity="0" Then
			ListUnitQuantity=""
		End if

	End if
	If TotalPrice="" Or IsNull(TotalPrice) Then
	TotalPrice=ListAlterUnitPrice
	Else
	TotalPrice=TotalPrice+(ListAlterUnitPrice*ListQuantity)
	End if
	Next
	GetExtended=TotalPrice
End Function


Function GetTotalPrice(OrderLista,TaxRate,ApplyTax)
	TaxRate=CDbl(TaxRate)
	TotalPrice= GetExtended(OrderLista)
	If ApplyTax="ApplyTax" Then
	TotalPrice=TotalPrice*(1+TaxRate*0.01)+ShippingCost*(1+TaxRate*0.01)
	Else

	TotalPrice=TotalPrice*(1+TaxRate*0.01)+ShippingCost
	End if

	GetTotalPrice=TotalPrice
	End Function

	Function GetTotaltax(extended, shippingCost,ApplyTax,TaxRate)
	TaxRate=CDbl(TaxRate)
	If ApplyTax="ApplyTax" Then
	GetTotaltax=extended*TaxRate*0.01+ShippingCost*(TaxRate*0.01)
	Else

	GetTotaltax=extended*TaxRate*0.01
	End if

End Function




Function recommendedID()
		orderIDCheckQ="select max(MyOrderID) as yy from TblOrder "
		Set orderIDCheckR=oConn.execute(orderIDCheckQ)
		If Not orderIDCheckR.eof Then
		mycount=orderIDCheckR("yy")
		End If
		orderIDCheckR.close
		intension=""
		recommendedID=CLng(mycount)+1
End Function

function getWeb(rqMthd,rqUrl,rqData) 'works
	dim h: if not isObject(oWeb) then set oWeb=server.createobject("winhttp.winhttprequest.5.1")	'initialize browser
	oWeb.settimeouts 3*1000,3*1000,3*1000,7*1000	' resolve,connect,send,receive :1*1000=1sec
	oWeb.open ucase(rqMthd),rqUrl,false				' handshake web server
	'Response.write "<BR><B>"&ucase(rqMthd)& "    " & rqUrl & "     " & rqData & "</B><BR>"
	oWeb.option(0)="Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)"		' useragent: client browser identification sent to web server !
	if lcase(rqMthd&"")="post" then oWeb.setrequestheader "Content-Type","application/x-www-form-urlencoded"' important when submiting forms
	on error resume next: oWeb.send rqData: oWeb.waitforresponse: h=oWeb.responsetext	' send request, catch timeout err
		if err then h="Error connecting to web server:"&vbcrlf &err.description &vbcrlf &"Html Response:" & h: err.clear
	on error goto 0: getWeb=h
end function


function debugErr(errmsg)
Dim s

s=Request.Cookies("userUnixName")

If True then  (s = "sergey") Or ( s = "winston")  then

	if isObject(oConn) then 
	oConn.close
	set oConn=nothing
	End If
	
	response.write "<br><br>"&errmsg
	response.end

End If

end Function


Function wx(errm)
If validateuser="winston" or validateuser="sergey"  Then

response.write errm
End if

end Function
%>
