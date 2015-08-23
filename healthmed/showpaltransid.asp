<%@ language=vbscript %>

<% 
pgTitle="View  Pay Pal Transaction "
%>
<!--#include file=header.asp-->


<%
paypalid=Request.QueryString("paypalid")

paypalidc=Request.QueryString("paypalidc")
bisppal=Request.QueryString("bisppal")

base="http://healthmed.asp.ca"

msgpaltrans=""

msgpaltrans=""

    url=base&"/ppal/transactiondetails.asp?transactionID="&paypalid&"&doc=1&bisppal="&bisppal


    msgpaltrans=getWeb("get",url,"")

    url=base&"/ppal/transactiondetails.asp?transactionID="&paypalidc&"&doc=2&bisppal="&bisppal

    msgpaltransc=getWeb("get",url,"")

If msgpaltrans <> "" Then

%>
<table>
<tr>
<td><%=msgpaltrans%>
</td>
</tr>
</table>
<%

End If


If msgpaltransc <> "" Then

%>
<table>
<tr>
<td><%=msgpaltransc%>
</td>
</tr>
</table>
<%

End If

%>




<!--#include file=footer.asp-->
