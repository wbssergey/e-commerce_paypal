<%
'----------------------------------------------------------------------------------
' PayPal DIMants File
' ====================================
' Authentication Credentials for making the call to the server
'----------------------------------------------------------------------------------
    DIM API_ENDPOINT	
	DIM API_USERNAME	
	DIM API_PASSWORD	
    DIM API_SIGNATURE	
	DIM PAYPAL_EC_URL	
	DIM ECURLLOGIN 


	DIM API_VERSION	
	
	DIM HTTPREQUEST_PROXYSETTING_SERVER 
	DIM HTTPREQUEST_PROXYSETTING_PORT 
	DIM USE_PROXY 


If  (Request.Form("bisppal")  = "") And (Request.QueryString ("bisppal") = "" )then
'Signature Security
	 API_ENDPOINT	= "https://api-3t.sandbox.paypal.com/nvp"
	 API_USERNAME	= "sdk-three_api1.sdk.com"
	 API_PASSWORD	= "QFZCWN5HZM8VBG7Q"
     API_SIGNATURE	= "A.d9eRKfd1yVkRrtmMfCFLTqa6M9AyodL0SJkhYztxUi8W9pCXF6.4NI"
	 PAYPAL_EC_URL	= "https://www.sandbox.paypal.com/webscr"
	 ECURLLOGIN = "https://developer.paypal.com"

	 API_VERSION	= "53.0"
	
	 HTTPREQUEST_PROXYSETTING_SERVER = ""
	 HTTPREQUEST_PROXYSETTING_PORT = ""
	 USE_PROXY = False

Else
'Certificate
'Credential:  API Certificate 
'Registrant:  Jeff Mason
'Wiztel USA Inc., 
'Wilmington, DE US 
'API Username:  paypal_api1.wiztel.ca 
'API Password:  Y4MZQ47XYEE4LS4A 
'Request Date:  Sep. 2, 2008 21:31:06 EDT 
 
	 API_ENDPOINT	= "https://api.paypal.com/nvp"
	 API_USERNAME	= "paypal_api1.wiztel.ca"   
	 API_PASSWORD	= "Y4MZQ47XYEE4LS4A"    
     API_SIGNATURE	= "A.d9eRKfd1yVkRrtmMfCFLTqa6M9AyodL0SJkhYztxUi8W9pCXF6.4NI"
	  PAYPAL_EC_URL	= "https://www.sandbox.paypal.com/webscr"
	 ECURLLOGIN = "https://www.paypal.com"

	 
'	 PAYPAL_EC_URL	= "https://www.paypal.com/webscr"
'	 ECURLLOGIN = "https://www.paypal.com"

	' API_ENDPOINT	= "https://api-3t.sandbox.paypal.com/nvp"
	' API_USERNAME	= "sdk-three_api1.sdk.com"
	' API_PASSWORD	= "QFZCWN5HZM8VBG7Q"
    ' API_SIGNATURE	= "A.d9eRKfd1yVkRrtmMfCFLTqa6M9AyodL0SJkhYztxUi8W9pCXF6.4NI"
	' PAYPAL_EC_URL	= "https://www.sandbox.paypal.com/webscr"
	' ECURLLOGIN = "https://developer.paypal.com"

	 API_VERSION	= "53.0"
	
	 HTTPREQUEST_PROXYSETTING_SERVER = ""
	 HTTPREQUEST_PROXYSETTING_PORT = ""
	 USE_PROXY = False

End if

%>
