<%
'----------------------------------------------------------------------------------
' PayPal Constants File
' ====================================
' Constants descriptions for EWP
'----------------------------------------------------------------------------------
	
		 '************************************************************************                                                                         
		 ' WARNING: Do not embed plaintext credentials in your application code.   
		 ' Doing so is insecure and against best practices.                        
		 '                                                                         
		 ' Your EWP Certificate credentials must be handled securely. Please consider          
		 ' encrypting them for use in any production environment, and ensure       
		 ' that only authorized individuals may view or modify them.               
		 ' ************************************************************************                                                                         
		 
		Const  BUY_NOW_PARAMS_SESSION_KEY = "buyNowParams"
		'this session key is used to keep BuyNow Parameters through out the user session.
		
		Const  STANDARD_IDENTITY_TOKEN = "*******"
		
		 ' This identity token is for the default account. You need to pass this identity token along with the transaction token 
		 ' to PayPal in order to receive information that confirms that a payment is complete.
		 ' 
		 ' Follow the steps given below to enable PDT and to generate Identity token for your account.
		 ' 1. Login to your PayPal Account
		 ' 2. Click the My Account tab
		 ' 3. Click the Profile subtab
		 ' 4. Click the Website Payment Preferences link
		 ' 5. Click the Payment Data Transfer On radio button
		 ' once you have enabled PDT, page displays identity token for your account below the radio button.
		 

		Const  STANDARD_EMAIL_ADDRESS = "sdk-seller@sdk.com"
		'The above default account we provide with this sample. You can use the above account
		' for testing purpose, or you can create your own paypal account on sandbox for testing.
		 

		Const  PAYPAL_URL = "https://www.sandbox.paypal.com"
		'The endpiont to where your buyers are directed to pay.  
		 ' Possible endpoints are:
				'the paypal sandbox at https://www.sandbox.paypal.com,
				'the paypal beta sandbox at https://www.beta-sandbox.paypal.com
				'the paypal production site at https://www.paypal.com
		 ' Each endpoint requires a different set of credentials.  
		 ' We provide you a default credentials for the sandbox in this constants file
		 
		
		Const  WEBSCR_URL = "https://www.sandbox.paypal.com/cgi-bin/webscr"
		'URL path appended to the endpoint.  Do not change this.
		
		Const  CERT_ID = "****************"
		'Certificate ID of the EWP Certificate
		
		Const  PAYPAL_IPN_LOG = "paypal_ipn_log.txt"
		'The name of the log file where IPNs are stored.  
%>
	

	
