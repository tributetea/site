<%@ Language=VBScript %>
<%	option explicit 
	Response.Buffer = True
	'@APPVERSION: 50.4011.0.2
set callback = Server.CreateObject("WorldPay.COMcallback")
callback.processCallback()

If callback.hadError() then
	
	Response.Write("<UL>")

	while callback.hasMoreErrors()
		Response.Write("<LI>" & callback.getNextError() & "</LI>")
	wend
	Response.Write("</UL>")

	Response.End

End If
If callback.didTransSuc then

Session("SessionID") = callback.getCartId()

wpresponse = "TransID="&callback.getTransId()&"&RawAuthMsg="&callback.getRawAuthMessage()&_
"&RawAuthCode="&callback.getRawAuthCode()&"&TransTime="&callback.getTransTime()&_
"&InstID="&callback.getInstallationId()&"&Company="&callback.getCompanyName()&_
"&AuthMode="&callback.getAuthMode()&"&Amount="&callback.getAmount()&_
"&CurrencyISO="&callback.getCurrencyISOCode()&"&AmountString="&callback.getAmountString()&_
"&Description="&callback.getDescription()&"&CustName="&callback.getName()&_
"&CustomerAddress="&callback.getAddress()&"&sCustZip="&callback.getPostalCode()&_
"&CustomerCountryISO="&callback.getCountryISOCode()&"&sCustPhone="&callback.getTelephone()&_
"&CustomerFax="&callback.getFax()&"&CustomerEmail="&callback.getEmail()&_
"&Auth="&callback.isAuth()&"&wpresponse=1"&_
"&AuthID="&callback.didTransSuc()&"&iCustID="&Trim(Request.Cookies("sfCustomer")("custID"))&_
"&sCustFirstName="&callback.getParameterString("M_sFName")&_
"&sCustLastName="&callback.getParameterString("M_sLName")&_
"&sCustCompany="&callback.getParameterString("M_sCompany")&_
"&sCustAddress1="&callback.getParameterString("M_sAddress1")&_
"&sCustAddress2="&callback.getParameterString("M_sAddress2")&_
"&sCustCity="&callback.getParameterString("M_sCity")&_
"&sCustState="&callback.getParameterString("M_sState")&_
"&sCustCountry="&callback.getParameterString("M_sCountry")&_
"&sCustFax="&callback.getParameterString("M_sFax")&_
"&sShipCustName="&callback.getParameterString("M_sShipName")&_
"&sShipCustCompany="&callback.getParameterString("M_sShipCompany")&_
"&sShipCustAddress="&callback.getParameterString("M_sShipAddress")&_
"&sShipCustCity="&callback.getParameterString("M_sShipCity")&_
"&sShipCustState="&callback.getParameterString("M_sShipState")&_
"&sShipCustCountry="&callback.getParameterString("M_sShipCountry")&_
"&sShipCustZip="&callback.getParameterString("M_sShipZip")&_
"&sShipCustPhone="&callback.getParameterString("M_sShipPhone")&_
"&ShipMethod="&callback.getParameterString("M_ShipMethod")&_
"&bPremiumShipping="&callback.getParameterString("M_bPremiumShipping")&_
"&sShipInstructions="&callback.getParameterString("M_sShipInstructions")&_
"&iShipID="&callback.getParameterString("M_iShipID")
'response.write "sReturn: " & sReturn
'response.end
merchURL = callback.getParameterString("M_merchURL")

merchURL = split(merchURL,"|")
protocal = merchURL(0)
sReturn	= merchURL(1)
If protocal = "0" Then
protocal = "http://"
ElseIf protocal ="1" Then
protocal = "https://"
End If

	Response.Write "<SCRIPT LANGUAGE=javascript>" & vbCrlf
	Response.Write "<!--" & vbCrlf
	Response.Write " window.location =" & Chr(34) & protocal&sReturn&"?"&wpresponse  & chr(34) &   vbcrlf
  	Response.Write "//-->" & vbcrlf
	Response.Write "</SCRIPT>"
	Response.end






Dim ProcErrMsg,  ProcResponse, iProcResponse, ProcMerchNumber, iTransactionID, ProcRefCode, ProcAvsCode, ProcAvsMsg

	If UCase(AuthID) = "FALSE" Then 
		ProcErrMsg = "This transaction has failed.  Please re-try  your payment"
	ElseIf UCase(AuthID) = "TRUE" Then
		ProcResponse = RawAuthMsg
		ProcMerchNumber = InstID
		iTransactionID = TransID
		ProcRefCode = RawAuthCode		
		ProcAvsCode = "not applicable"
		ProcAvsMsg	= "not applicable"
	End If
	' Write to response table	
	'Call setResponse("WorldPay",iOrderID,iTransactionID,ProcMerchNumber ,ProcAvsCode,ProcAVSMsg,ProcResponse,ProcRefCode,"",ProcErrMsg,iProcResponse)	
	WorldPay = ProcErrMsg
	End If
	%>








