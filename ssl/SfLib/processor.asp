<%

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'	The following constant, proc_live, controls the status of 
'	all payment processors supported:
'
'		Const proc_live = 1 is 'live' mode
'		Const proc_live = 0 is 'test' mode
'
'	Before your store goes live, you are encouraged to run 
'	an order through with the constant set to 0 for testing.	
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@BEGINVERSIONINFO

'@APPVERSION: 50.4011.0.6

'@FILENAME: processor.asp
	 


'@DESCRIPTION: Processes orders based on payment types

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000,2001 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

' #321 - MS

'@ENDVERSIONINFO
Function getProcStat(vData)
Dim tSQL, rsProcStatus, proc_status
  tSQL = "SELECT trnsmthTestServerPath FROM sfTransactionMethods WHERE  trnsmthdID = " & vdata
   Set rsProcStatus = cnn.execute (tSQL) 
getProcStat = Trim(rsProcStatus("trnsmthTestServerPath")) 

closeobj(rsProcStatus) 
 
End Function
'-------------------------------------------------------------------
' CyberCash subroutine 
' Requirement: CYCHMCK.DLL 2.0 or higher
' Last edited : October 3, 2000
'-------------------------------------------------------------------

Function CyberCash(proc_live)

	Dim Config, ccInput, Output, SocketObject, strMessage, CCID, MERCHANT_KEY, iProcResponse, ProcErrMsg
	Dim ProcCustNumber, ProcAddlData, ProcRefCode, ProcAuthCode, ProcMerchNumber, ProcActionCode, ProcErrLoc
	Dim ProcErrCode, ProcAvsCode, ProcAVSMsg

	Set Config = Server.CreateObject("CyberCashMCK.MessageBlock")
	Set ccInput = Server.CreateObject("CyberCashMCK.MessageBlock")
	Set Output = Server.CreateObject("CyberCashMCK.MessageBlock")

	ProcErrMsg = ""
	strMessage = "m" & sMercType
	CCID = trim(sLogin)
	MERCHANT_KEY = trim(sPassword)
	
	If trim(sPaymentServer) = "" or isNull(sPaymentServer) Then 
		sPaymentServer = "http://cr.cybercash.com/cgi-bin/cr21api.cgi/" 	
	End If	
	
	'Provide the config parameters required for CyberCash Transaction processing
	Config.Add "CYBERCASH_ID", CCID
	Config.Add "MERCHANT_KEY", MERCHANT_KEY
	Config.Add "CCPS_HOST", sPaymentServer & trim(strMessage)

	if len(trim(sCustCardExpiry)) > 5 then
	sCustCardExpiry= left(sCustCardExpiry,3) & right(sCustCardExpiry,2)
	end if

	'Provide the input parameters required for the CyberCash message (see developers guide)
	ccInput.Add "card-number",  sCustCardNumber
	ccInput.Add "card-exp",     sCustCardExpiry
	ccInput.Add "card-name",    sCustCardName
	ccInput.Add "card-Address", sCustAddress1 & "," & sCustAddress2
	ccInput.Add "card-city",    sCustCity
	ccInput.Add "card-state",   sCustState
	ccInput.Add "card-zip",     sCustZip
	ccInput.Add "card-country", sCustCountry
	ccInput.Add "order-id",     iOrderID
	ccInput.Add "amount",       "usd " & REPLACE(sGrandTotal,",","")

	Set SocketObject = Server.CreateObject("CyberCashMCK.socket.1")
	Set Output = SocketObject.SendMessageBlock(Config, ccInput)
		ProcResponse = Output.Item("MStatus") 'Approved/declined
			
		If ProcResponse = "success" Then
		 ProcResponse = "approved"
		 iProcResponse = 1
		Else
		 ProcResponse = "failed"
		 iProcResponse = 0
		End If
		
		ProcMessage		= Replace(Output.Item("aux-msg"),"'","''") 'Detailed Info
		ProcCustNumber	= Replace(Output.Item("cust-txn"),"'","''") 'Trans Number
		ProcAddlData	= Replace(Output.Item("addnl-response-data"),"'","''") ' AdditionalData
		ProcRefCode		= Replace(Output.Item("ref-code"),"'","''") 
		ProcAuthCode	= Replace(Output.Item("auth-code"),"'","''") 'Authorization Code
		ProcMerchNumber	= Replace(Output.Item("merch-txn"),"'","''") 'Merch Trans Number
		ProcActionCode	= Replace(Output.Item("action-code"),"'","''") 
		ProcErrMsg		= Replace(Output.Item("MErrMsg"),"'","''") 'Detailed info on failure
		ProcErrLoc		= Replace(Output.Item("MErrLoc"),"'","''") 'Location error occured
		ProcErrCode		= Replace(Output.Item("MErrCode"),"'","''") 'CyberCash error codes
		ProcAvsCode		= Replace(Output.Item("avs-code"),"'","''") 'CyberCash AVS code
		ProcAVSMsg			= AVSMsg(ProcAvsCode)

		Call setResponse("cybercash",iOrderID,ProcCustNumber,ProcMerchNumber,ProcAvsCode,ProcAVSMsg,ProcErrCode,ProcAuthCode,ProcRefCode,ProcErrMsg,iProcResponse)
		Set Config = Nothing
		Set ccInput = Nothing
		Set Output = Nothing
		CyberCash = ProcErrMsg
End Function

'---------------------------------------------------------------------
'   AuthorizeNet Send Sub-Routine 
'	Com Object 1.0
'---------------------------------------------------------------------
Function AuthNet(proc_live,iType)

	Dim AuthObj, sCustAddress, sShipCustAddress, strPost, MType, iResponseCode, iProcResponse, sResponseReasonText, sFailedReason, sErrorMessage, iAVSCode, iTransactionID, iAuthCode, GrandTotal, sAVSMsg, sTstRqst, sTax
	Dim ProcCustNumber, ProcAddlData, ProcRefCode, ProcAuthCode, ProcMerchNumber, ProcActionCode, ProcErrLoc
	Dim ProcErrCode, ProcAvsCode, ProcAVSMsg

	Set AuthObj = Server.CreateObject("AuthNetSSLConnect.SSLPost")  

	If sMercType = "authonly" Then MType = "AUTH_ONLY" 
	If sMercType = "authcapture" Then MType = "AUTH_CAPTURE" 
		
	sCustCardName = sCustCardName
	iOrderID =iOrderID
	sShipInstructions = sShipInstructions
    
    'Response.Write "proc_live = "& proc_live 
	If proc_live = 1 Then
		sTstRqst = "false"	
	ElseIf proc_live = 0 Then
		sTstRqst = "true"	
	End If	
	
	' Customer Info
	sCustFirstName = sCustFirstName
	sCustLastName = sCustLastName
	sCustName		  = sCustName
	sCustCompany	= sCustCompany
	sCustAddress1	= sCustAddress1
	sCustAddress2	= sCustAddress2
	sCustAddress	= sCustAddress1 & ";" & sCustAddress2
	sCustCity		= sCustCity
	sCustState		= sCustState
	sCustCountry	= sCustCountry
	sCustZip		= sCustZip
	sCustPhone		= sCustPhone
	sCustFax		= sCustFax
	sCustEmail		= sCustEmail
	
	' Payment variables
	sPaymentMethod			= sPaymentMethod
	sCustCardType			= sCustCardType
	sCustCardNumber			= Replace(sCustCardNumber," ","")
	sCustCardNumber			= Replace(sCustCardNumber,"-","")
	sCustCardExpiry			= Replace(sCustCardExpiry,"/","")
	'sCustCardExpiry =      Right(sCustCardExpiry,2)
	sBankName				= sBankName
	iRoutingNumber			= iRoutingNumber
	iCheckingAccountNumber	= iCheckingAccountNumber
	iPONumber				= iPONumber
	iCheckNumber			= iCheckNumber

	' Shipping Variables
	sShipCustFirstName 	= sShipCustFirstName
	sShipCustLastName 	= sShipCustLastName
	sShipCustName		= sShipCustName
	sShipCustCompany	= sShipCustCompany
	sShipCustAddress1	= sShipCustAddress1
	sShipCustAddress2	= sShipCustAddress2
	sShipCustAddress	= sShipCustAddress1 & ";" & sShipCustAddress2
	sShipCustCity		= sShipCustCity
	sShipCustState		= sShipCustState
	sShipCustCountry	= sShipCustCountry
	sShipCustZip		= sShipCustZip
	sShipCustPhone		= sShipCustPhone
	sLogin				= sLogin
	iShipMethod			= iShipMethod
	sShipping			= (cDbl(sHandling) + cDbl(sShipping))
	sTax 				= (cDbl(iSTax) + cDbl(iCTax))
	iCustID				= Request.Cookies("sfCustomer")("custID")	
	GrandTotal			= REPLACE(sGrandTotal,",","")

	' Post String
	If iType = 1 Then
		strPost = "x_Login="&sLogin&",x_Amount="&sGrandTotal&",x_Freight="&sShipping&",x_Tax="&sTax&",x_Card_Num="&sCustCardNumber&",x_Exp_Date="&sCustCardExpiry&",x_Password="&sPassword&",x_Method=CC,x_Type="&MType&",x_Cust_ID="&iCustID&",x_Test_Request="&sTstRqst&",x_Cust_ID="&iCustID&",x_Description=StoreFront Web Store Order,x_First_Name="&sCustFirstName&",x_Last_Name="&sCustLastName&",x_Company="&sCustCompany&",x_Address="&sCustAddress&",x_City="&sCustCity&",x_State="&sCustState&",x_Zip="&sCustZip&",x_Country="&sCustCountry&",x_Phone="&sCustPhone&",x_Fax="&sCustFax&",x_Ship_To_First_Name="&sShipCustFirstName&",x_Ship_To_Last_Name="&sShipCustLastName&",x_Ship_To_Company="&sShipCustCompany&",x_Ship_To_Address="&sShipCustAddress&",x_Ship_To_City="&sShipCustCity&",x_Ship_To_State="&sShipCustState&",x_Ship_To_Zip="&sShipCustZip&",x_Ship_To_Country="&sShipCustCountry&",x_Email="&sCustEmail						
	ElseIf iType = 2 Then
		strPost = "x_Login="&sLogin&",x_Amount="&sGrandTotal&",x_Freight="&sShipping&",x_Tax="&sTax&"x_Card_Num="&sCustCardNumber&",x_Exp_Date="&sCustCardExpiry&",x_Password="&sPassword & ",x_Method=ECHECK,x_Type=" & sMercType & ",x_Bank_Name=" & sBankName & ",x_Bank_ABA_Code=" & iRoutingNumber & ",x_Bank_Acct_Num=" & iCheckingAccountNumber & ",x_Test_Request=" & sTstRqst & ",x_Cust_ID="&iCustID&",x_Description=StoreFront Web Store Order,x_First_Name="&sCustFirstName&",x_Last_Name="&sCustLastName&",x_Company="&sCustCompany&",x_Address="&sCustAddress&",x_City="&sCustCity&",x_State="&sCustState&",x_Zip="&sCustZip&",x_Country="&sCustCountry&",x_Phone="&sCustPhone&",x_Fax="&sCustFax&",x_Ship_To_First_Name="&sShipCustFirstName&",x­_Ship_To_Last_Name="&sShipCustLastName&",x_Ship_To_Company="&sShipCustCompany&",x_Ship_To_Address="&sShipCustAddress&",x_Ship_To_City="&sShipCustCity&",x_Ship_To_State="&sShipCustState&",x_Ship_To_Zip="&sShipCustZip&",x_Ship_To_Country="&sShipCustCountry&",x_Email="&sCustEmail
	End If
	

	AuthObj.doSSLPost strPost
    sShipping = sShipping - sHandling '#430
	' Error Response

	If AuthObj.ErrorCode = 0 Then	
		If AuthObj.NumFields > 1 Then
				
			ProcActionCode = AuthObj.GetField(1)
			ProcAddlData = AuthObj.GetField(2)
			ProcErrCode = AuthObj.GetField(3)
			ProcMessage = AuthObj.GetField(4)
			ProcAuthCode = AuthObj.GetField(5)
			ProcAvsCode = AuthObj.GetField(6)
			ProcRefCode = AuthObj.GetField(7)

			If ProcActionCode = 1 Then
				iProcResponse = 1				
				
			ElseIf ProcActionCode = 2 Then
				iProcResponse = 0
				sErrorMessage = "The transaction has been declined.  Please check the account and resubmit, try another account or contact the card issuing bank.  Thank You.<br> Additional Information from the Processor: " & ProcMessage
				
			ElseIf ProcActionCode = 3 Then	 
				iProcResponse = 0
				sErrorMessage = "There was an error on the transaction processing network. Please check the account and resubmit the order.<br> Additional Information from the Processor: " & ProcMessage
				
			End If
		Else	
			' No connection
			iProcResponse = 0
			sFailedReason = "Connection Error"	
			sErrorMessage = "No Connection Error"	
		End If
	Else
		Select Case AuthObj.ErrorCode
		Case -1
			sErrorMessage = "A connection could not be established with the authorization network."
		Case -2
			sErrorMessage = "A connection could not be established with the authorization network."
		Case -3
			sErrorMessage = "A connection could not be established with the authorization."
		Case Else 
			sErrorMessage = "An error occured during processing."
		End Select
	End If					
	
	' Write to payment response table
	Call setResponse("AuthorizeNet",iOrderID,"","",ProcAvsCode,ProcMessage,ProcErrCode,ProcAuthCode,ProcRefCode,ProcErrMsg,ProcActionCode)
	
	closeobj(AuthObj)
	AuthNet = sErrorMessage
End Function

'-----------------------------------------------------------------------------
' PSIGate Payment Processor
' COM object integration
'-----------------------------------------------------------------------------
Function PSIGate(proc_live)
	Dim objPsiGate, iChargeType, nRetCode, sApproved, iProcResponse, sErrorMessage, iAuthNo, iTransactionID, sFailedReason, sPemPath
	
	' First set the response to 0
	iProcResponse = 0
	
	' Determine chargetype
	If sMercType = "authonly" Then
		iChargeType = 1
	Elseif sMercType = "authcapture" Then
		iChargeType = 2
	End If	

	If trim(sPaymentServer) = "" or isNull(sPaymentServer) Then	
		sPaymentServer = "secure.psigate.com"
	End If	

	Set objPsiGate = CreateObject("MyServer.PsiGate")
	
		' Get Path to PEM file.
		sPemPath = Server.MapPath(".")
		sPemPath = sPemPath & "\SFLib\" & sPassword  

		' Set up the request object
		objPsiGate.Configfile = sLogin 
		objPsiGate.Keyfile = sPemPath
		objPsiGate.Host = sPaymentServer
		objPsiGate.Port = 1139
			
		If proc_live = 1 Then
			objPsiGate.Result = 0 'live					
		ElseIf proc_live = 0 Then
			objPsiGate.Result = 1  'good
			'objPsiGate.Result = 2 'duplicate
			'objPsiGate.Result = 3 'declined
		End If	
		
		' Required Customer Info
		objPsiGate.Bname	= sCustName
		objPsiGate.Baddr1	= sCustAddress1
		objPsiGate.Baddr2	= sCustAddress2
		objPsiGate.Bcity	= sCustCity
		objPsiGate.Bstate	= sCustState
		objPsiGate.Bzip		= sCustZip
		objPsiGate.Bcountry	= sCustCountry
		objPsiGate.Oid 		= iOrderID
        objPsiGate.Userid       = sCustemail
        objPSIgate.Email = sCustEmail '# 325
		
		' Credit Card Info
		objPsiGate.Cardnumber = sCustCardNumber
		objPsiGate.Chargetype = iChargeType
		objPsiGate.Expmonth	  = sCustCardExpiryMonth
		objPsiGate.Expyear    =  right(trim(sCustCardExpiryYear),2) '#324
	
		' Required Shipping Info
		objPsiGate.Items = 1
		objPsiGate.Carrier = 1

		nRetCode = objPsiGate.AddItem("StoreFront","StoreFront Purchase", Cdbl(sGrandTotal),1,"",0,"")

		If Not nRetCode = 1 Then
			iProcResponse = 0
		ElseIf nRetCode = 1 Then 
			iProcResponse = 1
		End If

		' Send the order to PSiGate
		nRetCode = objPsiGate.ProcessOrder()
		' Error Checking
		If Not nRetCode = 1 Then
			iProcResponse = 0
		ElseIf 	nRetCode = 1 Then 
			iProcResponse = 1	
		End If

		' Get Response
		sApproved			= objPsiGate.Appr
		iAuthNo				= objPsiGate.Code
		iTransactionID		= objPsiGate.RefNo
		sFailedReason		= objPsiGate.Err
		sErrorMessage		= objPsiGate.ErrMsg		
		
		' If error occured
		If trim(sFailedReason) <> "" or trim(sErrorMessage) <> "" Then
			iProcResponse = 0
		End If

	' Write to response table	
	Call setResponse("PSIGate",iOrderID,iTransactionID,"","",sFailedReason,"",iAuthNo,"",sErrorMessage,iProcResponse)
	
	Set objPsiGate = nothing
	PSIGate = sFailedReason & " " &  sErrorMessage	
	
End Function


'------------------------------------------------------------
' LinkPoint
'------------------------------------------------------------
Function LinkPoint(proc_live)

	'----------------------------------------------------------------------
	' Modification for ASP by Dave Lambert, dlambert@infoponic.com
	' Modified for StoreFront by LaGarde Inc.
	' Created from ccapi_error.h from API 3.8
	'----------------------------------------------------------------------
ON ERROR RESUME NEXT
	Const Fail = 0
	Const Succeed = 1

	'Created from ccapi_client.h from API 3.8

	'Request types possible for OrderField_Chargetype
	Const Chargetype_Auth = 0
	Const Chargetype_Sale = 0
	Const Chargetype_Preauth = 1
	Const Chargetype_Postauth = 2
	Const Chargetype_Credit = 3
	Const Chargetype_Error = 0

	'Result types possible for OrderField_Result
	Const Result_Live = 0 
	Const Result_Good = 1
	Const Result_Duplicate = 2
	Const Result_Decline = 3

	'ESD types for ItemField_Esdtype
	Const Esdtype_None = 0
	Const Esdtype_Softgood = 1
	Const Esdtype_Key = 2

	' OrderField_t
	Const OrderField_Oid = 0   
	Const OrderField_Userid = 1
	Const OrderField_Bcompany = 2
	Const OrderField_Bcountry = 3
	Const OrderField_Bname = 4   
	Const OrderField_Baddr1 = 5  
	Const OrderField_Baddr2 = 6
	Const OrderField_Bcity = 7
	Const OrderField_Bstate = 8
	Const OrderField_Bzip = 9
	Const OrderField_Sname = 10
	Const OrderField_Saddr1 = 11
	Const OrderField_Saddr2 = 12
	Const OrderField_Scity = 13
	Const OrderField_Sstate = 14
	Const OrderField_Szip = 15
	Const OrderField_Scountry = 16
	Const OrderField_Phone = 17
	Const OrderField_Fax = 18
	Const OrderField_Refer = 19
	Const OrderField_Shiptype = 20
	Const OrderField_Shipping = 21
	Const OrderField_Tax = 22
	Const OrderField_Subtotal = 23 
	Const OrderField_Vattax = 24
	Const OrderField_Comments = 25
	Const OrderField_MaxItems = 26
	Const OrderField_Email = 27
	Const OrderField_Cardnumber = 28 
	Const OrderField_Expmonth = 29
	Const OrderField_Expyear = 30
	Const OrderField_Chargetype = 31
	Const OrderField_Chargetotal = 32
	Const OrderField_Referencenumber = 33 
	Const OrderField_Result = 34
	Const OrderField_Addrnum = 35
	Const OrderField_Ip = 36

	' Responses 
	Const OrderField_R_Time = 37
	Const OrderField_R_Ref = 38
	Const OrderField_R_Approved = 39
	Const OrderField_R_Code = 40
	Const OrderField_R_Ordernum = 41
	Const OrderField_R_Error = 42
	Const OrderField_R_FraudCode = 43

	' ReqField_t
	Const ReqField_Configfile = 0
	Const ReqField_Keyfile = 1
	Const ReqField_Appname = 2
	Const ReqField_Host = 3
	Const ReqField_Port = 4

	' ItemField_t
	Const ItemField_Itemid = 0   
	Const ItemField_Description = 1
	Const ItemField_Price = 2
	Const ItemField_Quantity = 3
	Const ItemField_Softfile = 4
	Const ItemField_Esdtype = 5
	Const ItemField_Serial = 6
	Const ItemField_MaxOptions = 7

	' ShippingField_t
	Const ShippingField_Country = 0
	Const ShippingField_State = 1
	Const ShippingField_Total = 2
	Const ShippingField_Items = 3
	Const ShippingField_Weight = 4
	Const ShippingField_Carrier = 5
	' Responses 
	Const ShippingField_R_Total = 6

	' TaxField_t
	Const TaxField_State = 0
	Const TaxField_Zip = 1
	Const TaxField_Total = 2
	' Responses 
	Const TaxField_R_Total = 3

	' OptionField_t
	Const OptionField_Option = 0
	Const OptionField_Choice = 1

	Dim total, ApiDriver, OrderCtx, ItemCtx, OptionCtx, ReqCtx, PemPath, Flag, ProcResponse, iProcResponse, ProcErrMessage
	Dim ProcMessage, ProcActionCode, ProcResponseCode, ProcAuthCode, ProcErrMsg, ProcCustNumber
	Dim ProcRefCode, ProcAvsCode, ProcAvsMsg, Result_Type, ProcAuth

	Set ApiDriver = Server.CreateObject("ComApi_3_8.ComApi")
	'Set ApiDriver = Server.CreateObject("ComApi_3_8.ComApi.1")
	
    
	If sPaymentServer = "" or IsNull(sPaymentServer) Then
	'If proc_live = 1 Then
		sPaymentServer = "secure.linkpt.net"
	'	ElseIf proc_live = 0 Then
	'	sPaymentServer = "staging.linkpt.net"
	'	End If
	End If
	OrderCtx	=		ApiDriver.csi_order_alloc()
	ReqCtx		=		ApiDriver.csi_req_alloc()
	' Get Path to PEM file.
	PemPath = Server.MapPath(".")
	PemPath = PemPath & "\SFLib\sf.pem"
	PemPath = Replace(PemPath, "\\", "\")
	
	Flag = ApiDriver.csi_req_set(ReqCtx, ReqField_Configfile, CStr(sLogin))
	Flag = ApiDriver.csi_req_set(ReqCtx, ReqField_Keyfile, CStr(PemPath))
	Flag = ApiDriver.csi_req_set(ReqCtx, ReqField_Host, CStr(sPaymentServer))
	Flag = ApiDriver.csi_req_set(ReqCtx, ReqField_Port, 1139)
	Flag = ApiDriver.csi_order_setrequest(OrderCtx, ReqCtx)

	If ApiDriver.bStat <> Succeed Then
		ProcResponse = "fail"
		iProcResponse = 0
		ProcErrMessage = "Error: " & ApiDriver.csi_util_errorstr(ApiDriver.csi_order_error(OrderCtx))
		'Set ApiDriver = nothing
		'Exit Function
	End If
	
	Dim sAddrNum
	Dim iPosit
	sAddrNum = Trim(sCustAddress1)
	iPosit = InStr(sAddrNum, " ")
    If iPosit > 0 Then
   	  sAddrNum = Left(sAddrNum,iPosit - 1)
   	End If    
	
	If Not IsNumeric(sAddrNum) then
	  	sAddrNum = "0000"
	End If


	' Get expiration date of credit card
	sCustCardExpiry = Replace(sCustCardExpiry,"/","")
	sCustCardExpiry = Replace(sCustCardExpiry,"-","")
	sCustCardExpiry = Replace(sCustCardExpiry," ","")

	Flag = ApiDriver.csi_order_set(OrderCtx, OrderField_Userid, CStr(iOrderID))
	Flag = ApiDriver.csi_order_set(OrderCtx, OrderField_Bname, CStr(sCustCardName))
	Flag = ApiDriver.csi_order_set(OrderCtx, OrderField_Bcompany, CStr(sCustCompany))
	Flag = ApiDriver.csi_order_set(OrderCtx, OrderField_Baddr1, CStr(sCustAddress1))
	
	Flag = ApiDriver.csi_order_set(OrderCtx, OrderField_Bcity, CStr(sCustCity))
	Flag = ApiDriver.csi_order_set(OrderCtx, OrderField_Bstate, CStr(sCustState))
	Flag = ApiDriver.csi_order_set(OrderCtx, OrderField_Bzip, CStr(sCustZip))
	Flag = ApiDriver.csi_order_set(OrderCtx, OrderField_Bcountry, CStr(sCustCountry))
	Flag = ApiDriver.csi_order_set(OrderCtx, OrderField_Sname, CStr(sShipCustName))
	Flag = ApiDriver.csi_order_set(OrderCtx, OrderField_Saddr1, CStr(sShipCustAddress1))
	Flag = ApiDriver.csi_order_set(OrderCtx, OrderField_Scity, CStr(sShipCustCity))
	Flag = ApiDriver.csi_order_set(OrderCtx, OrderField_Sstate, CStr(sShipCustState))
	Flag = ApiDriver.csi_order_set(OrderCtx, OrderField_Szip, CStr(sShipCustZip))
	Flag = ApiDriver.csi_order_set(OrderCtx, OrderField_Scountry, CStr(sShipCustCountry))
	Flag = ApiDriver.csi_order_set(OrderCtx, OrderField_Phone, CStr(sCustPhone))
	Flag = ApiDriver.csi_order_set(OrderCtx, OrderField_Fax, CStr(sCustFax))
	Flag = ApiDriver.csi_order_set(OrderCtx, OrderField_Comments, CStr(sShipInstructions))
	Flag = ApiDriver.csi_order_set(OrderCtx, OrderField_Cardnumber, CStr(sCustCardNumber))
    Flag = ApiDriver.csi_order_set(OrderCtx, OrderField_Addrnum, sAddrNum)
	'Set Flag for Authorization Only or Charge
	If sMercType = "authonly" Then
		Flag = ApiDriver.csi_order_set(OrderCtx, OrderField_ChargeType, Chargetype_Preauth)
	ElseIf sMercType = "authcapture" Then 
		Flag = ApiDriver.csi_order_set(OrderCtx, OrderField_ChargeType, Chargetype_Sale)
	End If		

	Flag = ApiDriver.csi_order_set(OrderCtx, OrderField_Expmonth, CStr(sCustCardExpiryMonth))
	Flag = ApiDriver.csi_order_set(OrderCtx, OrderField_Expyear, CStr(right(sCustCardExpiryYear,2)))
	Flag = ApiDriver.csi_order_set(OrderCtx, OrderField_Email, CStr(sCustEmail))

	' Testing or Live switch
	If proc_live = 1 Then
		Result_Type = Result_Live
	ElseIf proc_live = 0 Then
		Result_Type = Result_Good
	End If

	Flag = ApiDriver.csi_order_set(OrderCtx, OrderField_Result, 0)
	Flag = ApiDriver.csi_order_set(OrderCtx, OrderField_Chargetotal, CDbl(sGrandTotal))
	
	Flag = ApiDriver.csi_order_process(OrderCtx)

	If ApiDriver.bStat <> Succeed Then
		ProcResponse = "fail"
		iProcResponse = 0
		ProcErrMessage = "Error: " & ApiDriver.csi_util_errorstr(ApiDriver.csi_order_error(OrderCtx))
		'Set ApiDriver = nothing
		'Exit Function
	Else
		ProcMessage = ApiDriver.csi_order_get(OrderCtx, OrderField_R_Time)
		ProcActionCode = ApiDriver.csi_order_get(OrderCtx, OrderField_R_Ref)
		ProcResponseCode = ApiDriver.csi_order_get(OrderCtx, OrderField_R_Approved)
		ProcAuth = ApiDriver.csi_order_get(OrderCtx, OrderField_R_Code)
		ProcErrMessage = ApiDriver.csi_order_get(OrderCtx, OrderField_R_Error)
		ProcCustNumber = ApiDriver.csi_order_get(OrderCtx, OrderField_R_Ordernum)
	End If

	If ProcResponseCode = "APPROVED" Then
		ProcResponse = "approve"
		iProcResponse = 1
		ProcAuthCode = Mid(ProcAuth, 1, 6)
		ProcRefCode = Mid(ProcAuth, 7, 10)
		ProcAvsCode = Mid(ProcAuth, 17, 3)
		ProcAvsMsg	= AVSMsg(ProcAvsCode)
	Else
		ProcResponse = "fail"
		ProcErrMessage ="Your transaction was NOT successful. Please verify your payment information and try again."
		'ProcErrMessage = "Error: " & ApiDriver.csi_util_errorstr(ApiDriver.csi_order_error(OrderCtx))
		iProcResponse = 1
	End If

	Flag = ApiDriver.csi_order_drop(OrderCtx)
	Flag = ApiDriver.csi_req_drop(ReqCtx)
	Set ApiDriver = nothing
	
	Call setResponse("LinkPoint",iOrderID,ProcCustNumber,"",ProcAvsCode,ProcAvsMsg,ProcActionCode,ProcAuthCode,ProcResponseCode,ProcErrMessage,iProcResponse)
	LinkPoint = ProcErrMessage
	
End Function

'----------------------------------------------------------------------------
'	PayPal Transaction Processing
'----------------------------------------------------------------------------

Function PayPal
Dim sPath, ppPath, ppCmd, sPaymentServer, ppAmount, iShipID
Dim sReturn, ppBusiness, ppItem_Name, ppQuantity, PaymentString, SndPayment
Dim ppOrderID, ppCustom, ppReturn, ppItem_Number, sInstructions

	'ADMIN VARIABLES FOR PAYPAL
	If sPaymentServer = "" Then
		sPaymentServer = "http://www.paypal.com/xclick/?"
	End If

	sPath = "http://"&Request.ServerVariables("HTTP_HOST")&Request.ServerVariables("URL")
	sPath = Replace(sReturn,"verify.asp","/admin/sfReports1.asp?OrderID=" & iOrderID)
	
	If Request.ServerVariables("HTTPS") = "off" Then
		sReturn  = "http://"&Request.ServerVariables("HTTP_HOST")&Request.ServerVariables("URL")
	Else
		sReturn  = "https://"&Request.ServerVariables("HTTP_HOST")&Request.ServerVariables("URL")
	End If	
	
	sReturn  = Replace(sReturn,"verify.asp","confirm.asp")

	ppCmd = "_xclick"
	ppBusiness = sLogin
	ppReturn = sReturn
	ppItem_Name = Server.URLEncode(C_STORENAME & " Order")
	ppQuantity = "1"
	ppItem_Number = Session("SessionID")
	ppReturn = sReturn
	ppAmount = FormatCurrency(cDbl(sTotalPrice) + cDbl(sHandling) + cDbl(sShipping) + cDbl(iSTax) + cDbl(iCTax))
	ppOrderID = Server.URLEncode(iOrderID)
	
	ppCustom = Server.URLEncode(sCustCompany)&"|"&Server.URLEncode(sCustPhone)&_
	"|"&bCustSubscribed&"|"&Server.URLEncode(sShipCustName)&_
	"|"&Server.URLEncode(sShipCustCompany)	&"|"&Server.URLEncode(sShipCustAddress1)&_
	"|"&Server.URLEncode(sShipCustAddress2)&"|"&Server.URLEncode(sShipCustState)&_
	"|"&Server.URLEncode(sShipCustCity)&"|"&Server.URLEncode(sShipCustZip)&_
	"|"&Server.URLEncode(sShipCustCountry)&"|"&Server.URLEncode(sShipCustPhone)&_
	"|"&Server.URLEncode(sShipCustFax)&"|"&Server.URLEncode(sShipCustEmail)&_
	"|"&Server.URLEncode(sInstructions)&"|"&Server.URLEncode(iShipMethod)&_
	"|"&Server.URLEncode(sShipMethodName)&"|"&Server.URLEncode(iPremiumShipping)&_					
	"|"&Server.URLEncode(sLogin)&"|"&Server.URLEncode(sPassword)&"|"&iAddrID

	PaymentString = "cmd="&ppCmd&"&business="&ppBusiness&"&return="&ppReturn
	PaymentString = PaymentString&"&item_name="&ppItem_Name&"&amount="&ppAmount
	PaymentString = PaymentString&"&item_number="&ppItem_Number
	PaymentString = PaymentString&"&notify_url="&ppReturn&"&custom="&ppCustom
	PaymentString = sPaymentServer&PaymentString		
	'response.redirect PaymentString
	Response.Write "<SCRIPT LANGUAGE=javascript>" & vbCrlf
	Response.Write "<!--" & vbCrlf
	Response.Write " window.location =" & Chr(34) & PaymentString  & chr(34) &   vbcrlf
  	Response.Write "//-->" & vbcrlf
	Response.Write "</SCRIPT>"
	Response.End
End Function
'-------------------------------------------------------------------
' PayPal Response Function
'-------------------------------------------------------------------
Sub PayPalResp(iFlag)
	
			iCustID					= Trim(Request.Cookies("sfCustomer")("custID"))
			sCustFirstName			= Trim(Request.Form("first_name"))
			sCustMiddleInitial		= Trim(Request.Form("last_name"))
			sCustAddress1			= Trim(Request.Form("address_street"))
			sCustCity				= Trim(Request.Form("address_city"))
			sCustState				= Trim(Request.Form("address_state"))
			sCustZip				= Trim(Request.Form("address_zip"))
			sCustCountry			= Trim(Request.Form("address_country"))
			sCustEmail				= Trim(Request.Form("payer_email"))
			sPaymentMethod 			= "PayPal Transaction"
			Session("SessionID") 	= Trim(Request.Form("item_number"))

			arrCustom = Request.Form("Custom")
			arrCustom = Split(Request.Form("Custom"),"|")
			sCustCompany = Replace(arrCustom(0),"|","")
			sCustPhone = Replace(arrCustom(1),"|","")
			bCustSubscribed = Replace(arrCustom(2),"|","")
			sShipCustName = Replace(arrCustom(3),"|","")
			sShipCustCompany = Replace(arrCustom(4),"|","")
			sShipCustAddress1 = Replace(arrCustom(5),"|","")
			sShipCustAddress2 = Replace(arrCustom(6),"|","")
			sShipCustState = Replace(arrCustom(7),"|","")
			sShipCustCity = Replace(arrCustom(8),"|","")
			sShipCustZip = Replace(arrCustom(9),"|","")
			sShipCustCountry = Replace(arrCustom(10),"|","")
			sShipCustPhone = Replace(arrCustom(11),"|","")
			sShipCustFax = Replace(arrCustom(12),"|","")
			sShipCustEmail = Replace(arrCustom(13),"|","")
			sShipInstructions = Replace(arrCustom(14),"|","")
			iShipMethod = Replace(arrCustom(15),"|","")
			sShipMethodName = Replace(arrCustom(16),"|","")
			iPremiumShipping = Replace(arrCustom(17),"|","")
			sLogin = Replace(arrCustom(18),"|","")
			sPassword = Replace(arrCustom(19),"|","")
			iAddrID = Replace(arrCustom(20),"|","")

	Dim ProcErrMsg,  ProcResponse, iProcResponse, ProcMerchNumber, iTransactionID, ProcRefCode, ProcAvsCode, ProcAvsMsg

	If Request.Form("payment_status") = "Failed" Then 
		ProcErrMsg = "This Pay Pal transaction has failed.  Please re-try  your payment"
	ElseIf Request.Form("payment_status") = "Completed" OR Request.Form("payment_status") = "Pending" Then
		ProcResponse = Request.Form("payment_status")
		ProcMerchNumber = Request.Form("verify_sign")
		iTransactionID = Request.Form("txn_id")
		ProcRefCode = Request.Form("txn_id")		
		ProcAvsCode = "not applicable"
		ProcAvsMsg	= "not applicable"
	End If
	' Write to response table	
	' #321 Added the If Condition
	If Trim(iFlag) = "2" Then
		Call setResponse("paypal", iOrderID, iTransactionID, ProcMerchNumber, ProcAvsCode, ProcAvsMsg, ProcResponse, ProcRefCode, "", ProcErrMsg, iProcResponse)
  End If
	' Call setResponse("PayPal",iOrderID,iTransactionID,ProcMerchNumber ,ProcAvsCode,ProcAVSMsg,ProcResponse,ProcRefCode,"",ProcErrMsg,iProcResponse)	
	ProcErrMsg = ProcErrMsg
End Sub

'-----------------------------------------------------------------------------
'	WorldPay Processing Function
' Update on Oct 19,2001
'Compatible with version 1.07
'-----------------------------------------------------------------------------
Function WorldPay(proc_live)
	Dim wpAmount, wpOrderID, wpCustom, FromDate, ToDate, From, ToD, wpContinue, wpDescription, sTstRqst, sfprotocal
	Dim purchase, setInstallationId, setCartId, setShopperId, setCurrencyISOCode, setAmount, setAuthMode, setValidDates, setTestMode, process, hadError, hasMoreErrors
	Dim csCompany, csName, csAddress, csCity, csState, csCountry, csZip, csPhone, csFax, csEmail, sReturn, iShipID

	csCompany	= sCustCompany
	csName = sCustFirstName & " " & sCustLastName
	csAddress	= sCustAddress1 & " " & sCustAddress2 & " " & sCustCity & " " & sCustState
	csCountry	= sCustCountry
	csZip		= sCustZip
	csPhone		= sCustPhone
	csFax		= sCustFax
	csEmail		= sCustEmail

	If Request.ServerVariables("HTTPS") = "off" Then
		sReturn  = "http://"&Request.ServerVariables("HTTP_HOST")&Request.ServerVariables("URL")
		sfprotocal = "0"
		sReturn  = Replace(sReturn,"verify.asp","confirm.asp")
		sReturn  = Replace(sReturn,"http://","")
	Else
		sReturn  = "https://"&Request.ServerVariables("HTTP_HOST")&Request.ServerVariables("URL")
		sfprotocal = "1"
		sReturn  = Replace(sReturn,"verify.asp","confirm.asp")
		sReturn  = Replace(sReturn,"https://","")
	End If	
	
	
	If proc_live = 1 Then
		sTstRqst = "0"	
	ElseIf proc_live = 0 Then
		sTstRqst = "100"	
	End If	
	
	wpDescription = C_STORENAME & " Order"

	wpAmount = cDbl(sTotalPrice) + cDbl(sHandling) + cDbl(sShipping) + cDbl(iSTax) + cDbl(iCTax)
	wpOrderID = Server.URLEncode(iOrderID)
	set purchase = Server.CreateObject("WorldPay.COMpurchase")

Call purchase.init("")
Call purchase.setInstallationId(sLogin)

'Call purchase.setCartId(Session("SessionID"))
Call purchase.setShopperId(wpOrderID)
Call purchase.setCurrencyISOCode(CurrencyISO)
Call purchase.setAmount(wpAmount)
Call purchase.setDescription(wpDescription)
Call purchase.setName(csName)
Call purchase.SetAddress(csAddress)
Call purchase.setCountryISOCode(csCountry) 
Call purchase.SetPostCode(csZip)
Call purchase.SetTelephone(csPhone)
Call purchase.SetFax(csFax)
Call purchase.SetEmail(csEmail)
Call purchase.SetParameter("M_sFName",sCustFirstName)
Call purchase.SetParameter("M_sLName",sCustLastName)
Call purchase.SetParameter("M_sCompany",sCustCompany)
Call purchase.SetParameter("M_sAddress1",sCustAddress1)
Call purchase.SetParameter("M_sAddress2",sCustAddress2)
Call purchase.SetParameter("M_sCity",sCustCity)
Call purchase.SetParameter("M_sState",sCustState)
Call purchase.SetParameter("M_sCountry",sCustCountry)
Call purchase.SetParameter("M_sFax",sCustFax)
Call purchase.SetParameter("M_iShipID",iAddrID)
Call purchase.SetParameter("M_sShipName",sShipCustFirstName&"|"&sShipCustLastName)
Call purchase.SetParameter("M_sShipCompany",sShipCustCompany)
Call purchase.SetParameter("M_sShipAddress",sShipCustAddress1&"|"&sShipCustAddress2)
Call purchase.SetParameter("M_sShipCity",sShipCustCity)
Call purchase.SetParameter("M_sShipState",sShipCustState)
Call purchase.SetParameter("M_sShipCountry",sShipCustCountry)
Call purchase.SetParameter("M_sShipZip",sShipCustZip)
Call purchase.SetParameter("M_sShipPhone",sShipCustPhone)
Call purchase.SetParameter("M_ShipMethod",iShipMethod&"|"&sShipMethodName)
Call purchase.SetParameter("M_bPremiumShipping",bPremiumShipping)
Call purchase.SetParameter("M_sShipInstructions",sShipInstructions)
Call purchase.SetParameter("M_merchURL",sfprotocal&"|"&sReturn)
Call purchase.setAuthMode(purchase.AUTHMODE_full)

From = Year(FromDate) & "-" & Month(FromDate) & "-" & Day(FromDate) & "/" &	Hour(FromDate) & ":" & Minute(FromDate) & ":" & Second(FromDate)
ToD = Year(ToDate) & "-" & Month(ToDate) & "-" & Day(ToDate) & "/" &	Hour(ToDate) & ":" & Minute(ToDate) & ":" & Second(ToDate)

Call purchase.setValidDates(From, ToD)

'	Use test mode?
Call purchase.setTestMode(sTstRqst)

purchase.process()
If purchase.hadError() then

    REM -- Display the errors
	Response.Write("<UL>")
	
	While purchase.hasMoreErrors()
		Response.Write "<LI>" & purchase.getNextError() & "</LI>"
	Wend
	
	Response.Write("</UL>")
	Response.End
	
End If

'response.redirect purchase.produce 

	Response.Write "<SCRIPT LANGUAGE=javascript>" & vbCrlf
	Response.Write "<!--" & vbCrlf
	Response.Write " window.location =" & Chr(34) & purchase.produce  & chr(34) &   vbcrlf
  	Response.Write "//-->" & vbcrlf
	Response.Write "</SCRIPT>"
	Response.end
	

End Function

'----------------------------------------------------------------------------
'	WorldPay Response Function
'----------------------------------------------------------------------------
Sub WorldPayResp

Dim sShipCustAddress, sShipCustName, ShipMethod

			sPaymentMethod			= "WorldPay"
			iCustID					= Trim(Request.Cookies("sfCustomer")("custID")) 
			sCustName				= Trim(Request.QueryString("sCustName"))
			sCustCompany			= Trim(Request.QueryString("sCustCompany"))
			sCustAddress1			= Trim(Request.QueryString("sCustAddress1"))
			sCustAddress2			= Trim(Request.QueryString("sCustAddress2"))
			sCustCity				= Trim(Request.QueryString("sCustCity"))
			sCustState 				= Trim(Request.QueryString("sCustState"))
			sCustZip				= Trim(Request.QueryString("sCustZip"))
			sCustCountry			= Trim(Request.QueryString("sCustCountry"))
			sCustPhone				= Trim(Request.QueryString("sCustPhone"))
			sCustFax				= Trim(Request.QueryString("sCustFax"))
			sCustEmail				= Trim(Request.QueryString("CustomerEmail"))
			bCustSubscribed		    = Trim(Request.QueryString("bCustSubscribed"))
			iPremiumShipping		= Trim(Request.QueryString("iPremiumShipping"))		
			sShipCustName			= Split(Request.QueryString("sShipCustName"),"|")
			sShipCustFirstName		= sShipCustName(0)
			sShipCustLastName		= sShipCustName(1)
			sShipCustCompany		= Trim(Request.QueryString("sShipCustCompany"))
			sShipCustAddress		= Split(Request.QueryString("sShipCustAddress"),"|")
			sShipCustAddress1		= sShipCustAddress(0)
			sShipCustAddress2		= sShipCustAddress(1)
			sShipCustCity			= Trim(Request.QueryString("sShipCustCity"))
			sShipCustState			= Trim(Request.QueryString("sShipCustState"))
			sShipCustZip			= Trim(Request.QueryString("sShipCustZip"))
			sShipCustCountry		= Trim(Request.QueryString("sShipCustCountry"))	
			sShipCustPhone			= Trim(Request.QueryString("sShipCustPhone"))
			sShipCustFax			= Trim(Request.QueryString("sShipCustFax"))
			sShipCustEmail			= Trim(Request.QueryString("sShipCustEmail"))
			ShipMethod				= Split(Request.QueryString("ShipMethod"),"|")
			iShipMethod				= ShipMethod(0)
			sShipMethodName			= ShipMethod(1)
			sShipInstructions		= Trim(Request.QueryString("sShipInstructions"))	
			iAddrID					= Trim(Request.QueryString("iShipID"))
			
Dim ProcErrMsg,  ProcResponse, iProcResponse, ProcMerchNumber, iTransactionID, ProcRefCode, ProcAvsCode, ProcAvsMsg, Auth, TransTime, AuthMode

iTransactionID = Request.QueryString("TransID")
ProcResponse = Request.QueryString("RawAuthMessage")
ProcRefCode = Request.QueryString("RawAuthCode")
TransTime = Request.QueryString("TransTime")
AuthMode = Request.QueryString("AuthMode")
Auth = Request.QueryString("Auth")
ProcMerchNumber = Request.QueryString("InstID")
iCustID = Trim(Request.Cookies("sfCustomer")("custID"))

	If Auth Then
		ProcResponse = ProcResponse
		ProcMerchNumber = ProcMerchNumber
		iTransactionID = iTransactionID
		ProcRefCode = ProcRefCode		
		ProcAvsCode = "not applicable"
		ProcAvsMsg	= "not applicable" 
	Else
		ProcErrMsg = "This transaction has failed.  Please re-try  your payment"
	End If
	' Write to response table	
	Call setResponse("WorldPay",iOrderID,iTransactionID,ProcMerchNumber ,ProcAvsCode,ProcAVSMsg,ProcResponse,ProcRefCode,"",ProcErrMsg,iProcResponse)	
	ProcErrMsg = ProcErrMsg

End Sub

'-----------------------------------------------------------------------------
' Verisign's Signio PayProFlow Subroutine
'-----------------------------------------------------------------------------
Function SignioPayProFlow(proc_live)

Dim ccExp_Date, obj, ccAddress, Amt, TrxType, ObjResult, i, aPos, GrandTotal, iProcResponse, ProcRspMsg, ProcFailedMsg, strIn, aTemp,iTransactionID, iAuthCode, sAVSMsg
Dim sAVSADDR, sAVSZIP, sAVSMsg1, sAVSMsg2, ParmList, Ctx1, ProcMerchTransNum

	GrandTotal = FormatCurrency(sGrandTotal,2)
	ccExp_Date = sCustCardExpiryMonth & right(sCustCardExpiryYear,2) 'Issue #271

	ccAddress = sCustAddress1 & " " & sCustAddress2
	
	Amt = GrandTotal
	Amt = Mid(Amt,2)
		
	sMercType = UCase(sMercType)
		If sMercType = "AUTHCAPTURE" Then
			TrxType = "S"
		ElseIf sMercType = "AUTHONLY" Then
			TrxType = "A"
		End If

	' New for Faud Checking
	'ParmList = "TRXTYPE=" & TrxType & "&TENDER=C&USER=" & sLogin & "&PWD=" & sPassword & "&ACCT=" & sCustCardNumber & "&EXPDATE=" & ccExp_Date & "&AMT=" & Amt & "&ZIP=" & sCustZip & "&STREET=" & sCustAddress1 'Issue #260
	ParmList = "TRXTYPE=" & TrxType & "&TENDER=C&USER=" & sLogin & "&PWD=" & sPassword & "&ACCT=" & sCustCardNumber & "&EXPDATE=" & ccExp_Date & "&AMT=" & Amt & "&ZIP=" & sCustZip & "&City=" & sCustCity & "&State="& sCustState & "&STREET="& sCustAddress1 & "&PhoneNum="& sCustPhone & "&Country=" & sCustCountry & "&CUSTIP="& aReferer(2) & "&EMAIL=" & sCustEmail & "&FirstName=" & sCustFirstName & "&LastName="& sCustLastName & "&ShipFirstName=" & sShipCustFirstName & "&ShipLastName=" & sShipCustLastName & "&ShiptoCountry=" & sShipCustCountry & "&ShipToCity=" & sShipCustCity & "&ShiptoState=" & sShipCustState & "&ShipToZip=" & sShipCustZip 'Issue #260
	
	If sTransMethod = "17" Then 
		Set obj = Server.CreateObject("PFProSSControl.PFProSSControl2.1")
			If trim(sPaymentServer) = "" or isNull(sPaymentServer) Then 
				If proc_live = 1 Then 
					sPaymentServer = "payflow.verisign.com"
				ElseIf proc_live =  0 Then 
					sPaymentServer = "test-payflow.verisign.com"
				End If
			End If

		Ctx1 = obj.CreateContext(sPaymentServer, 443, 30, "", 0, "", "")
		strIn = obj.SubmitTransaction(Ctx1, ParmList, Len(ParmList))
		obj.DestroyContext (Ctx1)
			
	ElseIf sTransMethod = "3" Then
		Set obj = Server.CreateObject("PFProSSControl.PFProSSControl.1")
			If trim(sPaymentServer) = "" or isNull(sPaymentServer) Then 
				If proc_live = 1 Then 
					sPaymentServer = "payflow.verisign.com"
				ElseIf proc_live =  0 Then 
					sPaymentServer = "test-payflow.verisign.com"
				End If
			End If

		' Create object and process it
			obj.HostAddress = sPaymentServer
			obj.HostPort = 443
			obj.TimeOut = 30
			obj.DebugMode = 1
			obj.ParmList = ParmList
			
			obj.PNInit()
			obj.ProcessTransaction()
			strIn = obj.Response 	
			obj.PNCleanup 
	End If

		aPos = split(strIn, "&")
	For i = 0 to UBOUND(aPos)
		If Instr(aPos(i),"RESULT") > 0 Then			
			aTemp = split(aPos(i),"=")
			ProcResponse = Trim(aTemp(1))
		End If	
		If Instr(aPos(i),"RESPMSG") > 0 Then
			aTemp = split(aPos(i),"=")
			ProcRspMsg = Trim(aTemp(1))
		End If		
		If Instr(aPos(i),"PNREF") > 0 Then
			aTemp = split(aPos(i),"=")
			iTransactionID = Trim(aTemp(1))
		End If		
		If Instr(aPos(i),"AUTHCODE") > 0 Then
			aTemp = split(aPos(i),"=")
			iAuthCode = Trim(aTemp(1))
		End If		
		If Instr(aPos(i),"ERRMSG") > 0 Then
			aTemp = split(aPos(i),"=")
			ProcErrMsg = Trim(aTemp(1))
		End If		
		If Instr(aPos(i),"AVSADDR") > 0 Then
			aTemp = split(aPos(i),"=")
			sAVSADDR = Trim(aTemp(1))
			sAVSMsg1 = AVSMsg(sAVSADDR)	
		End If	
		If Instr(aPos(i),"AVSZIP") > 0 Then
			aTemp = split(aPos(i),"=")
			sAVSZIP = ProcAvsCode & Trim(aTemp(1))
			sAVSMsg2 = AVSMsg(sAVSZIP)	
		End If	

		' Added for Faud Checking
		If Instr(aPos(i),"ERRCODE") > 0 Then
			aTemp = split(aPos(i),"=")
			ProcFailedMsg = Trim(aTemp(1))
		End If		
		If Instr(aPos(i),"FRAUDCODE") > 0 Then
			aTemp = split(aPos(i),"=")
			ProcMerchTransNum = "<br>FraudCode: " & Trim(aTemp(1)) ' Item4
		End If
		If Instr(aPos(i),"SCORE") > 0 Then
			aTemp = split(aPos(i),"=")
			ProcMerchTransNum = ProcMerchTransNum & "<br>Score: " & Trim(aTemp(1)) 'Item4
		End If
		If Instr(aPos(i),"FRAUDMSG") > 0 Then
			aTemp = split(aPos(i),"=")
			ProcActionCode = "Fraud Msg: " & Trim(aTemp(1))
		End If
		If Instr(aPos(i),"REASON1") > 0 Then
			aTemp = split(aPos(i),"=")
			ProcActionCode = ProcActioncode & "<br>Reason 1: " & Trim(aTemp(1))
		End If
		If Instr(aPos(i),"REASON2") > 0 Then
			aTemp = split(aPos(i),"=")
			ProcActionCode = ProcActionCode & "<br>Reason 2: " & Trim(aTemp(1))
		End If
		If Instr(aPos(i),"REASON3") > 0 Then
			aTemp = split(aPos(i),"=")
			ProcActionCode = ProcActionCode & "<br>Reason 3: " & Trim(aTemp(1))
		End If
		If Instr(aPos(i),"EXCEPTION1") > 0 Then
			aTemp = split(aPos(i),"=")
			ProcActionCode = ProcActionCode & "<br>Exception 1: " & Trim(aTemp(1)) 
		End If
		If Instr(aPos(i),"EXCEPTION2") > 0 Then
			aTemp = split(aPos(i),"=")
			ProcActionCode = ProcActionCode & "<br>Exception 2: " & Trim(aTemp(1))
		End If
		If Instr(aPos(i),"EXCEPTION3") > 0 Then
			aTemp = split(aPos(i),"=")
			ProcActionCode = ProcActionCode & "<br>Exception 3: " & Trim(aTemp(1))
		End If
		If Instr(aPos(i),"EXCEPTION4") > 0 Then
			aTemp = split(aPos(i),"=")
			ProcActionCode = ProcActionCode & "<br>Exception 4: " & Trim(aTemp(1))
		End If
		If Instr(aPos(i),"EXCEPTION5") > 0 Then
			aTemp = split(aPos(i),"=")
			ProcActionCode = ProcActionCode & "<br>Exception 5: " & Trim(aTemp(1))
		End If
		If Instr(aPos(i),"EXCEPTION6") > 0 Then
			aTemp = split(aPos(i),"=")
			ProcActionCode = ProcActionCode & "<br>Exception 6: " & Trim(aTemp(1))
		End If
		If Instr(aPos(i),"EXCEPTION7") > 0 Then
			aTemp = split(aPos(i),"=")
			ProcActionCode = ProcActionCode & "<br>Exception 7: " & Trim(aTemp(1))
		End If
	Next	
	
	' The AVSCode is a combination of ADDR and ZIP AVS codes 
		ProcAvsCode = sAVSADDR & sAVSZIP
		sAVSMsg = sAVSMsg1 & ";" & sAVSMsg2

	' Failed or not  
		If ProcResponse = 0 Then
			iProcResponse = 1
		Else
			iProcResponse = 0
			If ProcResponse = "-1" Then
				ProcFailedMsg = "Server Socket Unavailable"			
			ElseIf ProcResponse = "-2" Then
				ProcFailedMsg = "Hostname lookup failed"			
			ElseIf ProcResponse = "-3" Then
				ProcFailedMsg = "Server Timed Out"			
			ElseIf ProcResponse = "-4" Then
				ProcFailedMsg = "Socket Initialization Error"			
			ElseIf ProcResponse = "-5" Then
				ProcFailedMsg = "SSL Context Initialization Failed"	
			ElseIf ProcResponse = "-6" Then
				ProcFailedMsg = "SSL Verification Policy Failure"								
			ElseIf ProcResponse = "-7" Then
				ProcFailedMsg = "SSL Verify Location Failed"		
			ElseIf ProcResponse = "-8" Then
				ProcFailedMsg = "X509 Certification Verification Error"		
			Else			
				ProcFailedMsg = "General Processing Error - Please try to notify the merchant"			
			End If	
			
			ProcFailedMsg = ProcFailedMsg & " " & ProcErrMsg & " " & ProcRspMsg
End If
	Set obj = Nothing

	' Write to response table		
	Call setResponse("Signio PFP",iOrderID,iTransactionID,ProcMerchTransNum,ProcAvsCode,sAVSMsg,ProcActionCode,iAuthCode,"",ProcFailedMsg ,iProcResponse)	
	SignioPayProFlow = ProcFailedMsg 
	
End Function	

'--------------------------------------------------------------------------
' SecurePay Function
'--------------------------------------------------------------------------
Function SecurePay(proc_live)
Dim objOrder, sActionCode, iProcResponse, iTransactionID, sErrorMessage, iAVSCode, sAvsMsg, orderAVSREQ
Dim SPCOM_Response, sNewVar, sSP_Array, sReturn_Code, sApprov_Num, sCard_Response, sAVS_Response
	
    If trim(sPaymentServer) = "" or isNull(sPaymentServer) Then 
       	sPaymentServer = "https://www.securepay.com/secure1/index.asp"
    End If	
    
	' Testing or live mode
	If proc_live = 1 Then
		orderAVSREQ = "1"
	ElseIf proc_live = 0 Then
		orderAVSREQ = "4"
		
		

	End If	
	
	'create instance of component 
    Set objOrder = Server.CreateObject("SPCOM.clsSecureSend")
   
    

    objOrder.URL		= sPaymentServer
    objOrder.Amount		= FormatNumber(sGrandTotal,2,0,0,0)
    objOrder.AVSREQ		= orderAVSREQ
    objOrder.Street			= sCustAddress1 & " " & sCustAddress2
    objOrder.City			= sCustCity
    objOrder.State			= sCustState
    objOrder.Zip			= sCustZip
    objOrder.NameOnCard	= sCustCardName
    objOrder.CreditCardNumber = sCustCardNumber 
    objOrder.Email			= sCustEmail
    objOrder.Month		= sCustCardExpiryMonth 
    objOrder.Year			= right(sCustCardExpiryYear,2)
    objOrder.Merch_ID		= sLogin 
    objOrder.tr_type		= "SALE" 
    objOrder.Send
    SPCOM_Response = objOrder.ReturnCode  
    
    Set objOrder = Nothing
	if instr(1,SPCOM_Response,",") then
	sNewVar = cStr(SPCOM_Response)
	sSP_Array = Split(sNewVar, ",")
	sActionCode  = sSP_Array(0)
	sApprov_Num  = sSP_Array(1)
	sCard_Response = sSP_Array(2)
	sAVS_Response = sSP_Array(3)
	else
	sCard_Response=SPCOM_Response
	end if
    If sActionCode = "Y" Then
		iProcResponse = 1
	Else 
		iProcResponse = 0
		sErrorMessage = sCard_Response
	End If 		
	
	iTransactionID =  sApprov_Num    
    iAVSCode = sAVS_Response
 
    sAvsMsg = AVSMsg(iAvsCode)

    Call setResponse("SecurePay",iOrderID,iTransactionID,"",iAVSCode,sAvsMsg,sActionCode,"","",sErrorMessage,iProcResponse)
    
    
	SecurePay = sErrorMessage
End Function

'-------------------------------------------------------------------------------------
' SurePay Sub-routine
'-------------------------------------------------------------------------------------
Function SurePay(proc_live)

Dim objXML, pp, auth, credit, addr, ship_addr, lineitem, ordertext, strXML, strMode, strDTD, ISO, strHeader, strResponse
Dim TOTAL, PROD_ID, PROD_DESC, PROD_QUANTITY, PROD_UNIT, strInsert, xmlResponse, strBool, objElement
Dim ProcErrMsg, ProcMessage, ProcErrLoc, ProcErrCode, ProcCustNumber, ProcMerchNumber, ProcAddlData, ProcRefCode, ProcAVSZip
Dim iProcResponse, strAuth, authMsg, iAVSCode, sAVSMsg, iTransactionID
Dim ORDER_TEXT

	' Fix up the total
	ISO = FindCurrencyIso(getLCID())
	TOTAL = trim(sGrandTotal & ISO)
		
	'workaround for lineitem in XML
	PROD_ID =  iOrderID
	PROD_QUANTITY = "1"
	PROD_UNIT = TOTAL
	PROD_DESC = "StoreFront Purchase"
	
	' check to see if there is a ship message or supply a default one
	If sShipInstructions <> "" Then
		ORDER_TEXT = mid(sShipInstructions,1,20)
	Else ORDER_TEXT = "No Order Text"
	End if

	If proc_live = 1 Then 	
		If trim(sPaymentServer) = "" or isNull(sPaymentServer) Then 
			sPaymentServer = "https://xml.surepay.com"			
		End If	
	ElseIf proc_live = 0 Then
		sPaymentServer = "https://xml.test.surepay.com"	
		sLogin = "5555"
		sPassword = "password"	
	End If     

'Create the XML object
Set objXML = Server.CreateObject("MSXML.DOMDocument")

	strHeader= "<!DOCTYPE pp.request PUBLIC """& chr(45) & chr(47)&chr(47) &"IMALL" & chr(47) &chr(47) &"DTD PUREPAYMENTS 1.0" & chr(47)&chr(47) & "EN"""_
			 &" ""http:" & chr(47)&chr(47) & "www.purepayments.com" & chr(47) & "dtd" & chr(47) & "purepayments.dtd"">" 
			
	'Set up XML
	Set pp = objXML.createElement("pp.request")
		pp.SetAttribute "merchant", sLogin
		pp.SetAttribute "password", sPassword
	
		Set auth = objXML.createElement("pp.auth")
		   auth.SetAttribute "ordernumber", iOrderID
		pp.appendChild auth

				Set credit = objXML.createElement("pp.creditcard")
					 credit.SetAttribute "number", sCustCardNumber
				 	 credit.SetAttribute "expiration", sCustCardExpiry
				auth.appendChild credit	 

					Set addr = objXML.createElement("pp.address")
						addr.SetAttribute "type", "billing"
						addr.SetAttribute "fullname", sCustCardName
						addr.SetAttribute "address1", sCustAddress1
						addr.SetAttribute "address2", sCustAddress2	
						addr.SetAttribute "city", sCustCity
						addr.SetAttribute "state", sCustState
						addr.SetAttribute "zip", sCustZip
						addr.SetAttribute "country", sCustCountry
						addr.SetAttribute "phone", sCustPhone
						addr.SetAttribute "email", sCustEmail
					credit.appendChild addr	

			   	Set ship_addr = objXML.createElement("pp.address")
					ship_addr.SetAttribute "type", "shipping"
					ship_addr.SetAttribute "fullname", sShipCustName
					ship_addr.SetAttribute "address1", sShipCustAddress1
					ship_addr.SetAttribute "address2", sShipCustAddress2
					ship_addr.SetAttribute "city", sShipCustCity
					ship_addr.SetAttribute "state", sShipCustState
					ship_addr.SetAttribute "zip", sShipCustZip
					ship_addr.SetAttribute "country", sShipCustCountry
					ship_addr.SetAttribute "phone", sShipCustPhone
					ship_addr.SetAttribute "email", sShipCustEmail
				auth.appendChild ship_addr				

					
				Set lineitem	= objXML.createElement("pp.lineitem")
					lineitem.SetAttribute "sku", iOrderID
					lineitem.SetAttribute "quantity", PROD_QUANTITY
					lineitem.SetAttribute "description", PROD_DESC
					lineitem.SetAttribute "unitprice", TOTAL 
					lineitem.SetAttribute "taxrate", "0"
				auth.appendChild lineitem
				
				Set ordertext = objXML.createElement("pp.ordertext")
					ordertext.SetAttribute "type", "description"
					ordertext.text = ORDER_TEXT			
				auth.appendChild ordertext

	objXML.appendChild pp

	' Server encode the string to Post
	strXML = "xml= " & Server.URLencode(strHeader) & Server.URLEncode(objXML.xml)

	strResponse = SubmitXML(strXML, sPaymentServer)
	strInsert = mid(strResponse,118)

	Set xmlResponse = CreateObject("MSXML.DOMDocument")
		strBool = xmlResponse.loadXML(strInsert)

	' Error checking and Response handling
	' check to see if the xml has been successfully loaded
	If Not strBool Then
		iProcResponse = 0
		ProcErrMsg = "loadXML has failed. Check the xml response string"	
	Else
			' check if xmlResponse has content
			If (xmlResponse.hasChildNodes) Then
			
			' get back the authresponse message 
		 	Set objElement = xmlResponse.documentElement.selectSingleNode("pp.authresponse")	 			 		

				If (objElement.getAttribute("failure") = "true") Then
					iProcResponse = 0
					
		 			strAuth = objElement.getAttribute("authstatus")
	     			
	     			' Get Authorization Status
	     			Select Case strAuth
						Case "DCL"
			  				AuthMsg = "Authorization Declined"
			  				ProcAuthCode =  objElement.getAttribute("transactionid")
						Case "ERR"
			  				AuthMsg = "Error occurred in Authorization"
			  				ProcAuthCode =  objElement.getAttribute("transactionid")
						Case "REF"
			  				AuthMsg = "Referred Authorization"
			  				ProcAuthCode =  objElement.getAttribute("transactionid")
						Case Else
					 		AuthMsg = "Unknown Authresponse occured"
					 		ProcAuthCode = ""
	   				End Select 
				
	 				ProcMessage = AuthMsg		
		 			ProcErrMsg = objElement.Text 	 			
	 				ProcCustNumber = objElement.getAttribute("ordernumber")
					ProcMerchNumber =  objElement.getAttribute("merchant")
				Else
					' success
					iProcResponse = 1
					strAuth = objElement.getAttribute("authstatus")
					iTransactionID =  objElement.getAttribute("transactionid")
					ProcCustNumber = objElement.getAttribute("ordernumber")
					ProcMerchNumber =  objElement.getAttribute("merchant")
				End if
	 	  
	 	  		' Get AVS info
	 	  		iAVScode = objElement.getAttribute("avs")
	 	  		sAVSMsg = AVSMsg(iAVScode)	 	  		
			Else
	   			ProcErrMsg = "No response came from SurePay"
	   			iProcResponse = 0
			End	 If    		
			
	End If

	' close objects
	Set objXML = Nothing
	Set xmlResponse = Nothing
	Set objElement = Nothing
			
	Call setResponse("SurePay",iOrderID,iTransactionID,ProcMerchNumber,iAVSCode,sAVSMsg,strAuth,"","",ProcErrMsg,iProcResponse)
	SurePay = ProcErrMsg
End Function

'-------------------------------------------------------------------------------------
' Function to submit the XML object, returning a response string
'-------------------------------------------------------------------------------------

Function SubmitXML(strXML, strPostURL)
	Dim XMLHttpRequest, strResponse
	
	Set XMLHttpRequest = Server.CreateObject("Microsoft.XMLHTTP")
		XMLHttpRequest.Open "POST", strPostURL , "false" , "" ,""
		XMLHttpRequest.Send strXML
	
	strResponse = XMLHttpRequest.responseText
	
	Set XMLHttpRequest = Nothing
	SubmitXML = strResponse    
End Function	


'-------------------------------------------------------------------------------------
' ISO Function and getLCID function
'-------------------------------------------------------------------------------------
Function FindCurrencyISO(LCID)
Dim SQL, rsCurr, SelCurrency
  
   SQL = "SELECT slctvalCurrencyISO FROM sfSelectValues WHERE slctvalLCID = '" & LCID & "' "         
   Set rsCurr = cnn.execute (SQL)
   SelCurrency = rsCurr("slctvalCurrencyISO")
   
   If (isNull(SelCurrency)) Then
    	Response.Write("Sorry, that country does not have an ISO currency type assigned to it")
    	Response.End
   End If	
   
   closeobj(rsCurr) 
   FindCurrencyISO = SelCurrency
End Function		


Function getLCID
	Dim SQL, rsAdmin, LCID
	SQL = "SELECT adminLCID FROM sfAdmin"
	Set rsAdmin = cnn.execute (SQL)
	LCID = rsAdmin("adminLCID")
	
	closeobj(rsAdmin)
	getLCID = LCID	
End Function

'-----------------------------------------------------------------------------
' AVS decoding function
' Returns the corresponding AVS message
'-----------------------------------------------------------------------------
Function AVSMsg(ProcAvsCode)
	Select Case ProcAvsCode
		Case "A"
		  AVSMsg = "Address matches, ZIP does not. (Code A)"
		Case "E"
		  AVSMsg = "Ineligible transaction. (Code E)"
		Case "N"
		  AVSMsg = "Neither address nor ZIP matches. (Code N)"
		Case "R"
		  AVSMsg = "Retry (system unavailable or timed out). (Code R)"
		Case "S"
		  AVSMsg = "Card type not supported. (Code S)"
		Case "U"
		  AVSMsg = "Address information unavailable. (Code U)"
		Case "W"
		  AVSMsg = "9 digit zip match, address does not. (Code W)"
		Case "X"
		  AVSMsg = "Exact match (9 digit zip and address). (Code X)"
		Case "Y"
		  AVSMsg = "Address and 5 digit zip match. (Code Y)"
		Case "Z"
		  AVSMsg = "5 digit zip matches, address does not. (Code Z)"
		Case Else
		  AVSMsg = "Unknown AVS Code."
	End Select	
End Function


'-----------------------------------------------------------------------------
' Write to payment response table after processing a transaction
' Returns nothing
'-----------------------------------------------------------------------------
Sub setResponse(sProcType,iOrderID,iTransactionID,iMercTransNo,iAVSCode,sAUXMsg,sActionCode,iAuthNo,iRetrievalCode,sErrorMessage,iProcResponse)
Dim rsResponse

Set rsResponse = Server.CreateObject("ADODB.RecordSet")
	rsResponse.Open "sfTransactionResponse", cnn, adOpenDynamic, adLockOptimistic, adCmdTable
	rsResponse.AddNew 
	rsResponse.Fields("trnsrspOrderId") = iOrderID
	rsResponse.Fields("trnsrspCustTransNo") = iTransactionID
	rsResponse.Fields("trnsrspMerchTransNo") = iMercTransNo
	rsResponse.Fields("trnsrspAVSCode") = iAVSCode
	rsResponse.Fields("trnsrspAUXMsg") = sAUXMsg
	rsResponse.Fields("trnsrspActionCode") = sActionCode
	rsResponse.Fields("trnsrspRetrievalCode") = iRetrievalCode
	rsResponse.Fields("trnsrspAuthNo") = iAuthNo
	rsResponse.Fields("trnsrspErrorMsg") = sErrorMessage
	rsResponse.Fields("trnsrspErrorLocation") = sProcType
	rsResponse.Fields("trnsrspSuccess") = iProcResponse
	rsResponse.Update
	closeobj(rsResponse)
End Sub
%>








