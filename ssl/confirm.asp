<%@ Language=VBScript %>
<%	option explicit 
	Response.Buffer = True
	%>
<!--#include file="SFLib/db.conn.open.asp"-->
<!--#include file="SFLib/adovbs.inc"-->
<!--#include file="error_trap.asp"-->
<!--#include file="SFLib/incDesign.asp"-->
<!--#include file="SFLib/incConfirm.asp"-->
<!--#include file="SFLib/incGeneral.asp"-->
<!--#include file="SFLib/mail.asp"-->
<!--#include file="SFLib/processor.asp"-->
<!--#include file="SFLib/incCC.asp"-->
<!--#include file="SFLIB/incAE.asp"-->
<%   
   Const vDebug = 0
 
'@BEGINVERSIONINFO

'@APPVERSION: 50.4011.0.4

'@FILENAME: confirm.asp

'@DESCRIPTION: Confirmation page, writes order to database

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000, 2001 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

' #321 MS

'@ENDVERSIONINFO

'@BEGINCODE
dim tmpOrderQty
If Application("AppName") = "StoreFrontAE" Then 'SFAE
	tmpOrderQty=GetTotalOrderQTY
else
	tmpOrderQty=GetTotalOrderQTYSE
end if	
if isnumeric(tmpOrderQty)=false or tmpOrderQty <= 0 then
'If Session("SessionID") = "" Then 'SFUPDATE\

	Session.Abandon
	Response.Redirect(C_HomePath & "search.asp")		
End If

If Application("AppName") = "StoreFrontAE" Then 'SFAE
	Confirm_CheckCartAndRedirect 
	ReleaseAppLock 'release lock over two minutes
				   'or any previous lock from this session
	LockApp 
'	Confirm_UpdateInventory 
	UnLockApp
End IF
if trim(Request.item("PaymentMethod"))="" and Request.item("custom") = "" and Request.QueryString("wpresponse") = "" Then
Response.redirect "process_order.asp"
'this is if paypal had an error and they hit continue on paypal site, it returns to confirm with no info
end if
IF lcase(Request.ServerVariables("HTTPS")) = "on" then '.3008
   If iConverion = 1 Then Response.Write "<script language=""JavaScript"" src=""https://www.oanda.com/cgi-bin/fxcommerce/fxcommerce.js?user=" & sUserName & """></script>" 
ELSE
   If iConverion = 1 Then Response.Write "<script language=""JavaScript"" src=""http://www.oanda.com/cgi-bin/fxcommerce/fxcommerce.js?user=" & sUserName & """></script>"
End IF



Dim CurrencyISO 
CurrencyISO = getCurrencyISO(Session.LCID )	
Dim aAdmin, sTransMethod, sPaymentMethod, bPremiumShipping, sPaymentServer, sLogin, sPassword,sMercType, sCustCardType, sCustCardName, sCustCardNumber
Dim sCustCardExpiryMonth, sCustCardExpiryYear,	sCustCardTypeName, iRoutingNumber, sBankName, iCheckNumber, iCheckingAccountNumber
Dim sPOName, iPONumber, iCustID, sCustFirstName, sCustMiddleInitial, sCustLastName, sCustCompany, sCustAddress1, sCustAddress2, sCustCity, sCustStateName, sCustZip, sCustCountryName
Dim sCustEmail, sCustPhone, sCustFax, bCustSubscribed, sShipCustFirstName, sShipCustMiddleInitial, sShipCustLastName, sShipCustCompany, rsAdminInfo, iHandling, iShipType
Dim sShipCustAddress1, sShipCustAddress2, sShipCustCity,sShipCustState,sShipCustStateName, sShipCustZip, sShipCustCountryName, sShipCustPhone, sShipCustEmail, sShipCustFax,sPhoneFaxPayType
Dim iShipMethod, sShipInstructions, iPayID, iAddrID, sShipMethodName, iTmpOrderID, iOrderID, aOrderID, iPhoneFaxType
Dim sBgColor, sFontFace,sFontColor,iFontSize, iRow, sMailMethod, sCustCardExpiry, sShipCustCountry, sCustState, sCustCountry, sShipCustName, sCustName	
Dim ProcResponse,ProcMessage,ProcCustNumber,ProcAddlData,ProcRefCode,ProcAuthCode,ProcMerchNumber,ProcActionCode,ProcErrMsg,ProcErrLoc,ProcErrCode,ProcAvsCode, iPremiumShipping
Dim sAuthResp,sMailPassword,arrCustom, proc_live
Dim sGrandTotalOut , sTotalPriceOut 'SFUPDATE
dim dUnitPrice,bTaxShipIsActive,dTaxAble_Amount


	iCustID	= Trim(Request.Cookies("sfCustomer")("custID"))
	' Collect values from admin table
	aAdmin = getAdminRow(sPaymentMethod)	
		' Trans method
		sTransMethod = aAdmin(0)
		' Payment server
		sPaymentServer = aAdmin(1)
		' Login 
		sLogin = aAdmin(2)
		' Password
		sPassword = aAdmin(3)
		' MerchantType
		sMercType = aAdmin(4)
		' Mail Method
		sMailMethod = aAdmin(5)
		' Handling Method 
		sHandling = aAdmin(6)
		
	'----------------------------------------------------
	' Recalculate Order
	'----------------------------------------------------
	Dim SQL, sProdID, sAttrUnitPrice, sUnitPrice, sProdName, sProdPrice, sProductSubtotal, iShipped, bCod
	Dim sTotalPrice, sShipping, dProductSubtotal, sGrandTotal, sHandling, sPrimaryEmail, sPath, sProcErrMsg, iHandlingType
	Dim iCounter, iSTax, iCTax, sTotalSTax, sTotalCTax, iQuantity, iProdAttrNum, iProdNumber, iCodAmount
	Dim rsAllOrders, aProduct, aProdAttrID, iProductCounter, aAllProd(), iAttrCounter, aPurchases, iAttrNumber, aReferer
	
	
	Set rsAdminInfo = Server.CreateObject("ADODB.RecordSet")
    rsAdminInfo.Open "sfAdmin", cnn, adOpenStatic,adLockReadOnly , adCmdTable
    
	With rsAdminInfo
		sHandling  = .Fields("adminHandling")
		iHandling  = .Fields("adminHandlingIsActive")
		iHandlingType = .Fields("adminHandlingType")
		iCodAmount    = .Fields("adminCODAmount")
		sPrimaryEMail = .Fields("adminPrimaryEmail")
		bTaxShipIsActive = .Fields("adminTaxShipIsActive")	'#383	
		sPath = .Fields("adminSSLPath")
	End With

	closeObj(rsAdminInfo)			
		
	If trim(iHandling) = "1" Then 
		If iHandlingType = 1 Then 
			iShipped = 1
		Else
			iShipped = 0
		End If
	Else
		sHandling = 0
	End If
	
	sPaymentMethod = Trim(Request.Form("PaymentMethod"))
	
	'COD Charge
	If sPaymentMethod <> "COD" Then			
		iCodAmount = 0
	End If	

	If sPaymentMethod = "PhoneFax" Then
		  iPhoneFaxType = CheckPhoneFax()		
	End If

If Session("SessionID") = "" Then
		Session.Abandon
		Response.Redirect(C_HomePath & "search.asp")		
End If

		If Request.QueryString("message") <> "" Then
			Call InternetCashResp			
		ElseIf Request.Form("custom") <> "" Then
			Call PayPalResp("1") ' #321
		ElseIf Request("CSVPOSRESPONSE") <> "" Then
	 	 	Call CSVPOSResp
		ElseIf Request.QueryString("wpresponse") <> "" Then 
			Call WorldPayResp("1") ' #321
		End If

		
	If Application("AppName") = "StoreFrontAE" Then 'SFAE
		Confirm_GetBillAmount 
		Confirm_GetBackOrderAmount
	End IF

			
	SQL = "SELECT * FROM sfTmpOrderDetails WHERE odrdttmpSessionID = " & Session("SessionID")

	Set rsAllOrders = Server.CreateObject("ADODB.RecordSet")
	rsAllOrders.Open SQL, cnn, adOpenKeyset, adLockReadOnly, adCmdText


	If rsAllOrders.EOF Then
		closeObj(rsAllOrders)
		closeObj(cnn)
		' redirect to neworder screen
		Session.Abandon
		Response.Redirect(C_HomePath & "search.asp")		
	Else
    	If not isnull(rsAllOrders("odrdttmpHttpReferer")) then
		  aReferer = Split(rsAllOrders("odrdttmpHttpReferer"), ",")
		Else
		 aReferer =""
		End If
		aPurchases = rsAllOrders.GetRows
		iProdNumber = rsAllOrders.RecordCount
		iAttrNumber = getAttributeNumber()
		rsAllOrders.MoveFirst
	
		Redim aOrderID(iProdNumber)
		Redim aAllProd(iAttrNumber,iProdNumber)
		If Application("AppName") = "StoreFrontAE" then 'SFAE
			Redim aProdInfoAE(iProdNumber,4)
		End If
		' Initialize total to 0
		sTotalPrice = "0"
		sShipping = "0"
		sProductSubtotal	= "0"	
		dProductSubtotal	= 0
		iProductCounter	= 0
		
		iSTax = 0 'SFUPDATE
		iCTax = 0 'SFUPDATE
		
		Do While NOT rsAllOrders.EOF
   
			' Get the ProdIDs
			iTmpOrderID = rsAllOrders.Fields("odrdttmpID")
			sProdID = rsAllOrders.Fields("odrdttmpProductID")
			iQuantity = rsAllOrders.Fields("odrdttmpQuantity")
			
	    
			' put orderid into an array for deletion
			aOrderID(iProductCounter) = iTmpOrderID
		    
			' Get an array of 3 values from getProduct()

			ReDim aProduct(3)
			aProduct = getProduct(sProdID)		
			sProdName = aProduct(0)
			sProdPrice = aProduct(1)
			'Confirm_SetProdPrice1 'SFAE
			
			iProdAttrNum = aProduct(2)
			If iShipped <> 1 Then iShipped = getShipped(sProdID)
			ReDim Preserve aProduct(6)
			aProduct(3)	= iQuantity
			aProduct(4) = sProdID			
		
			' Store values		
			aAllProd(0,iProductCounter) = aProduct		
			
			' If not an array, then the product does not exist 
			If NOT IsArray(aProduct) Then
				Response.Write "<br>Product Does Not Exist"
			Else
				If NOT IsNumeric(iProdAttrNum)Then 
					iProdAttrNum = 0
				End If	
							
				' Get Associated Attribute IDs in an array
				If iProdAttrNum > 0 Then							
					ReDim aProdAttrID(iProdAttrNum)
					aProdAttrID = getProdAttr("odrattrtmp",iTmpOrderID,iProdAttrNum)					
				End If
	
				sAttrUnitPrice = 0
						
				' Iterate Through Attributes
				If iProdAttrNum > 0 And IsArray(aProdAttrID) Then
					Dim sAttrSubtotal, aAttrDetails, sAttrName, sAttrPrice, iAttrType
						For iCounter = 0 To iProdAttrNum - 1
							aAttrDetails = getAttrDetails(aProdAttrID(iCounter))												
							sAttrName = aAttrDetails(0)
							sAttrPrice = aAttrDetails(1)
							iAttrType = aAttrDetails(2)							
								
							' Calculate Subtotal
							sAttrUnitPrice =  getAttrUnitPrice(sAttrUnitPrice,sAttrPrice,iAttrType)	
							' Store Values	
							
							aAllProd(iCounter+1,iProductCounter) = aAttrDetails								
						Next
				Elseif iProdAttrNum > 0 And NOT IsArray(aProdAttrID) Then 
						If vDebug = 1 Then Response.Write "<br>Error: No Attributes found for " & iTmpOrderID
						If vDebug = 1 Then Response.Write "<br>Deleting from Saved Orders. Sorry for the inconvenience."
														
						Call setDeleteOrder("odrdttmp",iTmpOrderID)
						If vDebug = 1 Then Response.Write "Deleted: " & iTmpOrderID & "</font>"						
				' End Product Attribute If
				End If	

				dUnitPrice = cdbl(cDbl(sAttrUnitPrice) + cDbl(sProdPrice))
				Confirm_SetProdPrice1 'SFAE
				
				' Store Attribute Affected Price
				aAllProd(0,iProductCounter)(5) = dUnitPrice	
               If Application("AppName") = "StoreFrontAE" Then 'SFAE
				aprodinfoAE(iProductCounter, 1) = dunitprice
			   End IF	
				
				' Set Unit Price for Product
				sUnitPrice = FormatCurrency(dUnitPrice)
				
				dProductSubtotal = iQuantity * (dUnitPrice)
				sProductSubtotal = FormatCurrency(dProductSubtotal)
				'sTotalPrice = sTotalPrice + cDbl(dProductSubtotal)	'SFUPDATE			
			' End IsArray If
			End If

			OVC_ShowGiftWrapValue 31 'SFAE
			OVC_ShowBackOrderMessage 31 'SFAE
		
			sTotalPrice = sTotalPrice + cDbl(dProductSubtotal) 'SFUPDATE
			
			
			
			
			iProductCounter = iProductCounter + 1
			
			rsAllOrders.MoveNext		
		Loop
End If 
		' object cleanup
		closeObj(rsAllOrders)
		 
		'sShipping = GetShipping()	'from incConfirm.asp	
		sShipping = cdbl(Session("sShipping")) 'SFUPDATE
		
		If iHandling <> 1 Or iShipped <> 1 Then sHandling = 0
		
		OVC_SaveSubTotalWOD 'SFAE 
		
		If Application("AppName") = "StoreFront" Then 'SFUPDATE
			sTotalPrice = getGlobalSalePrice(sTotalPrice) 
		End If
		
		If Application("AppName") = "StoreFrontAE" Then 'SFAE
			sTotalPrice =  ApplyALLDiscounts(sTotalPrice,"Total") '.3008
			Session("sTotalPrice") = cdbl(sTotalPrice)
		End IF
		
		
	
	 dTaxAble_Amount =  sTotalPrice

		'#383
			iSTax = iSTax + cDbl(getTax("State", sShipping, dTaxAble_Amount, sProdID))
			iCTax = iCTax + cDbl(getTax("Country", sShipping, dTaxAble_Amount, sProdID))
  
 
	'#383
		sGrandTotal = (cDbl(sTotalPrice) + cDbl(sHandling) + cDbl(sShipping) + cDbl(iSTax) + cDbl(iCTax) + cDbl(iCodAmount))
		
	
    
		
		
		If Application("AppName") = "StoreFront" Then 'SFUPDATE
			sTotalPriceOut = cdbl(sTotalPrice)
			sGrandTotalOut = cdbl(sGrandTotal) 
		End If
		
		If Application("AppName") = "StoreFrontAE" Then 'SFAE
			sTotalPriceOut = cdbl(sTotalPrice) '.3008
			sGrandTotalOut = cdbl(sGrandTotal) 
			If Session("SpecialBilling") =  1 then 
				sGrandTotal = cdbl(Session("BillAmount"))
				Confirm_DivideAmountDiscount 'SFAE '.3008
				'sGrandTotalOut = cdbl(sGrandTotal) 
			Else
				Session("BillAmount") = cdbl(sGrandTotal)
				Session("BackOrderAmount") = 0
			End If
			WriteVars 'SFAE debug mode only
		END IF
							
				
		sTotalSTax = cStr(iSTax)
		sTotalCTax = cStr(iCTax)
		
	If Not Request.QueryString("message") <> "" AND Not Request.Form("custom") <> "" AND Not Request("CSVPOSRESPONSE") <> "" AND Not Request.QueryString("wpresponse") <> "" Then 
			'--------------------------------------------------
			' Begin Payment processing
			'--------------------------------------------------
		
			sPaymentMethod			= Trim(Request.Form("PaymentMethod"))
			sShipMethodName		= Trim(Request.Form("ShipMethodName"))	   
			iCustID					= Trim(Request.Cookies("sfCustomer")("custID"))
			iAddrID					= Trim(Request.Form("ShipID"))
	
			' collect billing address
			sCustFirstName			= Trim(Request.Form("FirstName"))
			sCustMiddleInitial		= Trim(Request.Form("MiddleInitial"))
			sCustLastName			= Trim(Request.Form("LastName"))
			sCustCompany			= Trim(Request.Form("Company"))
			sCustAddress1			= Trim(Request.Form("Address1"))
			sCustAddress2			= Trim(Request.Form("Addrress2"))
			sCustCity				= Trim(Request.Form("City"))
			sCustState				= Trim(Request.Form("State"))
			sCustStateName			= Trim(Request.Form("StateName"))
			sCustZip				= Trim(Request.Form("Zip"))
			sCustCountry			= Trim(Request.Form("Country"))
			sCustCountryName		= Trim(Request.Form("CountryName"))
			sCustPhone				= Trim(Request.Form("Phone"))
			sCustFax				= Trim(Request.Form("Fax"))
			sCustEmail				= Trim(Request.Form("Email"))
			bCustSubscribed		= Trim(Request.Form("IsSubscribed"))
			sPassword				= Trim(Request.Form("Password"))
			sCustName				= sCustFirstName & " " & sCustLastName
					
			sShipCustFirstName		= Trim(Request.Form("ShipFirstName"))
			sShipCustMiddleInitial	= Trim(Request.Form("ShipMiddleInitial"))
			sShipCustLastName		= Trim(Request.Form("ShipLastName"))		
			sShipCustCompany		= Trim(Request.Form("ShipCompany"))
			sShipCustAddress1		= Trim(Request.Form("ShipAddress1"))
			sShipCustAddress2		= Trim(Request.Form("ShipAddress2"))
			sShipCustCity			= Trim(Request.Form("ShipCity"))
			sShipCustStateName	= Trim(Request.Form("ShipStateName"))
			sShipCustState			= Trim(Request.Form("ShipState"))
			sShipCustZip			= Trim(Request.Form("ShipZip"))
			sShipCustCountry		= Trim(Request.Form("ShipCountry"))
			sShipCustCountryName	= Trim(Request.Form("ShipCountryName"))
			sShipCustPhone			= Trim(Request.Form("ShipPhone"))
			sShipCustFax			= Trim(Request.Form("ShipFax"))
			sShipCustEmail			= Trim(Request.Form("ShipEmail"))
			sShipInstructions		= Trim(Request.Form("SpecialInstructions"))
						  'security chk #249
            sShipInstructions		= Replace(sShipInstructions,"<"," " )
		    sShipInstructions		= Replace(sShipInstructions,">"," " )
			sShipInstructions		= Replace(sShipInstructions,chr(34)," " )
			sShipInstructions		= Replace(sShipInstructions,"'"," " ) ' #307
	
                   
			
			
			sShipCustName			= sShipCustFirstName & " " & sShipCustLastName	
			iShipType				= Trim(Request.Form("ShipType"))
		
			'Password for Email
			sMailPassword = sPassword
				
			' collect customer payment info		
			If sPaymentMethod = "Credit Card" AND sTransMethod <> "15" AND sTransMethod <> "18" Then
				sCustCardType			= Trim(Request.Form("CardType"))
				sCustCardName			= Trim(Request.Form("CardName"))
				sCustCardNumber		= Trim(Request.Form("CardNumber"))
				sCustCardExpiryMonth	= Trim(Request.Form("CardExpiryMonth"))
				If len(sCustCardExpiryMonth) = 1 Then
					sCustCardExpiryMonth = "0" & sCustCardExpiryMonth
				End If
				sCustCardExpiryYear		= Trim(Request.Form("CardExpiryYear"))
				sCustCardTypeName			= Trim(getNameWithID("sfTransactionTypes",sCustCardType,"transID","transName",0))
				sCustCardExpiry			= sCustCardExpiryMonth & "/" & sCustCardExpiryYear   
				iPayID = setPayments(sCustCardType,sCustCardName,sCustCardNumber,sCustCardExpiryMonth,sCustCardExpiryYear,iCC)

			ElseIf sPaymentMethod	= "eCheck" Then
				iRoutingNumber		= Trim(Request.Form("RoutingNumber"))
				sBankName				= Trim(Request.Form("BankName"))
				iCheckNumber			= Trim(Request.Form("CheckNumber"))
				iCheckingAccountNumber	= Trim(Request.Form("CheckingAccountNumber"))
			ElseIf sPaymentMethod = "PO" Then
				sPOName				= Trim(Request.Form("POName"))
				iPONumber				= Trim(Request.Form("PONumber"))
			ElseIf sPaymentMethod = "PhoneFax" Then
				If Trim(Request.Form("CardNumber")) <> "" Then
					sPhoneFaxPayType = "Credit Card"
					sCustCardType		= Trim(Request.Form("CardType"))
					sCustCardName		= Trim(Request.Form("CardName"))
					sCustCardNumber	= Trim(Request.Form("CardNumber"))
					sCustCardExpiryMonth= Trim(Request.Form("CardExpiryMonth"))
					If len(sCustCardExpiryMonth) = 1 Then
						sCustCardExpiryMonth = "0" & sCustCardExpiryMonth
					End If
					sCustCardExpiryYear = Trim(Request.Form("CardExpiryYear"))
					sCustCardTypeName			= Trim(getNameWithID("sfTransactionTypes",sCustCardType,"transID","transName",0)) '#289
					sCustCardExpiry		= sCustCardExpiryMonth & "/" & sCustCardExpiryYear 
				ElseIf Trim(Request.Form("RoutingNumber")) <> "" Then
					sPhoneFaxPayType = "eCheck"
					iRoutingNumber	= Trim(Request.Form("RoutingNumber"))
					sBankName			= Trim(Request.Form("BankName"))
					iCheckNumber		= Trim(Request.Form("CheckNumber"))
					iCheckingAccountNumber	= Trim(Request.Form("CheckingAccountNumber"))
				ElseIf Trim(Request.Form("PONumber")) <> "" Then
					sPhoneFaxPayType = "PO"
					sPOName		= Trim(Request.Form("POName"))
					iPONumber	= Trim(Request.Form("PONumber"))
				End If			 
			End If
		End If
	
	  	
	  If sPaymentMethod = "Credit Card" AND (sTransMethod <> "15" AND sTransMethod <> "18") Then
		iOrderID = setOrderInitial(iCustID,iPayID,iAddrID,sPaymentMethod,"","","","","","",sShipMethodName,sTotalSTax,sTotalCTax,sHandling,sShipping,sTotalPrice,sGrandTotal,sShipInstructions,"",aReferer)		
	  Elseif sPaymentMethod = "eCheck" Then
		iOrderID = setOrderInitial(iCustID,0,iAddrID,sPaymentMethod,iRoutingNumber,sBankName,iCheckNumber,iCheckingAccountNumber,"","",sShipMethodName,sTotalSTax,sTotalCTax,sHandling,sShipping,sTotalPrice,sGrandTotal,sShipInstructions,"",aReferer)
	  Elseif sPaymentMethod = "PO" Then
		iOrderID = setOrderInitial(iCustID,0,iAddrID,sPaymentMethod,"","","","",iPONumber,sPOName,sShipMethodName,sTotalSTax,sTotalCTax,sHandling,sShipping,sTotalPrice,sGrandTotal,sShipInstructions,"",aReferer)'changed iponame to sponame
	  Elseif sPaymentMethod = "PhoneFax" Then
	     If  iPhoneFaxType <> "1" OR  Application("AppName") = "StoreFrontAE" Then
		   iOrderID = setOrderInitial(iCustID,0,iAddrID,sPaymentMethod & "_" & sPhoneFaxPayType,"","","","","","",sShipMethodName,sTotalSTax,sTotalCTax,sHandling,sShipping,sTotalPrice,sGrandTotal,sShipInstructions,"",aReferer)
	     End if
	  ElseIf sPaymentMethod = "COD" Then
		iOrderID = setOrderInitial(iCustID,0,iAddrID,sPaymentMethod,"","","","","","",sShipMethodName,sTotalSTax,sTotalCTax,sHandling,sShipping,sTotalPrice,sGrandTotal,sShipInstructions,iCODAmount,aReferer)
	  ElseIf (sPaymentMethod = "PayPal Transaction" OR sPaymentMethod = "PayPal" OR sPaymentMethod = "WorldPay") OR (sTransMethod = "15" OR sTransMethod = "18" OR sTransMethod = "5" OR sTransMethod = "6" OR sTransMethod = "12") Then
		iOrderID = setOrderInitial(iCustID,0,iAddrID,sPaymentMethod,"","","","","","",sShipMethodName,sTotalSTax,sTotalCTax,sHandling,sShipping,sTotalPrice,sGrandTotal,sShipInstructions,"",aReferer)
End If	
' Move from tmp order cart to orders   
	  If iPhoneFaxType = "" Then		 
	     Call setOrder(iOrderID)
	  ElseIf iPhoneFaxType  <> "1" OR Application("AppName") = "StoreFrontAE" Then 
		 Call setOrder(iOrderID)
	  End If 
		  
	Confirm_WriteAERecords 'SFAE	  
	  
  	' Collect values from admin table, again because some variables get lost
		aAdmin = getAdminRow(sPaymentMethod)	
		' Trans method
		sTransMethod = Trim(aAdmin(0))
		' Payment server
		sPaymentServer = Trim(aAdmin(1))
		' Login 
		sLogin = Trim(aAdmin(2))
		' Password
		sPassword = Trim(aAdmin(3))
		' MerchantType
		sMercType = Trim(aAdmin(4))
		' Mail Method
		sMailMethod = Trim(aAdmin(5))
		
	' Process Order

	  If sPaymentMethod = "Credit Card" or sPaymentMethod = "PayPal" Then
	     proc_live = getProcStat(sTransMethod)
		
		If cdbl(sGrandTotal) > 0 Then
		
		If sTransMethod = "15" OR sPaymentMethod = "PayPal" Then			
			sProcErrMsg = PayPal()	
		ElseIf sTransMethod = "1" Then
			sProcErrMsg =  CyberCash(proc_live)
		ElseIf sTransMethod = "16" Then
			'sProcErrMsg = CyberCash(proc_live,"1")
			sProcErrMsg = CyberCash(proc_live)
		ElseIf sTransMethod = "2" or sTransMethod = "11" or sTransMethod = "13" or sTransMethod = "19" Then
			sProcErrMsg =  AuthNet(proc_live,"1")
		ElseIf sTransMethod = "3" or sTransMethod = "17" Then
			sProcErrMsg =  SignioPayProFlow(proc_live)
		ElseIf sTransMethod = "4" Then
			sProcErrMsg =  SurePay(proc_live)
 	   	ElseIf sTransMethod = "7" Then
			sProcErrMsg =  LinkPoint(proc_live)			
		ElseIf sTransMethod = "8" Then
			sProcErrMsg = PSIGate(proc_live)
		ElseIf sTransMethod = "10" Then
			sProcErrMsg = SecurePay(proc_live)		
		ElseIf sTransMethod = "18" OR sPaymentMethod = "WorldPay" Then
			sProcErrMsg = WorldPay(proc_live)
		End If
	  End If	
	End If  

		If Request.item("custom") <> "" Then
			Call PayPalResp("2")
		ElseIf Request.QueryString("wpresponse") <> "" Then 
			Call WorldPayResp("2")
		End If
	
	  If Trim(sProcErrMsg) = "" Then

  	  
		' Expire CustID before processing order
		  Response.Cookies("sfCustomer").Expires = Now()       	

		    If Application("AppName") = "StoreFrontAE" Then 'SFAE
			Confirm_UpdateInventory 	
			end if

		' Delete from tmpOrders  
		  For iCounter = 0 to Ubound(aOrderID)-1
			Call setDeleteOrder("odrdttmp",aOrderID(iCounter))
			If vDebug = 1 Then Response.write "<br>Deleting:" & aOrderID(iCounter)
    				Next	  		    	
		
	' End If statement from Proc Response up top 	


			Confirm_SaveGrandTotal 'SFAE	  
		    Confirm_SaveAmounts 'SFAE
	If  iPhoneFaxType <> "1"  Then
     
     		' Set Order complete flag to 1
			Call setOrderComplete(iOrderID)
			
			' Begin email 
			 Call createMail("Confirm",sCustEmail)
	
			' Set Cookie For NewOrder Page
			  Response.Cookies("ReturningOrder") = iOrderID
			  Response.Cookies("ReturningOrder").Expires = Date() + 31   
	  End If
	  
	  
	  End If
	  
'@ENDCODE
%>
<html>
<head>
<script Language="Javascript">
function printthis(){ 
 	if (NS) {
		 window.print() ; 
 	} else {
 		var WebBrowser = '<OBJECT ID="WebBrowser1" WIDTH=0 HEIGHT=0 CLASSID="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2"></OBJECT>';
 		document.body.insertAdjacentHTML('beforeEnd', WebBrowser);
 		WebBrowser1.ExecWB(6, 2);//Use a 1 vs. a 2 for a prompting dialog box WebBrowser1.outerHTML = ""; 
 	}
 }
</script>

<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>TRIBUTE TEA CHECKOUT - 3. Complete Order</title>



<!--Header Begin -->
<link rel="stylesheet" href="sfCSS.css" type="text/css">
</head>

<body  bgproperties="fixed" link="<%= C_LINK %>" vlink="<%= C_VLINK %>" alink="<%= C_ALINK %>">
	<table border="0" cellpadding="1" cellspacing="0" class="tdbackgrnd" width="<%= C_WIDTH %>" align="center">
	  <tr>
    	<td>
	      <table width="100%" border="0" cellspacing="1" cellpadding="3">
    	    <tr>
          <td align="middle"  class="tdTopBanner"> 
            <%If C_BNRBKGRND = "" Then%>
            <img src="buttons/tt_blue.gif" border="0" width="275" height="36"> 
            <%Else%>
            <%End If%>
          </td>
    	    </tr>
<!--Header End -->
        	<tr>
	          <td colspan="2" align="center" class="tdMiddleTopBanner">
			   </td>
	        </tr>
    	    <tr>
        	  
          <td align="left" colspan="2" class="tdBottomTopBanner"> 
            <div align="center">Thank you for submitting your order.<br>
              Below is a summary of your transaction.<br>
              You will receive confirmation by e-mail shortly.</div>
          </td>
	        </tr>
    	    <tr>
        	  
          <td colspan="2" width="100%" align="center" class="tdContent" valign="center"><font color="#990000">Step 
            1: Customer Information | Step 2: Payment Details | <b>Step 3: <i>COMPLETE 
            ORDER</i></b> </font></td>
	        </tr>
    	    <tr>
	    	  <td align="center" class="tdContent2" width="100%" colspan="2">        
            	<table border="0" width="100%" cellspacing="0" cellpadding="4">            
	              <% If iPhoneFaxType <> 1 Then %>
					<tr align="center">
					  <td colspan="4" width="100%" align="center">				
						<table width="85%"  cellpadding="1" cellspacing="0" border="0" class="tdBottomTopBanner">   
						  <tr>
							<td align="center" width="100%">
							  <table width="100%"  cellpadding="8" cellspacing="0" class="tdContentBar">
								<tr>
								  
                            <td align="center" width="100%" class="tdAltBG1"><font class="Content_Large"> 
                              <% If Trim(sProcErrMsg) = "" Then %>
                              <b>Your Order ID# is <%= iOrderID %></b> </font> 
                              <br>
                              <font size="-1"> <b><br>
                              Please make a record of this number for future reference. 
                              </b></font> 
                              <% Else %>
                              <b><font color="#990000">An error occurred while 
                              processing your order</font></b><font color="#990000">. 
                              </font> 
                              <hr noshade width="100%" size="1">
								    <b>Error Message: <%= sProcErrMsg %> </b><br>
								    <a href="javascript:window.history.go(-2)">Resubmit</a>
								    <% End If %>
							     </td>
							  </tr>
						    </table>
					      </td>
					    </tr>
					  </table>
				    </td>
				  </tr>            
                  <% End If %>
                  <% If Trim(sProcErrMsg) = "" Then %>           
                  <tr>    
                    <td width="52%" class="tdContentBar">product</td>
                    <td width="16%" align="center" class="tdContentBar">unit price</td>
                    <td width="16%" align="center" class="tdContentBar">qty</td>
                    <td width="16%" align="center" class="tdContentBar">price</td>
                    
                  </tr>
			      <%
			iProductCounter = 0
			
			For iCounter = 0 To iProdNumber - 1	
				
				iProductCounter = iProductCounter + 1
		dim fontclass
				' Do alternating colors and fonts	
				If (iProductCounter mod 2) = 1 Then 
					fontclass="tdAltFont1"
				Else 	
					fontclass="tdAltFont2"
				End If	
				
				sProdName = aAllProd(0,iCounter)(0)
				sProdPrice = aAllProd(0,iCounter)(1)
				iProdAttrNum = aAllProd(0,iCounter)(2)
				iQuantity = aAllProd(0,iCounter)(3)
				sProdID = aAllProd(0,iCounter)(4)
				'Confirm_SetProdPrice2 'SFAE
				
   %>
			      <tr class='<%=fontClass%>'>	 
		            <td width="52%" valign="top"><b><%= sProdName %></b><br>
                      <%
						' Initially 0
						sAttrUnitPrice = 0
						
						' Iterate Through Attributes
					If (iProdAttrNum > 0) Then 
						For iRow = 1 to iProdAttrNum
								sAttrName = aAllProd(iRow,iCounter)(0)	
								sAttrPrice = aAllProd(iRow,iCounter)(1)	
								iAttrType = aAllProd(iRow,iCounter)(2)			
				
							   sAttrUnitPrice =  getAttrUnitPrice(sAttrUnitPrice,sAttrPrice,iAttrType)								
							  %>										                
				              &nbsp;&nbsp;<%=sAttrName%>
            		          <br>                									
				          <%		
						' iRow For loop
					    Next
					End If	
		
					dUnitPrice = cdbl(cDbl(sAttrUnitPrice) + cDbl(sProdPrice))
					Confirm_SetProdPrice2 'SFAE
										
					' Set Unit Price for Product
					If iConverion = 1 Then
						sUnitPrice = "<script> document.write(""" & FormatCurrency(dUnitPrice) & " = ("" + OANDAconvert(" & dUnitPrice & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script>" 
					Else
						sUnitPrice = FormatCurrency(dUnitPrice)
					End If
					dProductSubtotal = iQuantity * (dUnitPrice)
					If iConverion = 1 Then
						sProductSubtotal = "<script> document.write(""" & FormatCurrency(dProductSubtotal) & " = ("" + OANDAconvert(" & dProductSubtotal & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script>" 
					Else
						sProductSubtotal = FormatCurrency(dProductSubtotal)
					End If
					'sProductSubtotal = FormatCurrency(dProductSubtotal)
				
				
				%>	
                    </td>
                    <td nowrap width="16%" align="center" class='<%=fontClass%>' valign="top"><%= sUnitPrice %></td>
			        <td nowrap width="16%" align="center" class='<%=fontClass%>' valign="top"><%= iQuantity %></td>
			        <td nowrap width="16%" align="center" class='<%=fontClass%>' valign="top"><%= sProductSubtotal %></td>
				  </tr>
					<%OVC_ShowGiftWrapValue 3 'SFAE%>
	                <% 	Next %>
	            </table>
	          </td>
			</tr>
	        <tr>
				<td class="tdContent2"colspan="2">
		        	<table border="0" width="100%" cellspacing="0" cellpadding="2" class="tdContent2">
			         <%OVC_ShowOrderDiscounts'SFAE%>
		              
			            <tr>
							<td nowrap width="75%" align="right"><b>Sub Total:</b></td>
		                    <td nowrap width="25%" height="20"><b>
							 <%
							If iConverion = 1 Then
								Response.Write "<script> document.write(""" & FormatCurrency(sTotalPriceOut) & " = ("" + OANDAconvert(" & sTotalPriceOut & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script>" 
							Else
								Response.write FormatCurrency(sTotalPriceOut)
							End If
							'Response.End 'temp
							%></b>
							</td>
		              		<td width="5%" align="center"></td>
		            	</tr>
		            	<% If sHandling <> 0 Then %>
		            	<tr>
		              		<td width="75%" align="right">Handling:</td>
		             		<td nowrap width="25%" height="20">
											<%
								If iConverion = 1 Then
									Response.Write "<script> document.write(""" & FormatCurrency(sHandling) & " = ("" + OANDAconvert(" & sHandling & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script>" 
								Else
									Response.write FormatCurrency(sHandling)
								End If
								%>
							</td>
		              		<td width="5%" align="center"></td>
		            	</tr>
										<% 
							Else
								sHandling = 0
							End If 
							%>
										<% If iCodAmount <> 0 Then %>
		           	 <tr>
		             	<td width="75%" align="right">COD Charge:</td>
		                <td width="25%" height="20" nowrap>
									<%
						If iConverion = 1 Then
							Response.Write "<script> document.write(""" & FormatCurrency(iCodAmount) & " = ("" + OANDAconvert(" & iCodAmount & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script>" 
						Else
							Response.write FormatCurrency(iCodAmount)
						End If
						%>
						</td>
		                <td width="5%" align="center"></td>
		            </tr>
		            <% End If %>
		            <% If sShipping <> 0 OR trim(sShipMethodName) = "Free Shipping" Then %>
		            <tr>
		              <td width="75%" align="right"><%= sShipMethodName %>:</td>
		              <td nowrap width="25%" height="20">
		                <%
						If iConverion = 1 Then
							Response.Write "<script> document.write(""" & FormatCurrency(sShipping) & " = ("" + OANDAconvert(" & sShipping & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script>" 
						Else
							Response.write FormatCurrency(sShipping)
						End If
						%>		
					  </td>
		              <td width="5%" align="center"></td>
		            </tr>
		            <% End If %>
		            <% If iCTax <> 0 Then %>
		            <tr>
		              <td width="75%" align="right">Country Tax:</td>
		              <td nowrap width="25%" height="20">
									<%
						If iConverion = 1 Then
							Response.Write "<script> document.write(""" & FormatCurrency(iCTax) & " = ("" + OANDAconvert(" & iCTax & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script>" 
						Else
							Response.write FormatCurrency(iCTax)
						End If
						%>
						</td>
                        <td width="5%" align="center"></td>
		            </tr>
		            <% End If %>
		            <% If iSTax <> 0 Then %>
		            <tr>
		              <td width="75%" align="right">State Tax:</td>
		              <td nowrap width="25%" height="20">
									<%
						If iConverion = 1 Then
							Response.Write "<script> document.write(""" & FormatCurrency(iSTax) & " = ("" + OANDAconvert(" & iSTax & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script>" 
						Else
							Response.write FormatCurrency(iSTax)
						End If
						%>	
						</td>
		                <td width="5%" align="center"></td>
		            </tr>
		            <% End If %>
		            <tr>
		              <td width="75%" align="right"><b>Grand Total:</b></td>
		              <td nowrap width="25%" height="20"><b>
											<%
							If iConverion = 1 Then
								Response.Write "<script> document.write(""" & FormatCurrency(sGrandTotalOut) & " = ("" + OANDAconvert(" & sGrandTotalOut & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script>" 'sfae beta2
							Else
								Response.write FormatCurrency(sGrandTotalOut) 'sfaebeta2
							End If
							Confirm_ShowBillingInfo 'sfae beta2 %>	
							
		                </b>
						</td>
						  <% End If ' End ProcResponse Failed If %>	
     		            <td width="5%" align="center"></td>
		            </tr>
		            	<% 	If iShipType = 2 Then %>	          
	                <tr>
		         
		  		        <td width="80%" colspan="4">
							
                  <table>
                    <tr> 
                      <td height="55"> 
                        <% If InStr(sShipMethodName,"UPS") Then %>
                        <img src="images/logo_s.gif"> 
                        <% Else %>
                        &nbsp; 
                        <% End If %>
                      </td>
                      <td width="597" valign="top">
                        <div align="center"><font class="Content_Small"> </font><a href="http://www.tributetea.com/"> 
                          <script language="Javascript"> 
		 var NS = (navigator.appName == "Netscape");
		 var VERSION = parseInt(navigator.appVersion);
		 if (VERSION > 3) {
		 	document.write('<form><input type=button value="Print this Page" name="Print" onClick="printthis()"></form>'); 
		 }
	                                        </script>
                          <br>
                          <img src="buttons/logout.gif" width="75" height="23" border="0"></a></div>
                      </td>
								</tr>
                  </table>
						</td>
					</tr>
						<% End If %>

	          </table>
	
	       </td>
		</tr>
	          <% If sPaymentMethod = "eCheck" Then %>
	    <tr>
	       <td colspan="2" class="tdContent2">
	              <table border="1" width="95%" class="tdAltBG2" cellpadding="1" cellspacing="0" bordercolor="#000000" align="center">
				      <tr><td>
					      <table border="0" cellpadding="4" cellspacing="0" width="100%" align="center">
					        <tr>
						      <td align="left"><b><font class="ECheck"><%= sCustFirstName & " " & sCustLastName %></font></b></td>
						      <td align="right"><font class="ECheck2"><b>Check Number: </b><%= iCheckNumber %></font></td>
					        </tr>
					        <tr>
						      <td align="left" colspan="2"><font class="ECheck"><%= sCustAddress1 %>
						        <br><%= sCustAddress2 %></font></td>
					
					          <tr>
						        <td align="left" colspan="2"><font class="ECheck"><%= sCustCity & " "%> <%= sCustStateName %>, <%= sCustZip %></font></td>
					          </tr>
					          <tr>	
						        <td align="left" colspan="2"><font class="ECheck"><%= sCustCountryName %></font></td>						
					          </tr>
					          <tr>
						        <td align="left" colspan="2" height="10"></td>
					          </tr>
					          <tr>
					            <td align="center" colspan="2"><b><font class="ECheck">Pay the amount of : <%= FormatCurrency(cDbl(sTotalPrice) + cDbl(sHandling) + cDbl(sShipping) + cDbl(iSTax) + cDbl(iCTax))%></font></b>
					              <hr size="1" width="60%" color="#445566" align="center">
					            </td>
					          </tr> <%'added this close %>
					            <tr>
						          <td width="50%" height="20">&nbsp;</td>
						          <td align="center"><font class="ECheck2">									    
						Electronically Signed By: <b><%= sCustFirstName & " " & sCustLastName %></b></font>
						            <hr size="1" width="100%" class="tdAltBG1"></td>
					              </tr>	
					              <tr>
						            <td align="left" colspan="2"><b><font class="ECheck"><%= sBankName %></font></b></td>
					              </tr>
					              <tr>
				      					<td colspan="2" align="center" class="tdAltBG1"><font class="ECheck2">Payment Authorized by Account Holder. Indemnification Agreement Provided by Depositor.</font></td>
				      				</tr>	
					                <tr>
						              <td colspan="2" align="center"><font class="ECheck2"><b><%= iRoutingNumber %>::<%= iCheckingAccountNumber %> </b></font></td>
						              </tr>
					</table>
			       </td>
				  </tr>
			     </table>
	           </td>
	    </tr>
	                    <% End If %>

	                    <% If sPaymentMethod = "PhoneFax" Then %>
	    <tr>
	       <td colspan="2" class="tdContent2">
           		<table class="tdContent" border="0" width="100%" cellpadding="2" cellspacing="0">
		                      <tr><td width="100%" colspan="2">
			                      <table border="0" width="100%" cellspacing="0" cellpadding="3" class="tdContent2">
			                        <tr>
			                          <td width="100%" class="tdContentBar">Phone/Fax Printout</td>
			                        </tr>
			                      </table>
		                        </td></tr>	
	
	                          <!--Customer Information -->
	                          <tr><td width="100%" class="tdContent2" colspan="2">   
	                              <table border="0" width="100%" cellspacing="0" cellpadding="2">
	                                <tr>
		                              <td><b>Billing:</b></td>
		                              <td><b>Ship To:</b></td>
	                                </tr>
	                                <tr>
		                              <td><%= sCustFirstName %>&nbsp;&nbsp;<%= sCustMiddleInitial %>&nbsp;&nbsp;<%= sCustLastName %></td>
		                              <td><%= sShipCustFirstName %>&nbsp;&nbsp;<%= sShipCustMiddleInitial %>&nbsp;&nbsp;<%= sShipCustLastName %></td>
	                                </tr>
	                                <tr>
		                              <td><%= sCustCompany %></td>
		                              <td><%= sShipCustCompany%></td>
	                                </tr>
	                                <tr>
		                              <td><%= sCustAddress1 %></td>
		                              <td><%= sShipCustAddress1 %></td>
	                                </tr>       
                                    <%If sCustAddress2 <> "" Then%>
   	                                <tr>	
   		                              <td><%= sCustAddress2 %></td>
                                      <td><%= sShipCustAddress2%></td>
                                    </tr>    
                                    <%End If%>
	                                <tr>
		                              <td><%= sCustCity%>,&nbsp;<%= sCustStateName %>,&nbsp;<%= sCustZip%></td>
		                              <td><%= sShipCustCity%>,&nbsp;<%= sShipCustStateName %>,&nbsp;<%= sShipCustZip %></td>
	                                </tr>
	                                <tr>
		                              <td><%= sCustCountryName %></td>
		                              <td><%= sShipCustCountryName %></td>
	                                </tr>
	                                <tr>
		                              <td></td>
 		                              <td></td>
 	                                </tr>
	                                <tr>
		                              <td><%= sCustFax%></td>
		                              <td><%= sShipCustFax %></td>
	                                </tr>
	                                <tr>
		                              <td><%= sCustEmail%></td>
    	                              <td><%= sShipCustEmail %></td>
                                    </tr>
                                    <tr>
			                          <td width="100%" align="left" colspan="4">	
			                            <br><b>Special Instructions:</b></td>
		                            </tr>
		                            <tr>
		                              <td><%= sShipInstructions%></td>
    	                            </tr>
                                  </table>

	                            </td>
	                          </tr>
	
	                          <tr><td width="100%" class="tdContent2" colspan="2">   
	                              <table border="0" width="100%" cellspacing="0" cellpadding="2">
	                                <% 	If sPhoneFaxPayType = "Credit Card" Then %>
		                            <tr>
			                          <td width="100%" align="left" colspan="4">	
			                            <br><b>Credit Card information:</b></td>
		                            </tr>
		                            <tr>
			                          <td align="left">Card Type:</td><td align="left"><%= sCustCardTypeName %></td>
			                          <td align="left">Card Name:</td><td align="left"><%= sCustCardName %></td>
		                            </tr>
		                            <tr>
			                          <td align="left">Credit Card Number:</td><td align="left"><%= sCustCardNumber %></td>
			                          <td align="left">Credit Card Expiration Date:</td><td align="left"><%= sCustCardExpiry %></td>
		                            </tr>		
	                                <%	ElseIf sPhoneFaxPayType = "eCheck" Then %>
		                            <tr>
			                          <td width="100%" align="left" colspan="4">	
			                            <br><b>e-Check information:</b></td>
		                            </tr>
		                            <tr>
			                          <td>Account Number:</td> <td align="left"><%= iCheckingAccountNumber %></td>
			                          <td>Check Number:</td> <td align="left"><%= iCheckNumber %></td>
		                            </tr>
		                            <tr>
			                          <td>Bank Name:</td> <td align="left"><%= sBankName %></td>
			                          <td>Routing Number:</td> <td align="left"><%= iRoutingNumber %></td>
		                            </tr>	
		                            <tr>
			                          <td height="20" colspan="4">&nbsp;</td>
		                            </tr>
		                            <tr><td width="100%" colspan="4" align="center">	
			                            <table border="1" width="95%" class="tdAltBG2" cellpadding="1" cellspacing="0" bordercolor="#000000" align="center">
				                          <tr>
				                            <td align="center" class="tdAltBG1"><font class="ECheck2">Payment Authorized by Account Holder. Indemnification Agreement Provided by Depositor.</font></td>
				                            </tr>	
				                            <tr><td>
					                            <table border="0" cellpadding="4" cellspacing="0" width="100%" align="center">
					                              <tr>
						                            <td align="left"><b><font class="ECheck"><%= sCustFirstName & " " & sCustLastName %></font></b></td>
						                            <td align="right"><font class="ECheck2"><b>Check Number: </b><%= iCheckNumber %></font></td>
					                              </tr>
					                              <tr>
						                            <td align="left" colspan="2"><font class="ECheck"><%= sCustAddress1 %>
						                              <br><%= sCustAddress2 %></font></td>
					
					                                <tr>
						                              <td align="left" colspan="2"><font class="ECheck"><%= sCustCity & " "%> <%= sCustStateName %>, <%= sCustZip %></font></td>
					                                </tr>
					                                <tr>	
						                              <td align="left" colspan="2"><font class="ECheck"><%= sCustCountryName %></font></td>						
					                                </tr>
					                                <tr>
						                              <td align="left" colspan="2" height="10"></td>
					                                </tr>
					                                <tr>
					                                  <td align="center" colspan="2"><b><font class="ECheck">Pay the amount of : <%= FormatCurrency(cDbl(sTotalPrice) + cDbl(sHandling) + cDbl(sShipping) + cDbl(iSTax) + cDbl(iCTax))%></font></b>
					                                    <hr size="1" width="60%" color="#445566" align="center">
					                                  </td>
					                                  <tr>
						                                <td width="50%" height="20">&nbsp;</td>
						                                <td align="center"><font class="ECheck2">									    
						Electronically Signed By: <b><%= sCustFirstName & " " & sCustLastName %></b></font>
						                                  <hr size="1" width="100%" class="tdAltBG1"></td>
					                                    </tr>	
					                                    <tr>
						                                  <td align="left" colspan="2"><b><font class="ECheck"><%= sBankName %></font></b></td>
					                                    </tr>
					                                    <tr>
						                                  <td align="left"><font class="ECheck2"><b>Routing Number: </b> <%= iRoutingNumber %></font></td>
						                                  <td align="right"><font class="ECheck2"><b>Checking Account Number: </b> <%= iCheckingAccountNumber %> </font></td>
					                                    </tr>
					                                    <tr><td colspan="2" height="25"></td></tr>
					                                  </table>
			                                        </td></tr>
			                                    </table>
		                                      </td></tr></table>		
	                                    </td></tr>	
	                                    <% ElseIf sPhoneFaxPayType = "PO" Then %>
	                                    <tr>
		                                  <td width="100%" align="left" colspan="4">	
		                                    <br><b>Purchase Order information:</b></td>
	                                    </tr>
	                                    <tr>
		                                  <td width="25%" align="left">Name:</td><td width="25%" align="left"><%=	sPOName %></td>
		                                  <td width="25%" align="left">Purchase Order Number:</td><td width="25%" align="left"><%= iPONumber %></td>
	                                    </tr>
	                                    <% End If	%>
	                                    <tr>
	                                      <td align="center" colspan="4" height="60" valign="center">
		                                    <script Language="Javascript"> 
		 var NS = (navigator.appName == "Netscape");
		 var VERSION = parseInt(navigator.appVersion);
		 if (VERSION > 3) {
		 	document.write('<form><input type=button value="Print this Page" name="Print" onClick="printthis()"></form>'); 
		 }
	                                        </script>
	                                      </td>
	                                    </tr>
	      </table>
		  
		  

	                                </td>
	                              </tr>
	                            
	                            <% End If %>
	                    <!--Footer begin-->
                <!--#include file="foot.txt"-->
              </table>
            </td>
          </tr>
        </table>
       </body>

    </html>
<!--Footer End-->
   <% 
If  iPhoneFaxType = "1" then
 DeleteOrder iOrderId
end if   
   
If Trim(sProcErrMsg) = "" Then
   Response.Cookies("sfCustomer").Expires = Now()
   Response.Cookies(Session("GeneratedKey") & "sfOrder").Expires = Now() 
   Response.Cookies("EndSession") = Session("SessionID")
   Response.Cookies("EndSession").Expires = Date() + 25
   Session("SSLSession") = ""
   closeObj(cnn)
   Session.Abandon 
  End If
%>
