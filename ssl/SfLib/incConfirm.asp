<%
'@BEGINVERSIONINFO

'@APPVERSION: 550.4011.0.2
'@FILENAME: incConfirm.asp
	 

'@DESCRIPTION: Include File for Confirm.asp 

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws as an unpublished work, and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO
'------------------------------------------------------------------
' Checks PhoneFax type
'log #181 djp
'modified 10-29-01


' Returns 0 for recorded, 1 for non-recorded
'------------------------------------------------------------------
Function CheckPhoneFax
	Dim sLocalSQL, rsTrans, bType, Recorded
	
	sLocalSQL = "SELECT transName FROM sfTransactionTypes WHERE transType = 'PhoneFax' AND transIsActive = 1"
	Set rsTrans = Server.CreateObject("ADODB.RecordSet")
	rsTrans.Open sLocalSQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText

	If rsTrans.EOF Or rsTrans.BOF Then
		bType = 0
	Else
		If trim(rsTrans.Fields("transName")) = "Recorded" Then
			bType = "0"
		Else
			bType = "1"
		End If		
	End If
	closeObj(rsTrans)
	CheckPhoneFax = bType
End Function
'--------------------------------------------------------
' getAdminRow 
' Returns an array of Admin properties
'--------------------------------------------------------
Function getAdminRow(sPaymentMethod)
	Dim sLocalSQL1, sLocalSQL2, rsAdmin, rsPPAdmin, aAdminRows, aPPAdminRows, aAdmin(7)	
	
	sLocalSQL1 = "SELECT adminTransMethod, adminPaymentServer, adminLogin, adminPassword, adminMerchantType, adminMailMethod, adminHandling FROM sfAdmin"
	
	Set rsAdmin = Server.CreateObject("ADODB.RecordSet")
	rsAdmin.Open sLocalSQL1, cnn, adOpenDynamic, adLockOptimistic, adCmdText	
	
			If sPaymentMethod = "PayPal" Then

			sLocalSQL2 = "SELECT trnsmthdServerPath, trnsmthdLogin, trnsmthdPasswd 					FROM sfTransactionMethods WHERE trnsmthdID = 15"
	
			Set rsPPAdmin = Server.CreateObject("ADODB.RecordSet")
			rsPPAdmin.Open sLocalSQL2, cnn, adOpenDynamic, adLockOptimistic, adCmdText	
	
				If rsPPAdmin.EOF Then
					Response.Write "<br>PayPal is not correctly configured"
				Else	
					aPPAdminRows = rsPPAdmin.GetRows
				End If	

			End If

	If  rsAdmin.EOF Then
		Response.Write "<br>No records in Admin Table"
	Else
		aAdminRows = rsAdmin.GetRows
		
		' Transaction Method for credit cards
		aAdmin(0) = aAdminRows(0,0)
		' Payment server
		If sPaymentMethod = "PayPal" Then
		' Payment server
		aAdmin(1) = aPPAdminRows(0,0)
		' Login 
		aAdmin(2) = aPPAdminRows(1,0)
		' Password
		aAdmin(3) = aPPAdminRows(2,0)
		Else
		' Payment server
		aAdmin(1) = aAdminRows(1,0)
		' Login 
		aAdmin(2) = aAdminRows(2,0)
		' Password
		aAdmin(3) = aAdminRows(3,0)
		End If
		' MerchantType
		aAdmin(4) = aAdminRows(4,0)	
		' Mail Method
		aAdmin(5) = aAdminRows(5,0)
		' Handling Amount
		aAdmin(6) = aAdminRows(6,0)

	End If

	getAdminRow = aAdmin
	
	closeobj(rsAdmin)
End Function

'----------------------------------------------------------
' setOrderInitial
' Returns the id of the order written to
'----------------------------------------------------------
Function setOrderInitial(iCustID,iPayID,iAddrID,sPaymentMethod,iRoutingNumber,sBankName,iCheckNumber,iCheckingAccountNumber,iPONumber,iPOName,sShipMethodName,sTotalSTax,sTotalCTax,sHandling,sShipping,sTotalPrice,sGrandTotal,sShipInstructions,iCODAmount,aReferer)
	Dim rsOrder, iOrderID, bookMark,sTotalHandling
	
	Set rsOrder = Server.CreateObject("ADODB.RecordSet")
	rsOrder.CursorLocation = adUseClient
	rsOrder.Open "sfOrders Order By orderID", cnn, adOpenDynamic, adLockOptimistic, adCmdTable
	If iCODAmount > "0" Then
	sTotalHandling = (CDbl(sHandling)+CDbl(iCODAmount))
	else
	sTotalHandling =  CDbl(sHandling)
	End If

	rsOrder.AddNew
	rsOrder.Fields("orderCustId")				= trim(iCustID)
	rsOrder.Fields("orderPayId")				= trim(iPayID)
	If IsNumeric(iAddrID)Then
		rsOrder.Fields("orderAddrId")			= trim(iAddrID)
	End If
	rsOrder.Fields("orderDate")					= now()
	rsOrder.Fields("orderAmount")				= trim(sTotalPrice)
	rsOrder.Fields("orderComments")			= trim(sShipInstructions)
	rsOrder.Fields("orderShipMethod")			= trim(sShipMethodName)
	rsOrder.Fields("orderSTax")					= trim(sTotalSTax)
	rsOrder.Fields("orderCTax")					= trim(sTotalCTax)
	rsOrder.Fields("orderHandling")			= trim(sTotalHandling)
	rsOrder.Fields("orderShippingAmount")		= trim(sShipping)
	rsOrder.Fields("orderGrandTotal")			= trim(sGrandTotal)
	rsOrder.Fields("orderPaymentMethod")		= trim(sPaymentMethod)
	rsOrder.Fields("orderCheckAcctNumber")	= trim(iCheckingAccountNumber)
	rsOrder.Fields("orderCheckNumber")			= trim(iCheckNumber)
	rsOrder.Fields("orderBankName")			= trim(sBankName)
	rsOrder.Fields("orderRoutingNumber")		= trim(iRoutingNumber)
	rsOrder.Fields("orderPurchaseOrderName")	= trim(iPOName)
	rsOrder.Fields("orderPurchaseOrderNumber")	= trim(iPONumber)	
	
	if isArray(aReferer) then
	   	 on error resume next
    	rsOrder.Fields("orderRemoteAddress")		= aReferer(2)
    	rsOrder.Fields("orderTradingPartner")    = aReferer(0)
    	rsOrder.Fields("orderHttpReferrer")      = aReferer(1)
	else
	   rsOrder.Fields("orderRemoteAddress")	   =""
	   rsOrder.Fields("orderTradingPartner")    = ""
	   rsOrder.Fields("orderHttpReferrer")      = ""
    end if
	
	rsOrder.Update
	
	'bookMark = rsOrder.AbsolutePosition 
	'rsOrder.Requery 
	'rsOrder.AbsolutePosition = bookMark
	
	iOrderID = rsOrder.Fields("orderID")	
	closeobj(rsOrder)
	setOrderInitial = iOrderID
End Function

'-----------------------------------------------------------------
' Sets the address as the default id while setting the rest to 0
'-----------------------------------------------------------------
Sub SetActive(sPrefix,iID)
	Dim rsActive, sLocalSQL
	
	Select Case sPrefix 
		Case "cshpaddr"
			sLocalSQL = "SELECT cshpaddrIsActive,cshpaddrID FROM sfCShipAddresses WHERE cshpaddrCustID = " & Request.Cookies("sfCustomer")("custID") 
		Case "pay"
			sLocalSQL = "SELECT payIsActive,payID FROM sfCPayments WHERE payCustId = " & Request.Cookies("sfCustomer")("custID") 
	End Select
		
	If vDebug = 1 Then 	Response.Write "<br>SetActive SQL: " & sLocalSQL
	' Set the old ship addresses to 0
	
	Set rsActive = Server.CreateObject("ADODB.RecordSet")		
		rsActive.Open sLocalSQL, cnn, adOpenDynamic,adLockOptimistic, adCmdText
			
		Do While NOT rsActive.EOF
			If rsActive.Fields(sPrefix & "ID") <> cInt(iID) Then
				rsActive.Fields(sPrefix & "IsActive") = 0	
				rsActive.Update	
			Else
				rsActive.Fields(sPrefix & "IsActive") = 1
				rsActive.Update
			End If			
			rsActive.MoveNext			
		Loop
	closeobj(rsActive)	
End Sub

'----------------------------------------------------------------------
' parses through credit card payments
' Returns the id if old card is detected and a -1 if new card detected
'----------------------------------------------------------------------
Function CheckCardChange(sCardType,sCardName,sCardNumber,sCardExpiryMonth,sCardExpiryYear)

Dim sLocalSQL, rsPayment, iCardID, aPayArray, sCardExpiryDate, iRow

	sLocalSQL = "SELECT payCardType, payCardName, payCardNumber, payCardExpires, payID FROM sfCPayments WHERE payCustId = " & Session("CustID")
	
	If vDebug = 1 Then 	Response.Write "<br>CheckCardChange SQL: " & sLocalSQL

	Set rsPayment = Server.CreateObject("ADODB.RecordSet")
	rsPayment.Open sLocalSQL, cnn, adOpenKeyset, adLockOptimistic, adCmdText	
	iCardID = -1
	
	If NOT rsPayment.EOF Then
			aPayArray = rsPayment.GetRows
	
			
			iRow = 0
			sCardExpiryDate = sCardExpiryMonth & "/" & sCardExpiryYear
	
			If rsPayment.RecordCount > 0 Then 
					For iRow = 0 to rsPayment.RecordCount - 1
					' Debug
					If vDebug = 1 Then 	Response.Write "<br>PayArray: " & iRow & " : " & aPayArray(0,iRow)
	
						' Loop through, if anything changes, raise the 1 boolean
						If ((aPayArray(0,iRow) <> sCardType) Or (aPayArray(1,iRow) <> sCardName) Or (aPayArray(2,iRow) <> sCardNumber) Or (aPayArray(3,iRow) <> sCardExpiryDate)) Then
							If vDebug = 1 Then Response.Write "<br>No Match"
						Else
							iCardID = aPayArray(4,iRow)
							If vDebug = 1 Then Response.Write "<br>Match : " & iCardID
							
							Call SetActive("pay",iCardID)
							closeobj(rsPayment)
							CheckCardChange = iCardID
							Exit Function	
						End If
					' iRow next		
					Next
			End If
	End If		
	closeobj(rsPayment)
	CheckCardChange = iCardID
End Function

'--------------------------------------------------------
' Write into sfCPayments
' Returns ID written to
'--------------------------------------------------------
Function setPayments(sCardType,sCardName,sCardNumber,sCardExpiryMonth,sCardExpiryYear,iCC)
	Dim rsPayments, sLocalSQL,ccObj,rsAdmin,iEncode
	
	Set rsAdmin = Server.CreateObject("ADODB.Recordset")
	sLocalSQL = "SELECT adminEncodeCCIsActive FROM sfAdmin"
	rsAdmin.Open sLocalSQL,cnn,adOpenForwardOnly,adLockReadOnly,adCmdText
	iEncode = trim(rsAdmin.Fields("adminEncodeCCIsActive"))
	closeObj(rsAdmin)
	If iEncode = "1" Then
		Set ccObj = Server.CreateObject("SFServer.CCEncrypt")
		ccObj.putSeed(iCC)
	End If
	
	Set rsPayments = Server.CreateObject("ADODB.RecordSet")
		rsPayments.CursorLocation = adUseClient
		
		 rsPayments.Open "sfCPayments", cnn, adOpenKeyset, adLockOptimistic, adCmdTable		
			rsPayments.AddNew
			rsPayments.Fields("payCustId")		= Session("custID")
			rsPayments.Fields("payCardType") 	= trim(sCardType)
			rsPayments.Fields("payCardName") 	= trim(sCardName)
			On Error Resume Next
			If iEncode = "1" Then
				rsPayments.Fields("payCardNumber")		= trim(ccObj.encrypt(sCardNumber))
			Else
				rsPayments.Fields("payCardNumber")		= trim(sCardNumber)
			End If
			If Err.number <> 0 Then rsPayments.Fields("payCardNumber") = trim(sCardNumber)
			On Error GoTo 0
			rsPayments.Fields("payCardExpires")		= trim(sCardExpiryMonth & "/" & sCardExpiryYear)
			rsPayments.Fields("payIsActive")		= trim(1)
			rsPayments.Update
		setPayments = rsPayments.Fields("payID")				
	closeObj(ccObj)	
	closeobj(rsPayments)
End Function

'--------------------------------------------------------
' Gets attribute number ' worst case scenario
'--------------------------------------------------------
Function getAttributeNumber
	Dim sLocalSQL, rsAttr, iCount
	
	sLocalSQL = "SELECT odrattrtmpID FROM sfTmpOrderAttributes INNER JOIN sfTmpOrderDetails ON sfTmpOrderAttributes.odrattrtmpOrderDetailId = sfTmpOrderDetails.odrdttmpID WHERE odrdttmpSessionID = " & Session("SessionID")
	Set rsAttr = Server.CreateObject("ADODB.RecordSet")
	rsAttr.Open sLocalSQL, cnn, adOpenKeyset, adLockOptimistic, adCmdText
	iCount = rsAttr.RecordCount
	closeobj(rsAttr)
	getAttributeNumber = iCount
End Function

'---------------------------------------------------------------------
' This function checks whether a product exists and retrieves an array of info
'---------------------------------------------------------------------
Function get6ProdValues(sProdID)
Dim rsSelectProd, sLocalSQL, aLocalProdArray(6), sCategoryName, sManufacturerName,sVendorName, iSaleIsActive

	sLocalSQL = "SELECT prodCategoryID, prodManufacturerID, prodVendorID, prodName, prodPrice, "_
  				& " prodAttrNum, prodSaleIsActive, prodSalePrice " _
			  	& " FROM sfProducts WHERE prodID = '"& sProdID & "' AND prodEnabledIsActive=1"

  	  Set rsSelectProd = Server.CreateObject("ADODB.RecordSet")
  	  rsSelectProd.Open sLocalSQL, cnn

	  If vDebug = 1 Then Response.Write "<p>getProdValues SQL : " & sLocalSQL
	  
	  sCategoryName = getNameWithID("sfCategories",rsSelectProd.Fields("prodCategoryID"),"catID","catName",0)
	  sManufacturerName = getNameWithID("sfManufacturers",rsSelectProd.Fields("prodManufacturerID"),"mfgID","mfgName",0)
	  sVendorName = getNameWithID("sfVendors",rsSelectProd.Fields("prodVendorID"),"vendID","vendName",0)
	  
  	' Check if this record exists through prodID and price matches
      If rsSelectProd.EOF or rsSelectProd.BOF Then
  		  Response.Write "<br>Empty Recordset in rsSelectProd"
  	  Else
  		  aLocalProdArray(0) = sCategoryName
  		  aLocalProdArray(1) = sManufacturerName
  		  aLocalProdArray(2) = sVendorName
  		  aLocalProdArray(3) = rsSelectProd.Fields("prodName")
  		  aLocalProdArray(4) = rsSelectProd.Fields("prodAttrNum")
 		  iSaleIsActive = rsSelectProd.Fields("prodSaleIsActive")
  		  If iSaleIsActive = 1 Then
  			aLocalProdArray(5) = rsSelectProd.Fields("prodSalePrice")
  		  Else
  		    aLocalProdArray(5) = rsSelectProd.Fields("prodPrice")
  		  End If

  	  End If
  	  closeObj(rsSelectProd)
	  
  	  get6ProdValues = aLocalProdArray
End Function

'---------------------------------------------------------------------
' Copies record from tmpOrders to orderDetails
' Return OrderID
'---------------------------------------------------------------------
Function setOrder(iOrderID)

Dim rsCopy, rsOrderDetail, rsAttr, rsOrderAttr, sLocalSQL, rsSvdCart, rsSvdCartAttr, iLocalID,sReferer,sDateTime,sTmpAttrName,sAttrType,sAttrPrice 
Dim sCategoryName, sManufacturerName, sVendorName, sProdName, iProdAttrNum, sProdPrice, sSubtotal
Dim sTmpAttrDtName, sTmpAttrPrice, sTmpAttrType, iTmpOrderID, sTmpAttrID, aAttr, bookMark


	sLocalSQL = "Select odrdttmpID,odrdttmpQuantity,odrdttmpProductID,odrdttmpHttpReferer FROM sfTmpOrderDetails WHERE odrdttmpSessionID = " & Session("SessionID")

	If vDebug = 1 Then Response.Write "<p>setOrder SQL : " & sLocalSQL

	Set rsCopy = Server.CreateObject("ADODB.RecordSet")
	rsCopy.Open sLocalSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText

	If (rsCopy.BOF Or rsCopy.EOF)Then
		If vDebug = 1 Then Response.Write "<br> Empty RecordSet in rsCopy"
		Response.Redirect "abandon.asp"
	Else
		
		' First set referer
		'sReferer = rsCopy.Fields("odrdttmpHttpReferer")		
		'Call SetReferer(sReferer,iOrderID)
			
		Do While NOT rsCopy.EOF 

				' Collect from TmpOrderDetails
			iTmpOrderID		= Trim(rsCopy.Fields("odrdttmpID"))
			sProdID			= Trim(rsCopy.Fields("odrdttmpProductID"))
			iQuantity		= Trim(rsCopy.Fields("odrdttmpQuantity"))
			sDateTime		= Trim(FormatDateTime(Now))
			
			' Product array
			Redim aProduct(6)
			aProduct			= get6ProdValues(sProdID)			
 			sCategoryName		= aProduct(0)
			sManufacturerName	= aProduct(1)		
			sVendorName		= aProduct(2)		
			sProdName			= aProduct(3)
			iProdAttrNum		= aProduct(4)
			sProdPrice			= aProduct(5)
		dUnitPrice=sProdPrice
		Order_SetProdPrice 'SFAE
		sProdPrice=dUnitPrice	
			sSubtotal = getSubtotal(iTmpOrderID,sProdPrice,iQuantity,iProdAttrNum)					
		
		
		
			' Get the Id Key			
			Set rsOrderDetail = Server.CreateObject("ADODB.RecordSet")
				rsOrderDetail.CursorLocation = adUseClient
				rsOrderDetail.Open "sfOrderDetails Order By odrdtID", cnn, adOpenDynamic, adLockOptimistic, adCmdTable
				
					rsOrderDetail.AddNew
					rsOrderDetail.Fields("odrdtOrderId")		= trim(iOrderID)
					rsOrderDetail.Fields("odrdtQuantity")		= trim(iQuantity)
					rsOrderDetail.Fields("odrdtSubTotal")		= trim(sSubtotal)
					rsOrderDetail.Fields("odrdtCategory")		= trim(sCategoryName)
					rsOrderDetail.Fields("odrdtManufacturer")	= trim(sManufacturerName)
					rsOrderDetail.Fields("odrdtVendor")		= trim(sVendorName)
					rsOrderDetail.Fields("odrdtProductName")	= trim(sProdName)
					rsOrderDetail.Fields("odrdtPrice")			= trim(sProdPrice)
					rsOrderDetail.Fields("odrdtProductId")		= trim(sProdID)
					rsOrderDetail.Update
					
					'bookMark = rsOrderDetail.AbsolutePosition 
					'rsOrderDetail.Requery 
					'rsOrderDetail.AbsolutePosition = bookMark
					
					iLocalID  = rsOrderDetail.Fields("odrdtID")
				closeobj(rsOrderDetail)	
	
			If vDebug = 1 Then Response.Write "<p><font size=6>OrderDetail ID = " & iLocalID & "</font>"
		
			sLocalSQL = "SELECT odrattrtmpAttrID FROM sfTmpOrderAttributes WHERE odrattrtmpOrderDetailId = " & iTmpOrderID
			Set rsAttr = Server.CreateObject("ADODB.RecordSet")
			rsAttr.Open sLocalSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText
			
				' Copy Attributes
				Do While Not rsAttr.EOF		
					' Collect Attribute Info from sfTmpOrderAttributes
					sTmpAttrID = rsAttr.Fields("odrattrtmpAttrID")
					aAttr = getAttrDetails(sTmpAttrID)
					sTmpAttrDtName	= aAttr(0)
					sTmpAttrPrice	= aAttr(1)
					sTmpAttrType	= aAttr(2)
					sTmpAttrName	= aAttr(3)
					
						Set rsOrderAttr = Server.CreateObject("ADODB.RecordSet")
						rsOrderAttr.Open "sfOrderAttributes", cnn, adOpenDynamic, adLockOptimistic, adCmdTable
							rsOrderAttr.AddNew
							rsOrderAttr.Fields("odrattrOrderDetailId")	= trim(iLocalID)
							rsOrderAttr.Fields("odrattrAttribute")		= trim(sTmpAttrDtName)
							rsOrderAttr.Fields("odrattrName")			= trim(sTmpAttrName)
							rsOrderAttr.Fields("odrattrPrice")			= trim(sTmpAttrPrice)
							rsOrderAttr.Fields("odrattrType")			= trim(sTmpAttrType)
						rsOrderAttr.Update					
									
					rsAttr.MoveNext
				Loop		
		rsCopy.MoveNext	
		Loop		
	End If
	
  	closeObj(rsCopy)
  	closeObj(rsAttr)
  	closeobj(rsOrderAttr)
  	setOrder = iLocalID
End Function

'---------------------------------------------------------------------
' Calculates New SubTotal
'---------------------------------------------------------------------
Function getSubtotal(iTmpOrderID,sProdPrice,iQuantity,iProdAttrNum)	
Dim sLocalSQL, dSubtotal,iCounter, rsSelectRow, dAttrSubtotal, iAttrType, aAttrDetails, sTmpAttrPrice

	' initialize	
	dSubtotal = 0
    
    ' Determine which tables to write to 
	If iProdAttrNum > 0 Then
		  sLocalSQL = "SELECT odrattrtmpAttrID FROM sfTmpOrderAttributes WHERE odrattrtmpOrderDetailId = " &iTmpOrderID
			
		  Set rsSelectRow = Server.CreateObject("ADODB.RecordSet")
  		  rsSelectRow.Open sLocalSQL, cnn, adOpenDynamic,adLockOptimistic,adCmdText  		  
		 	  
  		' Check if this record exists through prodID and price matches
		  If rsSelectRow.EOF or rsSelectRow.BOF Then
		      If vDebug = 1 Then Response.Write "<br>Recordset for New Subtotal is empty."
  		  Else
  			  
  		  dSubtotal = cLng(iQuantity) * cDbl(sProdPrice)
  		  iCounter = 0
  		  dAttrSubtotal = 0
  				
  				Do While Not rsSelectRow.EOF
  					Redim aAttrDetails(3)
  					aAttrDetails = getAttrDetails(rsSelectRow.Fields("odrattrtmpAttrID"))
  					sTmpAttrPrice = aAttrDetails(1)
  					iAttrType = aAttrDetails(2)
  					  				
  						If iAttrType = 1 Then
  							dAttrSubtotal = dAttrSubtotal + cDbl(sTmpAttrPrice)*cLng(iQuantity)  							
  						ElseIf iAttrType = 2 Then
  							dAttrSubtotal = dAttrSubtotal - cDbl(sTmpAttrPrice)*cLng(iQuantity)							
  						End If			  				
  						rsSelectRow.MoveNext  				
  				Loop
  				 dSubtotal = dSubTotal + dAttrSubtotal
  		  ' End RecordSet If
  		  End If
  		  closeObj(rsSelectRow)  		  
  	Else 
  		  dSubtotal = sProdPrice*cLng(iQuantity)	  
	End If

    getSubtotal = dSubtotal  	  
End Function

'---------------------------------------------------------
' Sets referer
'---------------------------------------------------------
Sub setReferer(sReferer,iOrderID)
	Dim sLocalSQL, rsOrder
	On Error Resume Next
	sLocalSQL = "UPDATE sfOrders SET orderHttpReferrer = '" & sReferer & "' WHERE orderID =" & iOrderID 	
	Set rsOrder = cnn.Execute(sLocalSQL)
	closeobj(rsOrder)
	On Error GoTo 0
End Sub	

'----------------------------------------------------------
' Updates the Order with transaction Codes
'----------------------------------------------------------
Sub setTransactionResponse(iOrderID,ProcMessage,ProcCustNumber,ProcAddlData,ProcRefCode,ProcAuthCode,ProcMerchNumber,ProcActionCode,ProcErrMsg,ProcErrLoc,ProcErrCode,ProcAvsCode)
	Dim rsOrder
	
	Set rsOrder = Server.CreateObject("ADODB.RecordSet")
	rsOrder.Open "sfTransactionResponse", cnn, adOpenDynamic,adLockOptimistic,adCmdTable
	rsOrder.AddNew
	rsOrder.Fields("trnsrspOrderId")		= trim(iOrderID)
	rsOrder.Fields("trnsrspCustTransNo")	= trim(ProcCustNumber)
	rsOrder.Fields("trnsrspMerchTransNo")	= trim(ProcMerchNumber)
	rsOrder.Fields("trnsrspAVSCode")		= trim(ProcAvsCode)
	rsOrder.Fields("trnsrspAUXMsg")			= trim(ProcMessage)
	rsOrder.Fields("trnsrspActionCode")		= trim(ProcActionCode)
	rsOrder.Fields("trnsrspRetrievalCode")	= trim(ProcRefCode)
	rsOrder.Fields("trnsrspAuthNo")			= trim(ProcAuthCode)
	rsOrder.Fields("trnsrspErrorMsg")		= trim(ProcErrMsg)
	rsOrder.Fields("trnsrspErrorLocation")	= trim(ProcErrLoc)
	If ProcResponse <> "failed" Then
		rsOrder.Fields("trnsrspSuccess") = 1
	Else 
		rsOrder.Fields("trnsrspSuccess") = 0
	End If
	rsOrder.Update				
	closeobj(rsOrder)
End Sub
'----------------------------------------------------------
' Sets the order complete flag to 1
'----------------------------------------------------------
Sub setOrderComplete(iOrderID)
	Dim sLocalSQL, rsOrder
	
	sLocalSQL = "UPDATE sfOrders SET orderIsComplete = 1 WHERE orderID =" & iOrderID
	
	Set rsOrder = cnn.Execute(sLocalSQL)
	
	closeobj(rsOrder)
End Sub
'-----------------------------------------------------------
'Gets Shipping from TmpOrders Tbale
'-----------------------------------------------------------
Function getShipping()
	Dim rsTmpTable, sFilter, iShipping
	Set rsTmpTable = Server.CreateObject("ADODB.RecordSet")
	rsTmpTable.Open "sfTmpOrderDetails", cnn, adOpenForwardOnly, adLockReadOnly, adCmdTable
	sFilter = "odrdttmpSessionID = " & Session("SessionID")
	rsTmpTable.Filter = sFilter
	If rsTmpTable.EOF And rsTmpTable.BOF Then Response.Redirect "abandon.asp"
	iShipping = trim(rsTmpTable.Fields("odrdttmpShipping"))
	rsTmpTable.Close 
	Set rsTmpTable = nothing
	getShipping = iShipping
End Function

%>








