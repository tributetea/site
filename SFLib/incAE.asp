
<!--#include file="mail.asp"-->

<%
Const vDebugAE = 0
'----------------------------------------------------------------------------------------------------
'Warning: Do NOT remove or modify the following code block!
  Application("AppName")="StoreFrontAE"
  Application("CartName")="Wish List"
  Application("CartSaveButton")="ADD TO WISH LIST"
'----------------------------------------------------------------------------------------------------

'@BEGINVERSIONINFO

'@APPVERSION: 50.4011.0.2

'@FILENAME: 
	


'@DESCRIPTION: 

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws as an unpublished work, and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO
'Modified 10/23/01 
'Storefront Ref#'s: 157 'JF
'Modified 10/31/01 
'Storefront Ref#'s: 193 'JF

' Public Variables needed for AE functions in this module
Dim iShipQty,iboQty, igwQty
Dim gProductId
Dim gShowInStock
Dim gGiftWrap
Dim gGiftWrapTotal
Dim gGiftWrapGrandTotal
Dim gCouponCode
Dim gCouponDiscount
dim gStoreWideDiscount
Dim gNoThanks
Dim gTmpSQL,gOrderSQL
Dim gRecalcDone
Dim gShippedQTY
Dim sBackOrderShipping, sBillAmount, sBackOrderAmount
dim sTotalPrice_Bill
Dim sTotalPrice_Back
Dim gCouponDiscountOut
Dim gStoreWideDiscountOut
Dim aProdInfoAE() 'SFAE
' Public Initialization
gRecalcDone = 0
gtmpSQL = "Select * FROM sfTmpOrderDetails as A LEFT JOIN sfTmpOrderDetailsAE as B ON A.odrdttmpID = B.odrdttmpaeID "
gorderSQL = "Select * FROM sfOrderDetails as A LEFT JOIN sfOrderDetailsAE as B ON A.odrdtID = B.odrdtaeID "
Application("AppName") = "StoreFrontAE"
Session("AppName") = "StoreFrontAE"


Sub Confirm_DoInventory
If Application("AppName") = "StoreFrontAE" Then 'SFAE
	Confirm_CheckCartAndRedirect 
	ReleaseAppLock 'release lock over two minutes
			   'or any previous lock from this session
	LockApp 
	Confirm_UpdateInventory 
	UnLockApp
End If
End Sub


Sub LockApp
'For inventory tracking purpose
	IF Application("InventoryLock") = "Locked"   Then
		JS "window.history.go(-1)"	
		Response.End 
	Else
		Application.Lock 
		Application("InventoryLock") = "Locked"
		Application("LockTime") = Now()  'save date time of lock
		Application("LockSessionID") = Session("SessionId")
		Application.UnLock
	End IF

End Sub

Sub ShowLockValues
	'DEBug mode only
	Response.Write "<BR>" & vbcrlf
	Response.Write "<BR> InventoryLock:" & Application("InventoryLock") & vbcrlf
	Response.Write "<BR> LockTime:" & Application("LockTime") & vbcrlf
	Response.Write "<BR> LockSessionID:" & Application("LockSessionID") & vbcrlf
	Response.Write "<BR> My SessionID:" & Session("SessionID") & vbcrlf
	'Response.write "DateDiff:" & DateDiff("s",Application("LockTime"), Now)
End Sub
Sub ReleaseAppLock
'This routine prevents from infinite or bad locks 
	IF Application("InventoryLock") = "Locked" Then
		If  DateDiff("s",Application("LockTime"), Now) > 120 Or Application("LockSessionID") = Session("SessionID")  Then
			Application.Lock 
			Application("InventoryLock") = "Unlocked"
			Application("LockTime") = ""  
			Application("LockSessionID") = ""
			Application.UnLock
		End If
	End If
End Sub

Sub UnlockApp
	Application.Lock 
	Application("InventoryLock") = "Unlocked"
	Application("LockTime") = ""  
	Application("LockSessionID") = ""
	Application.UnLock
End Sub

'************************************************* COMMON ROUTINES *******************************
Sub SetBillingVariables
dim boqty,shpqty,ordqty
	boqty = gettotalbackorderqty()
	ordqty = gettotalorderqty()
	shpqty = clng(ordqty - boqty)
	
	Session("boqty") = clng(boqty)
	Session("shpqty") = clng(shpqty)
	Session("ordqty") = clng(ordqty)
	
	Session("BackOrderShipping") = 0
	Session("BillShipping") = 0
	Session("sTotalPrice") = 0
	
	If BackOrderBilling = 0 And Session("boqty") > 0 then 
		Session("SpecialBilling") = 1
	else
		Session("SpecialBilling") = 0
	End If

End Sub

Function BackOrderBilling
dim rst,sql

	Set rst = Server.CreateObject("ADODB.RecordSet")		
	sql = "Select * FROM sfadmin"
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText
	If rst.recordcount > 0 then
		rst.movefirst
		BackOrderBilling = rst("adminBackOrderBilling")	
	end if
	CloseObj (rst)
	If not isnumeric(backorderbilling) or isnull(backOrderBilling) or trim(backorderbilling) = "" then
		BackOrderBilling = 1 
	End If
	'Response.Write "<BR> BACKORDERBILLING:" & BackOrderBilling
End Function

Function getGlobalSalePriceAE(subtotal)
	Dim rsGlobalAdmin, sGlobalActive
	Set rsGlobalAdmin = Server.CreateObject("ADODB.RecordSet")
	rsGlobalAdmin.open "sfAdmin", cnn, adOpenForwardOnly, adLockReadOnly, adCmdTable
	
	sGlobalActive = Trim(rsGlobalAdmin.Fields("adminGlobalSaleIsActive"))
	If sGlobalActive = "1" Then
		getGlobalSalePriceAE = formatNumber(cDbl(subtotal)-(cDbl(subtotal)*cDbl(rsGlobalAdmin.Fields("adminGlobalSaleAmt"))), 2)
	Else
		getGlobalSalePriceAE = subTotal
	End If
	rsGlobalAdmin.Close
	Set rsGlobalAdmin = nothing
End Function

Sub Confirm_DivideAmountDiscount
	
On Error Resume Next
Dim AmountDiscount
Dim gTot, prodShip,rd
		
	If session("SpecialBilling") <> 1 then exit sub
	
	
	IF IsNumeric(Session("discAmount")) then 
		AmountDiscount = Session("discAmount")
	else
		AmountDiscount = 0 
	end if 
	
   IF AmountDiscount > 0 then '.3008
		'remove previously applied amount coupon discounts
		Session("BackOrderAmount") = cdbl(Session("BackOrderAmount") + AmountDiscount) 
		Session("BillAmount") = cdbl(Session("BillAmount") + AmountDiscount)
   End IF
	

	gTot = cdbl(Session("BackOrderAmount") + Session("BillAmount"))
	

	'Handle any dollar amount coupon
	If AmountDiscount => 0  then
		If gtot => AmountDiscount then
			If  Session("BackOrderAmount") => AmountDiscount Or  Session("BillAmount") => AmountDiscount Then
				If  Session("BackOrderAmount") => AmountDiscount then
					Session("BackOrderAmount") = cdbl(Session("BackOrderAmount") - AmountDiscount)
				Elseif Session("BillAmount") => AmountDiscount then
					Session("BillAmount") = cdbl(Session("BillAmount") - AmountDiscount)
				End If
			else
				'discount is bigger than all amounts
				rd = 0
				If Session("BackOrderAmount")=> Session("BillAmount") then
					rd = cdbl(AmountDiscount - Session("BackOrderAmount"))	
				end if 
				
				If Session("BillAmount") => Session("BackOrderAmount") then
					rd = cdbl(AmountDiscount - Session("BillAmount"))	
				end if 
				
				Session("BackOrderAmount") = 0
				Session("BillAmount") = cdbl(Session("BillAmount") - rd)
			End If
		Else
			Session("BackOrderAmount") = 0
			Session("BillAmount") = 0
		End If
		'response.Write "GS:" & sGrandtotal
		'Response.Write "got:" & gtot
	End If
'---------------------------------------------------
End Sub

Sub WriteVars
if vDebugAE = 1 then	
Dim i 
		For I = 0 to iProdNumber  - 1 
			Response.Write "<BR>------ Array Start----" & i
			Response.Write "<BR> MTP:" &  aProdInfoAE(i,1)
			Response.Write "<BR> BOQTY:" &  aProdInfoAE(i,2)
			Response.Write "<BR> GWQTY:" &  aProdInfoAE(i,3)
			Response.Write "<BR> GWPrice:" &  aProdInfoAE(i,4)
			Response.Write "<BR>----- Array End------"   
		Next
		Response.Write "<BR> ordqty: " & Session("ordQTY")
		Response.Write "<BR> boqty: " & Session("boQTY")
		Response.Write "<BR> shpqty: " & Session("shpQTY")
		
		Response.Write "<BR> SpecialBilling:" & Session("SpecialBilling") 
	
		Response.Write "<BR> sShipping:" & Session("sShipping") 
		Response.Write "<BR> backorder_shipping:" & Session("BackOrderShipping") 
		Response.Write  "<BR>shipped_shipping:" & Session("BillShipping") 
		
		Response.Write "<BR>"
		Response.Write "<BR>BillAmount:"  & Session("BillAmount")
		Response.Write "<BR>BackOrderAmount:"  & Session("BackOrderAmount")
		Response.Write "<BR>"
		Response.Write "<BR>sGrandTotal:"  & sGrandTotal
		Response.Write "<BR>sGrandTotalOut:"  & sGrandTotalOut
		Response.Write "<BR>"
		Response.Write "<BR>sTotalPrice:"  & stotalPrice
		Response.Write "<BR>sTotalPrice_bill:"  & Session("stotalPrice_bill")
		Response.Write "<BR>sTotalPrice_back:"  & Session("stotalPrice_back")
		Response.Write "<BR>"
		Response.Write "<BR>StoreWideDiscount:"  & Session("StoreWideDiscount")
		Response.Write "<BR>CouponDiscountPercent:"  &  Session("CouponDiscountPercent")
		Response.Write "<BR>CouponDiscountAmount:"  &  Session("CouponDiscountAmount")
		Response.End 
	end if
End Sub

Sub Confirm_SaveAmounts
dim rst,sql
	'Save amount in ordersAE
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	sql = "Select * FROM sfOrdersAE where orderAEID=" & iOrderID
	rst.Open sql, cnn, adOpenKeySet, adLockOptimistic, adCmdText
	if rst.recordcount > 0 then
		
	else
		rst.addnew
		rst("orderAEID") = iOrderID
	end If
	
	rst("orderBackOrderAmount") = cdbl(Session("BackOrderAmount"))
	rst("orderBillAmount")= cdbl(Session("BillAmount"))
 
	If cdbl(Session("CouponDiscountPercent")) > 0 then
		rst("orderCouponDiscount") = cdbl(Session("CouponDiscountPercent"))
	elseif cdbl(Session("CouponDiscountAmount")) > 0 then
		rst("orderCouponDiscount") = cdbl(Session("CouponDiscountAmount"))
	else
		rst("orderCouponDiscount") = 0
	End If
	
	rst.update
	
	CloseObj (rst)
	
End Sub

Sub Confirm_SaveGrandTotal
dim rst,sql

	'Save amount in ordersAE
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	sql = "Select * FROM sfOrders where orderID=" & iOrderID
	rst.Open sql, cnn, adOpenKeySet, adLockOptimistic, adCmdText
	if rst.recordcount > 0 then
		rst("orderGrandTotal") = cdbl(sGrandTotalOut)
		rst.update
	end if
	
	CloseObj (rst)
	
End Sub

Sub Confirm_SetProdPrice1 'get from cart
	dUnitPrice = GetMTPrice(sProdId,dUnitPrice,0)
	
	'confirm hidden
	aProdInfoAE(iProductCounter,1)  = cdbl(sProdPrice)
			
End Sub


Sub Confirm_SetProdPrice2 'get from orders
	dUnitPrice = GetMTPrice(sProdId,dUnitPrice,iOrderId)
End Sub


Sub Confirm_GetBillAmount
dim boqty,ordqty,dProdDiscountTotal,shpqty,gwprice

	If Session("SpecialBilling") <> 1 then exit sub
	
	
	Session("BillAmount") = 0
	Session("BackOrderAmount") = 0
	Session("sTotalPrice_Bill") = 0
	
	If Session("shpqty") <= 0 then exit sub
	
	sBillAmount = 0			
	sBackOrderAmount = 0
	
	boqty = Session("boqty")
	shpqty = Session("shpqty") 
	ordqty = Session("ordqty") 
		
	
	If ordqty <= 0 then  'safe check
		sBillAmount = 0
		exit sub	
	end if
		
	'
	
	sql = gtmpsql & "WHERE odrdttmpSessionID = " & Session("SessionID")
	
	Set rsAllOrders = Server.CreateObject("ADODB.RecordSet")
	rsAllOrders.Open SQL, cnn, adOpenKeyset, adLockReadOnly, adCmdText


	If rsAllOrders.EOF Then
		closeObj(rsAllOrders)
		closeObj(cnn)
		' redirect to neworder screen
		Session.Abandon
		Response.Redirect(C_HomePath & "search.asp")		
	Else
    	IF not isnull(rsAllOrders("odrdttmpHttpReferer")) then
		  aReferer = Split(rsAllOrders("odrdttmpHttpReferer"), ",")
		else
		 aReferer =""
		end if 
		aPurchases = rsAllOrders.GetRows
		iProdNumber = rsAllOrders.RecordCount
		iAttrNumber = getAttributeNumber()
		rsAllOrders.MoveFirst
	
		Redim aOrderID(iProdNumber)
		Redim aAllProd(iAttrNumber,iProdNumber)

		' Initialize total to 0
		sTotalPrice = "0"
		sShipping = "0"
		sProductSubtotal	= "0"	
		dProductSubtotal	= 0
		iProductCounter	= 0
		ictax = 0 'xx
		istax = 0  'xx
		
		Do While NOT rsAllOrders.EOF
   
			' Get the ProdIDs
			iTmpOrderID = rsAllOrders.Fields("odrdttmpID")
			sProdID = rsAllOrders.Fields("odrdttmpProductID")
			'iQuantity = rsAllOrders.Fields("odrdttmpQuantity")
			iQuantity = clng(rsAllOrders.Fields("odrdttmpQuantity")) - clng(rsAllOrders.Fields("odrdttmpBackOrderQty"))
			
		    
			' put orderid into an array for deletion
			aOrderID(iProductCounter) = iTmpOrderID
		    
			' Get an array of 3 values from getProduct()

			ReDim aProduct(3)
			aProduct = getProduct(sProdID)		
			sProdName = aProduct(0)
			sProdPrice = aProduct(1)
			dUnitPrice = sProdPrice
			Verify_SetProdPrice 'SFAE
			sProdPrice = dUnitPrice
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
				
				' Store Attribute Affected Price
				aAllProd(0,iProductCounter)(5) = cDbl(sAttrUnitPrice) + cDbl(sProdPrice)	

				' Set Unit Price for Product
				sUnitPrice = FormatCurrency(cDbl(sAttrUnitPrice) + cDbl(sProdPrice))
				
				dProductSubtotal = iQuantity * (cDbl(sAttrUnitPrice) + cDbl(sProdPrice))
				
				
				sProductSubtotal = FormatCurrency(dProductSubtotal)
				
				'sTotalPrice = sTotalPrice + cDbl(dProductSubtotal)				
			' End IsArray If
			End If
		
			'OVC_ShowGiftWrapValue 2 'SFAE
			
			If rsAllOrders("odrdttmpGiftWrapQTY") > 0 then
				gwprice = GetGiftWrapPrice(rsAllOrders("odrdttmpProductID"))
				If gwprice <> "X" Then 					
					igwqty = rsAllOrders("odrdttmpGiftWrapQTY")
					ishipqty= rsAllOrders.Fields("odrdttmpQuantity") - rsAllOrders.Fields("odrdttmpBackOrderQty")
					iboqty= rsAllOrders.Fields("odrdttmpBackOrderQty")

					if igwQty > ishipqty then igwqty = ishipqty
				
					if iShipqty  = 0 then igwqty = 0
									
					dProductSubtotal = cdbl(dProductSubtotal) + cdbl(igwqty * gwprice)					

				End If
			End if	
			sTotalPrice = sTotalPrice + cDbl(dProductSubtotal) 'SFUPDATE
			
			iSTax = iSTax + cDbl(getTax("State", sShipping, ApplyPercentDiscounts(dProductSubtotal,"Total"), sProdID))
			iCTax = iCTax + cDbl(getTax("Country", sShipping, ApplyPercentDiscounts(dProductSubtotal,"Total"), sProdID))
			
			iProductCounter = iProductCounter + 1
						
			rsAllOrders.MoveNext		
			
		Loop
		
		' object cleanup
		closeObj(rsAllOrders)
		
		sShipping = cdbl(Session("BillShipping"))
		'sShipping  = 0 'later
		
		If iHandling <> 1 Or iShipped <> 1 Then sHandling = 0
			
		iSTax = iSTax + cDbl(getTax("State", sShipping, "0", sProdID))
		iCTax = iCTax + cDbl(getTax("Country", sShipping, "0", sProdID))	

		sTotalPrice =  ApplyAllDiscounts(sTotalPrice,"Total") '.3008
							
		sBillAmount = (cDbl(sTotalPrice) + cDbl(sShipping) + cDbl(iSTax) + cDbl(iCTax)) 
		
		If shpqty > 0  then		
			sBillAmount = cdbl(sBillAmount) + cDbl(sHandling) + cDbl(iCodAmount) + cdbl( Session("iPremiumShipping"))
		End if
		
		Session("BillAmount") = cdbl(sBillAmount)
		Session("sTotalPrice_Bill") = sTotalPrice
		sShipping = 0
		sTotalPrice = 0
End If	

End Sub


Sub Confirm_GetBackOrderAmount
dim boqty,ordqty,dProdDiscountTotal,shpqty,gwprice
	
	If Session("SpecialBilling") <> 1 then exit sub
	
	Session("BackOrderAmount") = 0
	Session("sTotalPrice_Back") = 0
	
	
	boqty = Session("boqty")
	ordqty =  Session("ordqty")
	shpqty = SEssion("shpqty")
	
	If boqty <= 0 then 
		Session("BackOrderFlag") = "No"
		sBackOrderAmount = 0
		exit sub
	else
		Session("BackOrderFlag") = "Yes"
	end if
	
	
	sBackOrderAmount = 0
	'SQL = "SELECT * FROM sfTmpOrderDetails WHERE odrdttmpSessionID = " & Session("SessionID")
	sql = gtmpsql & "WHERE odrdttmpSessionID = " & Session("SessionID")
	Set rsAllOrders = Server.CreateObject("ADODB.RecordSet")
	rsAllOrders.Open SQL, cnn, adOpenKeyset, adLockReadOnly, adCmdText


	If rsAllOrders.EOF Then
		closeObj(rsAllOrders)
		closeObj(cnn)
		' redirect to neworder screen
		Session.Abandon
		Response.Redirect(C_HomePath & "search.asp")		
	Else
    	IF not isnull(rsAllOrders("odrdttmpHttpReferer")) then
		  aReferer = Split(rsAllOrders("odrdttmpHttpReferer"), ",")
		else
		 aReferer =""
		end if 
		aPurchases = rsAllOrders.GetRows
		iProdNumber = rsAllOrders.RecordCount
		iAttrNumber = getAttributeNumber()
		rsAllOrders.MoveFirst
	
		Redim aOrderID(iProdNumber)
		Redim aAllProd(iAttrNumber,iProdNumber)

		' Initialize total to 0
		sTotalPrice = 0
		sShipping = 0
		sProductSubtotal	= 0	
		dProductSubtotal	= 0
		iProductCounter	= 0
		IStax = 0
		ICtax = 0
		
		Do While NOT rsAllOrders.EOF
   
			' Get the ProdIDs
			iTmpOrderID = rsAllOrders.Fields("odrdttmpID")
			sProdID = rsAllOrders.Fields("odrdttmpProductID")
			'iQuantity = rsAllOrders.Fields("odrdttmpQuantity")
			iQuantity = rsAllOrders.Fields("odrdttmpBackOrderQty")
    
			' put orderid into an array for deletion
			aOrderID(iProductCounter) = iTmpOrderID
		    
			' Get an array of 3 values from getProduct()

			ReDim aProduct(3)
			aProduct = getProduct(sProdID)		
			sProdName = aProduct(0)
			sProdPrice = aProduct(1)
			dUnitPrice = sProdPrice
			Verify_SetProdPrice 'SFAE
			sProdPrice = dUnitPrice
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
				
				' Store Attribute Affected Price
				aAllProd(0,iProductCounter)(5) = cDbl(sAttrUnitPrice) + cDbl(sProdPrice)	

				' Set Unit Price for Product
				sUnitPrice = FormatCurrency(cDbl(sAttrUnitPrice) + cDbl(sProdPrice))
				
				dProductSubtotal = iQuantity * (cDbl(sAttrUnitPrice) + cDbl(sProdPrice))
				
				sProductSubtotal = FormatCurrency(dProductSubtotal)
				
			' End IsArray If
			End If
			
			'OVC_ShowGiftWrapValue 31 'SFAE
			If rsAllOrders("odrdttmpGiftWrapQTY") > 0 then
				gwprice = GetGiftWrapPrice(rsAllOrders("odrdttmpProductID"))
				If gwprice <> "X" Then 					
					igwqty = rsAllOrders("odrdttmpGiftWrapQTY")
					ishipqty= rsAllOrders.Fields("odrdttmpQuantity") - rsAllOrders.Fields("odrdttmpBackOrderQty")
					iboqty=iQuantity
					
					if iGwQTY <= iShipqty then 
						igwqty = 0 
					elseif igwqty > iShipqty then
						igwqty = igwqty - ishipqty
					end if 
					
					if iboqty = 0 then igwqty = 0
					
					dProductSubtotal = cdbl(dProductSubtotal) + cdbl(igwqty * gwprice)					
					
				End If
			End if
			sTotalPrice = sTotalPrice + cDbl(dProductSubtotal)  'SFUPDATE

			iSTax = iSTax + cDbl(getTax("State", sShipping, ApplyPercentDiscounts(dProductSubtotal,"Total"), sProdID))
			iCTax = iCTax + cDbl(getTax("Country", sShipping, ApplyPercentDiscounts(dProductSubtotal,"Total"), sProdID))
		
			rsAllOrders.MoveNext		

		Loop
		
		' object cleanup
		closeObj(rsAllOrders)
		
		'sShipping  = 0 'later
		sShipping = cdbl(Session("BackOrderShipping"))  
		
		If iHandling <> 1 Or iShipped <> 1 Then sHandling = 0
		
		iSTax = iSTax + cDbl(getTax("State", sShipping, "0", sProdID))
		iCTax = iCTax + cDbl(getTax("Country", sShipping, "0", sProdID))	
		
		
		sTotalPrice =  ApplyALLDiscounts(sTotalPrice,"Total") '.3008
		
		sBackOrderAmount = (cDbl(sTotalPrice) +  cDbl(sShipping) + cDbl(iSTax) + cDbl(iCTax) )
		'Response.Write "<BR>sBackOrderAmount " & sBAckOrderAmount
		If shpqty <= 0 and boqty > 0  then 
			sBackOrderAmount = cdbl(sBackOrderAmount) + cDbl(sHandling) + cDbl(iCodAmount) + cdbl( Session("iPremiumShipping"))
			'Response.Write "<BR>sBackOrderAmount " & sBackOrderAMount
		End If
		Session("BackOrderAmount") = cdbl(sBackOrderAmount)
		Session("sTotalPrice_Back") = sTotalPrice	
		sShipping = 0
		sTotalPrice = 0
End IF	

End Sub

Sub Confirm_ShowBillingInfo
If Session("SpecialBilling") <> 1 then exit sub

sBillAmount = cdbl(Session("BillAmount"))
sBackOrderAmount = cdbl(Session("BackOrderAmount"))

If Session("BackOrderFlag") = "No" then Exit sub
If BackOrderBilling = 1 then exit sub					
If cdbl(sBillAmount) = cdbl(sGrandTotalOut) then exit sub

dim sbillAmountOut
dim sBackOrderAmountOut
		If iConverion = 1 Then
			sBackOrderAmountOut = "<script> document.write(""" & FormatCurrency(sBackOrderAmount) & " = ("" + OANDAconvert(" & sBackOrderAmount & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script>" 'sfae beta2
			sBillAmountOut = "<script> document.write(""" & FormatCurrency(sBillAmount) & " = ("" + OANDAconvert(" & sBillAmount & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script>" 'sfae beta2
		Else
			sBackOrderAmountOut = FormatCurrency(sBackOrderAmount) 'sfaebeta2
			sBillAmountOut = FormatCurrency(sBillAmount) 'sfaebeta2
		End If
		
	Response.Write "<tr>" & vbcrlf
	Response.Write "<td width=""75%"" align=""right""><b>Billed Amount:</b></td>" & vbcrlf
   	Response.Write "<td nowrap width=""25%"" height=""20""><b>" & sBillAmountOut & vbcrlf
	Response.Write "</b></td> </tr>" & vbcrlf
		            
   	Response.Write "<tr>" & vbcrlf
	Response.Write "<td width=""75%"" align=""right""><b>Remaining Amount:</b></td>" & vbcrlf
   	Response.Write "<td nowrap width=""25%"" height=""20""><b>" & sBackOrderAmountOut & vbcrlf
	Response.Write "</b></td>  </tr>" & vbcrlf
	Response.Write "<tr>" & vbcrlf
	Response.Write "<td nowrap width=""16%"" align=""left"" valign=""top"">Note: Remaining amount will be billed upon shipment of backordered items.</td>"              & vbcrlf
	Response.Write "</tr>" & vbcrlf

	

End Sub

Function GetTotalBackOrderQTY
dim rst
dim sql
	'IF Session("SessionEnd")="Yes" then exit function
	sql = "Select SUM(odrdttmpBackOrderQty) as boqty  FROM sfTmpOrderDetails as A LEFT JOIN sfTmpOrderDetailsAE as B ON A.odrdttmpID = B.odrdttmpaeID "
	sql = SQL & " WHERE odrdttmpSessionID =" & Session("SessionID")
	
	Set rst = Server.CreateObject("ADODB.RecordSet")
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly,adcmdText
	If rst.recordcount > 0 then
	  If Not isNull(rst("boqty")) then
		GetTotalBackOrderQTY = clng(rst("boqty"))
	  Else
		GetTotalBackOrderQTY = 0
      End IF
 
	else
		GetTotalBackOrderQTY = 0
	End IF
	
	if NOT ISNUMERIC(GetTotalBackOrderQTY) THEN
		GetTotalBackOrderQTY = 0
	END IF
	closeobj (rst)

End Function

Function GetTotalOrderQTY
dim rst
dim sql
	sql = "Select SUM(odrdttmpQuantity) as ordqty  FROM sfTmpOrderDetails "
	sql = SQL & " WHERE odrdttmpSessionId =" & Session("SessionID") ' & " AND odrdttmpBackOrderQty <= 0 "
	
	Set rst = Server.CreateObject("ADODB.RecordSet")
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly,adcmdText
	If rst.recordcount > 0 then
		GetTotalOrderQTY = clng(rst("ordqty"))' - clng(GetTotalBackOrderQTY)
	else
		GetTotalOrderQTY = 0
	End IF
	
	closeobj (rst)
	

End Function



Function GetBackOrderQTY (iTmpOrderDetailID)
dim rst
dim sql
dim boqty

	boqty = 0	
	sql= gtmpSQL & "WHERE odrdttmpID=" & iTmpOrderDetailID
	
	Set rst = Server.CreateObject("ADODB.RecordSet")
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly,adcmdText
	If rst.recordcount > 0 then
		boqty = clng(rst("odrdttmpBackOrderQty"))
		If boqty < 0 then boqty = 0 
	End IF
	GetBackOrderQTY = boqty
	closeobj (rst)
	

End Function

Function GetAvlOrderQTY (iTmpOrderDetailID)
dim rst
dim sql
dim avlqty

	
	sql= gtmpSQL & "WHERE odrdttmpID=" & iTmpOrderDetailID
	avlqty = 0
	Set rst = Server.CreateObject("ADODB.RecordSet")
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly,adcmdText
	If rst.recordcount > 0 then
		avlqtY = clng(rst("odrdttmpQuantity")) - clng(rst("odrdttmpBackOrderQty"))
		If avlqty < 0 then avlqty = 0 
	End IF
	GetAvlOrderQTY = avlqty
	closeobj (rst)
	

End Function



Sub OVC_ShowBackOrderMessage (iPageID)
On Error Resume Next
dim rst
dim sql
dim boqty
dim sAttID	
	
	If CheckInventoryTracked(rsAllOrders("odrdttmpProductID")) <> 1 then 
		If PageId = 33 Then 'confirm hidden
			aProdInfoAE(iProductCounter,2)  = 0 'boqty
		End If
		Exit Sub
	End If
	
	If PageId = 31 Then 'confirm hidden
		aProdInfoAE(iProductCounter,2)  = 0
	End If
	
	
	sql= gtmpSQL & "WHERE odrdttmpID=" & rsAllOrders("odrdttmpid")
	
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly,adcmdText
	
	If rst.recordcount > 0 then
		'sAttID = GetAttDetailID(rst("odrdttmpID"),"tmp")
		'boqty = clng(rst("odrdttmpQuantity")) - clng(GetAvailableqty(rst("odrdttmpProductID"),sAttID))
		'If boqty < 0 then boqty = 0 
	
		IF iPageID <> 31 Then
			If rst("odrdttmpBackOrderQty") > 0 then 'beta 2

				Response.Write "<tr>" & vbcrlf
				Response.Write "<td colspan=""5"" width=""100%"" align=""left"" class='" & fontClass & "' valign=""top"" nowrap>" & vbcrlf
				Response.Write "<B>Back Ordered Qty: " & rst("odrdttmpBackOrderQty") & "</B> (of the above items) " & backorderbillingmessage & "" & vbcrlf
				Response.Write "</td> " & vbcrlf
				Response.Write "</tr>" & vbcrlf
		
			End If
		End If
		If PageId = 31 Then 'confirm hidden
			aProdInfoAE(iProductCounter,2)  = rst("odrdttmpBackOrderQty") 'boqty
		End If
	
	End IF
	
	CloseObj (rst)

End Sub

Sub BackOrderBillingMessage
	If BackOrderBilling = 0 then
		Response.write "<BR> Backordered items will be included in the order total but will not be billed until shipped."
	End If
End Sub


Sub OVC_ShowGiftWrapValue(PageID)  'NEW   OVC = order Verify Confirm

dim rst
dim sql
dim gGiftWrapTotal
Dim gwprice
dim gwqty
dim gwpricetotal
dim gwpriceOut
dim gwpriceTotalOut
	If PageId = 31 Then 'confirm hidden
		aProdInfoAE(iProductCounter,3)  = 0
		aProdInfoAE(iProductCounter,4)  = 0  
	End If	
	'gwpriceTotal = 0
	gwpriceTotal = 0
	'Session("gwprice") = 0
	'order page
	if pageid = 1 then sql= gtmpSQL & "WHERE odrdttmpID=" & rsAllOrders("odrdttmpid")
	
	'verify page 
	if pageid = 2 or  pageId = 31 then sql= gtmpSQL & "WHERE odrdttmpID=" & rsAllOrders("odrdttmpid")
		
	'confirm page
	if pageid = 3  then sql = gorderSQL &  " WHERE odrdtOrderId=" & iOrderID & " AND odrdtProductID='"&  sProdId & "'"	
	
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly,adcmdText
	
	'confirm page
	If pageid = 3  then
		If rst.recordcount > 0 then
			If  rst("odrdtGiftWrapQty") > 0 then
				gwpriceTotal = FormatCurrency(cdbl(rst("odrdtGiftWrapQty") * GetGiftWrapPrice(rst("odrdtProductID"))))
				gGiftWrapGrandTotal = cdbl(gGiftWrapGrandTotal + gwpriceTotal)
			end if
		else
			closeobj(rst)
			exit sub
		end if
		
		gwprice = GetGiftWrapPrice (sProdID)
		gwqty = rst("odrdtGiftWrapQty")
	else
		If rst.recordcount > 0 then
			If  rst("odrdttmpGiftWrapQty") > 0 then
				gwpriceTotal = FormatCurrency(cdbl(rst("odrdttmpGiftWrapQty") * GetGiftWrapPrice(rst("odrdttmpProductID"))))
				gGiftWrapGrandTotal = cdbl(gGiftWrapGrandTotal + gwpriceTotal)
			end if
			
			If PageId = 31 Then 'confirm hidden
				aProdInfoAE(iProductCounter,3)  = rst("odrdttmpGiftWrapQty") 'gwqty
			End If
		
		else
			closeobj(rst)
			exit sub
		end if
		gwprice = GetGiftWrapPrice (rsAllOrders("odrdttmpProductID"))
		gwqty = rst("odrdttmpGiftWrapQty")
	end if
	
	if pageid = 2 or pageid = 3 or pageid =31 then
	if gwqty <= 0 then 
		closeobj(rst)
		exit sub
	end if
	end if
	
		
	
	If gwprice <>  "X" then 
		If PageId = 31 Then 'confirm hidden
			aProdInfoAE(iProductCounter,4)  = cdbl(gwprice)
		End If	
		gwpriceout = cdbl(gwprice)
		
		dProductSubtotal = cdbl(dProductSubtotal) +  cdbl(gwpricetotal)
		
		If iConverion = 1 Then			
			gwpriceout = "<script>document.write(""" & FormatCurrency(gwpriceout) & " = ("" + OANDAconvert(" & gwpriceout & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script>"  
		else
			gwpriceout = FormatCurrency(gwprice)
		End If
		
		IF pageID <> 31 then
			Response.Write "<tr>" & vbcrlf
			Response.Write "<td width=""40%"" align=""left"" class='" & fontClass & "' valign=""top"" nowrap background=""""><b>Gift Wrap</B></td>" & vbcrlf
			Response.Write "<td width=""15%"" align=""center"" class='" & fontClass & "' valign=""top"" nowrap background="""">" & gwpriceout & "</td> " & vbcrlf
        End IF
        If PageId = 1 then 
			Response.Write "<td width=""15%"" align=""center"" class='" & fontClass & "' valign=""top"" nowrap background=""""><input type=""text"" style=""" & C_FORMDESIGN & """ name=""GWQTY" & iProductCounter & """ size=""2"" value=""" & gwqty & """></td>" & vbcrlf
		ELSEIF PageId = 2 or PageId = 3 Then
			Response.Write "<td width=""15%"" align=""center"" class='" & fontClass & "' valign=""top"" nowrap background="""">" & gwqty & "</td>" & vbcrlf
		End IF
		
		If  gwqty > 0 then
			gwpriceTotalOut = cdbl(gwpriceTotal)
		  If PageID <> 31 Then
			If iConverion = 1 Then			
				gwpriceTotalOut = "<script>document.write(""" & FormatCurrency(gwpriceTotalOut) & " = ("" + OANDAconvert(" & gwpriceTotalOut & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script>"  
			else
				gwpriceTotalOut = FormatCurrency(gwpriceTotal)
			End If				
				Response.Write "<td width=""15%"" align=""center"" class='" & fontClass & "' valign=""top"" nowrap background="""">" & gwpriceTotalOut & "</td> " & vbcrlf
		  ENd If
		Else
				Response.Write "<td width=""15%"" align=""center"" class='" & fontClass & "' valign=""top"" nowrap background="""">&nbsp;</td> " & vbcrlf
		End If

		If PageId = 1 then 
				Response.Write "<td width=""15%"" align=""center"" class='" & fontClass & "'  valign=""top"" nowrap background="""">&nbsp;</td>" & vbcrlf
		End IF
				Response.Write "</tr>" & vbcrlf
	
	else
		gwprice = 0
	end if 'gwprice <> "X" 
	
	CloseObj (rst)


End Sub



'******************************************* CONFIRM PAGE ***************************


Sub Verify_CalcOrderShipping
    'Session("BackOrder_Shipping") = GetShipping(sTotalShipPrice, iPremiumShipping, "", sShipCustCity,sShipCustState,sShipCustZip, sShipCustCountry, sTotalPrice,"BackOrder")
    'Session("Shipped_Shipping") =  GetShipping(sTotalShipPrice, iPremiumShipping, "", sShipCustCity,sShipCustState,sShipCustZip, sShipCustCountry, sTotalPrice,"Shipped")
End Sub


SUB Confirm_CheckCartAndRedirect
	If  CheckCartInventory = 0 then 'stock depleted !  
		Session("ShowInventoryMessage") = "1"
		response.Redirect(C_HomePath & "order.asp")
	End If

End SUb

Sub Confirm_UpdateInventory

dim sql
dim rst
dim i

	If vDebugAE = 1 then exit sub 'test
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	sql = gtmpSQL &  " WHERE odrdttmpSessionId=" & Session("SessionID")
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText
	
	If rst.recordcount > 0 then
		rst.movefirst
		For i = 1 to rst.recordcount
		
			UpdateAvailableQTY rst("odrdttmpProductId"),GetAttDetailID(rst("odrdttmpID"),"tmp"),clng(rst("odrdttmpQuantity"))
			If not rst.eof then rst.movenext
		Next
	End If
	CloseObj (rst)

End Sub


Sub Confirm_WriteAERecords

Dim rstOrderDetails,rstOrderDetailsAE,rstTmpOrderDetailsAE,rstOrdersAE,rstTmpOrdersAE
Dim sql
dim i 
Dim AttIDs 'b2
Dim rstDelete
dim sTmpids
	'Read from  sfTmpOrderDetailsAE
	sql = gTmpSQL & " WHERE odrdttmpSessionID=" & Session("SessionID") 
	Set rstTmpOrderDetailsAE = Server.CreateObject("ADODB.RecordSet")		
	rstTmpOrderDetailsAE.Open sql, cnn, adOpenKeySet, adLockOptimistic, adCmdText
	If not rstTmpOrderDetailsAE.recordcount > 0 then 
		closeobj(rst)
		exit sub
	End If
		   	stmpids = ""
		   	
	For i =  1 to rstTmpOrderDetailsAE.recordcount
		sTmpIds = sTmpIds & rstTmpOrderDetailsAE("odrdttmpAEID") & ","
		'Read sfOrderDetails
		sql = "Select * FROM sfOrderDetails WHERE odrdtOrderId=" & iOrderID & " AND odrdtProductID='"& rstTmpOrderDetailsAE("odrdttmpProductID") & "'"
		Set rstOrderDetails = Server.CreateObject("ADODB.RecordSet")		
		rstOrderDetails.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText
		
		If not rstOrderDetails.recordcount > 0 then 
			closeobj(rstTmpOrderDetailsAE)
			closeobj(rstOrderDetails)
		
		ELSE
		AttIds = GetAttDetailId(rstTmpOrderDetailsAE("odrdttmpAEID"),"tmp") 'b2
		Dim sOrderDTID ,k 
		
		rstOrderDetails.MoveFirst 
		For k = 1 to rstOrderDetails.RecordCount 
			sql = "Select * FROM sfOrderDetailsAE WHERE odrdtAEId=" & rstOrderDetails("odrdtId")
			Set rstOrderDetailsAE = Server.CreateObject("ADODB.RecordSet")		
			rstOrderDetailsAE.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText		
			If NOT rstOrderDetailsAE.recordcount > 0 then 	
				rstOrderDetailsAE.close
				sOrderDTID = rstOrderDetails("odrdtId")
				EXIT FOR
			end if 
			If not rstOrderDetails.eof then rstOrderDetails.MoveNext 
			CloseObj (rstOrderDetailsae)	
		Next
		CloseObj (rstOrderDetailsAE)	 
		CloseObj (rstOrderDetails)	
		
		'Write to sfOrderDetailsAE
		sql = "Select * FROM sfOrderDetailsAE WHERE odrdtAEId=" & sOrderDTID 'rstOrderDetails("odrdtId")
		rstOrderDetailsAE.Open sql, cnn, adOpenKeyset, adLockOptimistic, adCmdText					
		If not rstOrderDetailsAE.recordcount > 0 then
			rstOrderDetailsAE.AddNew 
			rstOrderDetailsAE.Fields("odrdtaeID") = sOrderDTID'rstOrderDetails("odrdtId")
		End If
		If rstTmpOrderDetailsAE("odrdttmpGiftWrapQty") > 0 then
			rstOrderDetailsAE.Fields("odrdtGiftWrapQty") = rstTmpOrderDetailsAE("odrdttmpGiftWrapQty")
			rstOrderDetailsAE.Fields("odrdtGiftWrapPrice") = cdbl(rstTmpOrderDetailsAE("odrdttmpGiftWrapQty") * GetGiftWrapPrice(rstTmpOrderDetailsAE("odrdttmpProductID")))
		Else
			rstOrderDetailsAE.Fields("odrdtGiftWrapQty") = 0
			rstOrderDetailsAE.Fields("odrdtGiftWrapPrice") = 0
		End if
		rstOrderDetailsAE.Fields("odrdtAttDetailID")= AttIds 'b2
		rstOrderDetailsAE.Fields("odrdtBackOrderQty") = rstTmpOrderDetailsAE("odrdttmpBackOrderQty")
		rstOrderDetailsAE.update
	
		CloseObj (rstOrderDetailsAE)	
		'CloseObj (rstOrderDetails)	
		End If
		
		'rstTmpOrderDetailsAE.Delete  adAffectCurrent 'delete tmp record
		'rstTmpOrderDetailsAE.Update 
		
		If not rstTmpOrderDetailsAE.eof then rstTmpOrderDetailsAE.movenext
	Next
		
	CloseObj (rstTmpOrderDetailsAE)
		
	on error resume next
'	'Delete records from tmporderdetails ae for this session
'	sql = " SELECT * FROM sfTmpOrderDetailsAE WHERE odrdttmpAEID in (" & stmpids & ")"
'	Set rstDelete = Server.CreateObject("ADODB.RecordSet")		
'	rstDelete.Open sql, cnn, adOpenKeySet, adLockOptimistic, adCmdText
'	'Response.Write rstDelete.RecordCount 
'	'Response.Write "<BR> " & stmpids
'	rstDelete.MoveFirst 
'	for i = 1 to rstDelete.RecordCount 
'		rstDelete.Delete adAffectcurrent
'		If not rstDelete.EOF then rstDelete.MoveNext  
'	next 
'	rstDelete.Update
'	closeobj(rstDelete)
	
	'Move Coupon
	sql = "Select * FROM sfTmpOrdersAE WHERE odrtmpSessionID=" & Session("SessionID")
	Set rstTmpOrdersAE = Server.CreateObject("ADODB.RecordSet")		
	rstTmpOrdersAE.Open sql, cnn, adOpenKeyset, adLockOptimistic, adCmdText		
	if rstTmpOrdersAE.RecordCount <= 0 then 
		closeobj(rstTmpOrdersAE)
		exit sub
	end if
		
	sql = "Select * FROM sfOrdersAE WHERE orderAEID=" & iOrderID
	Set rstOrdersAE = Server.CreateObject("ADODB.RecordSet")		
	rstOrdersAE.Open sql, cnn, adOpenKeyset, adLockOptimistic, adCmdText		
	if not rstOrdersAE.RecordCount > 0 then
		rstOrdersAE.AddNew 
		rstOrdersAE("orderAEID") = iOrderID
	end if
	
	If trim(rstTmpOrdersAE("odrtmpCouponCode")) <> "" then rstOrdersAE("orderCouponCode") = rstTmpOrdersAE("odrtmpCouponCode")
	rstOrdersAE.Update 
	closeobj(rstOrdersAE)
	
'	rstTmpOrdersAE.Delete adAffectCurrent 'delete tmp record 
'	rstTmpOrdersAE.Update 
'	
'	closeobj(rstTmpOrdersAE)
		
		
		
End Sub



'-------------------------------------------------------------------------------
'Purpose: gets attdetailid field based on order attributes - for inventory purposes
'Accepts: 
'Returns: a formulated invenAttDetailID field e.g. 48,51,89 
'Called From: 
'Comments:
'-------------------------------------------------------------------------------
Function GetAttDetailID (iDetailID,sType)
dim sql
dim rst
dim i
dim sAttID
	
	Set rst = Server.CreateObject("ADODB.RecordSet")		

	If sType = "svd" then 	sql = "Select * FROM sfSavedOrderAttributes WHERE odrattrsvdOrderDetailID=" & iDetailID & " ORDER BY odrattrsvdAttrID" 
	If sType = "tmp" then 	sql = "Select * FROM sfTmpOrderAttributes WHERE odrattrtmpOrderDetailId=" & iDetailID & " ORDER BY odrattrtmpAttrID" 
	If sType = "odr" then 	sql = "Select * FROM sfOrderDetailsAE WHERE odrdtAEID=" & iDetailID 'b2
	
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText
	If rst.recordcount <= 0 then
		GetAttDetailID = 0
		closeobj(rst)
		exit function
	End if
	
	sAttId = ""
	
	If sType ="odr" then   'b2
		sAttId = rst("odrdtAttDetailID")
		GetAttDetailID = sAttID						
		CloseObj (rst)
		Exit Function
	End If

	rst.movefirst
	For i = 1 to rst.recordcount
		If sType ="svd" then 
			If sAttID <> "" then
				sAttId = sAttId &  "," & rst("odrattrsvdAttrID") 
			else
				sAttId = rst("odrattrsvdAttrID")
			End if
				
		Elseif  sType ="tmp" then 
			if sAttID <> "" then
				sAttId = sAttId & "," & rst("odrattrtmpAttrID") 
			else
				sAttId = rst("odrattrtmpAttrID")
			end if
				
		End if
	
		If not rst.eof then rst.movenext
	Next
		
	GetAttDetailID = sAttID						
	CloseObj (rst)

End Function




'-------------------------------------------------------------------------------
'Purpose: 
'Accepts: 
'Returns:
'Called From: 
'Comments:
'-------------------------------------------------------------------------------
Function GetAttName (iDetailID,sType)
dim sql
dim rst
dim i
dim sAttID
dim sAttName
dim sProdId


	'First get product ID
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	If sType = "svd" then 	sql = "Select * FROM sfSavedOrderDetails WHERE odrdtsvdId=" & iDetailID
	If sType = "tmp" then 	sql = "Select * FROM sfTmpOrderDetails WHERE odrdttmpId=" & iDetailID
	If sType = "odr" then 	sql = "Select * FROM sfOrderDetails WHERE odrdtId=" & iDetailID
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText
	
	If rst.recordcount >  0 then
		if sType = "svd" then sProdId = rst("odrdtsvdProductId")
		if sType = "tmp" then sProdId = rst("odrdttmpProductId")
		if sType = "odr" then sProdId = rst("odrdtProductID")
	End if
	
	CloseObj(rst)
	
	
	
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	If sType = "svd" then 	sql = "Select * FROM sfSavedOrderAttributes WHERE odrattrsvdOrderDetailID=" & iDetailID
	If sType = "tmp" then 	sql = "Select * FROM sfTmpOrderAttributes WHERE odrattrtmpOrderDetailId=" & iDetailID
	If sType = "odr" then 	sql = "Select * FROM sfOrderAttributes WHERE odrattrOrderDetailId=" & iDetailID
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText
	
	If rst.recordcount <= 0 then
		GetAttName = ""
		closeobj(rst)
		exit function
	End if
	
	sAttId = ""
	
	For i = 1 to rst.recordcount
		
		If sType ="svd" then 
			If sAttID <> "" then
				sAttId = sAttId &  "," & rst("odrattrsvdAttrID") 
			else
				sAttId = rst("odrattrsvdAttrID")
			End if
		
				
		Elseif  sType ="tmp" then 
			if sAttID <> "" then
				sAttId = sAttId & "," & rst("odrattrtmpAttrID") 
			else
				sAttId = rst("odrattrtmpAttrID")
			end if
				
		Elseif sType ="odr" then 
			if sAttID <> "" then
				sAttId = sAttId & "," & rst("odrattrAttribute")
			else
				sAttId = rst("odrattrAttribute")
			end if
		End if
				
		If not rst.eof then rst.movenext
		
	Next
	
	
	closeobj(rst)
	
	
	Set rst = Server.CreateObject("ADODB.RecordSet")
	sql = "Select * FROM sfAttributeDetail WHERE attrdtID in (" & sAttId & ") order by attrdtAttributeId" 
	
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText
	
	if rst.recordcount > 0 then
		for i = 1 to rst.recordcount 
			sAttName = sAttName & " " & rst("attrdtName")
			rst.movenext
		next 
	else
		sAttName = ""
	end if
	
	GetAttName = sAttName
	CloseObj (rst) 
End Function



'******************************************* SAVE CART PAGE *********************************


'-------------------------------------------------------------------------------
'Purpose: 
'Accepts: 
'Returns:
'Called From: 
'Comments:
'-------------------------------------------------------------------------------

Sub SaveCart_ShowEmailWishListButton
 If iEmailActive = 1 Then
	Response.Write "<BR>" & vbcrlf
	Response.Write " <a href=""javascript:emailwishlist()""><img border=""0"" src=""" & C_BTN24 & """ alt=""Email your Wish List to Friend(s)""></a>" & vbcrlf
	Response.Write " <BR>" & vbcrlf
 End If 
 End Sub 
 
Sub SaveCart_WritesvdtmpAERecord 
Dim iTmpOrderDetailID,iSvdOrderDetailID
Dim rst
Dim rst2
Dim sql
Dim i		
dim attname
dim attid
	
	
	iTmpOrderDetailID = iTmpCartID
	iSvdOrderDetailID = iSvdOrderID
	
	'read from sfSvdOrderDetails to get values
	attname = GetAttName(iSvdOrderDetailID,"svd")
	attid=GetAttDetailID(iSvdOrderDetailID,"svd")		
	
	
	'write to tmpae
	sql = "Select * FROM sfTmpOrderDetailsAE WHERE odrdttmpAEID=" & iTmpOrderDetailID
	Set rst2 = Server.CreateObject("ADODB.RecordSet")		
	rst2.Open sql, cnn, adopenKeyset, adLockOptimistic, adCmdText
	If Not rst2.recordcount >  0 then
		rst2.AddNew 
		rst2.Fields("odrdttmpaeID") = iTmpOrderDetailID
	End if
		
	rst2.update 
	CloseObj (rst2)

		
End Sub



'********************************** SEARCH_RESULTS PAGE *************************************
Sub SearchResults_GetGiftWrap (strProductID)
Dim ret
	
	
	ret = GetGiftWrapPrice (strProductID)
	
	select case ret
		case "X" 'no gift wrap for this product
			
		case 0 'gift wrap for free if price is 0
		 	Response.Write "<p align=""left""> " & vbcrlf

			Response.Write "<br><INPUT name=chkGiftWrap type=checkbox value = 1 >Gift wrap (free of charge!)</br>" & vbcrlf
			Response.Write "</P>"  & vbcrlf
		case else
	    	Response.Write "<p align=""left"">" & vbcrlf
			Response.Write "<br><INPUT name=chkGiftWrap type=checkbox value = 1 >Gift wrap (add " & FormatCurrency(ret)  & " per item)</br>" & vbcrlf
			Response.Write "</P>"  & vbcrlf
    end select
    
    
End Sub



Sub Order_ShowCouponLink 
Dim spath,jsvar
	If not CouponOn then exit sub
		sPath = "Coupon.asp"
		jsvar = "javascript:show_page(" & "'" & sPath & "')"
		
		Response.Write "<p align = ""center"" ><a href=" & jsvar & ">Enter Coupon</a></p><BR>"	 & vbcrlf
		
End Sub


Sub SearchResults_ShowMTPricesLink(sProdId) 

Dim sql
Dim rst
Dim i

	
	sql = "Select * FROM sfMTPrices WHERE mtprodid= '" & sProdID & "' ORDER By mtIndex ASC"
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText
	If rst.recordcount > 0 then
		Dim spath,stype,jsvar
		sType = 0
		sPath = "MTPrices.asp?sProdId=" & sProdID
		jsvar = "javascript:show_page(" & "'" & sPath & "')"
		
		Response.Write "<BR> <a href=""" & jsvar & """>Check Volume Discounts</a>" & vbcrlf	
		
	End if

	CloseObj (rst)

End Sub

Sub SearchResults_CheckInventoryTracked
'Do we want to show out-of-stock with no backorders items ?
'If CheckInventoryTracked(arrProduct(0, iRec)) = 1  then 'b2
 	'If CheckInStock (arrProduct(0, iRec)) <= 0  AND  CheckBackOrder (arrProduct(0, iRec)) <> 1 then
	'	iRec = iRec + 1
	'End If
'End If	
End Sub



Sub SearchResults_GetProductInventory (strProduct)
Dim ret	,instock, sPath,jsvar

	
	'If gShowInStock <> 1 then Exit Sub
	
	
	ret = CheckInventoryTracked (strProduct)
	If ret = 1 then 
		
		If CheckShowStatus(strProduct) <> 1 then exit sub
			
		Instock = CheckInStock(strProduct)	
		Select case instock
			case "X" 'inventory not tracked for this product
				
				
			case 0 'inventory tracked
				Response.Write "<BR> " & vbcrlf
				Response.Write "Out of Stock!" & vbcrlf
				If checkbackorder(strProduct) = 1 then	
					Response.Write "<BR> Click ""Add to Cart"" to BackOrder!" & vbcrlf
				End If
				Response.Write "" & vbcrlf
													
			case else

				sPath = "StockInfo.asp?sProdId=" & strProduct
				jsvar = "javascript:show_stockinfo(" & "'" & sPath & "')"
				Response.Write "<BR> <a href=""" & jsvar & """>In Stock!</a>" & vbcrlf	
								
		End  Select
	end if
	
End sub



'********************************** VERIFY PAGE ****************************

Sub Verify_SetConfirmPath
	
	If  CheckCartInventory = 1  then
		sConfirmPath = sLPath & "ssl/confirm.asp"		
	else
		sConfirmPath = sLPath & "order.asp"		
	end if
End Sub

Sub Verify_SetProdPrice
	dUnitPrice = GetMTPrice(sProdId,dUnitPrice,0)'AE
End Sub

Sub OVC_ShowCouponDiscount
		           
					Response.Write "<tr>" & vbcrlf
					Response.Write "<td width=""75%"" align=""right"">Coupon Discount:</td>" & vbcrlf
					Response.Write "<td width=""25%"" align =""left"" height=""20"" nowrap>-" & gCouponDiscountOut & "</td>" & vbcrlf
	            	Response.Write "</tr>"  & vbcrlf

  
End Sub



Sub OVC_ShowOrderSubTotalWOD 

	Response.Write "<tr>" & vbcrlf
	Response.Write "<td width=""75%"" align=""right"">Sub Total:</td>" & vbcrlf
	If iConverion = 1 Then			
				'									 gCouponDiscountOut = "<script>document.write(""" & FormatCurrency(gCouponDiscountOut) &						" = ("" + OANDAconvert(" & gCouponDiscountOut &						   ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script>"  
	Response.Write "<td width=""25%"" align =""left"" height=""20"" nowrap><script>document.write(""" & FormatCurrency(FormatCurrency(Session("sTotalPriceWOD"))) & " = ("" + OANDAconvert(" & Session("sTotalPriceWOD") & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script></td>" & vbcrlf
			else
	Response.Write "<td width=""25%"" align =""left"" height=""20"" nowrap>" & FormatCurrency(Session("sTotalPriceWOD")) & "</td>" & vbcrlf		
	end if			

	
	Response.Write "</tr> " & vbcrlf
	

End Sub



Sub OVC_ShowStoreWideDiscount

		Response.Write "<tr>" & vbcrlf
		Response.Write "<td width=""75%"" align=""right"">Store Wide Discount:</td>" & vbcrlf
		Response.Write "<td width=""25%"" align =""left"" height=""20"" nowrap>-" & gStoreWideDiscountOut & "</td>" & vbcrlf
		Response.Write "</tr>"  & vbcrlf
	
			        	            
    
End Sub
Sub Order_InitializeDiscounts
	Session("CouponDiscountPercent") = 0
	Session("CouponDiscountAmount") = 0
End Sub
Sub OVC_ShowOrderDiscounts 
		If Session("CouponDiscountPercent") > 0 or Session("StoreWideDiscount") > 0 or Session("CouponDiscountAmount") > 0 then
			OVC_ShowOrderSubTotalWOD 
		End If 
		
	    If Session("CouponDiscountPercent") > 0 then 
			gCouponDiscountOut = cdbl( Session("CouponDiscountPercent"))
			If iConverion = 1 Then			
				gCouponDiscountOut = "<script>document.write(""" & FormatCurrency(gCouponDiscountOut) & " = ("" + OANDAconvert(" & gCouponDiscountOut & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script>"  
			else
				gCouponDiscountOut = FormatCurrency( Session("CouponDiscountPercent"))
			End If
			OVC_ShowCouponDiscount
		ELSEIf Session("CouponDiscountAmount") > 0 then
			gCouponDiscountOut = cdbl( Session("CouponDiscountAmount"))
			If iConverion = 1 Then			
				gCouponDiscountOut = "<script>document.write(""" & FormatCurrency(gCouponDiscountOut) & " = ("" + OANDAconvert(" & gCouponDiscountOut & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script>"  
			else
				gCouponDiscountOut = FormatCurrency( Session("CouponDiscountAmount"))
			End If
			OVC_ShowCouponDiscount
		END IF
		
		If Session("StoreWideDiscount") > 0 then
			gStoreWideDiscountOut  = Session("StoreWideDiscount")
			If iConverion = 1 Then			
				gStoreWideDiscountOut = "<script>document.write(""" & FormatCurrency(gStoreWideDiscountOut) & " = ("" + OANDAconvert(" & gStoreWideDiscountOut & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script>"  
			else
				gStoreWideDiscountOut = FormatCurrency(Session("StoreWideDiscount"))
			End If
			OVC_ShowStoreWideDiscount
		END If

End Sub

Sub OVC_AddGiftWrapTotal
	sTotalPrice = cdbl(sTotalPrice) + cdbl(GetOrderGiftWrapTotal_Tmp)
End Sub

Sub OVC_SaveSubTotalWOD
	Session("sTotalPriceWOD") = cdbl(sTotalPrice)
End Sub

Sub OVC_SaveValues
	Session("sTotalPrice") = cdbl(sTotalPrice)
	Session("sShipping") = cdbl(sShipping)
	Session("iSTax") = cdbl(iSTax)
	Session("iCTax") = cdbl(iCTax)
	Session("sHandling") = cdbl(sHandling)
	Session("iCODAmount")= cDbl(iCodAmount)
	
End Sub		
Sub OVC_GetValues
	sTotalPrice = cdbl(Session("sTotalPrice"))
	sShipping = cdbl(Session("sShipping"))
	iSTax = cdbl(Session("iSTax"))
	iCTax =cdbl(Session("iCTax"))
	sHandling = cdbl(Session("sHandling"))
	iCodAmount = cdbl(Session("iCODAmount"))
End Sub

Sub OVC_GetSubTotalWOD
	sTotalPrice = cdbl(Session("sTotalPriceWOD"))
End Sub

Function ApplyPercentDiscounts (xTotalPrice,sReturn)
Dim x,y,TotalDiscounts,NewTotal
	
	x=ApplyCouponDiscount(xTotalPrice,"Percent","Discount")
	y=ApplyStoreWideDiscount(xTotalPrice,"Discount")
	
	TotalDiscounts = cdbl(x) + cdbl(y)
	NewTotal = cdbl(xTotalPrice) - cdbl(TotalDiscounts)
	
	If TotalDiscounts < 0 then TotalDiscounts = 0
	If NewTotal < 0 then NewTotal = 0 
	
	If sReturn = "Discount" then
		ApplyPercentDiscounts = cdbl(TotalDiscounts)
	elseif sReturn = "Total" then
		ApplyPercentDiscounts = cdbl(NewTotal)
	End If
	'Response.Write "<BR> xTotalPrice(afterdiscount):" & xTotalPrice
	'Response.Write "<BR> NewTotal:" & NewTotal
	'Response.Write "<BR> Discount:" & totalDiscounts

End Function

'********************************* ORDER PAGE (inventory stuff) ***********************************************
Sub Order_ShowInventoryMessage

Dim spath
	Order_AdjustCart
	If Session("ShowInventoryMessage") = "1" then
		'Order_AdjustCart
		sPath = "invenMessage.asp"'
		js "show_page('" & sPath & "')"
	End if
	
	Session("ShowInventoryMessage") = "0"
End Sub

Sub Order_AdjustCart 'B2
Dim ret
	If  CheckCartInventory = 0  then
	
	dim rstAll,sql,i,sPath,sProdName,sAttName,bo
	dim inv,ordqty,itmporderdetailid,sprodid,avlqty,sattdetailid,sResponseMessage
	dim boqty
	SQL = "Select * FROM sfTmpOrderDetails as A LEFT JOIN sfTmpOrderDetailsAE as B ON A.odrdttmpID = B.odrdttmpaeID WHERE odrdttmpSessionID = " & Session("SessionID")
	
	Set rstAll = Server.CreateObject("ADODB.RecordSet")
	rstAll.Open SQL, cnn, adOpenKeyset, adLockReadOnly, adCmdText
	
	For i = 1 to rstall.RecordCount 
	
		inv =  CheckInventoryTracked(rstall("odrdttmpProductID"))  'inventory tracked for  this product?
		'bo = CheckBackOrder(rstall("odrdttmpProductID")) ' backorder allowed ?
		ordQTY = rstall("odrdttmpQuantity")
		boqty = rstAll("odrdttmpBackOrderQTY")
		iTmpOrderDetailID = rstall("odrdttmpID")
		sProdID = rstall("odrdttmpProductID")
		sAttDetailID = GetAttDetailID(iTmpOrderDetailID,"tmp")
		avlqty = GetAvailableQTY(sProdId,sAttDetailID) 
		sProdName = GetProductName(sProdId)
		sAttName = getattname(iTmpOrderDetailId,"tmp")
		
		If ordqty > 1 Then
			sResponseMessage = "have been added to your order."
		Else
			sResponseMessage = "has been added to your order."
		End If
		
		If  ordqty > avlqty AND boqty <= 0 then
			sPath = "InventoryOD.asp?sProdID=" & sProdID & "&iTmpOrderDetailID=" & iTmpOrderDetailID & "&iQuantity=" & iQuantity & "&sProdName=" & sProdName & "&sResponseMessage="& Server.URLEncode(sResponseMessage) 'AE
			js "show_page('" & sPath & "')"
		End if
		If not  rstAll.EOF then	rstAll.movenext
	
	next
	
	
	else
		ret = DeleteBadItems
		ValidateCartItems   ' corrects items qty according to backorder-flag and stock-qty
		ret = DeleteBadItems
		'Order_ProcessCoupon ' writes any coupon code entered in the box -- ONLY if RECALC pushed
		'Order_RecalcTotal   ' recalculate all the totals on this page
		
	End If
			
End Sub


Sub Order_FixTable  'b2
	Response.Write "</table>" & vbcrlf
	Response.Write "<table>" & vbcrlf
End Sub

Sub Order_ShowInStockValue
'If gShowInStock <> 1 then exit sub

dim rst
dim sql
dim avlqty	
dim sAttID	
	sql= gtmpSQL & "WHERE odrdttmpID=" & rsAllOrders("odrdttmpid")
	
	If CheckInventoryTracked(rsAllOrders("odrdttmpProductID")) <> 1 then 

		
		Response.Write "<td width=""15%"" align=""center"" class='" & fontClass & "' valign=""top"" nowrap></td> " & vbcrlf
		
		Exit sub
	End If	
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly,adcmdText
	If rst.recordcount > 0 then
	
	sAttID = GetAttDetailID(rst("odrdttmpID"),"tmp")
	avlqty = GetAvailableqty(rst("odrdttmpProductID"),sAttID) 
	
	End IF
			Response.Write "<td width=""15%"" align=""center"" class='" & fontClass & "' valign=""top"" nowrap>" & avlqty & "</td>" & vbcrlf

	CloseObj (rst)

End Sub


Sub InsertBlankRow
		Response.Write "<tr>" & vbcrlf
		Response.Write "<td width=""15%"" align=""left"" class='" & fontClass & "' valign=""top"" nowrap>&nbsp;</td>"  & vbcrlf
		Response.Write "<td width=""15%"" align=""left"" class='" & fontClass & "' valign=""top"" nowrap>&nbsp;</td> " & vbcrlf
		Response.Write "<td width=""15%"" align=""left"" class='" & fontClass & "' valign=""top"" nowrap>&nbsp;</td> " & vbcrlf
		Response.Write "<td width=""15%"" align=""left"" class='" & fontClass & "' valign=""top"" nowrap>&nbsp;</td> " & vbcrlf
		Response.Write "<td width=""15%"" align=""left"" class='" & fontClass & "' valign=""top"" nowrap>&nbsp;</td> " & vbcrlf
		Response.Write "</tr>" & vbcrlf

End Sub


Sub Order_ProcessCoupon
	If not CouponOn then exit sub
	
	gCouponCode = Request.form("FormCouponCode")
	'Response.Write gCouponCode
	'Response.End 
	Call Order_WriteCouponCode(gCouponCode)
	
End Sub


Sub Order_ShowCouponInput
	 
	if not CouponOn then exit sub  


	Response.Write "<table>" & vbcrlf
	Response.Write "<BR>" & vbcrlf
	Response.Write "<td width=""15%"" align=""left""  background=""" & sBkGrnd & """ valign=""top"" nowrap> Enter Coupon Code: <input type=""text"" style=""" & C_FORMDESIGN & """ size=""22"" name=""FormCouponCode""> " & vbcrlf
	Response.Write "</td>" & vbcrlf
	Response.Write "</table>" & vbcrlf


End  Sub

Sub Order_ShowGiftWrapValues

dim rst
dim sql
dim gGiftWrapTotal
	
	gGiftWrapTotal = 0
	
	
	sql= gtmpSQL & "WHERE odrdttmpID=" & rsAllOrders("odrdttmpid")
	
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly,adcmdText
	
	if rst.recordcount > 0 then
		if  rst("odrdttmpGiftWrapQty") > 0 then
	
			gGiftWrapTotal = FormatCurrency(cdbl(rst("odrdttmpGiftWrapQty") * GetGiftWrapPrice(rst("odrdttmpProductID"))))

			gGiftWrapGrandTotal = cdbl(gGiftWrapGrandTotal + gGiftWrapTotal)
		end if
	end if

	If GetGiftWrapPrice (rsAllOrders("odrdttmpProductID")) <>  "X" then 
	Response.Write "<td width=""15%"" align=""center"" class='" & fontClass & "' valign=""top"" nowrap><input type=""text"" style=""" & C_FORMDESIGN & """ name=""GWQTY" & iProductCounter &""" size=""2"" value=""" & rst("odrdttmpGiftWrapQty") & """></td>" & vbcrlf
	Response.Write "<td width=""15%"" align=""center"" class='" & fontClass & "' valign=""top"" nowrap>" & formatcurrency(gGiftWrapTotal) & "</td>" & vbcrlf
	sProductSubtotal = cdbl(sProductSubtotal) + cdbl(gGiftWrapTotal) 
	else 
		Response.Write "<td width=""15%"" align=""center"" class='" & fontClass & "' valign=""top"" nowrap></td> " & vbcrlf
		Response.Write "<td width=""15%"" align=""center"" class='" & fontClass & "' valign=""top"" nowrap></td> " & vbcrlf
	end if
	
	CloseObj (rst)


End Sub


Sub Order_Update_GiftWrapsBackOrder

dim rst
dim sql
dim Price
dim gwqty
dim boflag
dim ProdId,avlqty
	
	sql = "Select * FROM sfTmpOrderDetailsAE WHERE odrdttmpaeID=" & iTmpOrderID
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenKeyset, adLockOptimistic, adCmdText
	if rst.recordcount <= 0 then exit sub 'should never happen
	
	gwqty = Request.Form("GWQTY" & iCounter)
	'Response.Write  
	prodid = Request.Form("sProdID" & iCounter)
	If gwqty <> "" AND Not IsNull(gwqty) then
		If GetGiftWrapPrice(ProdId) <> "X" then
				If gwqty > iNewQuantity then gwqty = iNewQuantity
				rst.Fields("odrdttmpGiftWrapQTY") = gwqty
		End if
	End if
	
	
	avlqty = getavailableqty(prodid,getattdetailid(rst("odrdttmpaeID"),"tmp"))
	If avlqty <> "X" Then
	IF clng(avlqty) => clng(iNewQuantity) then
		rst.Fields("odrdttmpBackOrderQty") = 0 'beta 2
	End if
	End IF
	
	rst.update
	CloseObj (rst)
	'response.Redirect(C_HomePath & "order.asp")	
End Sub


'Sub Order_RecalcTotal
	'dTotalPrice = GetAETotals(dTotalPrice)
'End Sub

Sub Order_WriteCouponCode (sCouponCode)
dim sql,rst

	 
	If not CouponOn then exit sub
	'Write to sfTmpOrdersAE
	sql = "Select * FROM sfTmpOrdersAE where odrtmpsessionid=" & Session("SessionID")
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenKeyset, adLockOptimistic, adCmdText
	if not rst.recordcount > 0 then
		rst.AddNew
		rst("odrtmpSessionId") = Session("SessionID")
	end if
    if trim(sCouponCode) <> "" then
		rst("odrtmpCouponCode") = sCouponCode
	end if
    rst.update
    
    CloseObj (rst)
End sub

Sub Order_SetDeleteOrderAE (tmpOrderID)
dim sql
		
	If 	TmpOrderID <> "" then
	sql = "DELETE FROM sfTmpOrderDetailsAE WHERE odrdttmpaeID = " & tmpOrderID
	cnn.Execute(SQL)
	End if
	
End sub

Sub Order_SetProdPrice
	dUnitPrice = GetMTPrice(sProdId,dUnitPrice,0)'AE
End Sub



Sub Order_ShowGiftWrapHeading
dim sql,rst


		Response.Write "<td width=""15%"" align=""center"" class=""tdContentBar"">gift wrap items</td> " & vbcrlf
		Response.Write "<td width=""15%"" align=""center"" class=""tdContentBar"">gift wrap price</td>" & vbcrlf
End Sub
Sub Order_SetAEColSpan

		Response.Write "</TD>" & vbcrlf
		Response.Write "<td colspan=""7"" width=""40%"" class='" & fontClass & "'>" & vbcrlf

End Sub




'************************************** ADD PRODUCT PAGE ***************************************************
Sub AddProduct_SetsResponseMessage
	sResponseMessage = Replace(sResponseMessage,"cart","Wish List",1,1,1)
End Sub

Sub AddProduct_SetThanksMessage
		
	If Not (sActionType = "SaveProduct") AND CheckInventoryTracked(sProdID) = 1 then 
		sThanksPath = "inventoryAP.asp?sProdID=" & sProdID & "&iTmpOrderDetailID=" & iTmpOrderDetailID & "&iQuantity=" & iQuantity & "&sProdName=" & sProdName & "&sProdMessage=" & sProdMessage & "&sResponseMessage="& Server.URLEncode(sResponseMessage) 'AE
	End If
End Sub

Sub AddProduct_ToAddOrNotToAdd
' New Order
		  If (iTmpOrderDetailID < 1) Then
			  If vDebug = 1 Then Response.Write "<h2><font face=verdana color=#334455>New Order</h2>"
			  ' Write to TmpOrderDetails Table 
			  iTmpOrderDetailID = getTmpTable(aProdAttr,sProdID,iQuantity,sReferer,iShip)		 
			  
			  
		' Add to Existing Order
		  ElseIf (iTmpOrderDetailID > 0 AND iTmpOrderDetailID <> "" ) Then
			
			If  Request.Form("chkGiftWrap") <> CheckTmpOrderGiftWrap (iTmpOrderDetailID) then
				'Response.Write iTmpOrderDetailID
				'Response.Write Request.Form("chkGiftWrap") 
				'Response.Write CheckTmpOrderGiftWrap (iTmpOrderDetailID)
				'Response.End 
				
				'New record
				If vDebug = 1 Then Response.Write "<h2><font face=verdana color=#334455>New Order</h2>"
			  ' Write to TmpOrderDetails Table 
		  	   iTmpOrderDetailID = getTmpTable(aProdAttr,sProdID,iQuantity,sReferer,iShip)		 
			else 
				' Check to see if it is new product or adding to existing products
				If vDebug = 1 Then Response.Write "<h2><font face=""verdana"" color=""#334455"">Adding to Existing Cart</font></h2>"
				If vDebug = 1 Then Response.Write "<hr>Found OrderID = " & iTmpOrderDetailId
				
				' Update Quantity					
				Call setUpdateQuantity("odrdttmp",iQuantity,iTmpOrderDetailId)				
			end if
		
		' End Add Product If
		End If


End Sub
Sub AddProduct_WriteTmpOrderDetailsAE
Dim rst
Dim sql
Dim invenAttName
dim invenAttDetailID
Dim qty
	
	
	
	invenAttName=""
	invenAttDetailId=0
	qty =  iQuantity
	
	invenAttName = GetAttName(iTmpOrderDetailID,"tmp")
	invenAttDetailID = GetAttDetailID(iTmpOrderDetailID,"tmp")
	qty = GetTmpQTY(iTmpOrderDetailID)
	
	if not isnumeric(qty) then qty = iQuantity
	
		
	'Write to sfTmpOrderDetailsAE
	sql = "Select * FROM sfTmpOrderDetailsAE WHERE odrdttmpaeID=" & iTmpOrderDetailID
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenKeyset, adLockOptimistic, adCmdText
	If not rst.recordcount > 0 then
		rst.AddNew 
		rst.Fields("odrdttmpaeID") = iTmpOrderDetailID
	End If
    If  Request.Form("chkGiftWrap") = 1 then
		if rst.Fields("odrdttmpGiftWrapQty") > 0 then
			rst.Fields("odrdttmpGiftWrapQty") = rst.Fields("odrdttmpGiftWrapQty") + iQuantity
		else
			rst.Fields("odrdttmpGiftWrapQty") = iQuantity
		end if
	End If
		
	
	rst.update
	CloseObj (rst)
	
End Sub


'*****************************************************************************************
'*****************************************************************************************
'********************** Independent procedures and functions *****************************
'*****************************************************************************************
'*****************************************************************************************


'-------------------------------------------------------------------------------
'Purpose: 
'Accepts: 
'Returns:
'Called From: 
'Comments:
'-------------------------------------------------------------------------------
Sub AdjustTmpOrderDetails (iTmpOrderDetailID)

Dim sql,rst
dim gwqty
dim avlqty
dim ordqty


	sql = gtmpSQL & " where odrdttmpID=" & iTmpOrderDetailID 'rsAllOrders("odrdttmpProductID")
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn,  adOpenKeySet, adLockOptimistic, adCmdText
	
	If rst.RecordCount <= 0 then exit sub 'error this should never happen
	
	avlqty = GetAvailableQTY(rst("odrdttmpProductID"), GetAttDetailID (rst("odrdttmpId"),"tmp"))
	gwqty =  rst("odrdttmpGiftWrapQty")
	ordqty = rst("odrdttmpQuantity")
	
	IF CheckInventoryTracked(rst("odrdttmpProductID")) <> 1 or Not IsNumeric(avlqty) then 'b2
		closeobj(rst)
		Exit Sub
	End IF
	
	If avlqty = 0  AND CheckBackOrder(rst("odrdttmpProductID")) <> 1 then
		' Delete item if out of stock with no backorder
		closeobj(rst)
		DeletetmpOrderDetailsAE(rst("odrdttmpId"))
		exit sub	
	End if
	
	
	'If avlqty < ordqty then ordqty = avlqty
	
	'beta 2
	If avlqty < ordqty then 
		rst("odrdttmpBackOrderQty") = clng(ordqty)- clng(avlqty)
	else
		rst("odrdttmpBackOrderQty") = 0
	end if
		
	If gwqty > ordqty then gwqty =ordqty
	
	rst("odrdttmpGiftWrapQty") = gwqty
	rst.update
	
	rst("odrdttmpQuantity")	= ordqty
	rst.update
	CloseObj (rst)

End Sub



Function GetProdGiftWrapPrice(sProdID) '8/15
	GetProdGiftWrapPrice = 0
	GetProdGiftWrapPrice = getgiftwrapprice(sProdID)
	If GetProdGiftWrapPrice = "X" or GetProdGiftWrapPrice = "" then 
		GetProdGiftWrapPrice = 0
	elseIf GetProdGiftWrapPrice < 0 then 
		GetProdGiftWrapPrice = 0
	
	else
		GetProdGiftWrapPrice=  GetProdGiftWrapPrice 'in porgress
	End IF
End Function

Function ApplyStoreWideDiscount(xTotal,sReturn) '8/15
Dim oldamt ,DiscountAmount,NewTotal
	Session("StoreWideDiscount") = 0
	'gStoreWideDiscountOut 
	oldamt = xTotal
	DiscountAmount = 0
	
	NewTotal = getGlobalSalePriceAE(xTotal)
	
	DiscountAmount = cdbl(oldamt) - cdbl(NewTotal)
	
	If NewTotal < 0 then NewTotal = 0
	If DiscountAmount < 0 then DiscountAmount = 0
	
	If sReturn = "Discount" then
		ApplyStoreWideDiscount = DiscountAmount
	elseif sReturn = "Total" then
		ApplyStoreWideDiscount = cdbl(NewTotal)
	End if
	Session("StoreWideDiscount") = cdbl(DiscountAmount)
	
End Function

Function GetOrderGiftWrapTotal_TMP()

Dim rst,sql
Dim I
Dim oldamt

	GetOrderGiftWrapTotal_tmp = 0
		
	
	sql = gtmpSQL & " WHERE odrdttmpSessionID=" & Session("SessionID")
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly,adcmdText
	If rst.recordcount <= 0 then 
		GetOrderGiftWrapTotal_TMP = 0
		exit function
	end if
	
			
	rst.MoveFirst
	For i = 1 to rst.recordcount
		'GiftWrap
		If rst("odrdttmpGiftWrapQty") > 0 then
			GetOrderGiftWrapTotal_TMP = GetOrderGiftWrapTotal_TMP + cdbl(rst("odrdttmpGiftWrapQty") * GetGiftWrapPrice(rst("odrdttmpProductID")))
		End if
		
		If not rst.EOF then rst.MoveNext
	Next 
	
	CloseObj(rst)
	
End Function




Function GetOrderGiftWrapTotal_ODR (iOrderID)

Dim rst,sql
Dim I
Dim oldamt
	sql = gorderSQL & " WHERE odrdtOrderId=" & iOrderID
	
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly,adcmdText
	If rst.recordcount <= 0 then 
		GetOrderGiftWrapTotal_ODR = 0
		exit function
	end if
	
			
	rst.MoveFirst
	For i = 1 to rst.recordcount
		'GiftWrap
		If rst("odrdtGiftWrapQty") > 0 then
			GetOrderGiftWrapTotal_ODR = GetOrderGiftWrapTotal_ODR + cdbl(rst("odrdtGiftWrapQty") * GetGiftWrapPrice(rst("odrdtProductID")))
		End if
		
		If not rst.EOF then rst.MoveNext
	Next 
	
	CloseObj(rst)
	
End Function





'-------------------------------------------------------------------------------
'Purpose: Runs any javascript function
'Accepts: name of javascript
'Returns: nothing
'Called From: 
'Comments:
'-------------------------------------------------------------------------------
Sub JS(sfunction)
	sfunction = replace(sfunction,";","")
	Response.Write "<SCRIPT LANGUAGE=" & chr(34) & "javascript" & chr(34) & ">" & vbcrlf & vbcrlf
	Response.write sfunction & ";" & vbcrlf
	Response.Write "</SCRIPT>" & vbcrlf
End sub


'-------------------------------------------------------------------------------
'Purpose: 
'Accepts: 
'Returns:
'Called From: 
'Comments:
'-------------------------------------------------------------------------------
Sub CloseObj (objItem)
		On Error Resume Next
		objItem.Close
		Set objItem=nothing	
		On Error GoTo 0
End Sub



'-------------------------------------------------------------------------------
'Purpose: 
'Accepts: 
'Returns:
'Called From: 
'Comments:
'-------------------------------------------------------------------------------
Function CheckLogin

Dim iCustID,  iSessionID	
	    
	' Request Cookie for custID 
	iCustID		= Trim(Request.Cookies("sfCustomer")("custID"))
	iSessionID	= Trim(Request.Cookies("sfOrder")("SessionID"))
	    
	If iCustID = "" or iSessionID <> Session("SessionID") Then			
		CheckLogin = 0 'no login 
	Else
		CheckLogin = 1 
	End if
	
End Function


'-------------------------------------------------------------------------------
'Purpose: deletes any items from the cart with qty = 0 
'Accepts: 
'Returns:
'Called From: 
'Comments:
'-------------------------------------------------------------------------------
Function DeleteBadItems 'b2

dim rst
dim sql
dim i
	
	DeleteBadItems = 0

	sql= "Select * FROM sfTmpOrderDetails WHERE odrdttmpSessionID=" & Session("SessionID")
	
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenKeySet, adLockOptimistic,adcmdText
	
	If rst.recordcount > 0 then
		rst.movefirst
		For i = 1 to rst.recordcount
			
			If  rst("odrdttmpQuantity") <= 0 then
				rst.delete   'b2
				rst.update   'b2
				DeletetmpOrderDetailsAE rst("odrdttmpID") 
				DeleteBadItems = 1	
			End if
		if not rst.eof then rst.movenext
		Next 
				
	end if
	CloseObj (rst)

End Function




'-------------------------------------------------------------------------------
'Purpose: 
'Accepts: 
'Returns:
'Called From: 
'Comments:
'-------------------------------------------------------------------------------
Sub DeleteTmpOrderDetailsAE (iTmpOrderDetailID)
Dim sql,rst
	
	if Not isnumeric(iTmpOrderDetailID) then iTmpOrderDetailID = 0
	
	Set rst = Server.CreateObject("ADODB.Command")		
	
	'sftmpOrderDetailsAE
	sql = "DELETE FROM sftmporderdetailsAE WHERE odrdttmpaeID=" & iTmpOrderDetailID
	rst.ActiveConnection = cnn
	rst.CommandText = sql
	rst.Execute
	
	'sftmpOrderDetails
	sql = "DELETE FROM sftmporderdetails WHERE odrdttmpID=" & iTmpOrderDetailID
	rst.ActiveConnection = cnn
	rst.CommandText = sql
	rst.Execute
	
	'sftmpOrderAttributes
	sql = "DELETE FROM sfTmpOrderAttributes WHERE odrattrtmpOrderDetailId=" & iTmpOrderDetailID
	rst.ActiveConnection = cnn
	rst.CommandText = sql
	rst.Execute
	
	CloseObj(rst)	
	
End Sub




'************************************************************************************
'**********************  MULTI-TIER PRICING **********************************************
'************************************************************************************

'-------------------------------------------------------------------------------
'Purpose: 
'Accepts: 
'Returns:price for a single product based on volume discount (MTP)
'Called From: 
'Comments:
'-------------------------------------------------------------------------------
Function GetMTPrice (strProductID,sProdPrice,lOrderId) 
Dim rst,sql
Dim totqty
dim diff	
	If lOrderId > 0 then 
		'get from order details
		sql = "Select SUM(odrdtQuantity) as totqty FROM sfOrderDetails WHERE odrdtProductID= '" & strProductID & "' AND odrdtOrderId= " & lOrderID
	else 
		'get from temp order details
		sql = "Select SUM(odrdttmpQuantity) as totqty FROM sftmpOrderDetails WHERE odrdttmpProductID= '" & strProductID & "' AND odrdttmpSessionID= " & session("SessionID")
	End If
	
	Set rst = Server.CreateObject("ADODB.RecordSet")
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText	
	if rst.recordcount <= 0 then
		GetMTPrice = sProdPrice 'no mtp
		CloseObj (rst)
		exit function
	end if
	totqty = rst("totqty")
	CloseObj (rst)
	
	sql = "Select * FROM sfMTPrices WHERE mtprodid= '" & strProductID & "' AND mtQUANTITY  <= " & totqty & " ORDER BY mtValue DESC"
	Set rst = Server.CreateObject("ADODB.RecordSet")
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText	
	
	If rst.recordcount <= 0 then  
		GetMTPrice = sProdPrice 'no mtp
	else
		rst.movefirst
		If rst("mtType") = "Amount" then
			GetMTPrice = cdbl(sProdPrice) - cdbl(rst("mtValue"))
		else
			diff = cdbl(sProdPrice) * (cdbl(rst("mtValue"))/100)
			GetMTPrice = cdbl(sProdPrice) - cdbl(diff) 
		end if
	End If
	
	If cdbl(GetMTPrice) > cdbl(sProdPrice) then 
		GetMTPrice = sProdPrice
	End IF
	
	
	CloseObj (rst)	
	If GetMTPrice < 0 then GetMTPrice = 0

End Function



'************************************************************************************
'**********************  GIFT WRAPPING **********************************************
'************************************************************************************
'-------------------------------------------------------------------------------
'Purpose: 
'Accepts: 
'Returns:
'Called From: 
'Comments:
'-------------------------------------------------------------------------------
Function GetGiftWrapPrice (strProductID) 
Dim rst,sql
		
	sql = "Select * FROM sfgiftwraps WHERE gwProdID= '" & strProductID & "'"
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText	
	
	if rst.recordcount <= 0 then  
		GetGiftWrapPrice = "X"
		CloseObj (rst)
		exit function
	end if
	if rst("gwActivate") = 0 then 
		GetGiftWrapPrice = "X"
		CloseObj (rst)
		exit function
	end if
	
	GetGiftWrapPrice =  rst("gwPrice") 
	CloseObj (rst)
	   
End Function

'-------------------------------------------------------------------------------
'Purpose: 
'Accepts: 
'Returns:
'Called From: 
'Comments:
'-------------------------------------------------------------------------------
Function CheckTmpOrderGiftWrap (iTmpOrderDetailID) 
Dim rst,sql
		
	sql = "Select * FROM sftmpOrderDetailsAE WHERE odrdttmpaeID= " & iTmpOrderDetailID
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText	
	
	if rst.recordcount <= 0 then  
		CheckTmpOrderGiftWrap = 0
	else
		if rst("odrdttmpGiftWrapQty") > 0 then 
			CheckTmpOrderGiftWrap = 1
		else
			CheckTmpOrderGiftWrap = 0
		end if
	
	end if
	
	CloseObj (rst)
	   
End Function


'************************************************************************************
'**************************  COUPONING  **********************************************
'************************************************************************************


'-------------------------------------------------------------------------------
'Purpose: 
'Accepts: 
'Returns:
'Called From: 
'Comments:
'-------------------------------------------------------------------------------
Function GetSessionCouponCode
Dim rst,sql

	
	If not CouponOn then exit function
		
	sql = "Select * FROM sftmpOrdersAE WHERE odrtmpSessionID= " & Session("SessionID")
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText	
	
	if rst.recordcount > 0 then  
		GetSessionCouponCode = rst("odrtmpCouponCode")
		gCouponCode =rst("odrtmpCouponCode")
	else
	
	end if
	
	CloseObj (rst)
	
End Function



'-------------------------------------------------------------------------------
'Purpose: 
'Accepts: 
'Returns:
'Called From: 
'Comments: updated for .3008
'-------------------------------------------------------------------------------
Function ApplyCouponDiscount (xTotalPrice,sType,sReturn) 
Dim rst,sql,strCouponCode
Dim cpValue, oldamt, NewTotal,DiscountAmount,NoDiscount
	'Response.Write "<BR> dd"
	strCouponCode =GetSessionCouponCode()
	oldamt = xTotalPrice
	NewTotal = xTotalPrice
	DiscountAmount = 0
	NoDiscount = 0
	'Session("CouponDiscountPercent") = 0
	'Session("CouponDiscountAmount") = 0
	
	sql = "Select * FROM sfCoupons WHERE cpCouponCode= '" & strCouponCode & "' AND cpActivate = 1"
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText	
	
	If rst.recordcount <= 0 then 
		 NoDiscount = 1
	else
		If cdbl(rst("cpMin")) > cdbl(xTotalPrice) then NoDiscount = 1
		If rst("cpNeverExpire") = 0 AND rst("cpExpirationDate") < date then  NoDiscount = 1
	end if
	
	If NoDiscount = 1 then
		if sReturn = "Discount" then
			ApplyCouponDiscount = 0
		elseif sReturn ="Total"  Then 'total amount ?
			ApplyCouponDiscount = xTotalPrice
		end if	
	If  sType = "Percent" then
		Session("CouponDiscountPercent") = 0
	elseif sType = "Amount" then
		Session("CouponDiscountAmout") = 0
	end if		
		CloseObj (rst)
		exit function
	End If
			
	
	'percent based coupon
	If sType = "Percent"  Then
		cpValue = rst("cpValue")
		If  rst("cpType") = "Percent" then
			cpValue =  (xTotalPrice *  rst("cpValue"))/100
			NewTotal = xTotalPrice - cpValue
			DiscountAmount =  oldamt - NewTotal
			Session("CouponDiscountPercent") = cdbl(DiscountAmount)
		else
			DiscountAmount = 0
			Session("CouponDiscountPercent") = 0
			NewTotal = 0
		End If		
	END IF 
		
	'value based coupon	
	If sType = "Amount" then
		cpValue = rst("cpValue")
		If rst("cpType") = "Amount"  AND  (cdbl(cpvalue) <= cdbl(xTotalPrice)) then 
			NewTotal =  cdbl(xTotalPrice) - cdbl(cpValue)
			DiscountAmount =  oldamt - NewTotal
			Session("CouponDiscountAmount") = cdbl(DiscountAmount)
		else
			DiscountAmount =  0
			Session("CouponDiscountAmount") = 0
			NewTotal = 0
		End If
	End IF	
	
	If cdbl(NewTotal) < 0 then NewTotal = 0
	

	If DiscountAmount <= 0 then 
		DiscountAmount = 0
		NewTotal = xTotalPrice
		If sType ="Percent" then Session("CouponDiscountPercent") = 0
		If sType ="Amount" Then Session("CouponDiscountAmount") = 0
	End If
	
	If sReturn = "Total" then  ApplyCouponDiscount = NewTotal
	If sReturn = "Discount" then ApplyCouponDiscount = DiscountAmount
	'Response.Write "<BR> discount:" & discountamount
	'Response.Write "<BR> newtotal:" & newtotal
	
	CloseObj (rst)   
		
End Function



'-------------------------------------------------------------------------------
'Purpose: 
'Accepts: 
'Returns:
'Called From: 
'Comments:
'-------------------------------------------------------------------------------
Function CouponOn
Dim rst,sql
	
	sql = "Select * FROM sfCoupons WHERE cpActivate = 1"
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenStatic', adLockReadOnly, adCmdText	
	
	if rst.recordcount <= 0 then  
		CouponOn = false
	else
		CouponOn = true
	end if
	
	CloseObj (rst)
	   
End Function



'************************************************************************************
'**********************  INVENTORY TRACKING ************************************************
'************************************************************************************
	

'-------------------------------------------------------------------------------
'Purpose: checks all cart items availability
'Accepts: 
'Returns: 1 if all items in the cart are available, else returns 0
'Called From: 
'Comments:
'-------------------------------------------------------------------------------
Function CheckCartInventory
dim sql
dim rst
dim i
dim avlqty
	
	
	If session("SessionID") = "" then 
		CheckCartInventory = 1 
		exit function
	end if
	
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	sql = gtmpSQL & " WHERE odrdttmpSessionID=" & Session("SessionID")
	
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText
	if rst.recordcount > 0 then
	rst.movefirst
	For i = 1 to rst.recordcount
		'If rst("odrdttmpBackOrderQty") = 0 then 'skip back ordered items
		If rst("odrdttmpBackOrderQty") <= 0 then ' beta 2
			avlqty = GetAvailableqty(rst("odrdttmpproductid"),GetAttDetailID(rst("odrdttmpID"),"tmp"))
			If avlqty <> "X" AND avlqty < rst("odrdttmpQuantity") then
				CheckCartInventory = 0 'cart items need adjustment 
				CloseObj (rst)
				Exit Function
			End If
		End if	
		
		If not rst.eof then rst.movenext
	Next
		
	CheckCartInventory = 1 'cart items ok 	
	
	Else
	
	CheckCartInventory = 1
	
	End if
	
	CloseObj (rst)

End Function



Sub ValidateCartItems
dim sql
dim rst
dim i
	
	
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	sql = gtmpSQL & " WHERE odrdttmpSessionID=" & Session("SessionID")
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText
	
	If rst.recordcount > 0 then
		rst.movefirst
		For i = 1 to rst.recordcount
			'If rst("odrdttmpBackOrderFlag") <= 0 then 'beta 2
			AdjustTmpOrderDetails(rst("odrdttmpID"))
			'else
			'	AdjustTmpOrderDetailsBO(rst("odrdttmpID"))
			'End If
			If not rst.eof then rst.movenext
		Next
		
	End if
	
	CloseObj (rst)

End Sub



'-------------------------------------------------------------------------------
'Purpose: Get qty from tmporderdetails table for a product
'Accepts: 
'Returns:
'Called From: 
'Comments:
'-------------------------------------------------------------------------------
Function GetTMPQTY(iTmpOrderDetailID) 'incAE

Dim sql,rst
	
	
	sql = "Select * FROM sftmporderdetails  where odrdttmpID=" & iTmpOrderDetailID
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn,  adOpenKeySet, adLockOptimistic, adCmdText
	If rst.RecordCount <= 0 then 
		GetTMPQTY = "X"
		exit function
	end if
	
	GetTMPQTY = rst("odrdttmpQuantity")
	CloseObj (rst)
	
	
End Function


'-------------------------------------------------------------------------------
'Purpose: 
'Accepts: 
'Returns:
'Called From: 
'Comments:
'-------------------------------------------------------------------------------
Function GetTMPGiftWrapQTY(iTmpOrderDetailID) 'incAE
Dim sql,rst
	
	
	sql = "Select * FROM sftmporderdetailsAE  where odrdttmpaeID=" & iTmpOrderDetailID
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn,  adOpenKeySet, adLockOptimistic, adCmdText
	If rst.RecordCount <= 0 then 
		GetTMPGiftWrapQTY = "X"
		exit function
	end if
	
	GetTMPGiftWrapQTY = rst("odrdttmpGiftWrapQTY")
	CloseObj (rst)
	
	
End Function


Function GetTMPGiftWrapQTY_BO(iTmpOrderDetailID) 'incAE
	
Dim sql,rst
	
	
	sql = "Select * FROM sftmporderdetailsAE  where odrdttmpaeID=" & iTmpOrderDetailID
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn,  adOpenKeySet, adLockOptimistic, adCmdText
	If rst.RecordCount <= 0 then 
		GetTMPGiftWrapQTY_BO = "0"
		exit function
	end if
	
	IF rst("odrdttmpGiftWrapQTY") > rst("odrdttmpBackOrderQTY") then
		GetTMPGiftWrapQTY_BO =  rst("odrdttmpBackOrderQTY")
	else
		GetTMPGiftWrapQTY_BO =  rst("odrdttmpGiftWrapQTY")
	End if
	
	CloseObj (rst)
	
	
End Function



Function GetGiftWrapQTY(iTmpOrderDetailID,iType) 'incAE
	
Dim sql,rst,avlqty,gwqty,boqty

	sql = gtmpSQL & " where odrdttmpaeID=" & iTmpOrderDetailID
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn,  adOpenStatic, adLockreadOnly, adCmdText
	
	If rst.RecordCount <= 0 then 
		GetGiftWrapQTY = 0
		exit function
	end if
	
	avlqty= rst("odrdttmpQuantity")  - rst("odrdttmpBackOrderQTY")
	boqty = rst("odrdttmpBackOrderQTY")
	gwqty = rst("odrdttmpGiftWrapQTY")
	
	If iType = 1 then  'shipped gift wraps
		IF rst("odrdttmpGiftWrapQTY") > avlqty then
			GetGiftWrapQTY = avlqty 
		else
			GetGiftWrapQTY = gwqty
		End if
	End If 
	
	If iType = 1 then  'backordered gift wraps
		IF rst("odrdttmpGiftWrapQTY") => avlqty then
			GetGiftWrapQTY = 0
		else
			GetGiftWrapQTY = gwqty
		End if
	end if
	
	CloseObj (rst)
	
	
End Function



'-------------------------------------------------------------------------------
'Purpose: Checks to see if inventory is tracked for a product
'Accepts: product id
'Returns: 1 if inventory tracked, 0 if not tracked
'Called From: 
'Comments:
'-------------------------------------------------------------------------------
Function CheckInventoryTracked(strProductID)
Dim sql, rst
	
	sql = "Select * FROM sfInventoryInfo WHERE invenProdID= '" & strProductID & "'"
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText	
	
	If rst.recordcount <= 0 then 
		CheckInventoryTracked = 0
		exit function	
	End If
	
		
	If rst("invenbTracked") = 1 then
		CheckInventoryTracked = 1
	Else
		CheckInventoryTracked = 0 
	End If
			
	CloseObj (rst)
		
End Function


'-------------------------------------------------------------------------------
'Purpose: 
'Accepts: 
'Returns: 
'Called From: 
'Comments:
'-------------------------------------------------------------------------------
Function CheckShowStatus(strProductID)
Dim sql, rst
	
	sql = "Select * FROM sfInventoryInfo WHERE invenProdID= '" & strProductID & "'"
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText	
	
	If rst.recordcount <= 0 then 
		CheckShowStatus = 0
		exit function	
	End If
	
		
	If rst("invenbStatus") = 1 then
		CheckShowStatus = 1
	Else
		CheckShowStatus = 0
	End If
			
	CloseObj (rst)
		
End Function

'-------------------------------------------------------------------------------
'Purpose: 
'Accepts: 
'Returns: 
'Called From: 
'Comments:
'-------------------------------------------------------------------------------
Function CheckNotification(strProductID)
Dim sql, rst
	
	sql = "Select * FROM sfInventoryInfo WHERE invenProdID= '" & strProductID & "'"
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText	
	
	If rst.recordcount <= 0 then 
		CheckNotification = 0
		closeobj(rst)
		exit function	
	End If
	
		
	If rst("invenbNotify") = 1 then
		CheckNotification = 1
	Else
		CheckNotification = 0
	End If
			
	CloseObj (rst)
		
End Function



'-------------------------------------------------------------------------------
'Purpose: 
'Accepts: 
'Returns:
'Called From: 
'Comments:
'-------------------------------------------------------------------------------
Function GetAvailableQty(strProductID,AttIDs)
Dim sql, rst
dim ret
	
	
	ret = CheckInventoryTracked (strProductID)
	If ret =  0 then 
		GetAvailableQty = "X" 'No inventoryinfo record
		Exit function
	end if
	
	sql = "Select * FROM sfInventory WHERE invenProdID= '" & strProductID & "' AND  invenAttDetailID='" & AttIDs & "'"
	
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText	
	
	If rst.recordcount <= 0 then 
		GetAvailableQty = "X" 'Inventory record missing
		CloseObj (rst)
		exit function
	end if 
	
	GetAvailableQty = rst("invenInstock")
	
	CloseObj (rst)
	
End Function




'-------------------------------------------------------------------------------
'Purpose: 
'Accepts: 
'Returns:
'Called From: 
'Comments:
'-------------------------------------------------------------------------------
Sub UpdateAvailableQty(strProductID,AttIDs,subQTY)
Dim sql, rst
dim ret,avlqty
Dim bLow
Dim nProdID
Dim nSubject
Dim nBody
Dim nProdName
	ret = CheckInventoryTracked (strProductID)
	
	bLow = 0
 
    'Response.write "<BR> subqty:" & subqty
    'Response.write "<BR> prod-att:" & strProductID & "-" &  attids
    'Response.write "<BR> ret:" & ret
	
	If ret = 0 then Exit sub 'no inventory tracked so exit
	
	sql = "Select * FROM sfInventory WHERE invenProdID= '" & strProductID & "' AND  invenAttDetailID='" & AttIDs & "'"
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	'rst.CursorLocation = adUseClient
	rst.Open sql, cnn, adOpenKeyset, adLockOptimistic, adCmdText
	'rst.Open sql, cnn, adOpenDynamic, adLockOptimistic, adCmdText		

	if rst.recordcount <= 0 then
		closeobj(rst)
		exit sub
	end if


	avlqty = rst("invenInstock") 
	'Response.write "<BR> avlqty:" & avlqty
	'Response.write "<BR> rst(instock) after diff:" & rst("invenInstock")
	If (avlqty - subqty) < 0 then
		rst("invenInstock") = 0
		'rst("invenInstock") = (avlqty - subqty)
	else
		rst("invenInstock") = (avlqty - subqty)
	End if
		
	If rst("invenInstock") <= rst("invenLowFlag") then 
		bLow = 1
				
		nProdID = rst("InvenProdID")
		nProdName = GetProductName(nProdId)
		
		If  rst("invenInstock") <= 0 then
			nSubject = "Product Out of Stock!"
			nBody = vbcrlf & "Store: " & C_STORENAME
			nBody = nBody & vbcrlf & "Notification: Product Out of Stock!"  
			nBody = nBody & vbcrlf & "Product: " & nProdName
			nBody = nBody & vbcrlf & "Product Attributes: " & rst("invenAttName")
		
		Else 'qty is low 
			nSubject = "Product Stock Low!"
			nBody = vbcrlf & "Store: " & C_STORENAME
			nBody = nBody & vbcrlf & "Notification: Product Stock Low!"  
			nBody = nBody & vbcrlf & "Product: " & nProdName
			nBody = nBody & vbcrlf & "Product Attributes: " & rst("invenAttName")
			nBody = nBody & vbcrlf & "Quantity Remaining:" &  rst("invenInStock") 
		End If
		
	End if
'	DoEvents
	rst.update
	CloseObj (rst)
	
	
	'send notification if low
	If bLow =1 AND CheckNotification(nProdID)  = 1  then
		'sProdInfo = sProdInfo  & " " & GetProductName(strProductID)
		'Session("sNotification") = sProdInfo
		CreateMail "InvenNotification",nSubject & "|" & nBody 
	End if
	
	
	
	
End Sub



'-------------------------------------------------------------------------------
'Purpose: 
'Accepts: 
'Returns:
'Called From: 
'Comments:
'-------------------------------------------------------------------------------

Function GetProductName (strProductID)
dim sql
dim rst

	
	
	Set rst = Server.CreateObject("ADODB.RecordSet")		

	sql = "Select * FROM sfProducts WHERE ProdID= '" & strProductID & "'" 

	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText	

	If rst.recordcount >0 then
		GetProductName = rst("ProdName")
	else
		GetProductName= ""
	end If

	CloseObj (rst)

End function


	

'-------------------------------------------------------------------------------
'Purpose: Determines the general product available quantity (not attribute level)
'Accepts: 
'Returns:
'Called From: 
'Comments:
'-------------------------------------------------------------------------------
Function CheckInStock(strProductID)
Dim sql, rst

	sql = "Select Sum(invenInstock) as instock FROM sfInventory WHERE invenProdID= '" & strProductID & "'"
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText	
	If rst.recordcount <= 0 then 
		CheckInStock = "X" 'error
	else
		CheckInStock= rst("Instock")
		
	End if
	CloseObj (rst)
	
End Function


'-------------------------------------------------------------------------------
'Purpose: 
'Accepts: 
'Returns:
'Called From: 
'Comments:
'-------------------------------------------------------------------------------
Function CheckBackOrder(strProductID)
Dim sql, rst
	sql = "Select * FROM sfInventoryInfo WHERE invenProdID= '" & strProductID & "'"
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText	
	
	If rst.recordcount <= 0 then 
		CheckBackOrder = 0
		exit function
	End If
	
		
	IF rst("invenbBackOrder") <> 1 then 
		CheckBackOrder = 0
	Else
		CheckBackOrder = 1
	End If
	
	If rst("invenbTracked") <> 1 then 'if inventory not tracked then no backorder either
		CheckBackOrder = 0
	End If
				
	CloseObj(rst)
		
End Function
Sub sfReports1_ShowCouponDiscount

End Sub

Sub sfReports1_SQL1
	sSQL = "Select * FROM sfOrders as A LEFT JOIN sfOrdersAE as B ON A.orderID = B.orderaeID " & "WHERE orderID = " & sOrderId  & " and A.orderIsComplete = 1"
End Sub

Sub sfReports1_SQL2
	sSQL = "Select * FROM sfOrderDetails as A LEFT JOIN sfOrderDetailsAE as B ON A.odrdtID = B.odrdtaeID " & "WHERE odrdtOrderId = " & sOrderId
End Sub

Sub sfReports1_sql3
	sSQL = "Select * FROM sfOrderDetails as A LEFT JOIN sfOrderDetailsAE as B ON A.odrdtID = B.odrdtaeID " & "WHERE odrdtOrderId = " & rsOrderDetail.Fields("orderID")
End Sub

Sub sfReports1_ShowProductDetails

		Response.Write "<tr>" & vbcrlf
		Response.Write "<td></td>" & vbcrlf
		Response.Write "<td valign=""top"" align=""left"">Gift Wrap</td>" & vbcrlf
		Response.Write "<td valign=""top"" align=""right"">" & rsOrderProducts.Fields("odrdtGiftWrapQTY") & "</td>" & vbcrlf

	If rsOrderProducts.Fields("odrdtGiftWrapQTY") <> 0 Then 
			Response.Write "<td valign=""top"" align=""right"">" & FormatCurrency(rsOrderProducts.Fields("odrdtGiftWrapPrice")/ rsOrderProducts.Fields("odrdtGiftWrapQTY")) & "</td>" & vbcrlf
	else
   			Response.Write "<td valign=""top"" align=""right"">" & FormatCurrency(rsOrderProducts.Fields("odrdtGiftWrapPrice")) & "</td>" & vbcrlf
	end if
		Response.Write "<td valign=""top"" align=""right"">" & FormatCurrency(rsOrderProducts.Fields("odrdtGiftWrapPrice")) & "</td>" & vbcrlf
		Response.Write "</tr>" & vbcrlf
		Response.Write "<tr>" & vbcrlf
		Response.Write "<td> </td>" & vbcrlf
		Response.Write "<td valign=""top"" colspan = ""4"" align=""left"">BackOrdered Quantity: " & rsOrderProducts.Fields("odrdtBackOrderQTY") & "</td>" & vbcrlf
		Response.Write "</tr>" & vbcrlf

End Sub
Sub sfReports1_Coupon
On Error REsume Next
Dim sql, rst

   
	sql = "Select * FROM sfOrdersAE WHERE orderAEID= " & sOrderID
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText	
	If rst.recordcount > 0 then

		Response.Write "<tr>" & vbcrlf
		Response.Write "<td>Coupon Code:&nbsp;&nbsp;&nbsp;</td>" & vbcrlf
		Response.Write "<td align=""right"">" & rst("orderCouponCode") & "&nbsp;&nbsp;</td>" & vbcrlf
		Response.Write "</tr>" & vbcrlf
		Response.Write "<tr>" & vbcrlf
		Response.Write "<td>Coupon Discount:&nbsp;&nbsp;&nbsp;</td>" & vbcrlf
		Response.Write "<td align=""right"">- " & FormatCurrency(rst("orderCouponDiscount")) & "&nbsp;&nbsp;</td>" & vbcrlf
		Response.Write "</tr>" & vbcrlf
		

		iTempDisc = iTempDisc - cDbl(rst("orderCouponDiscount"))
	End If 
	
	rst.close
	set rst = nothing
End Sub

Sub sfReports1_Billing
Dim sql, rst


	sql = "Select * FROM sfOrdersAE WHERE orderAEID= " & sOrderID
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText	
	If rst.recordcount > 0 then
		

		Response.Write "<tr>" & vbcrlf
		Response.Write "<td><B>Billed Amount:&nbsp;&nbsp;&nbsp;</B></td>" & vbcrlf
		Response.Write "<td align=""right"">" & FormatCurrency(rst("orderBillAmount")) & "&nbsp;&nbsp;</td>" & vbcrlf
		Response.Write "</tr>" & vbcrlf
		Response.Write "<tr>" & vbcrlf
		Response.Write "<td><b>Remaining Amount:&nbsp;&nbsp;&nbsp;</B></td>" & vbcrlf
		Response.Write "<td align=""right"">" & FormatCurrency(rst("orderBackOrderAmount")) & "&nbsp;&nbsp;</td>" & vbcrlf
		Response.Write "</tr>" & vbcrlf
		

	End If 
	rst.close
	set rst = nothing
End Sub

Sub OVC_AddProductGiftWrapPrice
		dProductSubtotal = cdbl(dProductSubtotal) + cdbl(Session("gwprice"))
		sTotalPrice = cdbl(sTotalPrice) + cdbl(Session("gwprice"))
End Sub

'-------------------------------------------------------------------------------
'Purpose: 
'Accepts: 
'Returns:
'Called From: 
'Comments: new function for  .3008
'-------------------------------------------------------------------------------
Function ApplyALLDiscounts (xTotalPrice,sReturn)
Dim x,y,z,TotalDiscounts,NewTotal
	
	Session("discTotal") = 0
    Session("discAmount") = 0
    Session("discPercent") = 0
    Session("discStoreWide") = 0
    
	x=ApplyCouponDiscount(xTotalPrice,"Percent","Discount")
	y=ApplyCouponDiscount(xTotalPrice,"Amount","Discount")
	z=ApplyStoreWideDiscount(xTotalPrice,"Discount")
	
	TotalDiscounts = cdbl(x) + cdbl(y) + cdbl(z)
	NewTotal = cdbl(xTotalPrice) - cdbl(TotalDiscounts)
	
	If TotalDiscounts < 0 then TotalDiscounts = 0
	If NewTotal < 0 then NewTotal = 0 
	
	If sReturn = "Discount" then
		ApplyALLDiscounts = cdbl(TotalDiscounts)
	elseif sReturn = "Total" then
		ApplyALLDiscounts = cdbl(NewTotal)
	End If
	
    Session("discTotal") = cdbl(TotalDiscounts)
    Session("discAmount") = cdbl(y)
    Session("discPercent") = cdbl(x)
    Session("discStoreWide") = cdbl(z)
    
	'Response.Write  "<BR> Amount Discounts:" & Session("discAmount") & "<BR>"
End Function





%>





















