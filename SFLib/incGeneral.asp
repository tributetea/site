
<%
'@BEGINVERSIONINFO

'@APPVERSION: 50.3008.0.4

'@FILENAME: incgeneral.asp
	 

'

'@DESCRIPTION: multple functions used in the web application

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000,2001 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO
'Modified 10/24/01 
'Storefront Ref#'s: 168,163 'JF

'Modified 10/29/01 
'Storefront Ref#'s: 180 'djp

'Modified 10/31/01 
'Storefront Ref#'s: 194 'jf

'Modified 12/4/01 
'Storefront Ref#'s: 241 djp

'Modified 3/14/02
'Storefront Ref# 359 mls

'Modified 3/28/02
'Storefront Ref# 315

Dim rsAdminGen, C_STORENAME, C_HomePath, C_SecurePath,iConverion,sUserName,iEzeeHelp,sEzeeHelp,iSaveCartActive,iEmailActive,iBrandActive,sAffID,sLCID
Set rsAdminGen = Server.CreateObject("ADODB.Recordset")
rsAdminGen.Open "SELECT adminStoreName, adminDomainName, adminSSLPath, adminOandaID, adminActivateOanda,adminEzeeLogin,adminEzeeActive,adminSaveCartActive,adminEmailActive,adminSFActive,adminSFID,adminLCID FROM sfAdmin", cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
C_STORENAME  = trim(rsAdminGen.Fields("adminStoreName"))
C_HomePath   = trim(rsAdminGen.Fields("adminDomainName"))
C_SecurePath = trim(rsAdminGen.Fields("adminSSLPath"))
iConverion   = trim(rsAdminGen.Fields("adminActivateOanda"))
sUserName    = trim(rsAdminGen.Fields("adminOandaID"))
sEzeeHelp    = trim(rsAdminGen.Fields("adminEzeeLogin"))
iEzeeHelp    = trim(rsAdminGen.Fields("adminEzeeActive"))
iSaveCartActive = trim(rsAdminGen.Fields("adminSaveCartActive"))
iEmailActive = trim(rsAdminGen.Fields("adminEmailActive"))
iBrandActive = trim(rsAdminGen.Fields("adminSFActive"))
sAffID       = trim(rsAdminGen.Fields("adminSFID"))
sLCID = trim(rsAdminGen.Fields("adminLCID"))
closeObj(rsAdminGen)

If Session("LCID") <> "" Then
Session.LCID = Session("LCID")
Else
Session.LCID = sLCID
Session("LCID") = sLCID
End If

If Mid(C_HomePath, Len(C_HomePath), 1) <> "/" Then
    C_HomePath = C_HomePath & "/"
End If

	'Referal Varables
'dim REFERER,HTTP_REFERER,REMOTE_ADDRESS
'if trim(Request.QueryString("REFERER"))<>"" then
'	REFERER = Request.QueryString("REFERER")
'	Response.Cookies("sfHTTP_REFERER")("REFERER") = REFERER
'	end if
'	HTTP_REFERER = Request.ServerVariables("HTTP_REFERER")
'	REMOTE_ADDRESS = Request.ServerVariables("REMOTE_ADDR")
'	
'	Response.Cookies("sfHTTP_REFERER")("HTTP_REFERER") = HTTP_REFERER
'	Response.Cookies("sfHTTP_REFERER")("REMOTE_ADDRESS") = REMOTE_ADDRESS
'	Response.Cookies("sfHTTP_REFERER").Expires = Date() + 1
'--------------------------------------------------------
' MakeUSDate converts all date inputs to US date format
'--------------------------------------------------------
	
Function MakeUSDate(InDate)
	If Not IsDate(InDate) Then Exit Function
	MakeUSDate = Month(InDate)&"/"&Day(InDate)&"/"&Right(Year(InDate),2)
End Function

Function GetTotalOrderQTYSE
dim rst
dim sql
If Session("SessionID")<> "" Then 
	sql = "Select SUM(odrdttmpQuantity) as ordqty  FROM sfTmpOrderDetails "
	sql = SQL & " WHERE odrdttmpSessionId =" & Session("SessionID") ' 
	
	Set rst = Server.CreateObject("ADODB.RecordSet")
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly,adcmdText
	If rst.recordcount > 0 then
		GetTotalOrderQTYSE = clng(rst("ordqty"))
	else
		GetTotalOrderQTYSE = 0
	End IF
	closeobj (rst)
	set rst = Nothing
Else	
	GetTotalOrderQTYSE = 0
End if	
	
	

End Function

'----------------------------------------
' getShippingSaleText 
'----------------------------------------

Function getShippingSaleText(sShipping)
	Dim rsShippingAdmin, sText
	sText = ""
	Set rsShippingAdmin = Server.CreateObject("ADODB.RecordSet")
	rsShippingAdmin.Open "sfAdmin", cnn, adOpenForwardOnly, adLockReadOnly, adCmdTable
	If Trim(rsShippingAdmin.Fields("adminFreeShippingIsActive")) = "1" Then
		If sShipping = 0 Then
			sText = "Free Shipping on orders over  <b>" & FormatCurrency(rsShippingAdmin.Fields("adminFreeShippingAmount")) & "</b>!</font> "
		End If
	End If
	rsShippingAdmin.Close
	Set rsShippingAdmin = nothing
	getShippingSaleText = sText
End Function
'------------------------------------------------------------------
'These two functions handle Global Sales
'------------------------------------------------------------------
Function getGlobalSaleText()
	Dim rsGlobalAdmin, sGlobalActive, sText
	sText = ""
	Set rsGlobalAdmin = Server.CreateObject("ADODB.RecordSet")
	rsGlobalAdmin.open "sfAdmin", cnn, adOpenForwardOnly, adLockReadOnly, adCmdTable
	
	sGlobalActive = Trim(rsGlobalAdmin.Fields("adminGlobalSaleIsActive"))
	If sGlobalActive = "1" Then
				sText = "All items discounted <b>" & cDbl(rsGlobalAdmin.Fields("adminGlobalSaleAmt")) * 100 & "%</B>! "
	End If
	rsGlobalAdmin.Close
	Set rsGlobalAdmin = nothing
	getGlobalSaleText = sText
End Function
function getGlobalSalePrice(subtotal)
	Dim rsGlobalAdmin, sGlobalActive
	Set rsGlobalAdmin = Server.CreateObject("ADODB.RecordSet")
	rsGlobalAdmin.open "sfAdmin", cnn, adOpenForwardOnly, adLockReadOnly, adCmdTable
	
	sGlobalActive = Trim(rsGlobalAdmin.Fields("adminGlobalSaleIsActive"))
	If sGlobalActive = "1" Then
		getGlobalSalePrice = formatNumber(cDbl(subtotal)-(cDbl(subtotal)*cDbl(rsGlobalAdmin.Fields("adminGlobalSaleAmt"))), 2)
	Else
		getGlobalSalePrice = subTotal
	End If
	rsGlobalAdmin.Close
	Set rsGlobalAdmin = nothing
End Function

'---------------------------------------------------------------------
' Purpose: Deletes recordset from TmpOrders and associated child relations
'---------------------------------------------------------------------
Sub setDeleteOrder(sPrefix,iOrderDetailId)	
Dim rsDelete, sLocalSQL, rsDelete2, rsDelete3,rsDelete4,sLocalSQL2,sLocalSQL3,sLocalSQL4

	Select Case sPrefix
		Case "odrdttmp"
			sLocalSQL = "DELETE FROM sfTmpOrderDetails WHERE odrdttmpID = " & iOrderDetailId	
			sLocalSQL2 = "DELETE FROM sfTmpOrderAttributes WHERE odrattrtmpOrderDetailId = " & iOrderDetailId	
			
			If Application("AppName")= "StoreFrontAE" Then
				'Delete records from tmporderdetails ae for this session
				sLocalSQL3 = " Delete FROM sfTmpOrderDetailsAE WHERE odrdttmpAEID = " & iOrderDetailId & ""
				'Move Coupon
				sLocalSQL4 = "Delete FROM sfTmpOrdersAE WHERE odrtmpSessionID=" & Session("SessionID")
				Set rsDelete3 = cnn.Execute(sLocalSQL3)
				Set rsDelete4 = cnn.Execute(sLocalSQL4)
			
		  		closeObj(rsDelete3)
		  		closeObj(rsDelete4)
	  		end if
		Case "odrdtsvd"
			sLocalSQL = "DELETE FROM sfSavedOrderDetails WHERE odrdtsvdID = " & iOrderDetailId	
			sLocalSQL2 = "DELETE FROM sfSavedOrderAttributes WHERE odrattrsvdOrderDetailId = " & iOrderDetailId	
	End Select	
		If vDebug = 1 Then Response.Write "<br> DeleteTmp SQL : " & sLocalSQL & "<br>SQL2: " & sLocalSQL2	
		
		Set rsDelete2 = cnn.Execute(sLocalSQL2)
		Set rsDelete = cnn.Execute(sLocalSQL)
			
	  	closeObj(rsDelete)
	  	closeObj(rsDelete2)
End Sub 	

Function getTax(choice, sShipping, sTotalPrice, sProdID)

	Dim sState, sCountry, SQL, rsTax, iTax, rsAdmin, iTaxAmt, rsProd
	
	Set rsProd = Server.CreateObject("ADODB.Recordset")
	Set rsTax = Server.CreateObject("ADODB.RecordSet")
	Set rsAdmin = Server.CreateObject("ADODB.RecordSet")
	
	SQL = "SELECT prodCountryTaxIsActive, prodStateTaxIsActive FROM sfProducts WHERE prodID = '" & sProdID & "'"
	rsProd.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
	
	SQL = "SELECT adminTaxShipIsActive FROM sfAdmin"
	If vDebug = 1 Then Response.Write SQL & "<br><br>"
	rsAdmin.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
	Select Case choice

			Case "State"
			' Shipping State is Taxed
			If Request("ShipState") <> "" Then
				sState = Request("ShipState")
			ElseIf	Request("sShipCustState") <> "" Then
				sState = Request("sShipCustState")
			ElseIf sShipCustState <> "" Then
				sState = sShipCustState
			ElseIf Request("State") <> "" Then
				sState = Request("State")
			ElseIf sCustState <> "" Then
				sState = sCustState	
			End If

			SQL = "SELECT loclstTax FROM sfLocalesState WHERE loclstAbbreviation = '" & sState & "' AND loclstLocaleIsActive = 1 AND loclstTaxIsActive = 1"
			If vDebug = 1 Then Response.Write SQL & "<br><br>"
			rsTax.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
			If Not rsTax.EOF Then
				If trim(rsProd.Fields("prodStateTaxIsActive")) = "1" Then
					If rsAdmin.Fields("adminTaxShipIsActive") = 1 Then
						iTaxAmt = CDbl(sShipping) + CDbl(sTotalPrice)
					Else
						iTaxAmt = CDbl(sTotalPrice)
					End If 
				End If
			'	iTax = iTaxAmt * CDbl(rsTax.Fields("loclstTax"))
			'#353
			   	iTax = Formatnumber(iTaxAmt * CDbl(rsTax.Fields("loclstTax")),2)
			
			Else
				iTax = 0
			End If

			Case "Country"
'			if Request("Country") <> "" Then
'				sCountry = Request("Country")
'			ElseIf sCustCountry <> "" Then
'				sCountry = sCustCountry
'			ElseIf Request("ShipCountry") <> "" Then
			
			If Request("ShipCountry") <> "" Then
				sCountry = Request("ShipCountry")
			ElseIf Request("sShipCustCountry") <> "" Then
				sCountry = Request("sShipCustCountry")
'			ElseIf sShipCustCountry <> "" Then
'				sCountry = sShipCustCountry
			Elseif Request("Country") <> "" Then
				sCountry = Request("Country")
			ElseIf sCustCountry <> "" Then
				sCountry = sCustCountry
			End if
	

			SQL = "SELECT loclctryTax FROM sfLocalesCountry WHERE loclctryAbbreviation = '" & sCountry & "' AND loclctryLocalIsActive = 1 AND loclctryTaxIsActive = 1"
			If vDebug = 1 Then Response.Write SQL & "<br><br>"
			rsTax.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
			If Not rsTax.EOF Then
				If rsProd.Fields("prodCountryTaxIsActive") = 1 Then
					If rsAdmin.Fields("adminTaxShipIsActive") = 1 Then
						iTaxAmt = CDbl(sShipping) + CDbl(sTotalPrice)
					Else
						iTaxAmt = CDbl(sTotalPrice)
					End If 
					iTax = iTaxAmt * CDbl(rsTax.Fields("loclctryTax"))
				End If
			Else
				iTax = 0
			End If
		End Select
	closeObj(rsAdmin)
	closeObj(rsTax)
	getTax = formatNumber(iTax, 2)
End Function
'---------------------------------------------------------------------
' Collect Attribute IDs
'---------------------------------------------------------------------
Function getProdAttr(sPrefix,sOrderID,iProdAttrNum)
Dim sLocalSQL, rsAttrID, iCounter, aLocalArray

Select Case sPrefix
	Case "odrattrtmp"
		sLocalSQL = "SELECT odrattrtmpAttrID FROM sfTmpOrderAttributes WHERE odrattrtmpOrderDetailId = " & sOrderID
	Case "odrattrsvd"	
		sLocalSQL = "SELECT odrattrsvdAttrID FROM sfSavedOrderAttributes WHERE odrattrsvdOrderDetailId = " & sOrderID
	Case "odr"
		sLocalSQL = "SELECT odrattrID FROM sfOrderAttributes WHERE odrattrOrderDetailId = " & sOrderID 
End Select 
	
	Set rsAttrID = Server.CreateObject("ADODB.RecordSet")
  	rsAttrID.Open sLocalSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText 

	  If vDebug = 1 Then Response.Write "<p>getProdAttr SQL : " & sLocalSQL
	  
  	' Check if this record exists through prodID and price matches
      If rsAttrID.EOF or rsAttrID.BOF Then  		 
  		 If vDebug = 1 Then Response.Write "<p>Empty Recordset in rsAttrID"  		 
  	  Else
  		  Redim aLocalArray(iProdAttrNum)	
  		  For iCounter = 0 to iProdAttrNum - 1
  			aLocalArray(iCounter) = rsAttrID.Fields(sPrefix & "AttrID")
  			If vDebug = 1 Then Response.Write "<br>AttrID: " & aLocalArray(iCounter)
  		  rsAttrID.MoveNext
  		  Next
   	  ' End RecordSet If
  	  End If  	  
  	  
  	  closeObj(rsAttrID)  	 
  	  getProdAttr = aLocalArray
End Function
'---------------------------------------------------------------------
' This function checks whether a product exists and retrieves an array of info
'---------------------------------------------------------------------
Function getProduct(sProdID)
Dim sLocalSQL, aLocalProdArray(3), rsSelectProd

	sLocalSQL = "SELECT prodName, prodNamePlural, prodPrice, prodAttrNum,prodSaleIsActive,prodSalePrice FROM sfProducts WHERE prodEnabledIsActive=1 AND prodID = '"& sProdID & "'"

  	  Set rsSelectProd = Server.CreateObject("ADODB.RecordSet")
  	  rsSelectProd.Open sLocalSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText 

	  If vDebug = 1 Then Response.Write "<p>getProdValues SQL : " & sLocalSQL
	  
  	' Check if this record exists through prodID and price matches
      If rsSelectProd.EOF or rsSelectProd.BOF Then  		 
  		 If vDebug = 1 Then Response.Write "<p>Empty Recordset in rsSelectProd. Product " & sProdID & " possibly not activated."  		 
  	  Else
		  aLocalProdArray(0) = rsSelectProd.Fields("prodName")		  
			' Check if sale price is active 
			If rsSelectProd.Fields("prodSaleIsActive") = 1 Then
 					aLocalProdArray(1) = rsSelectProd.Fields("prodSalePrice")
 			Else 	
 					aLocalProdArray(1) = rsSelectProd.Fields("prodPrice")	
 			End If		
  		  aLocalProdArray(2) = rsSelectProd.Fields("prodAttrNum")

   	  ' End RecordSet If
  	  End If
  	  closeObj(rsSelectProd)

  	  getProduct = aLocalProdArray
End Function

'-------------------------------------------------------
' Update saved cart customers' info in sfCustomers
'-------------------------------------------------------
Sub setUpdateCustomer(sNewEmail,sFirstName,sMiddleInitial,sLastName,sCompany,sAddress1,sAddress2,sCity,sState,sZip,sCountry,sPhone,sFax,bSubscribed)
	Dim	sLocalSQl, rsUpdate, iOldNum
	
	sLocalSQL = "Select custFirstName, custMiddleInitial, custLastName, custCompany, custAddr1, custAddr2, custCity, custState, custZip, custCountry, "_
				& "custPhone, custFax, custTimesAccessed, custLastAccess, custEmail, custIsSubscribed FROM sfCustomers WHERE custID = " & Trim(Request.Cookies("sfCustomer")("custID"))
	
	Set rsUpdate = SErver.CreateObject("ADODB.RecordSet")
		rsUpdate.Open sLocalSQL,cnn,adOpenDynamic,adLockOptimistic,adCmdText
		
		If Not rsUpdate.EOF Then
				iOldNum = (rsUpdate.Fields("custTimesAccessed"))
				If iOldNum = "" or isnull(iOldNum) Then 
					iOldNum = 1 
				Else
					iOldNum = cInt(iOldNum)	
				End If		
				rsUpdate.Fields("custFirstName")		= sFirstName
				rsUpdate.Fields("custMiddleInitial")	= sMiddleInitial
				rsUpdate.Fields("custLastName")		= sLastName
				rsUpdate.Fields("custCompany")			= sCompany
				rsUpdate.Fields("custAddr1")			= sAddress1
				rsUpdate.Fields("custAddr2")			= sAddress2
				rsUpdate.Fields("custCity")				= sCity
				rsUpdate.Fields("custState")			= sState
				rsUpdate.Fields("custZip")				= sZip
				rsUpdate.Fields("custCountry")			= sCountry
				rsUpdate.Fields("custPhone")			= sPhone
				rsUpdate.Fields("custFax")				= sFax	
				rsUpdate.Fields("custTimesAccessed")	= iOldNum + 1
				rsUpdate.Fields("custLastAccess")		= Date()
				If sNewEmail <> "" Then
					rsUpdate.Fields("custEmail")		= sNewEmail
				End If		
				If CStr(bSubscribed) = "" Or CStr(bSubscribed) = "0" Then
                             rsUpdate.Fields("custissubscribed") = 0
                Else
                             rsUpdate.Fields("custissubscribed") = 1
                End If
				rsUpdate.Update		
		End If
		closeObj(rsUpdate)	
End Sub

'---------------------------------------------------------------------
' This function returns one specific value associated with a single id
' Used for lookup of VendorID, ManufacturerID, CategoryID, etc
'---------------------------------------------------------------------
Function getNameWithID(sLocalTableName,sLocalFindKey,sLocalFindKeyLabel,sLocalSearchName,bStringOrNot)
	Dim sLocalSQL, rsGetNameFromID, sLocalGetResult
if trim(sLocalFindKey) <> "" then
	    ' build SQL string based on whether the key is a string or not
		If (bStringOrNot = 0) Then
			sLocalSQL = "SELECT " & sLocalSearchName & " FROM " & sLocalTableName & " WHERE " & sLocalFindKeyLabel & "= " & Trim(sLocalFindKey) 
		ElseIf (bstringOrNot = 1) Then
			sLocalSQL = "SELECT " & sLocalSearchName & " FROM " & sLocalTableName & " WHERE " & sLocalFindKeyLabel & "= '" & Trim(sLocalFindKey) & "'"
		Else
		    Response.Write("The boolean parameter is not valid. Please input either 1 for true or 0 for false")
		    Exit Function
  		End If

	If vDebug = 1 Then Response.Write "<br>" & sLocalSQL
		
  	Set rsGetNameFromID = Server.CreateObject("ADODB.RecordSet")
  	rsGetNameFromID.Open sLocalSQL, cnn

  		If rsGetNameFromID.EOF Or rsGetNameFromID.BOF Then
  			'If vDebug = 1 Then Response.Write "Either the recordset doesn't exit or the field name is not typed correctly :<Br>" & sLocalSQL			
  		Else
  		  sLocalGetResult = rsGetNameFromID.Fields("" &sLocalSearchName& "")
  		End If
  	closeObj(rsGetNameFromID)

    getNameWithID = sLocalGetResult
 else  
   getNameWithID = "" 
 end if
End Function
'---------------------------------------------------------------------
' Enters record svdOrders, returns the ID of the SvdOrder
'---------------------------------------------------------------------
Function getSavedTable(aProdAttr,sProdID,iNewQuantity,iCustID,sReferer)
Dim rsCopy, sLocalSQL, rsSvdCart, rsSvdCartAttr, iKeyID, sDateTime,sTmpAttrName, sTmpAttrID, aTmpOrderArray, bookMark

	' Write to svd cart			
	Set rsSvdCart = Server.CreateObject("ADODB.RecordSet")
		rsSvdCart.CursorLocation = adUseClient
		rsSvdCart.Open "sfSavedOrderDetails Order By odrdtsvdID", cnn, adOpenDynamic, adLockOptimistic
			rsSvdCart.AddNew
			rsSvdCart.Fields("odrdtsvdCustID") = iCustID
			rsSvdCart.Fields("odrdtsvdQuantity") = iNewQuantity
			rsSvdCart.Fields("odrdtsvdProductID") = sProdID
			rsSvdCart.Fields("odrdtsvdDate") = Now
			rsSvdCart.Fields("odrdtsvdSessionID") = Session("SessionID")
			rsSvdCart.Fields("odrdtsvdHttpReferer") = left(sReferer,255)
			rsSvdCart.Update
			
			'bookMark = rsSvdCart.AbsolutePosition 
			'rsSvdCart.Requery 
			'rsSvdCart.AbsolutePosition = bookMark
			
			iKeyID  = rsSvdCart.Fields("odrdtsvdID")
	
		If vDebug = 1 Then Response.Write "<p><font size=4><b>SvdCart Key ID = " & iKeyID & "</b></font>"
		' Copy Attributes
		iCounter = 0
			
		' Collect Attribute Info from sfTmpOrderAttributes
		If IsArray(aProdAttr) Then
			Do While NOT aProdAttr(iCounter) = ""
				sTmpAttrID = aProdAttr(iCounter)
				If vDebug = 1 Then Response.Write "<p> sTmpAttrID = " & sTmpAttrID	
					Set rsSvdCartAttr = Server.CreateObject("ADODB.RecordSet")
						rsSvdCartAttr.Open "sfSavedOrderAttributes", cnn, adOpenDynamic, adLockOptimistic
						rsSvdCartAttr.AddNew
						rsSvdCartAttr.Fields("odrattrsvdOrderDetailId") = iKeyID
						rsSvdCartAttr.Fields("odrattrsvdAttrID") = sTmpAttrID
					rsSvdCartAttr.Update							
				iCounter = iCounter + 1			
			Loop
		' End IsArray If 	
		End If
		
		If vDebug = 1 Then 	Response.Write "<p><font color=""red"" face=""verdana"" size=""2"">Copied Record To SavedOrder</font>"			
	
  	closeObj(rsCopy)
  	closeObj(rsSvdCart)
  	closeobj(rsSvdCartAttr)
  	getSavedTable = iKeyId
End Function
'---------------------------------------------------------------------
' Enters record TmpOrders, returns the ID of the TmpOrder
'---------------------------------------------------------------------
Function getTmpTable(aProdAttr,sProdID,iNewQuantity,sReferer,iShip)
Dim  sLocalSQL, rsTmpCart, rsTmpCartAttr, iKeyID, sTmpAttrName, sTmpAttrID, aTmpOrderArray, bookMark

	' Write to tmp cart			
	Set rsTmpCart = Server.CreateObject("ADODB.RecordSet")
		rsTmpCart.CursorLocation = adUseClient
		rsTmpCart.Open "sfTmpOrderDetails Order By odrdttmpID", cnn, adOpenDynamic, adLockOptimistic, adCmdTable
			rsTmpCart.AddNew
			rsTmpCart.Fields("odrdttmpQuantity")	= iNewQuantity
			rsTmpCart.Fields("odrdttmpProductID")	= sProdID
			rsTmpCart.Fields("odrdttmpSessionID")	= Session("SessionID")
			If sReferer <> "" and NOT isNull(sReferer) Then
				rsTmpCart.Fields("odrdttmpHttpReferer") = left(sReferer,255)
			End If	
			rsTmpCart.Fields("odrdttmpShipping") = iShip
			rsTmpCart.Update
			'bookMark = rsTmpCart.AbsolutePosition 
			'rsTmpCart.Requery 
			'rsTmpCart.AbsolutePosition = bookMark			
			iKeyID  = rsTmpCart.Fields("odrdttmpID")
	
		If vDebug = 1 Then Response.Write "<p><font size=4><b>TmpCart Key ID = " & iKeyID & "</b></font>"
		' Copy Attributes
		iCounter = 0
			
		' Collect Attribute Info from sfTmpOrderAttributes
		If IsArray(aProdAttr) Then
			Do While NOT aProdAttr(iCounter) = ""
				sTmpAttrID = aProdAttr(iCounter)
				If vDebug = 1 Then Response.Write "<p> <b>sTmpAttrID</b> = " & sTmpAttrID	
					Set rsTmpCartAttr = Server.CreateObject("ADODB.RecordSet")
						rsTmpCartAttr.Open "sfTmpOrderAttributes", cnn, adOpenDynamic, adLockOptimistic, adCmdTable
						rsTmpCartAttr.AddNew
						rsTmpCartAttr.Fields("odrattrtmpOrderDetailId") = iKeyID
						rsTmpCartAttr.Fields("odrattrtmpAttrID") = sTmpAttrID
					rsTmpCartAttr.Update							
				iCounter = iCounter + 1			
			Loop
		' End IsArray If 	
		End If
		
		If vDebug = 1 Then 	Response.Write "<p><font color=""red"" face=""verdana"" size=""2"">Copied Record To TmpOrder</font>"			
	
  	closeObj(rsTmpCart)
  	closeobj(rsTmpCartAttr)
  	getTmpTable = iKeyId
End Function
'---------------------------------------------------------------------
' Purpose: Updates the Quantity field with associated prodId and CartID
'---------------------------------------------------------------------
Sub setUpdateQuantity(sPrefix,iQuantity,iTmpOrderID)
Dim rsUpdate, sLocalSQL, iOldQuantity, iNewQuantity, rsGetQuantity

	Select Case sPrefix
		Case "odrdttmp"		
				sLocalSQL = "SELECT odrdttmpQuantity FROM sfTmpOrderDetails WHERE odrdttmpID=" &iTmpOrderID & " AND odrdttmpSessionID=" & Session("SessionID")
				If vDebug = 1 Then Response.Write "<br> setUpdateQuantity SQL : " & sLocalSQL

		Case "odrdtsvd"
				sLocalSQL = "SELECT odrdtsvdQuantity FROM sfSavedOrderDetails WHERE odrdtsvdID=" & iTmpOrderID & " AND odrdtsvdCustID=" & Request.Cookies("sfCustomer")("custID") 
	End Select


	Set rsGetQuantity = Server.CreateObject("ADODB.RecordSet")
	rsGetQuantity.Open sLocalSQL, cnn
	If rsGetQuantity.EOF And rsGetQuantity.BOF Then Response.Redirect "abandon.asp"
	' Get Old Quantity
	iOldQuantity = rsGetQuantity.Fields(sPrefix & "Quantity")
	rsGetQuantity.Close

	iNewQuantity = cInt(iOldQuantity) + cInt(iQuantity)

	' Now Update
	Set rsUpdate = Server.CreateObject("ADODB.RecordSet")
		rsUpdate.Open sLocalSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText
	
		rsUpdate.Fields(sPrefix & "Quantity") = iNewQuantity
  		rsUpdate.Update

  	closeObj(rsGetQuantity)
  	closeObj(rsUpdate)
End Sub

'---------------------------------------------------------------------
' Purpose: Updates the Quantity field with associated prodId and CartID
'---------------------------------------------------------------------
Sub setReplaceQuantity(sPrefix,iQuantity,iTmpOrderID)
Dim rsUpdate, sLocalSQL

	Select Case sPrefix
		Case "odrdttmp"		
				sLocalSQL = "SELECT odrdttmpQuantity FROM sfTmpOrderDetails WHERE odrdttmpID=" &iTmpOrderID & " AND odrdttmpSessionID=" & Session("SessionID")
				If vDebug = 1 Then Response.Write "<br> setUpdateQuantity SQL : " & sLocalSQL

		Case "odrdtsvd"
				sLocalSQL = "SELECT odrdtsvdQuantity FROM sfSavedOrderDetails WHERE odrdtsvdID=" & iTmpOrderID & " AND odrdtsvdCustID=" & Request.Cookies("sfCustomer")("custID") 
				If vDebug = 1 Then Response.Write "<br> setSvdUpdateQuantity SQL : " & sLocalSQL
	End Select

	' Now Update
	Set rsUpdate = Server.CreateObject("ADODB.RecordSet")
	rsUpdate.Open sLocalSQL, cnn, adOpenDynamic, adLockOptimistic
	If rsUpdate.EOF And rsUpdate.BOF Then Response.Redirect "abandon.asp"
	rsUpdate.Fields(sPrefix & "Quantity") = iQuantity
  	rsUpdate.Update
  		If rsUpdate.EOF Or rsUpdate.BOF Then
			Response.Write "<br>Empty Recordset in rsUpdate"
		Else
			If vDebug = 1 Then Response.Write "<p><font color=""red"" face=""verdana"" size=""2"">Successful update of Quantity to: " & iQuantity & "</font>"
		End If

  	closeObj(rsUpdate)
End Sub
 

'---------------------------------------------------------------------
' Checks for existence of same product and attributes (if any)
' Returns the OrderDetail ID or -1 if record DNE
'---------------------------------------------------------------------
Function getOrderID(sPrefix,sAttrPrefix,sProdID,aProdAttr,iProdAttrNum)
Dim sTmpVar, bHasAttributes, iLocalResult, rsSelectProd, sTmpPrefixID, sTmpAttrName, sTmpAttr
Dim sLocal, sSQL, sLocalSQL, sAttrName, bMatch, iUpperBound
	iLocalResult = -1
	bHasAttributes = (iProdAttrNum > 0)
	bMatch = 0

	' SQL select
	Select Case sPrefix
		Case "odrdttmp"	
			If bHasAttributes Then
				sLocalSQL = "SELECT odrdttmpID, odrattrtmpAttrID FROM sfTmpOrderAttributes INNER JOIN sfTmpOrderDetails ON sfTmpOrderAttributes.odrattrtmpOrderDetailId = sfTmpOrderDetails.odrdttmpID" _
					& " WHERE odrdttmpSessionID = " & Session("SessionID") & " AND odrdttmpProductID = '" & sProdID & "'"
			Else
				sLocalSQL = "SELECT odrdttmpID FROM sfTmpOrderDetails WHERE odrdttmpSessionID = " & Session("SessionID") & " AND odrdttmpProductID = '" & sProdID & "'"
			End If
		Case "odrdtsvd" 
			If bHasAttributes Then
				sLocalSQL = "SELECT odrdtsvdID, odrattrsvdAttrID FROM sfSavedOrderDetails INNER JOIN sfSavedOrderAttributes ON sfSavedOrderDetails.odrdtsvdID = sfSavedOrderAttributes.odrattrsvdOrderDetailId " _
					& " WHERE odrdtsvdCustID=" & Request.Cookies("sfCustomer")("custID") & " AND odrdtsvdProductID = '" & sProdID & "'"
			Else
				sLocalSQL = "SELECT odrdtsvdID FROM sfSavedOrderDetails WHERE odrdtsvdCustID=" & Request.Cookies("sfCustomer")("custID") & " AND odrdtsvdProductID = '" & sProdID & "'"
			End If	
	End Select	

	Set rsSelectProd = Server.CreateObject("ADODB.RecordSet")
  	    rsSelectProd.Open sLocalSQL, cnn, adOpenKeyset, adLockOptimistic, adCmdText

  	' Check if this record exists through prodID and price matches
      If (rsSelectProd.BOF And rsSelectProd.EOF) Then
		  'No Records Matching, return -1
		  iLocalResult = -1
	  Else
 		  If bHasAttributes Then
 		  	  	' -- Debug Use -- Look what has been collected 
  				If vDebug = 1 Then 
  					Do While Not rsSelectProd.EOF
						Response.Write "<p>ID : " & rsSelectProd.Fields(sPrefix & "Id") & " AttrID :" & rsSelectProd.Fields(sAttrPrefix & "AttrID") 
						rsSelectProd.MoveNext
					Loop
				End If
				rsSelectProd.MoveFirst
 			
 				iUpperBound = UBound(aProdAttr)
 			
					' Check that there are at least as many product attributes as there are rows
					If rsSelectProd.RecordCount < cInt(iProdAttrNum) Then
						getOrderID = -1
					Else
					' Start comparison of product attributes
 							Do While Not rsSelectProd.EOF
									For iCounter = 0 to iUpperBound-1 					
  									sTmpAttr = aProdAttr(iCounter)   
  								
  									' If sTmpAttr is empty, the attribute specified is no longer available in the db											
  										If sTmpAttr = "" or rsSelectProd.EOF Then 
  											getOrderID = ""
 											Exit Function
 										Else  								
											If cStr(sTmpAttr) = cStr(rsSelectProd.Fields(sAttrPrefix & "AttrID")) Then
													bMatch = bMatch + 1
											End If		
											
											If vDebug = 1 Then Response.Write "<p>" & sTmpAttr & " VS " & rsSelectProd.Fields(sAttrPrefix & "AttrID")  						
											If vDebug = 1 Then Response.Write "<br>bMatch = " & bMatch 			
			
											If bMatch = cInt(iProdAttrNum) Then
												' Return the Found Record
												getOrderID = rsSelectProd.Fields(sPrefix & "ID")
												Exit Function					
	  										End If					
	  									' End sTmpAttr Empty If
	  									End If
									rsSelectProd.MoveNext
									Next

							' Reset Match at end of Recordset
							  bMatch = 0  		
  				
  							' Loop through recordset
							Loop

				' End iProdAttrNum if		
				End If		
				
  			' Matched Product with No attributes
  			Else 
  				getOrderID = rsSelectProd.Fields(sPrefix & "ID")			
    			Exit Function
  			' End Has Attributes If
  			End If

 	  ' End RecordSet If
  	  End If

  	  closeObj(rsSelectProd)
  	  getOrderID = -1
End Function
'---------------------------------------------------------------------
' Returns the name, price, and type associated with the attribute ID
'---------------------------------------------------------------------
Function getAttrDetails(iAttrID)
Dim sLocalSQL, rsFindAttr, aLocalAttr

	sLocalSQL = "SELECT attrName, attrdtName, attrdtPrice, attrdtType FROM sfAttributeDetail INNER JOIN sfAttributes ON sfAttributes.attrID = sfAttributeDetail.attrdtAttributeId WHERE attrdtID = " & iAttrID
		Set rsFindAttr = Server.CreateObject("ADODB.RecordSet")
			rsFindAttr.Open sLocalSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText
			If rsFindAttr.BOF Or rsFindAttr.EOF Then
				If vDebug = 1 Then Response.Write "<br>Empty Recordset in getAttrNames"
			Else
				Redim aLocalAttr(4)
				aLocalAttr(0) = rsFindAttr.Fields("attrdtName")
				aLocalAttr(1) = rsFindAttr.Fields("attrdtPrice")
				aLocalAttr(2) = rsFindAttr.Fields("attrdtType")
				aLocalAttr(3) = rsFindAttr.Fields("attrName")
			End If
		
	closeObj(rsFindAttr)
	getAttrDetails = aLocalAttr
End Function
'---------------------------------------------------------------------
' Returns the name, price, and type associated with the attribute ID of Old Order
'---------------------------------------------------------------------
Function getAttrDetailsRetriveOrder(iAttrID)
Dim sLocalSQL, rsFindAttr, aLocalAttr

	sLocalSQL = "SELECT odrattrAttribute, odrattrName, odrattrPrice, odrattrType FROM sfOrderAttributes WHERE odrattrID = " & iAttrID
		Set rsFindAttr = Server.CreateObject("ADODB.RecordSet")
			rsFindAttr.Open sLocalSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText
			If rsFindAttr.BOF Or rsFindAttr.EOF Then
				If vDebug = 1 Then Response.Write "<br>Empty Recordset in getAttrNames"
			Else
				Redim aLocalAttr(4)
				aLocalAttr(0) = rsFindAttr.Fields("odrattrAttribute")
				aLocalAttr(1) = rsFindAttr.Fields("odrattrPrice")
				aLocalAttr(2) = rsFindAttr.Fields("odrattrType")
				aLocalAttr(3) = rsFindAttr.Fields("odrattrName")
			End If
		
	closeObj(rsFindAttr)
	getAttrDetailsRetriveOrder = aLocalAttr
End Function

'---------------------------------------------------------------------
' This function calculates the subtotal for attributes
'---------------------------------------------------------------------
Function getAttrUnitPrice (dAttrTotal,sAttrPrice,iAttrType)
	' Recalculate Price
	If iAttrType = 1 Then
		dAttrTotal = dAttrTotal + cDbl(sAttrPrice)
	ElseIf iAttrType = 2  Then
		dAttrTotal = dAttrTotal + cDbl(sAttrPrice)*(-1)	
	End If
getAttrUnitPrice = dAttrTotal
End Function
'-------------------------------------------------------------------
' Returns the recordset corresponding to a custId identifier
'-------------------------------------------------------------------
Function getRow(sTableName,sIdName,iID,cnn)
	Dim sLocalSQL, rsSet
		
	sLocalSQL = "SELECT * FROM " & sTableName & " WHERE " & sIdName & " = " & iID
		
	' Object Creation
	Set rsSet = Server.CreateObject("ADODB.RecordSet")
	rsSet.Open sLocalSQL, cnn, adOpenForwardOnly, adLockOptimistic, adCmdText
				
	Set getRow = rsSet
End Function
'-------------------------------------------------------------------
' Gets records for tables with multiple records for one customer ID
' Returns the recordset
'-------------------------------------------------------------------
Function getRowActive(sTableName,sIdName,sActiveName,iID,cnn)
	Dim sLocalSQL, rsSet
		
	sLocalSQL = "SELECT * FROM " & sTableName & " WHERE " & sIdName & " = " & iID & " AND " & sActiveName  & " = 1"
		
	' Object Creation
	Set rsSet = Server.CreateObject("ADODB.RecordSet")
	rsSet.Open sLocalSQL, cnn, adOpenForwardOnly, adLockOptimistic, adCmdText
				
	Set getRowActive = rsSet
End Function

'--------------------------------------------------------------------
' Function : getCreditCardList 
' This returns the credit list in HTML format for dropdown box.
'--------------------------------------------------------------------	
Function getCreditCardList()	
	Dim rsCCList, sLocalSQL, sCCList, iCounter
	
	sLocalSQL = "Select transID, transName From sfTransactionTypes WHERE transType = 'Credit Card' AND transIsActive = 1"
	
	Set rsCCList = Server.CreateObject("ADODB.RecordSet")
	rsCCList.Open sLocalSQL, cnn, adOpenKeyset, adLockOptimistic, adCmdText
	
	sCCList = ""
	For iCounter = 1 to rsCCList.RecordCount
		sCCList = sCCList & "<option value=""" & Trim(rsCCList.Fields("transID")) &""">" & Trim(rsCCList.Fields("transName")) & "</option>"
		rsCCList.MoveNext
	Next	
	
	getCreditCardList = sCCList
	closeObj(rsCCList)
End Function


'-------------------------------------------------------
' Compares email and password, then returns the ID of the customer
' Returns -1 for failed authentication
'-------------------------------------------------------
Function customerAuth(sEmail,sPassword,sType)
	Dim sLocalSQL, iCustID, rsGetID
	Select Case sType
		Case "strict"
			sLocalSQL = "SELECT custID FROM sfCustomers WHERE custEmail = '" & sEmail & "' AND custPasswd = '" & sPassword & "' AND custID = " & Session("custID") 
		Case "loose"
			sLocalSQL = "SELECT custID FROM sfCustomers WHERE custEmail = '" & sEmail & "' AND custPasswd = '" & sPassword & "'"
		Case "loosest"
			sLocalSQL = "SELECT custID FROM sfCustomers WHERE custEmail = '" & sEmail & "'" 
		Case else	
			sLocalSQL = "SELECT custID FROM sfCustomers WHERE custEmail = '" & sEmail & "'"
	End Select
			
		If sEmail = "" Or sPassword = "" Then		
			iCustID = -1		
		Else	
			Set rsGetID = Server.CreateObject("ADODB.RecordSet")
			rsGetID.Open sLocalSQL,cnn,adOpenForwardOnly,adLockReadOnly,adCmdText				 
	
			If rsGetID.BOF Or rsGetID.EOF Or sEmail = "" Or sPassword = "" Then
				iCustID = -1
			Else	
				iCustID = rsGetID.Fields("custID")
			End If
		End If		
	
	customerAuth = iCustID
	closeobj(rsGetID)		
End Function

'------------------------------------------------------------------
' Gets the InternetCash Merchant ID
'------------------------------------------------------------------
Function getICashMercID()
	Dim sLocalSQL, rsICash, iID
	
	sLocalSQL = "SELECT trnsmthdLogin FROM sfTransactionMethods WHERE trnsmthdName = 'InternetCash'"
	Set rsICash = Server.CreateObject("ADODB.RecordSet")
	rsICash.Open sLocalSQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
	
	If rsICash.EOF or rsICash.BOF Then
		Response.Write "Error: No merchant ID set for Internet Cash in table sfTransactionMethods"
	Else
		iID = trim(rsICash.Fields("trnsmthdLogin"))
	End If
	
	closeobj(rsICash)
	getICashMercID = iID	
End Function

'------------------------------------------------------------------
' Gets shipping types
'------------------------------------------------------------------
Function getShipped(sProdID)
	Dim rsProdShipped, SQL
	SQL = "SELECT prodShipIsActive FROM sfProducts WHERE prodID = '" & sProdID & "'"
	Set rsProdShipped = Server.CreateObject("ADODB.Recordset")
	rsProdShipped.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
	getShipped = rsProdShipped(0)
	closeObj(rsProdShipped)
End Function


'---------------------------------------------------------------
' To see if it is a saved cart customer
' Returns a boolean value
'---------------------------------------------------------------
Function CheckSavedCartCustomer(iCustID)
	Dim sSQL, rsTmp, bTruth
	sSQL = "SELECT custFirstName FROM sfCustomers WHERE custID=" & iCustID
	
	bTruth = false
	
	Set rsTmp = Server.CreateObject("ADODB.RecordSet")
		 rsTmp.Open sSQL,cnn,adOpenDynamic,adLockOptimistic,adCmdText
		 If NOT rsTmp.EOF Then
		 	If trim(rsTmp.Fields("custFirstName")) = "Saved Cart Customer" Then
		 		bTruth = true
		 		
		 	Else
		 		bTruth = false
		 	End If		
		 End If
		
	closeobj(rsTmp)	
	CheckSavedCartCustomer = bTruth
End Function

'--------------------------------------------------------
' Checks if Customer exists in customer table
'--------------------------------------------------------
Function CheckCustomerExists(iCustID)
	Dim sSQL, rsCust, bExists
	sSQL = "SELECT custID FROM sfCustomers WHERE custID = " & iCustID
	Set rsCust = Server.CreateObject("ADODB.RecordSet")
		rsCust.Open sSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText
		
		If NOT rsCust.EOF Then
			If cInt(rsCust.Fields("custID")) > 0 Then
				bExists = true
			Else
				bExists = false	
			End If
		Else
			bExists = false	
		End If
		
	CheckCustomerExists = bExists
End Function
Function getCurrencyISO(slcid)
Dim rsSelect
dim sSql ,strLcid
set rsSelect = server.CreateObject ("ADODB.Recordset")
sSql = "Select slctvalLCID,slctvalCurrencyISO From sfSelectValues Where slctvalLCID = " & "'" &  slcid & "'" 
     rsSelect.Open sSql ,cnn,adOpenForwardOnly ,adLockReadOnly,adcmdtext
     
     getCurrencyISO = trim(rsSelect.Fields("slctvalCurrencyISO")) 
    'Response.Write getCurrencyISO &  "what the" 
    rsSelect.Close 
    set rsSelect = nothing
End Function
Function DeleteOrder(sID)

Dim rsDelete 'As New ADODB.Recordset
Dim rsDelete1 'As New ADODB.Recordset
Dim rsDelete2 'As New ADODB.Recordset
Dim rsDelete3 'As New ADODB.Recordset
Dim vOrderId 'As Variant
Dim sSql
On Error Resume Next
Set rsDelete = Server.CreateObject("ADODB.RecordSet")
Set rsDelete1 = Server.CreateObject("ADODB.RecordSet")
Set rsDelete2 = Server.CreateObject("ADODB.RecordSet")
Set rsDelete3 = Server.CreateObject("ADODB.RecordSet")

sSql = "SELECT * FROM sfOrders" _
        & " WHERE orderID = " & sID
    rsDelete.Open sSql, Cnn, adOpenKeyset, adLockOptimistic, adCmdText
vOrderId = rsDelete("orderAddrId")
sSql = "SELECT * FROM sfOrderDetails WHERE odrdtOrderId = " & rsDelete.Fields("orderID")
    rsDelete2.Open sSql, Cnn, adOpenKeyset, adLockOptimistic, adCmdText
'    '''''rsOrderCredit
sSql = "SELECT * FrOM sfCPayments WHERE payID = " & Trim(rsDelete.Fields("orderPayId"))
rsDelete3.Open sSql, Cnn, adOpenKeyset, adLockOptimistic, adCmdText
rsDelete.Delete 
'rsDelete1.Delete adAffectCurrent
rsDelete2.Delete 
rsDelete3.Delete 

Set rsDelete = Nothing
Set rsDelete1 = Nothing
Set rsDelete2 = Nothing
Set rsDelete3 = Nothing
End Function

Function Reset_Shipping()
	Dim sSql,RstProd,rsttmpOrder
	Set rstProd = Server.CreateObject("ADODB.RecordSet")
	Set rsttmpOrder = Server.CreateObject("ADODB.RecordSet")
	sSql = "SELECT * FROM sfTmpOrderDetails" _
	        & " WHERE odrdttmpSessionID = " & Session("SessionID")
	rsttmpOrder.Open sSql, Cnn, adOpenKeyset, adLockOptimistic, adCmdText
	
While rsttmpOrder.EOF =False
	 sSql = "SELECT prodShipIsActive FROM sfProducts " _
	        & " WHERE prodID = '" & rsttmpOrder("odrdttmpProductID") & "'"
	RstProd.Open sSql,cnn,adOpenStatic ,adLockReadOnly ,1
	If Not isNull(RstProd("prodShipIsActive")) then
	  rsttmpOrder("odrdttmpShipping") = RstProd("prodShipIsActive")
	Else
	  rsttmpOrder("odrdttmpShipping") = 0
	end if
	rsttmpOrder.Update 
	rsttmpOrder.MoveNext 
	rstProd.Close '#309
Wend	        

On error Resume Next
rsttmpOrder.Close 

Set rstProd =Nothing
Set rsttmpOrder = Nothing

End Function

Function get_Invalid_eMail(sData)
Dim iLoop,rst,sSql,sTemp,aCHK
aCHK = split(sData,",")
Set rst = Server.CreateObject("ADODB.RecordSet")
sTemp = ""
For iLoop = 0 to uBound(aCHK)
  sSql = "Select custId From sfCustomers Where CustEmail = '" & aCHK(iLoop) & "'"
 If rst.State = 1 then rst.Close 
  rst.Open sSql,cnn,adOpenStatic ,adLockReadOnly ,1
   If rst.EOF AND rst.BOF Then
    sTemp = sTemp & aCHK(iLoop) & ","
   
   End if
     
next
 If right(stemp,1) = ";" then
   sTemp = left(len(sTemp)-1)
 End If   
 
get_Invalid_eMail = sTemp
on error resume next
rst.Close 
set rst = nothing
End Function

%>








