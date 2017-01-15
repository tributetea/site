<%
'@BEGINVERSIONINFO

'@APPVERSION: 50.4011.0.2

'@FILENAME: incorder.asp
	 
'

'@DESCRIPTION: Include File for Order.asp

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws as an unpublished work, and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO
'---------------------------------------------------------------------
' Generates a list for payment methods -- to move to sflists
'---------------------------------------------------------------------
Function getPaymentMethods()
Dim sLocalSQL, rsTransType, sList, sTempTransType

	sLocalSQL = "SELECT DISTINCT transtype FROM sfTransactionTypes WHERE transIsActive = 1"
	
	Set rsTransType = Server.CreateObject("ADODB.RecordSet")
		rsTransType.Open sLocalSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText
		
		If rsTransType.EOF or rsTransType.BOF Then
			Response.Write "<p>Admin Payment Error: Please contact store owner"
		Else	
			Do While NOT rsTransType.EOF
				sTempTransType = rsTransType.Fields("transtype")
				If sTempTransType = "Credit Card" Then
					sList = sList & "<option selected value= """ & sTempTransType  & """>" & sTempTransType & "</option>"
					rsTransType.MoveNext			
				Else
					sList = sList & "<option value= """ & sTempTransType  & """>" & sTempTransType & "</option>"
					rsTransType.MoveNext			
				End If  
			Loop	
			
		' End RecordSet If	
		End If	
		closeObj(rsTransType)
	getPayMentMethods = sList	
End Function


'--------------------------------------------------------------------
' Get back an array of attribute ids
'--------------------------------------------------------------------
Function getProdAttrID(sPrefix,iOrderID,iProdAttrNum)
Dim sLocalSQL, rsGetAttr, aAttrIDArray, iCounter, iArrayBound

Select Case sPrefix
	Case "odrattrtmp"
	sLocalSQL = "SELECT odrattrtmpID FROM sfTmpOrderAttributes WHERE odrattrtmpOrderDetailId = " & iOrderID
	Case "odrattrsvd"	
	sLocalSQL = "SELECT odrattrsvdID FROM sfSavedOrderAttributes WHERE odrattrsvdOrderDetailId = " & iOrderID
End Select
If vDebug = 1 Then Response.Write "<p> getProdAttrID SQL: " & sLocalSQL & "<br> iProdAttrNum: " & iProdAttrNum
	Set rsGetAttr = Server.CreateObject("ADODB.RecordSet")
		rsGetAttr.Open sLocalSQL, cnn, adOpenKeyset, adLockOptimistic, adCmdText
	
	iCounter = 0
	iArrayBound = rsGetAttr.RecordCount
		
		rsGetAttr.MoveFirst
		If iArrayBound < iProdAttrNum Then
				Response.Write "<p>An Error in Attributes has occured. Possible DB Error"
				Response.Write "<br>iArrayBound = " & iArrayBound & "<br>iProdAttrNum = " & iProdAttrNum
				Exit Function
		Else
			Redim aAttrIDArray(iArrayBound)
				Do While Not rsGetAttr.BOF Or rsGetAttr.EOF
					aAttrIDArray(iCounter) = rsGetAttr(sPrefix & "ID")
					iCounter = iCounter + 1
					rsGetAttr.MoveNext
				Loop
		End If	
	closeObj(rsGetAttr)
	getProdAttrID = aAttrIDArray
End Function
 

'---------------------------------------------------------------------
' Select manufacturer, vendor, category Names from ids associated
' Returns an arrary with manufacturer,vendor,category name values
'---------------------------------------------------------------------
Function getIdNames(iProdManufacturerID,sProdVendorID,iProdCategoryID)
Dim rsGetIDNames, aIdNames(3), sLocalSQL

	  sLocalSQL = "SELECT sfManufacturers.mfgName, sfVendors.vendName, sfCategories.catName FROM sfManufacturers, sfVendors, sfCategories WHERE "_
			 & " sfManufacturers.mfgID = " & iProdManufacturerID & " AND sfVendors.vendID = " & sProdVendorID & " AND sfCategories.catID = " & iProdCategoryID

	  If vDebug = 1 Then Response.Write "<p> getIDName SQL : " & sLocalSQL

      Set rsGetIdNames = Server.CreateObject("ADODB.RecordSet")
  	  rsGetIdNames.Open sLocalSQL, cnn

  	' Check if this record exists through prodID and price matches
      If rsGetIdNames.BOF Or rsGetIdNames.EOF Then
  		  Response.Write "<p>Product " & sProdID & " could not find accompanying ID names for vendors, manufacturers, and category"  		    		  
  	  Else
  	    ' get variables
  		  aIdNames(0) = rsGetIdNames.Fields("catName")
  		  aIdNames(1) = rsGetIdNames.Fields("mfgName")
  		  aIdNames(2) = rsGetIdNames.Fields("vendName")
   	  rsGetIdNames.Close
	  End If

	closeObj(rsGetIdNames)
	getIdNames = aIdNames
End Function


'---------------------------------------------------------------------
' This writes into Order Details, Returns ID of OrderDetails +++++ TO CHANGE TO ORDER
'---------------------------------------------------------------------
Function getTmpOrderDetails(iOrderDetID,iQuantity,sSubtotal,sCategory,sManufacturer,sVendor,sProdName,sProdPrice,sProdID,sReferer)
Dim rsOrderInput, iLocalID

Set rsOrderInput = Server.CreateObject("ADODB.RecordSet")
	rsOrderInput.CursorLocation = adUseClient
	rsOrderInput.Open "sfTmpOrderDetails", cnn, adOpenDynamic, adLockOptimistic,adCmdTable

	'Input base product information
	rsOrderInput.AddNew
	rsOrderInput.Fields("odrdttmpQuantity")		= iQuantity
	rsOrderInput.Fields("odrdttmpSubtotal")		= sSubtotal
	rsOrderInput.Fields("odrdttmpCategory")		= sCategory
	rsOrderInput.Fields("odrdttmpManufacturer") = sManufacturer
	rsOrderInput.Fields("odrdttmpVendor")		= sVendor
	rsOrderInput.Fields("odrdttmpProductName")	= sProdName
	rsOrderInput.Fields("odrdttmpPrice")		= sProdPrice
	rsOrderInput.Fields("odrdttmpProductID")	= sProdID
	rsOrderInput.Fields("odrdttmpSessionID")	= Session.SessionID
	rsOrderInput.Fields("odrdttmpHttpReferer")	= sReferer
	rsOrderInput.Update
	iLocalID = rsOrderInput.Fields("odrdttmpID")
	
	If vDebug = 1 Then Response.Write "<p><font color=""#335544"" size=""6""> Tmp Order ID = "  & iLocalID & "</font>"
		
	rsOrderInput.Close
	Set rsOrderInput = Nothing
	getTmpOrderDetails = iLocalID
End Function


'---------------------------------------------------------------------
' This subroutine writes into sfOrderAttributes, returns the ID number ++ TO WRITE TO ORDER ATTRIBUTES
'---------------------------------------------------------------------
Function setTmpOrderAttributes(iOrderDetailID,sAttribute,sAttrName,sAttrPrice,iType)
Dim rsAttributeInput, iLocalResult, sLocalSQL

	Set rsAttributeInput = Server.CreateObject("ADODB.RecordSet")

	'Input information into the database
	rsAttributeInput.Open "sfTmpOrderAttributes", cnn, adOpenDynamic, adLockOptimistic
	rsAttributeInput.AddNew
	rsAttributeInput.Fields("odrattrtmpOrderDetailID")	= iOrderDetailID
	rsAttributeInput.Fields("odrattrtmpAttribute")		= sAttribute
	rsAttributeInput.Fields("odrattrtmpName")			= sAttrName
	rsAttributeInput.Fields("odrattrtmpPrice")			= sAttrPrice
	rsAttributeInput.Fields("odrattrtmpType")			= iType
	rsAttributeInput.Update
	rsAttributeInput.Close

	sLocalSQL = "SELECT odrattrtmpID FROM sfTmpOrderAttributes WHERE odrattrtmpName ='"_
				 & sAttrName & "' AND odrattrtmpOrderDetailID = " & iOrderDetailID
	rsAttributeInput.Open sLocalSQL,cnn

		If rsAttributeInput.BOF Or rsAttributeInput.EOF Then
			Response.Write "<br> Empty recordset in rsAttributeInput" 
		Else
			iLocalResult = rsAttributeInput.Fields("odrattrtmpID")
		End If

	closeObj(rsAttributeInput)
	setTmpOrderAttributes = iLocalResult
End Function

'---------------------------------------------------------------------
' Copies record from tmpOrders to svdOrders, returns the key of the SvdOrder
'---------------------------------------------------------------------
Function setCopyToSavedTable(iTmpOrderDetailID,sProdID,iNewQuantity,iCustID)
Dim rsCopy, sLocalSQL, rsSvdCart, rsSvdCartAttr, iKeyID, sReferer,sDateTime,sTmpAttrName, sTmpAttrID, aTmpOrderArray

	sLocalSQL = "Select * FROM sfTmpOrderDetails INNER JOIN sfTmpOrderAttributes ON odrdttmpID = odrattrtmpOrderDetailId "_
				& " WHERE odrdttmpID = " & iTmpOrderDetailID & " And odrdttmpSessionID = " & Session.SessionID

	If vDebug = 1 Then Response.Write "<p>getCopyToSvdTable SQL : " & sLocalSQL

	Set rsCopy = Server.CreateObject("ADODB.RecordSet")
	rsCopy.Open sLocalSQL, cnn, adOpenStatic, adLockOptimistic

	If vDebug = 1 Then
		Dim aVarArray, iCounter
		iCounter = 0
		Do While NOT rsCopy.EOF		
			Response.Write "<br> TmpOrder ID : " & rsCopy.Fields(0)
			Response.Write "<br> Attr? " & rsCopy.Fields("odrattrtmpID")
			rsCopy.MoveNext
			iCounter = iCounter + 1	
			Response.Write "<br>Loop"
		Loop
		rsCopy.MoveFirst		
	End If
		

	If (rsCopy.BOF Or rsCopy.EOF)Then
		Response.Write "<br> Empty RecordSet in rsCopy"
	Else
	
		' Collect from TmpOrderDetails
			aTmpOrderArray = rsCopy.GetRows		
			iQuantity = aTmpOrderArray(1,0) + iNewQuantity
			sReferer = aTmpOrderArray(4,0)	
			sDateTime = FormatDateTime(Now)
			rsCopy.MoveFirst			
		
		' Copy to svd cart			
		Set rsSvdCart = Server.CreateObject("ADODB.RecordSet")
			rsSvdCart.CursorLocation = adUseClient
			rsSvdCart.Open "sfSavedOrderDetails", cnn, adOpenDynamic, adLockOptimistic
				rsSvdCart.AddNew
				rsSvdCart.Fields("odrdtsvdCustID")		= iCustID
				rsSvdCart.Fields("odrdtsvdQuantity")	= iQuantity
				rsSvdCart.Fields("odrdtsvdProductID")	= sProdID
				rsSvdCart.Fields("odrdtsvdDate")		= sDateTime
				rsSvdCart.Fields("odrdtsvdSessionID")	= Session.SessionID
				rsSvdCart.Fields("odrdtsvdHttpReferer") = sReferer
				rsSvdCart.Update
				iKeyID  = rsSvdCart.Fields("odrdtsvdID")
	
		If vDebug = 1 Then Response.Write "<p><font size=4><b>SvdCart Key ID = " & iKeyID & "</b></font>"
		' Copy Attributes
		Do While Not rsCopy.EOF		
			' Collect Attribute Info from sfTmpOrderAttributes
			sTmpAttrID = Trim(rsCopy.Fields("odrattrtmpAttrID"))
				Set rsSvdCartAttr = Server.CreateObject("ADODB.RecordSet")
				rsSvdCartAttr.Open "sfSavedOrderAttributes", cnn, adOpenDynamic, adLockOptimistic
					rsSvdCartAttr.AddNew
					rsSvdCartAttr.Fields("odrattrsvdOrderDetailId") = iKeyID
					rsSvdCartAttr.Fields("odrattrsvdAttrID") = sTmpAttrID
				rsSvdCartAttr.Update							
			rsCopy.MoveNext
		Loop	
	End If
	
  	closeObj(rsCopy)
  	closeObj(rsSvdCart)
  	closeobj(rsSvdCartAttr)
  	setCopyToSavedTable = iKeyId
End Function

'---------------------------------------------------------------------
' This function checks whether a product exists and retrieves an array of info
'---------------------------------------------------------------------
Function getProdValues(sProdID,iQuantity)
Dim sLocalSQL, aLocalProdArray(3), rsSelectProd

	sLocalSQL = "SELECT prodName, prodNamePlural, prodMessage, prodAttrNum FROM sfProducts WHERE prodEnabledIsActive=1 AND prodID = '"& sProdID & "'"

  	  Set rsSelectProd = Server.CreateObject("ADODB.RecordSet")
  	  rsSelectProd.Open sLocalSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText

	  If vDebug = 1 Then Response.Write "<p>getProdValues SQL : " & sLocalSQL
	  
  	' Check if this record exists through prodID and price matches
      If rsSelectProd.EOF or rsSelectProd.BOF Then  		 
  		 If vDebug = 1 Then Response.Write "<p>Empty Recordset in rsSelectProd. Product " & sProdID & " possibly not activated."  		 
  	  Else
  		  If rsSelectProd.Fields("prodNamePlural") <> "" And iQuantity > 1 Then
  			 aLocalProdArray(0) = rsSelectProd.Fields("prodNamePlural")
  		  Else	 
  			 aLocalProdArray(0) = rsSelectProd.Fields("prodName")
  		  End If	  		 
  		  aLocalProdArray(1) = rsSelectProd.Fields("prodMessage")
  		  aLocalProdArray(2) = rsSelectProd.Fields("prodAttrNum")

   	  ' End RecordSet If
  	  End If
  	  closeObj(rsSelectProd)

  	  getProdValues = aLocalProdArray
End Function
%>



