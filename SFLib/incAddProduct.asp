<%
'@BEGINVERSIONINFO

'@APPVERSION: 50.4011.0.2

'@FILENAME: incAddProduct.asp
	 

'

'@DESCRIPTION: Include File For addproduct.asp

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
' Add To Cart Functions ----------------------------------------------
'---------------------------------------------------------------------

'---------------------------------------------------------------------
' This function checks whether a product exists and retrieves an array of info
'---------------------------------------------------------------------
Function getProdValues(sProdID,iQuantity)
Dim sLocalSQL, aLocalProdArray(3)
Dim rsSelectProd
	sLocalSQL = "SELECT prodName, prodNamePlural, prodMessage, prodAttrNum, prodShipIsActive FROM sfProducts WHERE prodEnabledIsActive=1 AND prodID = '"& sProdID & "'"

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
  		  aLocalProdArray(3) = rsSelectProd.Fields("prodShipIsActive")

   	  ' End RecordSet If
  	  End If
  	  closeObj(rsSelectProd)

  	  getProdValues = aLocalProdArray
End Function


'---------------------------------------------------------------------
' Retrieves an array of the product's attributes' id
' This is to accommodate old product pages, so only three attributes are needed

'---------------------------------------------------------------------
Function getAttrID(sProdID,aProdAttr)
	Dim sLocalSQL, aLocalArray,  iArrayBound, rsGetAttrID
	
	sLocalSQL = "SELECT attrdtID FROM sfAttributeDetail INNER JOIN sfAttributes ON sfAttributeDetail.attrdtAttributeId = sfAttributes.attrID WHERE trim(attrProdId) = '"& sProdID & "'"_
	& " AND (trim(sfAttributeDetail.attrdtName) = '" & aProdAttr(0) & "' OR trim(sfAttributeDetail.attrdtName) = '" & aProdAttr(1) & "' OR trim(sfAttributeDetail.attrdtName) = '" & aProdAttr(2) & "')"

		If vDebug = 1 Then  Response.Write "<br>Attribute SQL: " & sLocalSQL
  		Set rsGetAttrID = Server.CreateObject("ADODB.RecordSet")
  		rsGetAttrID.Open sLocalSQL, cnn, adOpenKeyset, adLockOptimistic, adCmdText
	
	  	' Check if this record exists through prodID and price matches
		If rsGetAttrID.EOF or rsGetAttrID.BOF Then
			If vDebug = 1 Then Response.Write "<br> Empty Recordset for rsGetAttrID"
		Else
  			iCounter = 0
  			iArrayBound = rsGetAttrID.RecordCount
  			If vDebug = 1 Then Response.Write "<br>Array Bound in getAttrID" & iArrayBound
  			ReDim aLocalArray(iArrayBound)
			  			  			
  			Do While NOT rsGetAttrID.EOF
				aLocalArray(iCounter) = rsGetAttrID.Fields("attrdtID")
				If vDebug = 1 Then Response.Write "<br>Array ID = " & aLocalArray(iCounter)
  				iCounter = iCounter + 1
  				rsGetAttrID.MoveNext   		
   			Loop
		   	' End RecordSet if	
  		End If
  	  
  	  closeObj(rsGetAttrID)
  	  getAttrID = aLocalArray
End Function

'-------------------------------------------------------------------
' This combines identical items in saved cart
'-------------------------------------------------------------------
Sub setCombineProducts(iCustID)
Dim	sLocalSQL,sLocalSQL2,rsSaved,iTotalRecords, tmpProdID, i, j, tmpProdQuantity, tmpProdSvdID, rsDetails, rsDetails2, cmpProdID
Dim aProd, rsCheckExists, sCheckExists, sAttrString, sAttrString2	
	sLocalSQL = "SELECT odrdtsvdID, odrdtsvdQuantity, odrdtsvdProductID FROM sfSavedOrderDetails WHERE odrdtsvdCustID = " & iCustID
	
	Set rsSaved = Server.CreateObject("ADODB.RecordSet")
	rsSaved.Open sLocalSQL, cnn, adOpenKeySet, adLockOptimistic, adCmdText
	iTotalRecords = rsSaved.RecordCount
	i = 0
	Redim aProd(3,iTotalRecords-1)
	
	Do While Not rsSaved.EOF
		tmpProdID		= Trim(rsSaved.Fields("odrdtsvdProductID"))
		tmpProdQuantity = Trim(rsSaved.Fields("odrdtsvdQuantity"))
		tmpProdSvdID	= Trim(rsSaved.Fields("odrdtsvdID"))
	
		aProd(0,i) = tmpProdID
		aProd(1,i)	= tmpProdQuantity
		aProd(2,i) = tmpProdSvdID		
		
	i = i + 1
	rsSaved.MoveNext
	Loop
	
	
	For i = 0 to UBOUND(aProd,2)
		tmpProdID = aProd(0,i)
		tmpProdSvdID = aProd(2,i)
		
		For j = i + 1 to UBOUND(aProd,2)
				cmpProdID = aProd(0,j)
				If cmpProdID = tmpProdID Then

					sLocalSQL = "SELECT odrattrsvdID, odrattrsvdOrderDetailId, odrattrsvdAttrID "_
					& "FROM sfSavedOrderDetails INNER JOIN sfSavedOrderAttributes on sfSavedOrderAttributes.odrattrsvdOrderDetailId = sfSavedOrderDetails.odrdtsvdID "_
					& "WHERE sfSavedOrderAttributes.odrattrsvdOrderDetailId = " & aProd(2,j)
						
					sLocalSQL2 = "SELECT odrattrsvdID, odrattrsvdOrderDetailId, odrattrsvdAttrID "_
					& "FROM sfSavedOrderDetails INNER JOIN sfSavedOrderAttributes on sfSavedOrderAttributes.odrattrsvdOrderDetailId = sfSavedOrderDetails.odrdtsvdID "_
					& "WHERE sfSavedOrderAttributes.odrattrsvdOrderDetailId = " & aProd(2,i)	
					
					sCheckExists = "SELECT odrdtsvdID FROM sfSavedOrderDetails WHERE sfSavedOrderDetails.odrdtsvdID = " & aProd(2,i)
					
					Set rsDetails = Server.CreateObject("ADODB.RecordSet")
					rsDetails.Open sLocalSQL, cnn, adOpenKeySet, adLockOptimistic, adCmdText
					
					Set rsCheckExists = Server.CreateObject("ADODB.RecordSet")
						rsCheckExists.Open sCheckExists, cnn, adOpenKeySet, adLockOptimistic, adCmdText
						If NOT rsCheckExists.EOF Then
								Set rsDetails2 = Server.CreateObject("ADODB.RecordSet")
								rsDetails2.Open sLocalSQL2, cnn, adOpenKeySet, adLockOptimistic, adCmdText
							
								If rsDetails.EOF AND rsDetails2.EOF Then
									' combine the two since there are no attributes
									Call setUpdateQuantity("odrdtsvd",aProd(1,j),tmpProdSvdID)
									Call setDeleteOrder("odrdtsvd",aProd(2,j))	
								Else
									' compare the attributes
									Do While Not rsDetails.EOF
										sAttrString = sAttrString & "[" & Trim(rsDetails.Fields("odrattrsvdAttrID")) & "]"							
										rsDetails.MoveNext
									Loop
									Do While Not rsDetails2.EOF
										sAttrString2 = sAttrString2 & "[" & Trim(rsDetails2.Fields("odrattrsvdAttrID")) & "]"
										rsDetails2.MoveNext
									Loop
							
									If sAttrString	= sAttrString2 Then
										Call setUpdateQuantity("odrdtsvd",aProd(1,j),tmpProdSvdID)
										Call setDeleteOrder("odrdtsvd",aProd(2,j))
									End If	
								End If	
						End If					
				End If					
			Next
	Next		
closeObj(rsDetails)
closeObj(rsCheckExists)
closeObj(rsSaved)				
End Sub

'-----------------------------------------------------------------------
' Deletes saved customer row
'-----------------------------------------------------------------------
Sub DeleteCustRow(iCustID)
	Dim rsDelete, sSQL
	
	sSQL = "DELETE FROM sfCustomers WHERE custID= " & iCustID	& " AND custFirstName = 'Saved Cart Customer'"
	Set rsDelete = cnn.Execute(sSQL)
	closeObj(rsDelete)
End Sub


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

'-------------------------------------------------------
' Add new customer's info in sfCustomers
' Returns ID of inserted row
'-------------------------------------------------------
Function getSvdCustomer(sCustEmail,sPassword)			
Dim	sLocalSQl, rsUpdate, iKeyID, bookMark

	Set rsUpdate = Server.CreateObject("ADODB.RecordSet")
		rsUpdate.CursorLocation = adUseClient
		rsUpdate.Open "sfCustomers ORDER BY custID",cnn,adOpenDynamic,adLockOptimistic,adCmdTable
		rsUpdate.AddNew
		rsUpdate.Fields("custFirstName")		= "Saved Cart Customer"	
		rsUpdate.Fields("custPasswd")			= trim(sPassword)
		rsUpdate.Fields("custEmail")			= trim(sCustEmail)
		rsUpdate.Fields("custTimesAccessed")	= 1
		rsUpdate.Fields("custLastAccess")		= Date()
		
		rsUpdate.Update		
		'bookMark = rsUpdate.AbsolutePosition 
		'rsUpdate.Requery 
		'rsUpdate.AbsolutePosition = bookMark			
		iKeyID  = rsUpdate.Fields("custID")
		closeObj(rsUpdate)	
		
		getSvdCustomer = iKeyID
End Function

'-------------------------------------------------------------------
' Subroutine setUpdateSavedCartCustID
'-------------------------------------------------------------------
Sub setUpdateSavedCartCustID(iCustID,iDeletedCustID)
	Dim sSQL, rsTmpCust
	sSQL = "Select odrdtsvdCustID FROM sfSavedOrderDetails WHERE odrdtsvdCustID=" & iDeletedCustID
	Set rsTmpCust = Server.CreateObject("ADODB.RecordSet")		
		rsTmpCust.Open sSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText					
			Do While NOT rsTmpCust.EOF
					rsTmpCust.Fields("odrdtsvdCustID")	= trim(iCustID)
					rsTmpCust.Update	
					rsTmpCust.MoveNext
			Loop
		closeobj(rsTmpCust)		
End Sub

%>



