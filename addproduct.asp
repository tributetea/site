<%@ Language=VBScript %>
<%	option explicit 
	Response.Buffer = True
	%>
      
<!--#include file="SFLib/db.conn.open.asp"-->
<!--#include file="SFLib/adovbs.inc"-->
<!--#include file ="SFLib/incAddProduct.asp"-->
<!--#include file="error_trap.asp"-->
<!--#include file="SFLib/incGeneral.asp"-->
<!--#include file="SFLIB/incAE.asp"-->
<%   
   Const vDebug = 0

'@BEGINVERSIONINFO

'@APPVERSION: 50.4011.0.2

'@FILENAME: addproduct.asp
 

'

'@DESCRIPTION: adds product to order

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000, 2001 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO

'@BEGINCODE

Dim sProdID, sCategory, sManufacturer, iQuantity, sProdPrice
Dim sSaveCart, sAddProduct, iOldPage, sSearchPath, sActionType, sResponseMessage	
Dim rsSelectProd, iProdCategoryID, iProdMaufacturerID, iProdVendorID
Dim sProdName, sProdMessage, iSaleIsActive, iProdAttrNum, aProdAttr, aProdValues, aAttrValues
Dim aManVenCatNames, sVendor, iTmpOrderDetailId, iAttrDtType
Dim dAttrTotal, dProdTotal, sSubtotal, dSubTotal, sReferer, iShip
Dim sSQL, iCounter, iGeneratedCartID, sLogedin, sThanksPath, sThanksType

' Collect variables passed from the Form 
sAddProduct	 			= Request.Form("AddProduct.x")
sSaveCart	 			= Request.Form("SaveCart.x")
sProdID		 			= Trim(Request("PRODUCT_ID"))
iQuantity	 				= Trim(Request("QUANTITY"))
iOldPage	 			= Trim(Request.Form("Order_Flag"))
iGeneratedCartID 		= Session.SessionID  ' Open to Other methods of generating IDs
sReferer	 			= Request.Cookies("sfHTTP_REFERER")("REFERER") & "," & Request.Cookies("sfHTTP_REFERER")("HTTP_REFERER") & "," & Request.Cookies("sfHTTP_REFERER")("REMOTE_ADDRESS")
sLogedin            			= Request.QueryString("logedin") 
If Session("SessionID")	 = "" or IsNull(Session("SessionID")) Then
Session("SessionID")	 = Session.SessionID
End If

If iQuantity = "" Or Not IsNumeric(iQuantity) Then iQuantity = 1

' Set the Search Path in Sessions for redirection in thnks
	If vDebug = 1 Then Response.Write "<br>" & Request.ServerVariables("HTTP_REFERER")

'Used in order and saved cart for shopping path
Response.Cookies("sfSearch")("SearchPath") = Request.ServerVariables("HTTP_REFERER")
Response.Cookies("sfSearch").Expires = Date() + 1

If Request.Cookies("sfAddProduct")("Path") = "" Then
	Response.Cookies("sfAddProduct")("Path") = Request.ServerVariables("HTTP_REFERER")
	Response.Cookies("sfAddProduct").Expires = Date() + 1
End If


' Get ActionType
If sSaveCart <> "" Then
	sActionType = "SaveProduct"
	If iQuantity > 1 Then
	sResponseMessage = "have been saved to your cart."
	Else
	sResponseMessage = "has been saved to your cart."
	End If
	AddProduct_SetsResponseMessage 'SFAE
Else
	sActionType = "AddProduct"	
	If iQuantity > 1 Then
	sResponseMessage = "have been added to your order."
	Else
	sResponseMessage = "has been added to your order."
	End If
End If	

'------------------------------------------------------------------------------
' Preliminary Data Collection Work 
'------------------------------------------------------------------------------

' Old Prod Page, Check input and convert sProd_Name to sProdID

If iOldPage <> "" Then
	Dim bResult, sProd_Name, element
    bResult = IsNumeric(iQuantity)	
	    If bResult = 0 Then
			Response.Write "<br>Quantity was not numeric"				
			Response.End				
		End If   
		
		For Each element In Request.Form
			If InStr(element,"PRODUCT_ID") Then 
				sProd_Name = element
			End	If
		Next
	sProdID = Trim(Request.Form(sProd_Name))
End If   

If sLogedin = "" Then
if trim(getProduct(sProdID)(0))<>"" then 

	' Get an array of 4 values from getProdValues()
    ReDim aProdValues(3)
			'On Error Resume Next
			aProdValues = getProdValues(sProdID,iQuantity)		
  			If aProdValues(0) <> "" Then sProdName = Server.URLEncode(aProdValues(0))
  			If aProdValues(1) <> "" Then sProdMessage = Server.URLEncode(aProdValues(1))
  			If aProdValues(2) <> "" Then iProdAttrNum = Server.URLEncode(aProdValues(2))
			iShip = aProdValues(3)
			'Call CheckForError()
   
	 ' Collect attributes into an array		
	 ' Old page compaitability

	    If iOldPage <> "" AND Request.Form("AttributeA") <> "" Then
				 ReDim aProdAttr(3)
					aProdAttr(0) = trim(Request.Form("AttributeA"))
					If (Request.Form("AttributeB") <> "") Then
						aProdAttr(1) = trim(Request("AttributeB"))
							If(Request.Form("AttributeC") <> "") Then
								aProdAttr(2) = trim(Request("AttributeC"))
							End If	
					End If	
				' get an array of the unique keys of the atttributes
	
				  aProdAttr = getAttrID(sProdID,aProdAttr)

				' At least one attribute is no longer in the attribute table for the product
				  If vDebug = 1 Then Response.Write "<p>iProdAttrNum = " & iProdAttrNum & "<br>aProdAttr Bound : " & Ubound(aProdAttr)
				
				 If CDbl(trim(Ubound(aProdAttr))) < CDbl(trim(iProdAttrNum)) Then
					 Response.Write "<p><b><font face=""verdana"" size=""2"">Product no longer has one of the attributes listed in the product pages. Please contact store owner.</font></b>"
					 ' Redirect to error page later --
					 Response.End
				  End If	 
				
	   Else If (Trim(Request("attr1")) <> "") Then 
	   		
			Dim sTmpAttr, sNewVar			
  			ReDim aProdAttr(iProdAttrNum)						
			iCounter = 1
			For iCounter = 1 to iProdAttrNum
			    sNewVar = "attr" & iCounter				    		     
				If (Request(sNewVar) <> "") Then	
					sTmpAttr = Request(sNewVar)
					aProdAttr(iCounter-1) = sTmpAttr				
				End If							
			Next	 		         
		End If  
	' End Attributes If		 
	  End If 	 


	'-------------------------------------------------------------------------------
    ' This shows whether there is a previous order or a new order. 
    ' New Products are treated like new orders but can be gathered together through
    ' the session varaible SessionCartID     
    '-------------------------------------------------------------------------------
	iTmpOrderDetailId = getOrderID("odrdttmp","odrattrtmp",sProdID,aProdAttr,iProdAttrNum)
	If vDebug = 1 Then Response.Write "<p>Found or Not Found -- Record " & iTmpOrderDetailID
	
	If iTmpOrderDetailID <> ""  Then

		If (sActionType = "AddProduct") Then

		'----------------------------------------------------------------------------
		' Insert Into TmpOrder Table 
		'----------------------------------------------------------------------------
		' New Order
		  If (iTmpOrderDetailID < 1) Then 
		   	  Dim sTmpAttrName, iUpperBound, iTmpAttrID, iTmpAttrDtID, sTmpAttribute
			  If vDebug = 1 Then Response.Write "<h2><font face=verdana color=#334455>New Order</h2>"

			  ' Write to TmpOrderDetails Table 
			  iTmpOrderDetailID = getTmpTable(aProdAttr,sProdID,iQuantity,sReferer,iShip)		 
		' Add to Existing Order
		  ElseIf (iTmpOrderDetailID > 0 AND iTmpOrderDetailID <> "") Then
				' Check to see if it is new product or adding to existing products
				If vDebug = 1 Then Response.Write "<h2><font face=""verdana"" color=""#334455"">Adding to Existing Cart</font></h2>"
				If vDebug = 1 Then Response.Write "<hr>Found OrderID = " & iTmpOrderDetailId
				
				' Update Quantity					
				Call setUpdateQuantity("odrdttmp",iQuantity,iTmpOrderDetailId)				

		' End Add Product If
		  End If
		  AddProduct_WriteTmpOrderDetailsAE  'SFAE
		'------------------------------------------------------------------------------
		' End Add Product -------------------------------------------------------------
		'------------------------------------------------------------------------------
		
		
		'------------------------------------------------------------------------------	
		' ActionType Save Cart
		'------------------------------------------------------------------------------
		ElseIf (sActionType = "SaveProduct") Then
			
		    Dim iCustID, iSvdCartID, sPathString, iSessionID	
		    
		    ' Request Cookie for custID 
			iCustID		= Trim(Request.Cookies("sfCustomer")("custID"))
		    iSessionID	= Trim(Request.Cookies("sfOrder")("SessionID"))
		    
		    ' If no cookie with custID, direct to Login
			If iCustID = "" or iSessionID <> Session("SessionID") Then			
				  ' Write to saved with custID of 0			  
				  Call getSavedTable(aProdAttr,sProdID,iQuantity,0,sReferer)
				  
				  ' Write to a cookie for thank you redirection
				  Response.Cookies("sfThanks")("PreviousAction") = "SaveCart"
				  Session("qString") = "thanks.asp?iQuantity=" & iQuantity & "&sProdName=" & sProdName & "&sProdMessage=" & sProdMessage & "&sResponseMessage="& Server.URLEncode(sResponseMessage)
				  'Response.Cookies("sfThanks")("qString") = "thanks.asp?iQuantity=" & iQuantity & "&sProdName=" & sProdName & "&sProdMessage=" & sProdMessage & "&sResponseMessage="& Server.URLEncode(sResponseMessage)
				  Response.Cookies("sfThanks").Expires = Date() + 1
				  
				' Redirect to login					  		
				  Response.Redirect("login.asp")				  
			Else		
				
				' Check for existing SessionCartId
				' Get SavedCart ID, -1 is returned if not found
					iSvdCartID = getOrderID("odrdtsvd","odrattrsvd", sProdID,aProdAttr,iProdAttrNum)

						If vDebug = 1 Then Response.Write "<p>Saved Cart Found or Not Found -- Record " & iSvdCartID
			
						If iSvdCartID <> "" Then
								If iSvdCartID < 0 Then						
									If vDebug = 1 Then Response.Write "<h2><font face=verdana color=#334455>Prior Record Not Found</h2></font>"						
									' Write as new row
									  Call getSavedTable(aProdAttr,sProdID,iQuantity,iCustID,sReferer)
								' Existing cart
								Else				
									If vDebug = 1 Then Response.Write "<h2><font face=verdana color=#334455>Adding to Existing Saved Cart</h2> <br>SvdCartID = " & iSvdCartID  
									' Update Quantity
										Call setUpdateQuantity("odrdtsvd",iQuantity,iSvdCartID)
								' End existing saved cart If	
								End If		
						Else
							Response.Write "<p>Number of attributes not equal to the product specs or database writing error"
				
						' End iSvdCartID Null If
						End If	
						
				
			' End No Cookie If	
			End If	
			
		' End Save To Cart If
		End If
	  
		 '---------------------------------------------------------------------------
		 ' End Save To Cart ---------------------------------------------------------
		 '---------------------------------------------------------------------------  
	 
	Else
		Response.Write "<br>Unknown ActionType Occurred or Database writing error"	
		Response.End	
	' End iTmpOrderDetail Null If
	End If 
	sThanksPath = "thanks.asp?iQuantity=" & iQuantity & "&sProdName=" & sProdName & "&sProdMessage=" & sProdMessage & "&sResponseMessage="& Server.URLEncode(sResponseMessage)
	sThanksType = 1
	AddProduct_SetThanksMessage 'SFAE
else
	sResponseMessage="This product is not currently available."
	sThanksPath = "thanks.asp?iQuantity=0&sProdName=&sProdMessage=&sResponseMessage="& Server.URLEncode(sResponseMessage)
	sThanksType = 1
end if
Else
	sThanksPath = Session("qString") 'Request.Cookies("sfThanks").Item("qString")
	sThanksType = Request.Cookies("sfAddProduct")("Path")
	Response.Cookies("sfAddProduct").Expires = Now()
	Response.Cookies("sfThanks").Expires = Now()
End If
cnn.Close
Set cnn = Nothing
%>  
<html>
<head>
<SCRIPT>
function redirect(path, type) {
	var sFeatures, h, w, myThanks, i
	h = window.screen.availHeight 
	w = window.screen.availWidth 
<%if sResponseMessage="This product is not currently available." then %>
	sFeatures = "height=" + h*.44 + ",width=" + h*.50 + ",screenY=" + (h*.30) + ",screenX=" + (w*.33) + ",top=" + (h*.30) + ",left=" + (w*.33) + ",resizable,scrollbars=yes"
<%else%>
	sFeatures = "height=" + h*.44 + ",width=" + h*.50 + ",screenY=" + (h*.30) + ",screenX=" + (w*.33) + ",top=" + (h*.30) + ",left=" + (w*.33) + ",resizable,scrollbars=yes"
<%end if%>
	
	myThanks = window.open(path,"",sFeatures)
		if ('<%=request("logedin")%>' == "1") {
		window.history.go(-2)
		//this makes it so that if you had to log in to save to cart, hitting back does not take yu to Thank you pop up 'JF
	}

	if (type == "1") {
		window.history.go(-1)
	}
	else {
		window.location = type	
	}
}
</SCRIPT>

<link rel="stylesheet" href="sfCSS.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title><%= C_STORENAME %>-SF Adds Order to Shopping Cart</title>
<meta http-equiv="Pragma" content="no-cache">


</head>
<body onload="javascript:redirect('<%= sThanksPath %>', '<%= sThanksType %>')">
</body>
</html>



