<%@ Language=VBScript %>
<%	option explicit 
	Response.Buffer = True
	%>
	<!--#include file="SFLib/db.conn.open.asp"-->
<!--#include file="SFLib/adovbs.inc"-->
<!--#include file="error_trap.asp"-->

<!--#include file="SFLib/incConfirm.asp"-->
<!--#include file="SFLib/incDesign.asp"-->
<!--#include file="SFLib/incVerify.asp"-->
<!--#include file="SFLib/incGeneral.asp"-->
<!--#include file="SFLib/incCC.asp"-->
<!--#include file="SFLIB/incAE.asp"-->
	<%
   Const vDebug = 0
   ' #52 Start 
If InStr(1, Request.ServerVariables("HTTP_REFERER"), "process_order.asp",1) = 0 and Request.querystring("optionID")="" then
	response.end
   end if
   ' #52 End
%>
<html>
<head>
<script LANGUAGE="javascript">
					<!--

function ValidateMe(txtbox)
  {                     
     if(txtbox.value != "")
      {  
       var sval;
       if(txtbox.name == 'CardExpiryMonth')
        {
         sval =txtbox.value
         if((sval < 1) || (sval > 12))
           {
            alert("Please enter a valid expiration month");
            txtbox.focus();
            return false;
           } 
         if(sval < 10 && sval.length == 1 )
           {   
             txtbox.value = "0" + sval;           
           } 
      }
     if(txtbox.name == 'CardExpiryYear')
      {   
        var d = new Date();                           
        var yy = (d.getFullYear());  
        var mm = (d.getMonth() +1);
        sval = txtbox.value
        if(document.frmVerify.CardExpiryMonth.value == "")
          {
           document.frmVerify.CardExpiryMonth.focus();
           return false;
          }
        if(txtbox.length < 4)
         {
          alert("Enter a valid 4 Digit Date ");
          txtbox.focus();
          return false;
         }  
        if(sval < yy)
         { 
          alert("Date is not Valid");
          txtbox.focus();
          return false;
         }
       if((sval == yy)&& (mm > document.frmVerify.CardExpiryMonth.value))
           alert("Date is not Valid");
           txtbox.focus();
           return false;
      }
   }
       return true;
}
					

//-->
</SCRIPT>
<script language="javascript" src="SFLib/sfCheckErrors.js"></script>
<link rel="stylesheet" href="sfCSS.css" type="text/css">
<%

 Response.CacheControl = "no-cache"
 Response.AddHeader "Pragma", "no-cache"
 Response.Expires = -1


%>

<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">

<%
'@BEGINVERSIONINFO

'@APPVERSION: 50.4011.0.6

'@FILENAME: verify.asp
	 



'@DESCRIPTION: Verify all user information 

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000,2001 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO
'modified 10/23/01
'Storefront Ref#'s:150,158
'modified 10/31/01
'Storefront Ref#'s:195
'modified 12/4/01
'Storefront Ref#'s:249
%>
<%
dim tmpOrderQty
If Application("AppName") = "StoreFrontAE" Then 'SFAE
	tmpOrderQty=GetTotalOrderQTY
else
	tmpOrderQty=GetTotalOrderQTYSE
end if	
if isnumeric(tmpOrderQty)=false or tmpOrderQty <= 0 then
'If Session("SessionID") = "" Then 'SFUPDATE
		' redirect to neworder screen
		Session.Abandon
		Response.Redirect(C_HomePath & "search.asp")				
End If

IF lcase(Request.ServerVariables("HTTPS")) = "on" then '.3008
   If iConverion = 1 Then Response.Write "<script language=""JavaScript"" src=""https://www.oanda.com/cgi-bin/fxcommerce/fxcommerce.js?user=" & sUserName & """></script>" 
ELSE
   If iConverion = 1 Then Response.Write "<script language=""JavaScript"" src=""http://www.oanda.com/cgi-bin/fxcommerce/fxcommerce.js?user=" & sUserName & """></script>"
End IF

Dim CurrencyISO 
Dim sSubmitActionAE ,sLTL,bLtl,arrLTL,bFromVerify'SFAE

CurrencyISO = getCurrencyISO(Session.LCID )	
		
	Dim rsAdminInfo, sCustFirstName, sCustMiddleInitial, sCustLastName, sCustCompany, sCustAddress1, sCustAddress2, sPaymentMethod, iShipMethod, sInstructions, sShipMethodName
	Dim sCustCity, sCustState, sCustStateName, sCustZip, sCustCountry, sCustCountryName, sCustPhone, sCustFax, sCustEmail, sCustCardType, sCustCardTypeName, sShipCustFirstName, sShipCustMiddleInitial, sShipCustLastName
	Dim sShipCustAddress1, sShipCustAddress2, sShipCustCity, sShipCustState, sShipCustStateName, sShipCustZip, sShipCustCountry,sShipCustCountryName, sShipCustPhone
	Dim sShipCustFax, sShipCustCompany, sShipCustEmail,sCustCardName,sCustCardNumber,sCustCardExpiry,sCustCardExpiryMonth,sCustCardExpiryYear, rsGetCustCardDetails
	Dim sBankName, iCheckNumber, iRoutingNumber, iCheckingAccountNumber, sPOName, iPONumber, iCardID, iCustID, iShipID, bCustSubscribed,iHandlingType, aAdmin
	Dim iShipType, iShipType2, sHandling, iHandling, iPremiumShipping, iPhoneFaxType, sPassword, strOut, iShipped, iCodAmount, bCod, sCCList, sSubmitAction, Path
	Dim sTransMethod,bshipped, sLogin, sTotalShipPrice, iShip, odrdttmpShipping,sFreeShip, sShipCode
	Dim sGrandTotal,sGrandTotalOut,sTotalPriceOut, sPaymentServer 'SFUPDATE
	dim dUnitPrice,bTaxShipIsActive,dTaxAble_Amount
        bFromVerify = false
        'If Application("AppName") = "StoreFront" Then 'SFAE
          if Request.Form("FromVerify") = "1" then
             bFromVerify = true
         end if 
       'end if
        iShipped = 0
        
 ' Response.Write  Request.Form("OptionID")
	    Set rsAdminInfo = Server.CreateObject("ADODB.RecordSet")
		rsAdminInfo.Open "sfAdmin", cnn, adOpenStatic, adLockReadOnly, adCmdTable
		With rsAdminInfo
			iShipType		 = .Fields("adminShipType")
			iShipType2		 = .Fields("adminShipType2")
			sHandling		 = .Fields("adminHandling")
			iHandling		 = .Fields("adminHandlingIsActive")
			iHandlingType    = .Fields("adminHandlingType")
			iCodAmount       = .Fields("adminCODAmount")	
			bTaxShipIsActive = .Fields("adminTaxShipIsActive")	'#383			
		End With
		closeObj(rsAdminInfo)	
			
		aAdmin = getAdminRow(sPaymentMethod)	
		sTransMethod = aAdmin(0)
		' Payment server
		sPaymentServer = aAdmin(1)
		' Login 
		sLogin = aAdmin(2)
		bshipped = Request.Form ("bShip")		
	  
	   if iHandlingType = 2 and bshipped = "0" then
			
			iShipped = 0
	   else
	    
	     iShipped = 1		
	   End If

	
	
	'---------------SFUPDATE ------------------------------------------
	   If  iShipped <> 1 Then 
	     sHandling = 0
	   end if
	   
	   If trim(iHandling) = "1" Then 
		If iHandlingType = 1 Then 
			iShipped = 1
		Else
			iShipped = 0
		End If
			Else
			sHandling = 0
		End If
		'

	 '----------------------------------------------------------
	   
	   
	   'bCod = False 
       sFreeShip =False
	   sPaymentMethod = Request.Form("paymentmethod")
	   iShipMethod = Request.Form("Shipping")

	   'COD Charge
	If sPaymentMethod <> "COD" Then			
		iCodAmount = 0
	End If	
	
	   
	    dim posit
	     posit =instr(iShipMethod,",")
	    if posit > 0 then
	    iShipMethod = left(trim(iShipMethod),posit-1) 'FreeShipping''''''''''''''''''''''
	      sFreeShip =True
	    end if 
	   sInstructions = Request.Form("Instructions")
		If sPaymentMethod = "COD" Then
		bCod = "True"
		End If
		
	    ' If PhoneFax, determine What kind it is
	    ' iPhoneFaxType = 0 is written to the database except for specific payment information
	    ' iPhoneFaxType = 1 is the "no database write at all kind of phonefax

	    If sPaymentMethod = "PhoneFax" Then
		  iPhoneFaxType = CheckPhoneFax()			
	    End If
		
 If sPaymentMethod = "Credit" then
   sPaymentMethod = "Credit Card"
 End If

	
If (sPaymentMethod = "Credit Card" AND (sTransMethod <> "15" AND sTransMethod <> "18" AND sPaymentMethod <> "PayPal")) Then
		sSubmitAction = "this.CardNumber.creditCardNumber = true;this.CardExpiryMonth.creditCardExpMonth = true;this.CardExpiryYear.creditCardExpYear = true;return sfCheck(this);"
Elseif sPaymentMethod = "PhoneFax" OR sPaymentMethod = "COD" Then 
		sSubmitAction ="" ' "this.CardType.optional = true;this.CardName.special = true;this.CardNumber.special = true;this.CardExpiryMonth.special = true;this.CardExpiryYear.special = true;this.CheckNumber.special = true;this.BankName.special = true;this.RoutingNumber.special = true;this.CheckingAccountNumber.special = true;this.POName.special = true;this.PONumber.special = true;return sfCheck(this);"
Elseif sPaymentMethod = "eCheck"  Then 
        'sSubmitAction = "return Check_EC_PO('CheckNumber', this);"
        sSubmitAction = ""
Elseif sPaymentMethod = "PO"  Then 
      sSubmitAction = "return POCheck(POName.value,PONumber.value);"  ' #303
Else
	    sSubmitAction = "return sfCheck(this);"
End If
	    ' Determine Ship Method
	  if cstr(iship) <> "0" then
	    If (iShipType = 1 Or iShipType = 3) Then
			If iShipMethod = 1 Then
				sShipMethodName = "Premium Shipping"
			ElseIf iShipMethod = 0 Then	
				sShipMethodName = "Regular Shipping"
			ElseIf iShipMethod = 3 Then	
				sShipMethodName = "Free Shipping"
			End If							
		ElseIf iShipType = 2 Then
			sShipMethodName = getNameWithID("sfShipping",iShipMethod,"shipID","shipMethod",0)				
		End If		
	  End If	
			
		' Update Customer Record
		' Gather info to put into customer table		
		sCustFirstName			= Trim(Request.Form("FirstName"))
		sCustMiddleInitial		= Trim(Request.Form("MiddleInitial"))
		sCustLastName			= Trim(Request.Form("LastName"))
		sCustCompany			= Trim(Request.Form("Company"))
		sCustAddress1			= Trim(Request.Form("Address1"))
		sCustAddress2			= Trim(Request.Form("Address2"))
		sCustCity				= Trim(Request.Form("City"))
		sCustState				= Trim(Request.Form("State"))
		sCustStateName		    = Trim(getNameWithID("sfLocalesState",sCustState,"loclstAbbreviation","loclstName",1))
		sCustZip				= Trim(Request.Form("Zip"))
		sCustCountry			= Trim(Request.Form("Country"))
		sCustCountryName		= Trim(getNameWithID("sfLocalesCountry",sCustCountry,"loclctryAbbreviation","loclctryName",1))	
		sCustPhone				= Trim(Request.Form("Phone"))
		sCustFax				= Trim(Request.Form("Fax"))
		sCustEmail		    	= Trim(Request.Form("Email"))
		bCustSubscribed     	= Trim(Request.Form("Subscribe"))
		iCustID			    	= Trim(Request.Cookies("sfCustomer")("custID"))		
		iPremiumShipping    	= Trim(Request.Form("Shipping"))
		If bCustSubscribed 	= "" Then
			bCustSubscribed 	= 0
		End If
					
		If iCustID = "" Then 
			iCustID = Session("CustID")			
		End If	
		sPassword = getPassword(iCustID) 
		
		If iPhoneFaxType <> "1" Then
			If iCustID = "" Then
				If Trim(Request.Form("Password")) <> ""  Then
					sPassword = Trim(Request.Form("Password"))
					
					' Check if customer already exists
					iCustID = customerAuth(sCustEmail,sPassword,"loose")
					
					If iCustID <> -1 Then						
						Response.Cookies("sfCustomer")("custID") = iCustID
						Response.Cookies("sfCustomer").Expires = Date() + 730
						Call setUpdateCustomer(sCustEmail,sCustFirstName,sCustMiddleInitial,sCustLastName,sCustCompany,sCustAddress1,sCustAddress2,sCustCity,sCustState,sCustZip,sCustCountry,sCustPhone,sCustFax,bCustSubscribed)
					Else
						iCustID = getNewCustomer(sCustEmail,sPassword,sCustFirstName,sCustMiddleInitial,sCustLastName,sCustCompany,sCustAddress1,sCustAddress2,sCustCity,sCustState,sCustZip,sCustCountry,sCustPhone,sCustFax,bCustSubscribed)						
						Response.Cookies("sfCustomer")("custID") = iCustID
						Response.Cookies("sfCustomer").Expires = Date() + 730
					End If
				ElseIf trim(sPassword) <> "" Then
					Response.Cookies("sfCustomer")("custID") = iCustID
					Response.Cookies("sfCustomer").Expires = Date() + 730
					Call setUpdateCustomer(sCustEmail,sCustFirstName,sCustMiddleInitial,sCustLastName,sCustCompany,sCustAddress1,sCustAddress2,sCustCity,sCustState,sCustZip,sCustCountry,sCustPhone,sCustFax,bCustSubscribed)						
				ElseIf Trim(Request.Form("Password")) = "" AND sPassword = "" Then					
					sPassword = generatePassword()
					iCustID = getNewCustomer(sCustEmail,sPassword,sCustFirstName,sCustMiddleInitial,sCustLastName,sCustCompany,sCustAddress1,sCustAddress2,sCustCity,sCustState,sCustZip,sCustCountry,sCustPhone,sCustFax,bCustSubscribed)			
					Response.Cookies("sfCustomer")("custID") = iCustID
					Response.Cookies("sfCustomer").Expires = Date() + 730				
				End If	
				
				' logged in
				Response.Cookies(Session("GeneratedKey") & "sfOrder")("SessionID") = Session("SessionID")
				Response.Cookies(Session("GeneratedKey") & "sfOrder").Expires = Date() + 1
			Else
				Dim bSvdCartCust
				bSvdCartCust = CheckSavedCartCustomer(iCustID)
				
				If trim(Request.Form("Password")) = "" AND sPassword = "" Then
						sPassword = generatePassword()
				End If	
				
				If trim(Request.Form("Password")) <> "" Then
 							sPassword = trim(Request.Form("Password")) 										
				End If		
								
				If bSvdCartCust Then									
					Call setUpdateSvdCustomer(sCustEmail,sPassword,sCustFirstName,sCustMiddleInitial,sCustLastName,sCustCompany,sCustAddress1,sCustAddress2,sCustCity,sCustState,sCustZip,sCustCountry,sCustPhone,sCustFax,bCustSubscribed,iCustID)									
				Else
					If Request.Cookies(Session("GeneratedKey") & "sfOrder")("SessionID") = Session("SessionID") Then
						Call setUpdateCustomer(sCustEmail,sCustFirstName,sCustMiddleInitial,sCustLastName,sCustCompany,sCustAddress1,sCustAddress2,sCustCity,sCustState,sCustZip,sCustCountry,sCustPhone,sCustFax,bCustSubscribed)
					Else				
						iCustID = getNewCustomer(sCustEmail,sPassword,sCustFirstName,sCustMiddleInitial,sCustLastName,sCustCompany,sCustAddress1,sCustAddress2,sCustCity,sCustState,sCustZip,sCustCountry,sCustPhone,sCustFax,bCustSubscribed)			
						Response.Cookies("sfCustomer")("custID") = iCustID
						Response.Cookies("sfCustomer").Expires = Date() + 730
					End If	
										
					' logged in
					Response.Cookies(Session("GeneratedKey") & "sfOrder")("SessionID") = Session("SessionID")
					Response.Cookies(Session("GeneratedKey") & "sfOrder").Expires = Date() + 1	
				End If
			End If	
			
 		End If	
 		
 		If Trim(Request.Form("ShipFirstName")) <> "" or Trim(Request.Form("ShipMiddleInitial"))  <> "" or Trim(Request.Form("ShipLastName"))  <> "" or Trim(Request.Form("ShipCompany"))  <> "" or Trim(Request.Form("ShipAddress1"))  <> "" or Trim(Request.Form("ShipAddress2"))  <> "" or Trim(Request.Form("ShipCity"))  <> "" or Trim(Request.Form("ShipState"))  <> "" or Trim(getNameWithID("sfLocalesState",sShipCustState,"loclstAbbreviation","loclstName",1))  <> "" or Trim(Request.Form("ShipZip"))  <> "" or Trim(Request.Form("ShipCountry"))  <> "" or Trim(getNameWithID("sfLocalesCountry",sShipCustCountry,"loclctryAbbreviation","loclctryName",1))  <> "" or Trim(Request.Form("ShipPhone"))  <> "" or Trim(Request.Form("ShipFax"))  <> "" or Trim(Request.Form("ShipEmail")) <> "" Then
			sShipCustFirstName			= Trim(Request.Form("ShipFirstName"))
			sShipCustMiddleInitial		= Trim(Request.Form("ShipMiddleInitial"))
			sShipCustLastName			= Trim(Request.Form("ShipLastName"))
			sShipCustCompany			= Trim(Request.Form("ShipCompany"))
			sShipCustAddress1			= Trim(Request.Form("ShipAddress1"))
			sShipCustAddress2			= Trim(Request.Form("ShipAddress2"))
			sShipCustCity				= Trim(Request.Form("ShipCity"))
			sShipCustState				= Trim(Request.Form("ShipState"))
			sShipCustStateName		= Trim(getNameWithID("sfLocalesState",sShipCustState,"loclstAbbreviation","loclstName",1))
			sShipCustZip				= Trim(Request.Form("ShipZip"))
			sShipCustCountry			= Trim(Request.Form("ShipCountry"))
			sShipCustCountryName		= Trim(getNameWithID("sfLocalesCountry",sShipCustCountry,"loclctryAbbreviation","loclctryName",1))
			sShipCustPhone				= Trim(Request.Form("ShipPhone"))
			sShipCustFax				= Trim(Request.Form("ShipFax"))
			sShipCustEmail				= Trim(Request.Form("ShipEmail"))
		Else
			sShipCustFirstName			= Trim(sCustFirstName)
			sShipCustMiddleInitial		= Trim(sCustMiddleInitial)
			sShipCustLastName			= Trim(sCustLastName)
			sShipCustCompany			= Trim(sCustCompany)
			sShipCustAddress1			= Trim(sCustAddress1)
			sShipCustAddress2			= Trim(sCustAddress2)
			sShipCustCity				= Trim(sCustCity)
			sShipCustState				= Trim(sCustState)
			sShipCustStateName		= Trim(getNameWithID("sfLocalesState",sShipCustState,"loclstAbbreviation","loclstName",1))
			sShipCustZip				= Trim(sCustZip)
			sShipCustCountry			= Trim(sCustCountry)
			sShipCustCountryName		= Trim(getNameWithID("sfLocalesCountry",sShipCustCountry,"loclctryAbbreviation","loclctryName",1))
			sShipCustPhone				= Trim(sCustPhone)
			sShipCustFax				= Trim(sCustFax)
			sShipCustEmail				= Trim(sCustEmail)	
		End If
        'Response.Write  "Type  " & iPhoneFaxType  
		'Response.End 
		If iPhoneFaxType = 0 Then
				iShipID = CheckShippingChange(sShipCustFirstName,sShipCustMiddleInitial,sShipCustLastName,sShipCustCompany,sShipCustAddress1,sShipCustAddress2,sShipCustCity,sShipCustState,sShipCustZip,sShipCustCountry,sShipCustPhone,sShipCustFax,sShipCustEmail)
				' Update database
				'Response.Write "ShipId  " & iShipId
				If iShipID = 0 Then
				   	iShipID = setShipping(sShipCustFirstName,sShipCustMiddleInitial,sShipCustLastName,sShipCustCompany,sShipCustAddress1,sShipCustAddress2,sShipCustCity,sShipCustState,sShipCustZip,sShipCustCountry,sShipCustPhone,sShipCustFax,sShipCustEmail)
				End If
		End If
		
		sCCList = getCreditCardList()
%>
<title><%= C_STORENAME %>-SF Verification Page/Third Step in Checkout</title>






</head>
<body bgproperties="fixed"  link="<%= C_LINK %>" vlink="<%= C_VLINK %>" alink="<%= C_ALINK %>">
<table border="0" cellpadding="1" cellspacing="0" class="tdbackgrnd" width="600" align="center">
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
        <tr>
          <td align="center" class="tdMiddleTopBanner">Payment Details</td>
        </tr>
        <tr>
          <td class="tdBottomTopBanner"> Please review your order and enter payment 
            details below. When finished click &quot;<b>checkout</b>&quot; to 
            submit your transaction for processing. 
        </tr>
        <tr>
          <td width="100%" align="center" class="tdContent2"  valign="center"><font color="#990000">Step 
            1: Customer Information | <b>Step 2: <i>PAYMENT DETAILS</i></b> | 
            Step 3: Complete Order</font></td>
        </tr>
        <tr>
          <td class="tdContent2" width="100%">        
            <table border="0" width="100%" cellspacing="0" cellpadding="4">
              <tr>
                <td width="52%" class="tdContentBar">product</td>
                <td width="16%" align="center" class="tdContentBar">unit price</td>
                <td width="16%" align="center" class="tdContentBar">qty</td>
                <td width="16%" align="center" class="tdContentBar">price</td>
                </tr>
              <%
'@BEGINCODE
Dim SQL, sProdID, sAttrUnitPrice, sUnitPrice, sProdName, sProdPrice, sProductSubtotal
Dim sBgColor, sFontFace, sFontColor, sTotalPrice, sShipping, dProductSubtotal
Dim iCounter, iSTax, iCTax
Dim iFontSize, iOrderID, iQuantity, iProdAttrNum, iProductCounter 
Dim rsAllOrders, aProduct, aProdAttrID

	

SQL = "SELECT * FROM sfTmpOrderDetails WHERE odrdttmpSessionID = " & Session("SessionID")

Set rsAllOrders = Server.CreateObject("ADODB.RecordSet")
rsAllOrders.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
sTotalShipPrice = 0 
sTotalPrice =0
If  sFreeShip =False And cint(iShip) < 1 Or Not isnumeric(Iship) or not isnull(iShip) Then
 iShip = "1"
 putShipping(iShip)
End if
Do While NOT rsAllOrders.EOF
    
	' Get the ProdIDs
	iOrderID = rsAllOrders.Fields("odrdttmpID")
	sProdID = rsAllOrders.Fields("odrdttmpProductID")
	iQuantity = rsAllOrders.Fields("odrdttmpQuantity")
	iShip = iShip + rsAllOrders.Fields("odrdttmpShipping")
	' Get an array of 3 values from getProduct()
    '++ On Error Resume Next
	ReDim aProduct(3)
		aProduct = getProduct(sProdID)		
		sProdName = aProduct(0)
		sProdPrice = aProduct(1)
		'Verify_SetProdPrice 'SFAE
		iProdAttrNum = aProduct(2)
		
		If iShipped <> 1 Then iShipped = getShipped(sProdID)
	' ++ Call CheckForError()
			
		' If not an array, then the product does not exist 
		If NOT IsArray(aProduct) Then
			Response.Write "<br>Product Does Not Exist"
		Else
				If NOT IsNumeric(iProdAttrNum)Then 
					iProdAttrNum = 0
				End If	
						
				' Get Associated Attribute IDs in an array
				If iProdAttrNum <> "" Then							
					ReDim aProdAttrID(iProdAttrNum)
					aProdAttrID = getProdAttr("odrattrtmp",iOrderID,iProdAttrNum)	
				End If
								
			
				' Response Write all Output for debugging
				If vDebug = 1 And IsArray(aProdAttrID) Then 
					Response.Write "<p>Product = " & sProdID & "<br>ProdName = " & sProdName & "<br>ProdPrice = " & sProdPrice & "<br>ProdAttrNum = " & iProdAttrNum
						
					For iCounter = 0 To iProdAttrNum - 1 
						Response.Write "<br>Attribute :" & aProdAttrID(iCounter)
					Next			
					
				End If	 
				
				iProductCounter = iProductCounter + 1
				dim fontclass
				' Do alternating colors and fonts	
				If (iProductCounter mod 2) = 1 Then 
							fontclass="tdAltFont1"
				Else 	
							fontclass="tdAltFont2"
				End If	
		
		%>
              <tr class='<%=fontclass%>'>
                <td width="52%" valign="top" background=""><b><%= sProdName %></b><br>
                  <%
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
			%>										                
            &nbsp;&nbsp;<%=sAttrName%>
                  <br>                									
			      <%		
					' ProdAttr Loop
					Next
				Elseif iProdAttrNum > 0 And NOT IsArray(aProdAttrID) Then 
					Response.Write "<br>Error: No Attributes found for " & iOrderID
					Response.Write "<br>Deleting from Saved Orders. Sorry for the inconvenience."
													
					Call setDeleteOrder("odrdttmp",iOrderID)
					If vDebug = 1 Then Response.Write "<p><font color=""red""> Deleted: " & iOrderID & "</font>"						
				' End Product Attribute If
				End If	
			    
			    dUnitPrice = cDbl(sAttrUnitPrice) + cDbl(sProdPrice)
				Verify_SetProdPrice 'SFAE
					    
				If Trim(iShip) => 1 and getShipped(sProdID) <> "0" Then		
					' Set Unit Price for Product
					If iConverion = 1 Then
						sUnitPrice = "<script> document.write(""" & FormatCurrency(dUnitPrice) & " = ("" + OANDAconvert(" & dUnitPrice & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script>"  'yaz
					Else
						sUnitPrice = FormatCurrency(dUnitPrice)
					End If
					dProductSubtotal = iQuantity * (dUnitPrice)
				
					If iConverion = 1 Then
						sProductSubtotal = "<script> document.write(""" & FormatCurrency(dProductSubtotal) & " = ("" + OANDAconvert(" & dProductSubtotal & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script>"  'yaz
					Else
						sProductSubtotal = FormatCurrency(dProductSubtotal)
					End If
									
					sTotalShipPrice = sTotalShipPrice + cDbl(dProductSubtotal)
					'sTotalPrice = cdbl(sTotalPrice) + cDbl(dProductSubtotal) 'B2 se/ae fix
					
'				ElseIf Trim(iShip) = 0 Then
				else
				
					' Set Unit Price for Product
					If iConverion = 1 Then
						sUnitPrice = "<script> document.write(""" & FormatCurrency(dUnitPrice) & " = ("" + OANDAconvert(" & dUnitPrice & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script>"  'yaz
					Else
						sUnitPrice = FormatCurrency(dUnitPrice)
					End If
					dProductSubtotal = iQuantity * (dUnitPrice)
					If iConverion = 1 Then
						sProductSubtotal = "<script> document.write(""" & FormatCurrency(dProductSubtotal) & " = ("" + OANDAconvert(" & dProductSubtotal & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script>"  'yaz
					Else
						sProductSubtotal = FormatCurrency(dProductSubtotal)
					End If

					'sTotalPrice = cdbl(sTotalPrice) + cDbl(dProductSubtotal) 'SFUPDATE
				End If 				
				
			%>
                </td>
                <td width="16%" align="center" class='<%=fontClass%>' valign="top" nowrap background=""><%= sUnitPrice %></td>
			    <td width="16%" align="center" class='<%=fontClass%>' valign="top" nowrap background=""><%= iQuantity %></td>
			    <td width="16%" align="center" class='<%=fontClass%>' valign="top" nowrap background=""><%= sProductSubtotal %></td>
                </tr>
                <%
                OVC_ShowGiftWrapValue 2 'SFAE
		        OVC_ShowBackOrderMessage 2'SFAE 
		' End IsArray If
		End If
		
		sTotalPrice = cdbl(sTotalPrice) + cDbl(dProductSubtotal) 'SFUPDATE
		
		
	rsAllOrders.MoveNext		
	Loop

	OVC_SaveSubTotalWOD 'SFAE
	
	If Application("AppName") = "StoreFront" Then 'SFUPDATE
		sTotalPrice = getGlobalSalePrice(sTotalPrice) 
		sShipping = getShipping(sTotalShipPrice, iPremiumShipping, "", sShipCustCity,sShipCustState,sShipCustZip, sShipCustCountry, sTotalPrice,"All")' SFUPDATE
	End If
	
	IF Application("AppName") = "StoreFrontAE" Then 'SFAE
		sTotalPrice =  ApplyALLDiscounts(cdbl(sTotalPrice),"Total") '.3008
		SetBillingVariables
		If Session("SpecialBilling") =0  then
    	   if bFromVerify = false then	
	    	   sShipping = getShipping(sTotalShipPrice, iPremiumShipping, "", sShipCustCity,sShipCustState,sShipCustZip, sShipCustCountry, sTotalPrice,"All")
	       else
		     bLtl = true
		     sShipping =  session("sLTL")
		     arrLTL = split(sShipping,"|")
	         sShipping = Request.Form("ltlrate")
	         sShipMethodName = Request.Form("LTLCarrier")
	         sShipping = CDbl(sShipping)
	         sLTL = arrltl(1)
		     
		   end if 
		Else	
			if instr(1,sShipMethodName,"LTL") then
			
				Session("LTLIndex") = Request.QueryString("OptionID")
				sLTL=getShipping(sTotalShipPrice, iPremiumShipping, "", sShipCustCity,sShipCustState,sShipCustZip, sShipCustCountry, sTotalPrice,"All")
				
				
				If left(sLTL,1) = "@" then 
	              	bLtl=True
		          	sLTL =  mid(sLTL,2,len(sLTL))
			       	arrLTL = split(sLTL,"|")
				    sLTL = arrLTL(1)
		 		end if 
		 		
				If left( Session("BackOrderShipping"),1) = "@" then 
	              	bLtl=True
		          	Session("BackOrderShipping") =  mid(Session("BackOrderShipping"),2,len(Session("BackOrderShipping")))
			       	arrLTL = split(Session("BackOrderShipping"),"|")
				    Session("BackOrderShipping") = arrLTL(0)
				end if   
				If left( Session("BillShipping"),1) = "@" then    
			 	    Session("BillShipping") = mid(Session("BillShipping"),2,len(Session("BillShipping")))
			 	    arrLTL = split(Session("BillShipping"),"|")
			        Session("BillShipping") = arrLTL(0)
				 end if  
				 
				Session("LTLIndex") = ""
				
			ELSE

			Session("BackOrderShipping") = getShipping(sTotalShipPrice, iPremiumShipping, "", sShipCustCity,sShipCustState,sShipCustZip, sShipCustCountry, sTotalPrice,"BackOrder")
			Session("BillShipping") = getShipping(sTotalShipPrice, iPremiumShipping, "", sShipCustCity,sShipCustState,sShipCustZip, sShipCustCountry, sTotalPrice,"Shipped")
			END IF
	
			If isnumeric(Session("BillShipping")) AND  isnumeric(Session("BackOrderShipping")) Then
				sShipping = cdbl(Session("BillShipping")) + cdbl(Session("BackOrderShipping"))
			Else
				sShipping = 0 
			End If
	
		End If
  	End If
	
	IF Application("AppName") = "StoreFrontAE" Then 'SFAE
		'LTL stuff
	  if bFromVerify = false then
		 If left(sShipping,1) = "@" then 
		    bLtl = true
		    sShipping = mid(sShipping,2,len(sShipping))
		    arrLTL = split(sShipping,"|")
		    sShipping = arrltl(0)
		    sShipping = CDbl(sShipping)
		    sLTL = arrltl(1)
		 end if
		end if	
	end if	
	  
    If isNumeric(sShipping) Then 'SFUPDATE
		Session("sShipping") = cdbl(sShipping) 
	Else
		Session("sShipping") = 0
	End If
	
	Session("sShipping") = cdbl(sShipping) 'SFUPDATE
	putShipping(sShipping)
	
	
	If iHandling <> 1 Or iShipped <> 1 Then sHandling = 0
	
 	'#383
	 dTaxAble_Amount =  sTotalPrice
	
		'#383
			iSTax = iSTax + cDbl(getTax("State", sShipping, dTaxAble_Amount, sProdID))
			iCTax = iCTax + cDbl(getTax("Country", sShipping, dTaxAble_Amount, sProdID))
  
 
	'#383
		sGrandTotal = (cDbl(sTotalPrice) + cDbl(sHandling) + cDbl(sShipping) + cDbl(iSTax) + cDbl(iCTax) + cDbl(iCodAmount))
		
	
        sTotalPriceOut =  cdbl(sTotalPrice) 
		sGrandTotalOut = cdbl(sGrandTotal)

	
closeObj(rsAllOrders)
closeObj(rsAdminInfo)
'djp ltl 8-25-01


'Response.Write sSubmitAction
'Response.End 		
%>

              </table>
            </td></tr>
          <tr><td class="tdContent2">
		      <table border="0" width="100%" cellspacing="0" cellpadding="2" class="tdContent2">
		      <%OVC_ShowOrderDiscounts 'SFAE%>			
		            <tr>
		            <td width="75%" align="right"><b>Sub Total:</b></td>
		            <td width="25%" height="20" nowrap><b>
            <%
        
		If iConverion = 1 Then
			Response.Write "<script> document.write(""" & FormatCurrency(sTotalPriceOut) & " = ("" + OANDAconvert(" & sTotalPriceOut & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script>"  'yaz
		Else
			Response.write FormatCurrency(sTotalPriceOut)
		End If
		%>
		              </b></td>
		            <td width="5%" align="center"></td>
		          </tr>
		          <% If iHandling = 1 And iShipped = 1 Then %>
		          <tr>
		            <td width="75%" align="right">Handling:</td>
		            <td width="25%" height="20" nowrap>
			          <%
			If iConverion = 1 Then
				Response.Write "<script> document.write(""" & FormatCurrency(sHandling) & " = ("" + OANDAconvert(" & sHandling & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script>"  'yaz
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
		          <% If bCod Then %>
		          <tr>
		            <td width="75%" align="right">COD Charge:</td>
		            <td width="25%" height="20" nowrap>
			          <%
			If iConverion = 1 Then
				Response.Write "<script> document.write(""" & FormatCurrency(iCodAmount) & " = ("" + OANDAconvert(" & iCodAmount & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script>"  'yaz
			Else
				Response.write FormatCurrency(iCodAmount)
			End If
			%>
		             </td>
		            <td width="5%" align="center"></td>
		          </tr>
		          <% 
		Else
			iCodAmount = 0
		End If 
		%>
		<% 	

		If sShipping <> 0 OR iShipMethod = 19 or sFreeShip =True Then 
		         if bltl = false then
		           %>  
		            <tr>
		            <td width="75%" align="right"><%= sShipMethodName %>:</td>
		            <td width="25%" height="20" nowrap>
		          <%
	              	If iConverion = 1  Then
		            	Response.Write "<script> document.write(""" & FormatCurrency(sShipping) & " = ("" + OANDAconvert(" & sShipping & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script>"  'yaz
		            Else
			        Response.write FormatCurrency(sShipping)
		            End If
		%>		
		              </td>
		            <td width="5%" align="center"></td>
		          </tr>
		          <% 
		        else
		           Response.write sLTL
'		          If Session("SpecialBilling")<> 0  then
		           if bFromVerify = true then
		            dim sReset
		            sReset =   "<script id=callreset language='javascript'>" & CHR(13) & vbCrlf
		            sReset = sReset & "<!--" & CHR(13) & vbCrlf
		            sReset = sReset & "   resetme(" & Request.QueryString("OptionID") & ")" & vbCrlf 
		            sReset = sReset & "//-->" & CHR(13) & vbCrlf
	                sReset = sReset & "</script>" & vbCrlf 
		            Response.Write sReset
		           
		           end if
'		          End if 
		        end if    
		    End If %>

		          <% If iCTax <> 0 Then %>
		          <tr>
		            <td width="75%" align="right">Country Tax:</td>
		            <td width="25%" height="20" nowrap>
			          <%
			If iConverion = 1 Then
				Response.Write "<script> document.write(""" & FormatCurrency(iCTax) & " = ("" + OANDAconvert(" & iCTax & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script>"  'yaz
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
		            <td width="25%" height="20" nowrap>
			          <%
			If iConverion = 1 Then
				Response.Write "<script> document.write(""" & FormatCurrency(iSTax) & " = ("" + OANDAconvert(" & iSTax & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script>"  'yaz
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
		            <td width="25%" height="20" nowrap><b>
		              <%
		If iConverion = 1 Then
			'Response.Write "<script> document.write(""" & FormatCurrency(cDbl(sTotalPrice) + cDbl(sHandling) + cDbl(sShipping) + cDbl(iSTax) + cDbl(iCTax)+ cDbl(iCodAmount)) & " = ("" + OANDAconvert(" & cDbl(sTotalPrice) + cDbl(sHandling) + cDbl(sShipping) + cDbl(iSTax) + cDbl(iCTax) + cDbl(iCodAmount) & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script>"  'yaz
			Response.Write "<script> document.write(""" & FormatCurrency(sGrandTotalOut) & " = ("" + OANDAconvert(" & sGrandTotalOut & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script>"  'SFUPDATE
		Else
			'Response.write FormatCurrency(cDbl(sTotalPrice) + cDbl(sHandling) + cDbl(sShipping) + cDbl(iSTax) + cDbl(iCTax) + cDbl(iCodAmount))
			Response.write FormatCurrency(sGrandTotalOut) 'SFUPDATE
		End If
		
		%>	
		              </b></td>
		            <td width="5%" align="center"></td>
		          </tr>
	<% 	If iShipType = 2 Then %>	          
		          <tr>
		         
		          <td width="80%" colspan="4"><table><tr>
		           <td><% If InStr(sShipMethodName,"UPS") Then %><img src="images/logo_s.gif"><% Else %>&nbsp;<% End If %></td>
                      <td><font class="Content_Small"> </font></td>
                    </tr>
</table ></td>
</tr>
<% End If %>
		        </table>
              </td></tr>

            <!--Customer Information -->
            <tr><td width="100%" class="tdContent2">   
				<HIDE_SF_PRESERVE><form method="post" name="frmVerify" action="confirm.asp" onSubmit="javascript: <%=sSubmitAction %>"></HIDE_SF_PRESERVE>
           <table border="0" width="100%" cellspacing="0" cellpadding="2">
	              <tr>
		            <td colspan="2" width="100%" class="tdContentBar">Customer Information</td>
	              </tr>
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
                  <%If sCustAddress2 <> "" or sShipCustAddress2 <>"" Then%>
   	              <tr>	
   	              <%If sCustAddress2 <> "" Then%>
   		            <td><%= sCustAddress2 %></td>
   		          <%else%>
   		          <td></td>
   		          <%End If%>
   		          <%If sShipCustAddress2 <> "" Then%>
                    <td><%= sShipCustAddress2%></td>
                  <%else%>
   		          <td></td>
                  <%End If%>
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
		            <td><% If sCustPhone <> "" Then %>Phone: <%= sCustPhone%><% Else %>&nbsp;<% End If %></td>
		            <td><% If sShipCustPhone <> "" Then %>Phone: <%= sShipCustPhone %><% Else %>&nbsp;<% End If %></td>
	              </tr>
								<% If sCustFax <> "" Or sShipCustFax <> "" Then %>
	              <tr>
		            <td><% If sCustFax <> "" Then %>Fax: <%= sCustFax%><% Else %>&nbsp;<% End If %></td>
		            <td><% If sShipCustFax <> "" Then %>Fax: <%= sShipCustFax %><% Else %>&nbsp;<% End If %></td>
	              </tr>
								<% End If %>
	              <tr>
		            <td><%= sCustEmail%></td>
    	            <td><%= sShipCustEmail %></td>
                  </tr>
                  <%If sShipping <> 0 OR iShipMethod = 19 or sFreeShip =True Then %>
	              <tr>
		            <td></td>
    	            <td><b>Ship Via:</b> <%= sShipMethodName %></td>
                  </tr>
                  <%end if%>
                </table>
              </td></tr>
            <!-- Payment Selection -->  
            <% If ((sTransMethod <> "15" or sTransMethod <> "18") AND (sPaymentMethod <> "Credit Card" OR sPaymentMethod <> "PayPal")) Then %>  
                <!-- Start Payment Selection Info -->
           <tr>
           <td width="100%" class="tdContent2"> 
		    	             
                  <%If (sPaymentMethod = "Credit Card" AND sTransMethod <> "15" AND sTransMethod <> "18") Then %>
			      
			      

                 <table class="tdContent2" border="0" width="100%" cellpadding="2" cellspacing="0">
			        <tr>
			          <td width="100%" colspan="2">
			            <table border="0" width="100%" cellpadding="2" cellspacing="0" class="tdContentBar">
			              <tr>
			                <td width="100%" class="tdContentBar"> Credit Card Information</td>
			              </tr>
			            </table>
			          </td>
			        </tr>
			        <tr>
			          <td width="50%"><b>Card Type</b><font color="#FF0000">*</font><b>:</b></td>
			        <td width="50%"><b>Name as it appears on card<font color="#FF0000">*</font>:</b></td>
			        </tr>
			        <tr>
			          <td width="50%"><select name="CardType" title="Credit Card Type" style="<%= C_FORMDESIGN %>"><%= sCCList %></select>
			          </td>
			          <td width="50%"><input type="text" name="CardName" title="Name on Card" size="30" Style="<%= C_FORMDESIGN%>"></td>
			        </tr>
			        <tr>
			          
                <td width="50%"><b>Card Number<font color="#FF0000">*</font>:</b></td>
			          
                <td width="50%"><b>Expiration Date (e.g. 11 - 2005)<font color="#FF0000">*</font>:</b></td>
			        </tr>
			        <tr>
			          <td width="50%"><input type="text" name="CardNumber" title="Credit Card Number" size="30" style="<%= C_FORMDESIGN%>"></td>
			          
                <td width="50%">Month 
                  <select name=cc-exp-month>
<option value="01"> 01<option value="02"> 02<option value="03"> 03<option value="04"> 04<option value="05"> 05<option value="06"> 06<option value="07"> 07<option value="08"> 08<option value="09"> 09<option value="10"> 10<option value="11"> 11<option value="12"> 12
</select>
                  Year 
                  <select name=cc-exp-year>
<option value="2002"> 2002<option value="2003"> 2003<option value="2004"> 2004<option value="2005"> 2005<option value="2006"> 2006<option value="2007"> 2007<option value="2008"> 2008<option value="2009"> 2009<option value="2010"> 2010<option value="2011"> 2011<option value="2012"> 2012<option value="2013"> 2013<option value="2014"> 2014<option value="2015"> 2015<option value="2016"> 2016<option value="2017"> 2017<option value="2018"> 2018
</select>
                </td>
			        </tr>
			        <tr><td colspan="2" height="20"></td></tr>
			      </table>
<%ElseIf sPaymentMethod = "PhoneFax" Then%>	
			      <table class="tdContent2" border="0" width="100%" cellpadding="2" cellspacing="0">
			        <tr>
			          <td width="100%" colspan="2">
			            <table border="0" width="100%" cellpadding="2" cellspacing="0" class="tdContentBar">
			              <tr>
			                <td width="100%" class="tdContentBar">Phone Fax Order Information</td>
			              </tr>
			              <tr>
			              <td colspan="2" align="center" valign="center">
          			        <p style="margin-top:20pt">
		        	        Phone Fax method
			                <p style="margin-top:20pt">
			                </td>
			            </tr>
			            </table>
			          </td>
			        </tr>
			        <tr><td height="20">&nbsp;</td></tr>
			        <tr><td align="center">
				        <table width="80%">
				        <% if CheckPaymentMethod("Credit Card")=1 then%>
				          <tr><td colspan="2" align="center"><b><font class="ECheck">Complete this section for Credit Card purchases
					          <hr width="90%" size="1" noshade class="tdAltBG1"></font></b></td>
				            </tr>
				            <tr>
				              <td width="50%"><b>Card Type</b><font color="#FF0000">*</font>:</td>
				            <td width="50%"> <b>Name as it appears on card<font color="#FF0000">*</font>:</b></td>
				            </tr>
				            <tr>
				              <td width="50%"> <select name="CardType" title="Credit Card Type" style="<%= C_FORMDESIGN %>"><option></option><%= sCCList %></select>
				            </td>
				            <td width="50%"><input type="text" name="CardName" title="Name on Card" size="30" Style="<%= C_FORMDESIGN%>"></td>
				          </tr>
				          <tr>
				            <td width="50%"><b>Card Number<font color="#FF0000">*</font>:</b></td>
				            <td width="50%"><b>Expiration Date<font color="#FF0000">*</font>:</b></td>
				          </tr>
				          <tr>
				            <td width="50%"><input type="text" name="CardNumber" title="Credit Card Number" size="30" style="<%= C_FORMDESIGN%>"></td>
				            <td width="50%"> <b>Month</b> <input type="text" size="2" name="CardExpiryMonth" title="Credit Card Month" Style="<%= C_FORMDESIGN%>">   
				              <b>Year</b> <input type="text" size="4" name="CardExpiryYear" title="Credit Card Year" Style="<%= C_FORMDESIGN%>"></td>
				        </tr>
				        <%end if
				        if CheckPaymentMethod("eCheck")=1 then
				        %>
				        <tr><td height="40">&nbsp;</td></tr>
			
				        <tr><td colspan="2" align="center"><b><font class="ECheck">Complete this section for eCheck purchases
					        <hr width="90%" size="1" noshade class="tdAltBG1"></font></b></td>
				          </tr>
				          <tr>
					        <td align="left"><b><font class="ECheck">Check Number:</font></b></td>
					        <td align="left"><b><font class="ECheck">Bank Name <font color="#FF0000">*</font>:</font></b></td>
				          </tr>
				          <tr>
					        <td align="left"><font class="ECheck"><input type="Text" name="CheckNumber" title="Check Number" size="30" style="<%= C_FORMDESIGN%>"></font></td>
					    <td align="left"><font class="ECheck"><input type="Text" name="BankName" title="Bank Name" size="30" style="<%= C_FORMDESIGN%>"></font></td>
				  </tr>
				  <tr>
					<td align="left"><font class="ECheck"><b>Bank Routing Number <font color="#FF0000">*</font>:</b></font></td>
					<td align="left"><font class="ECheck"><b>Checking Account Number <font color="#FF0000">*</font>:</b></font></td>
							
				  </tr>
				  <tr>
					<td align="left"><font class="ECheck"><input type="Text" name="RoutingNumber" title="Routing Number" size="30" style="<%= C_FORMDESIGN%>"></font></td>		
					<td align="left"><font class="ECheck"><input type="Text" name="CheckingAccountNumber" title="Checking Account Number" size="30" style="<%= C_FORMDESIGN%>"></font></td>
				</tr>
				<%
				 end if
				 if CheckPaymentMethod("PO")=1 then
				%>
				<tr>
					<td colspan="2" height="40">&nbsp;</td>
				</tr>	
				<tr><td colspan="2" align="center"><b><font class="ECheck">Complete this section for Purchase Order purchases
					<hr width="90%" size="1" noshade class="tdAltBG1"></font></b></td>
				  </tr>
				  <tr>	
				    <td colspan="2" align="center" valign="center">
				      <b>Purchase Order Name <font color="#FF0000">*</font>:</b> 
				      <br><input type="text" size="25" name="POName" title="PO Name" style="<%= C_FORMDESIGN %>"></td>
			
				  <tr>
				    <td colspan="2" align="center" valign="center">
				      <p style="margin-top:10pt"><b>PO Purchase Number <font color="#FF0000">*</font>:</b>  
				      <br><input type="text" size="25" name="PONumber" title="PO Number" style="<%= C_FORMDESIGN %>"></td>
				
				    </tr>
				    <tr>
				      <td colspan="2"><p style="margin-top:10pt">&nbsp;</td></tr>
				    </table>
			      </td>	
			</tr>	
			<%end if %>
			</table>
        <%ElseIf sPaymentMethod = "eCheck" Then%>				
		<table class="tdContent2" border="0" width="100%" cellpadding="4" cellspacing="0">
			<tr><td width="100%" colspan="2">
			    <table border="0" width="100%" cellspacing="0" class="tdContentBar">
			      <tr>
			        <td width="100%" colspan="2" class="tdContentBar">eCheck Information</td>
			      </tr>
			    </table>
			  </td></tr>
			<tr>
				<td align="left"><b><font class="ECheck">Check Number <font color="#FF0000">*</font>:</font></b></td>
				<td align="left"><b><font class="ECheck">Bank Name <font color="#FF0000">*</font>:</font></b></td>
			</tr>
			<tr>
				
				<td align="left"><font class="ECheck"><input type="Text" name="CheckNumber" title="Check Number" size="30" style="<%= C_FORMDESIGN%>"></font></td>
				<td align="left"><font class="ECheck"><input type="Text" name="BankName" title="Bank Name" size="30" style="<%= C_FORMDESIGN%>"></font></td>
			</tr>
			<tr>
				<td align="left"><font class="ECheck"><b>Bank Routing Number <font color="#FF0000">*</font>:</b></font></td>
				<td align="left"><font class="ECheck"><b>Checking Account Number <font color="#FF0000">*</font>:</b></font></td>
						
			</tr>
			<tr>
				<td align="left"><font class="ECheck"><input type="Text" name="RoutingNumber" title="Routing Number" size="30" style="<%= C_FORMDESIGN%>"></font></td>		
				<td align="left"><font class="ECheck"><input type="Text" name="CheckingAccountNumber" title="Checking Account Number" size="30" style="<%= C_FORMDESIGN%>"></font></td>
			</tr>
			<tr>
				<td colspan="2" height="20"></td>
			</tr>	
			</table>  
    <%ElseIf sPaymentMethod = "PO" Then %>
		<table class="tdContent2" border="0" width="100%" cellpadding="2" cellspacing="0">
		  <tr>
			<td colspan="2" width="100%" class="tdContentBar">Purchase Order Payment Information</td>
		  </tr>
		  <tr>	
			<td colspan="2" align="center" valign="center">
			  <p style="margin-top:20pt"><b>Purchase Order Name <font color="#FF0000">*</font>:</b>
			  <br><input type="text" size="25" name="POName" title="PO Name" style="<%= C_FORMDESIGN %>" value="<%= sCustFirstName & " " & sCustLastName %>"></td>
		    </tr>	
		    <tr>
			  <td colspan="2" align="center" valign="center">
			    <p style="margin-top:10pt"><b>PO Purchase Number <font color="#FF0000">*</font>:</b>  
			    <br><input type="text" size="25" name="PONumber" title="PO Number" style="<%= C_FORMDESIGN %>"></td>
			
		      </tr>
		      <tr>
			    <td colspan="2"><p style="margin-top:10pt">&nbsp;</td>
		        </tr>
		      </table>  			            
              <%ElseIf sPaymentMethod = "COD" Then %>
		      <table class="tdContent2" border="0" width="100%" cellpadding="2" cellspacing="0">
		        <tr>
			      <td colspan="2" width="100%" class="tdContentBar">COD Payment Method</td>
		        </tr>
		        <tr>	
			      <td colspan="2" align="center" valign="center">
			        <p style="margin-top:20pt">
			        COD payment method
			        <p style="margin-top:20pt">
			        </td>
		            </tr>
		          </table>    
                  <%End If%>
				</td></tr>       
              <% End If %>
              <!--Shipping--><!--Special Instructions-->
                <tr><td width="100%" class="tdContent2">   
	                <table border="0" width="100%" cellspacing="0" cellpadding="2">
	                  <tr>
	                    <td width="100%" colspan="2" class="tdContentBar">Special Instructions</td>
	                  </tr>           
                      <%If sInstructions = "" Then %>
                      <tr><td width="100%" colspan="2" align="center" height="40" valign="center"><i>None Specified</i></td></tr>
                      <% Else 
                      'security chk
                        sInstructions		= Replace(sInstructions,"<"," " )
						sInstructions		= Replace(sInstructions,">"," " )
						sInstructions		= Replace(sInstructions,chr(34)," " )
						sInstructions		= Replace(sInstructions,"'"," " ) ' #307
	
                      %>  
                         
		              <tr><td width="100%" colspan="2" align="left" height="40" valign="center"><font class="ECheck"><%= sInstructions %></font></td></tr>    
                      <% End If %>
	
                      <tr><td height="80" valign="top" align="center" colspan="2">
		                  <input type="hidden" name="FirstName" value="<%= sCustFirstName %>">
		                  <input type="hidden" name="MiddleInitial" value="<%= sCustMiddleInitial %>">
		                  <input type="hidden" name="LastName" value="<%= sCustLastName %>">
		                  <input type="hidden" name="Company" value="<%= sCustCompany %>">
		                  <input type="hidden" name="Address1" value="<%= sCustAddress1 %>">
		                  <input type="hidden" name="Address2" value="<%= sCustAddress2%>">
		                  <input type="hidden" name="City" value="<%= sCustCity %>">
		                  <input type="hidden" name="State" value="<%= sCustState %>">
		                  <input type="hidden" name="StateName" value="<%= sCustStateName %>">
		                  <input type="hidden" name="Zip" value="<%= sCustZip %>">
		                  <input type="hidden" name="Country" value="<%= sCustCountry %>">
		                  <input type="hidden" name="CountryName" value="<%= sCustCountryName %>">
		                  <input type="hidden" name="Phone" value="<%= sCustPhone %>">
		                  <input type="hidden" name="Fax" value="<%= sCustFax %>">
		                  <input type="hidden" name="Email" value="<%= sCustEmail %>">
		                  <input type="hidden" name="IsSubscribed" value="<%= bCustSubscribed %>">
		                  <input type="hidden" name="Password" value="<%= sPassword %>">

		                  <input type="hidden" name="PaymentMethod" value="<%= sPaymentMethod %>">
		                  <input type="hidden" name="ShipType" value="<%= iShipType %>">	
		                  <input type="hidden" name="ShipMethod" value="<%= iShipMethod %>">	
		                  <input type="hidden" name="ShipMethodName" value="<%= sShipMethodName %>">
		
		                  <input type="hidden" name="ShipID" value="<%= iShipID %>">
		                  <input type="hidden" name="ShipFirstName" value="<%= sShipCustFirstName %>">
		                  <input type="hidden" name="ShipMiddleInitial" value="<%= sShipCustMiddleInitial %>">
		                  <input type="hidden" name="ShipLastName" value="<%= sShipCustLastName %>">
		                  <input type="hidden" name="ShipCompany" value="<%= sShipCustCompany %>">
		                  <input type="hidden" name="ShipAddress1" value="<%= sShipCustAddress1 %>">
		                  <input type="hidden" name="ShipAddress2" value="<%= sShipCustAddress2 %>">
		                  <input type="hidden" name="ShipCity" value="<%= sShipCustCity %>">
		                  <input type="hidden" name="ShipState" value="<%= sShipCustState %>">
		                  <input type="hidden" name="ShipStateName" value="<%= sShipCustStateName %>">
		                  <input type="hidden" name="ShipZip" value="<%= sShipCustZip %>">
		                  <input type="hidden" name="ShipCountry" value="<%= sShipCustCountry %>">
		                  <input type="hidden" name="ShipCountryName" value="<%= sShipCustCountryName %>">
		                  <input type="hidden" name="ShipPhone" value="<%= sShipCustPhone %>">
		                  <input type="hidden" name="ShipFax" value="<%= sShipCustFax %>">
		                  <input type="hidden" name="ShipEmail" value="<%= sShipCustEmail %>">	
		                  <input type="hidden" name="SpecialInstructions" value="<%= sInstructions %>">				
	
                          <br>    
                          <% If sPaymentMethod <> "InternetCash" OR  (sTransMethod <> "15" AND sPaymentMethod <> "Credit Card")Then %>
		                  <input type="image" src="<%= C_BTN05 %>" border="0" name="verify">
	                      <% ElseIf sPaymentMethod = "InternetCash" Then %>		
		                  <font class="ECheck">You are using <b>InternetCash</b> to pay for your purchase, please enter payment information in the popup window and press continue.</font>
                     
                          <% End If %>                        
                      </td>
                      </tr> 

                  </table>
 <HIDE_SF_PRESERVE></form></HIDE_SF_PRESERVE> 
                                 </td>
                      </tr> 

<%if bltl = false then%>                 
              <!--#include file="foot.txt"-->
<%end if%>              
               </table>
       
   			     
				              </td></tr>
<%if bltl = true then%>                 
              <!--#include file="foot.txt"-->
<%end if%>              

       
            </table>
       </body>
</html>
<%

closeObj(cnn)
%>








