<% 

'@BEGINVERSIONINFO

'@APPVERSION: 50.4011.0.2

'@FILENAME: incAdmin.asp
	 

'

'@DESCRIPTION: Multiple administrative functions 

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000, 2001 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT 

'@ENDVERSIONINFO

	'-------------------------------------------------------------------
	' Sub to update admin table 	
	'-------------------------------------------------------------------
	Sub setUpdateAdminPayment(sPaymentType,sPaymentServerPath,sPaymentLogin,sPaymentPassword,sMerchantType,bCardEncode,bCreditCard,bECheck,bCOD,sCODAmount,bPhoneFaxRecorded,bPhoneFaxNon_Recorded,bPO,bPP) 'JF 9/25/01
		Dim sLocalSQL, rsAdmin, aAdmin, rsTransTypes, rsTransMthd
		
		' Update admin
		sLocalSQL = "SELECT adminTransMethod,adminPaymentServer,adminEncodeCCIsActive,adminLogin,adminPassword,adminMerchantType,adminCODAmount FROM sfAdmin"
		Set rsAdmin	= Server.CreateObject("ADODB.RecordSet")
			rsAdmin.Open sLocalSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText
			rsAdmin.Fields("adminTransMethod")		= sPaymentType
			rsAdmin.Fields("adminPaymentServer")	= sPaymentServerPath
			rsAdmin.Fields("adminEncodeCCIsActive") = bCardEncode
			rsAdmin.Fields("adminLogin")			= sPaymentLogin
			rsAdmin.Fields("adminPassword")			= sPaymentPassword
			rsAdmin.Fields("adminMerchantType")		= sMerchantType
			rsAdmin.Fields("adminCODAmount")		= sCODAmount
			rsAdmin.Update		
		closeobj(rsAdmin)
		
		' Update trans method table
		sLocalSQL = "SELECT trnsmthdServerPath, trnsmthdLogin, trnsmthdPasswd FROM sfTransactionMethods WHERE trnsmthdID =" & sPaymentType
		Set rsTransMthd = Server.CreateObject("ADODB.RecordSet")
			rsTransMthd.Open sLocalSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText
			
			If NOT rsTransMthd.EOF Then
					rsTransMthd.Fields("trnsmthdServerPath") = sPaymentServerPath
					rsTransMthd.Fields("trnsmthdLogin") = sPaymentLogin
					rsTransMthd.Fields("trnsmthdPasswd") = sPaymentPassword
					rsTransMthd.Update				
			End If	
			
			
			'JF 9/25/01 removed Internet Cash stuff
			closeobj(rsTransMthd)
			
			
		' update trans type table	 
		sLocalSQL = "SELECT transType, transName, transIsActive FROM sfTransactionTypes"
		Set rsTransTypes = Server.CreateObject("ADODB.RecordSet")
			rsTransTypes.Open sLocalSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText
			
			Do While NOT rsTransTypes.EOF
				If rsTransTypes.Fields("transType")	= "Credit Card" Then
					If bCreditCard = 1 Then
						rsTransTypes.Fields("transIsActive") = 1
					Else
						rsTransTypes.Fields("transIsActive") = 0	
					End If	
				ElseIf rsTransTypes.Fields("transType")	= "eCheck" Then
					If bECheck = 1 Then
						rsTransTypes.Fields("transIsActive") = 1
					Else
						rsTransTypes.Fields("transIsActive") = 0	
					End If	
				ElseIf rsTransTypes.Fields("transType")	= "COD" Then
					If bCOD = 1 Then
						rsTransTypes.Fields("transIsActive") = 1
					Else
						rsTransTypes.Fields("transIsActive") = 0	
					End If	
				ElseIf rsTransTypes.Fields("transType")	= "PO" Then
					If bPO = 1 Then
						rsTransTypes.Fields("transIsActive") = 1
					Else 
						rsTransTypes.Fields("transIsActive") = 0	
					End If	
				
				'JF 9/25/01
				ElseIf rsTransTypes.Fields("transType")	= "PayPal" Then
					If bPP = 1 Then
						rsTransTypes.Fields("transIsActive") = 1
					Else 
						rsTransTypes.Fields("transIsActive") = 0	
					End If
						
				ElseIf rsTransTypes.Fields("transType")	= "PhoneFax" Then
					If rsTransTypes.Fields("transName")	= "Recorded" Then
						If bPhoneFaxRecorded = 1  Then
							rsTransTypes.Fields("transIsActive")= 1
						Else
							rsTransTypes.Fields("transIsActive")= 0	
						End If	
					ElseIf rsTransTypes.Fields("transName") = "Non-Recorded" Then
						If bPhoneFaxNon_Recorded = 1  Then
							rsTransTypes.Fields("transIsActive")= 1
						Else
							rsTransTypes.Fields("transIsActive")= 0	
						End If	
					End If
				'JF 9/25/01 removed Internet Cash stuff
				End If
				
				rsTransTypes.Update
				rsTransTypes.MoveNext				
			Loop
			closeobj(rsTransTypes)
	End Sub
	
	'-------------------------------------------------------------------
	' Function to retrieve active payment methods
	' Returns a comma-delimited string
	'-------------------------------------------------------------------
	Function CheckCardMerchant()
		Dim sLocalSQL, rsAdmin, aAdmin
		
		sLocalSQL = "SELECT adminEncodeCCIsActive, adminMerchantType, adminCODAmount, adminTransMethod, adminPaymentServer, adminLogin, adminPassword FROM sfAdmin"
		Set rsAdmin	= Server.CreateObject("ADODB.RecordSet")
		rsAdmin.Open sLocalSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText
		
		If NOT rsAdmin.EOF Then
			aAdmin = rsAdmin.GetRows			
		End If		
		closeobj(rsAdmin)
		CheckCardMerchant = aAdmin
	End Function
'JF 9/25/01 removed Internet Cash stuff
	

	'-------------------------------------------------------------------
	' Set/Get Credit Cards
	'-------------------------------------------------------------------
	Sub setUpdateCC(bVisa,bMC,bDiscover,bAMEX,bDiners,bCarteBlanche,bCreditCard)
		Dim rsCC, sLocalSQL
		
		sLocalSQL = "SELECT transName, transIsActive FROM sfTransactionTypes"
		
		Set rsCC = Server.CreateObject("ADODB.RecordSet")
			rsCC.Open sLocalSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText
		
		Do While NOT rsCC.EOF And bCreditCard = 1
				If Trim(rsCC.Fields("transName"))	= "Visa" Then
					If bVisa <> "" Then
						rsCC.Fields("transIsActive") = 1
					Else 
						rsCC.Fields("transIsActive") = 0	
					End If	
				ElseIf Trim(rsCC.Fields("transName"))	= "MasterCard" Then
					If bMC <> "" Then
						rsCC.Fields("transIsActive") = 1
					Else 
						rsCC.Fields("transIsActive") = 0		
					End If	
				ElseIf Trim(rsCC.Fields("transName"))= "American Express" Then
					If bAMEX <> "" Then
						rsCC.Fields("transIsActive") = 1
					Else 
						rsCC.Fields("transIsActive") = 0		
					End If	
				ElseIf Trim(rsCC.Fields("transName"))	= "Discover" Then
					If bDiscover <> "" Then
						rsCC.Fields("transIsActive") = 1
					Else 
						rsCC.Fields("transIsActive") = 0		
					End If	
				ElseIf Trim(rsCC.Fields("transName")) =  "Diners Club" Then
					If bDiners <> "" Then
						rsCC.Fields("transIsActive") = 1 
					Else 
						rsCC.Fields("transIsActive") = 0	
					End If	
				ElseIf Trim(rsCC.Fields("transName")) =  "Carte Blanche" Then
					If bCarteBlanche <> "" Then
						rsCC.Fields("transIsActive") = 1 
					Else 
						rsCC.Fields("transIsActive") = 0		
					End If		
				End If
				rsCC.Update
				rsCC.MoveNext				
			Loop
		closeobj(rsCC)	
	End Sub

	'-------------------------------------------------------------------
	' Function to retrieve active credit cards
	' Returns a comma-delimited string -- meaning you gotta split it
	'-------------------------------------------------------------------
	Function getCreditCards
		Dim sLocalSQL, rsCC, aRows, iRow, sString	
		sLocalSQL = "SELECT transName FROM sfTransactionTypes WHERE transIsActive = 1 AND transType = 'Credit Card'"
		sString = ""	
		Set rsCC = Server.CreateObject("ADODB.RecordSet")
		rsCC.Open sLocalSQL, cnn, adOpenDynamic,adLockOptimistic, adCmdText
		If Not (rsCC.BOF And rsCC.EOF) Then	
			aRows = rsCC.GetRows
			For iRow = 0 to UBound(aRows, 2)
				sString = sString & "," & aRows(0,iRow)
			Next
		End If
		closeobj(rsCC)
		
		getCreditCards = sString			
	End Function

	'-------------------------------------------------------------------
	' Function to retrieve active payment methods
	' Returns a comma-delimited string -- meaning you gotta split it
	'-------------------------------------------------------------------
	Function getPaymentMethods
		Dim sLocalSQL, rsGetPay, aRows, iRow, sString	
		sLocalSQL = "SELECT DISTINCT transType,transName FROM sfTransactionTypes WHERE transIsActive = 1"
		sString = ""	
		Set rsGetPay = Server.CreateObject("ADODB.RecordSet")
		rsGetPay.Open sLocalSQL, cnn, adOpenDynamic,adLockOptimistic, adCmdText
		If Not (rsGetPay.BOF And rsGetPay.EOF) Then
			aRows = rsGetPay.GetRows
			For iRow = 0 to UBound(aRows, 2)
				If Trim(aRows(0,iRow)) = "PhoneFax" Then
					sString = sString & "," & Trim(aRows(0,iRow)) & Trim(aRows(1,iRow)) 
				Else					
					sString = sString & "," & aRows(0,iRow)
				End If
			Next
		End If
		closeobj(rsGetPay)
		getPaymentMethods = sString			
	End Function
	
	'-------------------------------------------------------------------
	' M0002 functions and subroutines ----------------------------------
	'-------------------------------------------------------------------
		
	'-------------------------------------------------------------------
	' Delete row in value shipping table
	'-------------------------------------------------------------------
	Sub setDeleteValShip(valShipAmount,iType)
		Dim sLocalSQL, delete
		
		If  iType = 0 Then
			sLocalSQL = "DELETE FROM sfValueShipping WHERE valShpPurTotal = '" & valShipAmount & "'"		
		ElseIf iType = 1 Then
			sLocalSQL = "DELETE FROM sfValueShipping"
		End If
		
		Set delete = cnn.Execute(sLocalSQL)		
		closeobj(delete)		
	End Sub
	
		
	'------------------------------------------------------------------
	' Redo the val-ship 
	'------------------------------------------------------------------
	Sub setUpdateNewVal(iPurchaseVal,iNewVal)
		Dim sLocalSQL, rsVal
				
		Set rsVal = Server.CreateObject("ADODB.RecordSet")
		rsVal.Open "sfValueShipping", cnn, adOpenDynamic,adLockOptimistic, adCmdTable
			rsVal.AddNew
			rsVal.Fields("valShpPurTotal")	= 	iPurchaseVal
			rsVal.Fields("valShpAmt")		= 	iNewVal
			rsVal.Update
		closeobj(rsVal)		
	End Sub
	
	'------------------------------------------------------------------
	' Set new ship value amount to value shipping table 
	'------------------------------------------------------------------
	Sub setNewValShip(sNewTotal,sNewShip)
		Dim rsValShip
		
		Set rsValShip = Server.CreateObject("ADODB.RecordSet")
		rsValShip.Open "sfValueShipping",cnn,adOpenDynamic,adLockOptimistic,adCmdTable		
		rsValShip.AddNew
		rsValShip.Fields("valShpPurTotal") = sNewTotal
		rsValShip.Fields("valShpAmt") = sNewship
		rsValShip.Update
		
		closeobj(rsValShip)		
	End Sub
		
	'------------------------------------------------------------------
	' Updates shipping 
	'------------------------------------------------------------------
	Sub setUpdateShipping(sShipID)
		Dim sLocalSQL, rsShip, aShipID, i, j
		sLocalSQL = "SELECT shipID, shipIsActive FROM sfShipping"		
		aShipID = split(sShipID,",")
		
		Set rsShip = Server.CreateObject("ADODB.RecordSet")
		rsShip.Open sLocalSQL, cnn, adOpenKeyset, adLockOptimistic, adCmdText
		
		' First Set all to 0
		Do While NOT rsShip.EOF
			rsShip.Fields("shipIsActive") = 0
			rsShip.MoveNext
		Loop
		rsShip.MoveFirst
		
		' Now set the ones still activated	
		Do While NOT rsShip.EOF
			For i = 0 to UBound(aShipID)
				If 	aShipID(i) = Trim(rsShip.Fields("shipID")) Then
					rsShip.Fields("shipIsActive") = 1						
				End If					
			Next	
			rsShip.MoveNext
		Loop		
				
		closeobj(rsShip)
	End Sub
	
	'-------------------------------------------------------------------
	' Function to retrieve the category list
	'-------------------------------------------------------------------
	Function getCategoryRow()
		Set getCategoryRow = openTable("sfCategories",cnn)
	End Function
	'-------------------------------------------------------------------
	' Function to retrieve the Manufacterers list
	'-------------------------------------------------------------------
	Function getManufacterersList()
		Set getManufacterersList = openTable("sfManufacturers", cnn)
	End Function
	'-------------------------------------------------------------------
	' Function that opens a full table
	'-------------------------------------------------------------------
	Function OpenTable(tableName, cnn)
		' Variable Declarations
		Dim recSet
		Dim sList
		Dim intId

		' Object Creation
		Set recSet = Server.CreateObject("ADODB.RecordSet")
		recSet.CursorLocation = adUseClient
		recSet.Open tableName, cnn, adOpenForwardOnly, adLockOptimistic, adCmdTable
				
		Set OpenTable = recSet
	End Function	
	
	'-------------------------------------------------------------------
	' Function to check if a column is empty and fill with something
	'-------------------------------------------------------------------
	Function chkColumnValue(strColumn, objRS)
		' Variable Declarations
		Dim strCV
		
		strCV = objRS(strColumn)
		If strCV = "" or IsNull(strCV) Then strCV = "---"
		
		chkColumnValue = strCV			
	End Function
	
	'-------------------------------------------------------------------
	' Function to Display Boolean Values with Yes/No
	'-------------------------------------------------------------------
	Function getColumnBValue(strColumn, objRS)
		' Variable Declarations
		Dim strCV
		
		strCV = objRS(strColumn)
		If strCV = 1 Then
			strCV = "Yes"
		Else 
			strCV = "No"
		End If
		
		getColumnBValue = strCV	
	End Function	
	
	
	'--------------------------------------------------------------------
	' Updates admin Shipping-related info
	'--------------------------------------------------------------------
	Sub updateAdminShip(iShippingMethod,iShipType2,sOriginCountry,sOriginZip,iShipTax,iHandlingActive,sHandling,iHandlingType,iPremShip,sPremShipCharge,sMinShipCharge,sUspspassword,sUspsUserName)
		Dim sLocalSQL, rsAdmin
		
		sLocalSQL = "SELECT adminOriginZip,adminOriginCountry,adminShipType,adminShipType2,adminPrmShipIsActive,adminHandlingIsActive,adminTaxShipIsActive,adminHandlingType,adminHandling,adminShipMin,adminSpcShipAmt,adminUsPsPassword,adminUsPsUserName FROM sfAdmin"
		Set rsAdmin = Server.CreateObject("ADODB.RecordSet")
			rsAdmin.Open sLocalSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText
			rsAdmin.Fields("adminOriginZip")		= sOriginZip
			rsAdmin.Fields("adminOriginCountry")	= sOriginCountry
			rsAdmin.Fields("adminShipType")			= iShippingMethod
			If iShipType2 <> "" Then
				rsAdmin.Fields("adminShipType2")		= iShipType2
			End If	
			rsAdmin.Fields("adminPrmShipIsActive")	= iPremShip
			rsAdmin.Fields("adminHandlingIsActive") = iHandlingActive
			rsAdmin.Fields("adminTaxShipIsActive")	= iShipTax
			If iHandlingActive = 1  Then
				rsAdmin.Fields("adminHandlingType")	= iHandlingType
				rsAdmin.Fields("adminHandling")		= sHandling
			End If	
			rsAdmin.Fields("adminShipMin")			= sMinShipCharge
			If iPremShip = 1 Then
				rsAdmin.Fields("adminSpcShipAmt")	= sPremShipCharge
			End If	
			rsAdmin.fields("adminUsPsPassword") =			sUspspassword
			rsAdmin.fields("adminUsPsUserName") = 			sUspsUserName
		rsAdmin.Update
		closeobj(rsAdmin)		
	End Sub

	'--------------------------------------------------------------------
	' Function : getCategoryList
	' This returns the category list in HTML format for dropdown box.
	'--------------------------------------------------------------------	
	Function getCategoryList(iValue)
		
		' Variable Declarations
		Dim rsCategoryList
		Dim sList
		Dim intId
		
		' Object Creation and Query
		Set rsCategoryList = Server.CreateObject("ADODB.RecordSet")
		rsCategoryList.Open "SELECT DISTINCT catID, catName  FROM sfCategories", cnn, adOpenForwardOnly, adLockOptimistic, adCmdText
		sList = ""
		If iValue = "" Then
			Do While Not rsCategoryList.EOF
				sList = sList & "<OPTION value=" & rsCategoryList.Fields("catID") & ">" & rsCategoryList.Fields("catName") & "</OPTION>"
				rsCategoryList.MoveNext
			Loop
		Else
			Do While Not rsCategoryList.EOF
				intId = trim(rsCategoryList.Fields("catID"))
				If iValue = intId Then
					sList = sList & "<OPTION value=" & intId & " selected>" & rsCategoryList.Fields("catName") & "</OPTION>"
				Else
					sList = sList & "<OPTION value=" & intId & ">" & rsCategoryList.Fields("catName") & "</OPTION>"
				End If 
				rsCategoryList.MoveNext
			Loop				
		End If

		' Object Cleanup
		rsCategoryList.Close
		Set rsCategoryList = nothing 

		' Return Value	
		getCategoryList = sList
	End Function

	'-------------------------------------------------------------------
	' Function : getManufacturersList
	' This returns the mfg list in HTML format for dropdown box.
	'-------------------------------------------------------------------
	Function getManufacturersList(iValue)
	' Variable Declarations
	Dim rsManufacturersList
	Dim sList
	Dim intId
	
	' Object Creation and Query
	Set rsManufacturersList = Server.CreateObject("ADODB.RecordSet")
	rsManufacturersList.Open "sfManufacturers", cnn, adOpenForwardOnly, adLockOptimistic, adCmdTable
	sList = ""
	If iValue = "" Then
		Do While Not rsManufacturersList.EOF
			sList = sList & "<OPTION value=" & rsManufacturersList.Fields("mfgID") & ">" & rsManufacturersList.Fields("mfgName") & "</OPTION>"
			rsManufacturersList.MoveNext
		Loop
	Else
		Do While Not rsManufacturersList.EOF
			intId = trim(rsManufacturersList.Fields("mfgID"))
			If iValue = intId Then
				slist = slist & "<OPTION value=" & intId & " selected>" & rsManufacturersList.Fields("mfgName") & "</OPTION>"
			Else
				slist = slist & "<OPTION value=" & intId & ">" & rsManufacturersList.Fields("mfgName") & "</OPTION>"
			End If 
			rsManufacturersList.MoveNext
		Loop				
	End If
	'object cleanup
	rsManufacturersList.Close 
	Set rsManufacturersList = nothing
	
	'return value
	getManufacturersList = sList
	End Function
	
	'-------------------------------------------------------------------
	' Function : getVendorList
	' This returns the vendor list in HTML format for dropdown box.
	'-------------------------------------------------------------------
	Function getVendorList(iValue)

	' Variable Declarations
	Dim rsVendorList
	Dim sList
	Dim intId
	
	' Object Creation and Query
	Set rsVendorList = Server.CreateObject("ADODB.RecordSet")
	rsVendorList.Open "sfVendors", cnn, adOpenForwardOnly, adLockOptimistic, adCmdTable
	sList = ""
	If iValue = "" Then
		Do While Not rsVendorList.EOF
			sList = sList & "<OPTION value=" & rsVendorList.Fields("vendID") & ">" & rsVendorList.Fields("vendName") & "</OPTION>"
			rsVendorList.MoveNext
		Loop
	Else
		Do While Not rsVendorList.EOF
			intId = trim(rsVendorList.Fields("vendID"))
			If iValue = intId Then
				sList = sList & "<OPTION value=" & intId & " selected>" & rsVendorList.Fields("vendName") & "</OPTION>"
			Else
				sList = sList & "<OPTION value=" & intId & ">" & rsVendorList.Fields("vendName") & "</OPTION>"
			End If 
			rsVendorList.MoveNext
		Loop				
	End If
	
	'object cleanup
	rsVendorList.Close
	Set rsVendorList = nothing 
	
	'return value
	getVendorList = sList
	End Function

	'-------------------------------------------------------------------
	' Function to retrieve the product list
	'-------------------------------------------------------------------
	Function getProductList()
		Set getProductList = openTable("sfProducts",cnn)
	End Function
	
	
	'-------------------------------------------------------------------
	' Function that opens a full table
	'-------------------------------------------------------------------
	Function OpenTable(tableName, cnn)
		' Variable Declarations
		Dim recSet
		Dim sList
		Dim intId

		' Object Creation
		Set recSet = Server.CreateObject("ADODB.RecordSet")
		recSet.CursorLocation = adUseClient
		recSet.Open tableName, cnn, adOpenForwardOnly, adLockOptimistic, adCmdTable
				
		Set OpenTable = recSet
	End Function	
	
	'-------------------------------------------------------------------
	' Function that releases an object
	'-------------------------------------------------------------------
	Sub closeObj(objItem)
		On Error Resume Next
		objItem.Close
		Set objItem=nothing	
		On Error GoTo 0
	End Sub
	
	'-------------------------------------------------------------------
	' M0003 functions and subroutines ----------------------------------	
	'-------------------------------------------------------------------
	
	'-------------------------------------------------------------------
	' setCtryTax sets the tax for the country
	'-------------------------------------------------------------------
	Sub setCtryTax(sCtryAbbr,dCtryTax)
		Dim sLocalSQL, rsCtry
		sLocalSQL = "SELECT loclctryTax, loclctryTaxIsActive FROM sfLocalesCountry WHERE loclctryAbbreviation = '" & sCtryAbbr & "'"

		Set rsCtry = Server.CreateObject("ADODB.RecordSet")
		rsCtry.Open sLocalSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText
		
		If NOT rsCtry.EOF Then
			rsCtry.Fields("loclctryTax")		= dCtryTax
			rsCtry.Fields("loclctryTaxIsActive")= 1
		rsCtry.Update
		End If
		
		closeobj(rsCtry)		
	End Sub
	
	'-------------------------------------------------------------------
	' Makes Ctry Tax inactive
	'-------------------------------------------------------------------
	Sub setResetTaxCtry(sCtryAbbr)
		Dim sLocalSQL, rsCtry
		sLocalSQL = "SELECT loclctryTaxIsActive FROM sfLocalesCountry WHERE loclctryAbbreviation = '" & sCtryAbbr & "'"

		Set rsCtry = Server.CreateObject("ADODB.RecordSet")
		rsCtry.Open sLocalSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText
		
		If NOT rsCtry.EOF Then
			rsCtry.Fields("loclctryTaxIsActive")= 0
		rsCtry.Update
		End If
		
		closeobj(rsCtry)			
	End Sub
	
	'-------------------------------------------------------------------
	' setStateTax sets the tax for the state
	'-------------------------------------------------------------------
	Sub setStateTax(sStateAbbr,dStateTax)
		Dim sLocalSQL, rsState
		sLocalSQL = "SELECT loclstTax, loclstTaxIsActive FROM sfLocalesState WHERE loclstAbbreviation = '" & sStateAbbr & "'"

		Set rsState = Server.CreateObject("ADODB.RecordSet")
		rsState.Open sLocalSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText
		If NOT rsState.EOF Then
			rsState.Fields("loclstTax")			= dStateTax
			rsState.Fields("loclstTaxIsActive")	= 1
		rsState.Update
		End If		
		closeobj(rsState)		
	End Sub 
	
	'-------------------------------------------------------------------
	' Resets Tax for the state, making it inactive
	'-------------------------------------------------------------------
	Sub setResetTaxState(sStateAbbr)
		Dim sLocalSQL, rsState
		sLocalSQL = "SELECT loclstTaxIsActive FROM sfLocalesState WHERE loclstAbbreviation = '" & sStateAbbr & "'"

		Set rsState = Server.CreateObject("ADODB.RecordSet")
		rsState.Open sLocalSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText
			rsState.Fields("loclstTaxIsActive")	= 0
		rsState.Update
		
		closeobj(rsState)			
	End Sub
	
	'-------------------------------------------------------------------
	' M0007 functions and subroutines ----------------------------------	
	'-------------------------------------------------------------------
	
	'-------------------------------------------------------------------
	' getActiveCountryList generates lists of active countries depending on type wanted
	' returns a string
	'-------------------------------------------------------------------
	Function getActiveCountryList(iType)
		Dim sLocalSQL, rsCtry, sCtryList
		
		If iType = 1 Then
			sLocalSQL = "SELECT loclctryAbbreviation,loclctryName,loclctryTax,loclctryTaxIsActive, loclctryLocalIsActive FROM sfLocalesCountry WHERE loclctryTaxIsActive = 0 AND loclctryLocalIsActive = 1 ORDER BY loclctryName"
		ElseIf iType = 2 Then
			sLocalSQL = "SELECT loclctryAbbreviation,loclctryName,loclctryTax,loclctryTaxIsActive, loclctryLocalIsActive FROM sfLocalesCountry WHERE loclctryTaxIsActive = 1 AND loclctryLocalIsActive = 1 ORDER BY loclctryName"
		ElseIf iType = 3 Then
			sLocalSQL = "SELECT loclctryAbbreviation,loclctryName,loclctryLocalIsActive FROM sfLocalesCountry WHERE loclctryLocalIsActive = 0 ORDER BY loclctryName"
		ElseIf iType = 4 Then
			sLocalSQL = "SELECT loclctryAbbreviation,loclctryName,loclctryLocalIsActive FROM sfLocalesCountry WHERE loclctryLocalIsActive = 1 ORDER BY loclctryName"			
		End If	
		
		Set rsCtry = Server.CreateObject("ADODB.RecordSet")
			rsCtry.Open sLocalSQL,cnn,adOpenDynamic,adLockOptimistic,adCmdText

			Do While Not rsCtry.EOF
				sCtryList = sCtryList & "<option value=""" & rsCtry.Fields("loclctryAbbreviation") & """>" & rsCtry.Fields("loclctryName") & "</option>"
				rsCtry.MoveNext
			Loop		
			
		getActiveCountryList = sCtryList
	End Function

	'-------------------------------------------------------------------
	' getActiveStateList generates lists of active states depending on type wanted
	' returns a string
	'-------------------------------------------------------------------
	Function getActiveStateList(iType)
		Dim sLocalSQL, rsState, sStateList
		
		If iType = 1 Then
			sLocalSQL = "SELECT loclstAbbreviation,loclstName,loclstTax,loclstTaxIsActive, loclstLocaleIsActive FROM sfLocalesState WHERE loclstTaxIsActive = 0 AND loclstLocaleIsActive = 1 ORDER BY loclstName"
		ElseIf iType = 2 Then
			sLocalSQL = "SELECT loclstAbbreviation,loclstName,loclstTax,loclstTaxIsActive, loclstLocaleIsActive FROM sfLocalesState WHERE loclstTaxIsActive = 1 AND loclstLocaleIsActive = 1 ORDER BY loclstName"
		ElseIf iType = 3 Then
			sLocalSQL = "SELECT loclstAbbreviation,loclstName FROM sfLocalesState WHERE loclstLocaleIsActive = 0 ORDER BY loclstName"
		ElseIf iType = 4 Then
			sLocalSQL = "SELECT loclstAbbreviation,loclstName FROM sfLocalesState WHERE loclstLocaleIsActive = 1 ORDER BY loclstName"								
		End If	
		
		Set rsState = Server.CreateObject("ADODB.RecordSet")
			rsState.Open sLocalSQL,cnn,adOpenDynamic,adLockOptimistic,adCmdText

			Do While NOT rsState.EOF
				sStateList = sStateList & "<option value=""" & rsState.Fields("loclstAbbreviation") & """>" & rsState.Fields("loclstName") & "</option>"
				rsState.MoveNext
			Loop	
			
		getActiveStateList = sStateList
	End Function
	
	'-------------------------------------------------------------------
	' Resets active flag for the state, making it inactive
	'-------------------------------------------------------------------
	Sub setState(sStateAbbr,iType)
		Dim sLocalSQL, rsState
		sLocalSQL = "SELECT loclstLocaleIsActive FROM sfLocalesState WHERE loclstAbbreviation = '" & sStateAbbr & "'"

		Set rsState = Server.CreateObject("ADODB.RecordSet")
		rsState.Open sLocalSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText
			rsState.Fields("loclstLocaleIsActive")	= iType
		rsState.Update
		
		closeobj(rsState)			
	End Sub
	
	'-------------------------------------------------------------------
	' Resets active flag for the country, making it inactive
	'-------------------------------------------------------------------
	Sub setCountry(sCtryAbbr,iType)
		Dim sLocalSQL, rsCtry
		sLocalSQL = "SELECT loclctryLocalIsActive FROM sfLocalesCountry WHERE loclctryAbbreviation = '" & sCtryAbbr & "'"
		Set rsCtry = Server.CreateObject("ADODB.RecordSet")
		rsCtry.Open sLocalSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText
		
		If NOT rsCtry.EOF Then
			rsCtry.Fields("loclctryLocalIsActive")= iType
			rsCtry.Update
		End If
		
		closeobj(rsCtry)			
	End Sub	
	
	'-------------------------------------------------------------------
	' M0012 functions and subroutines ----------------------------------	
	'-------------------------------------------------------------------

	'--------------------------------------------------------------------
	' Find previous record of Category Name	
	' Returns ID of found record or -1 for not found record
	'--------------------------------------------------------------------	
	Function findRecord(sCategoryName)
		Dim sLocalSQL, rsCategorySettings
		sLocalSQL = "Select catID,catName FROM sfCategories WHERE catName = '" & sCategoryName & "'"
	
		Set rsCategorySettings = Server.CreateObject("ADODB.recordSet")
		rsCategorySettings.Open sLocalSQL, cnn, adOpenStatic, adLockOptimistic, adCmdText	
	
		Do While NOT rsCategorySettings.EOF 
			If Trim(rsCategorySettings.Fields("catName")) = sCategoryName Then
				findRecord = Trim(rsCategorySettings.Fields("catID"))
				closeobj(rsCategorySettings)
				Exit Function				
			Else
				rsCategorySettings.moveNext
			End If
		Loop
		closeobj(rsCategorySettings)
		findRecord = -1	
	End Function
	
	'--------------------------------------------------------------------
	' Find previous record of mfg Name	
	' Returns ID of found record or -1 for not found record
	'--------------------------------------------------------------------	
	Function findMfgRecord(sMfgName)
		Dim sLocalSQL, rsMfgSettings
		sLocalSQL = "Select mfgID,mfgName FROM sfManufacturers WHERE mfgName = '" & sMfgName & "'"
	
		Set rsMfgSettings = Server.CreateObject("ADODB.recordSet")
		rsMfgSettings.Open sLocalSQL, cnn, adOpenStatic, adLockOptimistic, adCmdText	
	
		Do While NOT rsMfgSettings.EOF 
			If Trim(rsMfgSettings.Fields("mfgName")) = sMfgName Then
				findMfgRecord = Trim(rsMfgSettings.Fields("mfgID"))
				closeobj(rsMfgSettings)
				Exit Function				
			Else
				rsMfgSettings.moveNext
			End If
		Loop
		closeobj(rsMfgSettings)
		findMfgRecord = -1	
	End Function
	
	'--------------------------------------------------------------------
	' Find previous record of Vendor Name	
	' Returns ID of found record or -1 for not found record
	'--------------------------------------------------------------------	
	Function findVendRecord(sVendName)
		Dim sLocalSQL, rsVendSettings
		sLocalSQL = "Select VendID,VendName FROM sfVendors WHERE vendName = '" & sVendName & "'"
	
		Set rsVendSettings = Server.CreateObject("ADODB.recordSet")
		rsVendSettings.Open sLocalSQL, cnn, adOpenStatic, adLockOptimistic, adCmdText	
	
		Do While NOT rsVendSettings.EOF 
			If Trim(rsVendSettings.Fields("VendName")) = sVendName Then
				findVendRecord = Trim(rsVendSettings.Fields("vendID"))
				closeobj(rsVendSettings)
				Exit Function				
			Else
				rsVendSettings.moveNext
			End If
		Loop
		closeobj(rsVendSettings)
		findVendRecord = -1	
	End Function
	
	'--------------------------------------------------------------------
	' Find previous record of Category Name	
	'--------------------------------------------------------------------	
	Function getCategoryInfo(iCatID)
		Dim sLocalSQL, rsCategory, aCat(5), aCategory
		sLocalSQL = "Select catName,catDescription,catImage,catIsActive,catHttpAdd FROM sfCategories WHERE catID=" & iCatID 
	
		Set rsCategory = Server.CreateObject("ADODB.recordSet")
		rsCategory.Open sLocalSQL, cnn, adOpenForwardOnly, adLockOptimistic, adCmdText	
		aCategory = rsCategory.getRows	
		aCat(0)	= aCategory(0,0)
		aCat(1) = aCategory(1,0)
		aCat(2)	= aCategory(2,0)
		aCat(3)	= aCategory(3,0)
		aCat(4)	= aCategory(4,0)
		'Response.Write "<br>" & aCat(0) & " " & aCat(1) & " " & aCat(2) & " " & aCat(4)
		closeobj(rsCategory)
		getCategoryInfo = aCat
	End Function
	
	'-------------------------------------------------------------------
	' Add to new category
	'-------------------------------------------------------------------
	Sub setNewCategory(sCategoryName,sDescription,sImageURL,sIsActive,sPageURL)
		Dim rsCategory		
		Set rsCategory = Server.CreateObject("ADODB.RecordSet")
			rsCategory.Open "sfCategories", cnn, adOpenDynamic, adLockOptimistic, adCmdTable
			rsCategory.AddNew
			rsCategory.Fields("catName")		=	sCategoryName
			rsCategory.Fields("catDescription") =	sDescription
			rsCategory.Fields("catImage")		=   sImageURL
			rsCategory.Fields("catIsActive")	=	sIsActive
			rsCategory.Fields("catHttpAdd")		=   sPageURL
			rsCategory.Update
		closeobj(rsCategory)	
	End Sub	
	
	'-------------------------------------------------------------------
	' Edit category
	'-------------------------------------------------------------------
	Sub setEditCategory(iCatID,sCategoryName,sDescription,sImageURL,sIsActive,sPageURL)
		Dim rsCategory, sLocalSQL
		
		sLocalSQL = "SELECT catName, catDescription, catImage, catHasSubCategory, catIsActive, catHttpAdd FROM sfCategories WHERE catID =" &  iCatID
				
		Set rsCategory = Server.CreateObject("ADODB.RecordSet")
			rsCategory.Open sLocalSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText
		
		If Not rsCategory.EOF Then
		
			rsCategory.Fields("catName")		=	sCategoryName
			rsCategory.Fields("catDescription") =	sDescription
			rsCategory.Fields("catImage")		=   sImageURL
			rsCategory.Fields("catIsActive")	=	sIsActive
			rsCategory.Fields("catHttpAdd")		=   sPageURL
			rsCategory.Update
		End If	
		closeobj(rsCategory)	
	End Sub		
	
	'-------------------------------------------------------------------
	' getCategoriesList generates a list of categories
	'-------------------------------------------------------------------
	Function getCategoriesList
		Dim sLocalSQL, rsCategory, sCat
		
		sLocalSQL = "SELECT catID, catName FROM sfCategories"
		Set rsCategory = Server.CreateObject("ADODB.RecordSet")
			rsCategory.Open sLocalSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText
			
		Do While NOT rsCategory.EOF 
			sCat = sCat & "<option value=" & Trim(rsCategory.Fields("catID")) & ">" &  Trim(rsCategory.Fields("catName")) & "</option>"
			rsCategory.MoveNext
		Loop	
		getCategoriesList = sCat
	End Function
	
	'------------------------------------------------------------------
	' Returns an array of all supported countries and states
	'------------------------------------------------------------------
	Function getCtryStates(iType)
		Dim sLocalSQL, rsCS, aAll
		If iType = 1 Then
			sLocalSQL = "SELECT loclctryAbbreviation,loclctryName FROM sfLocalesCountry WHERE loclctryLocalIsActive = 1 ORDER BY loclctryName"
		ElseIf iType = 2 Then
			sLocalSQL = "SELECT loclstAbbreviation,loclstName FROM sfLocalesState WHERE loclstLocaleIsActive = 1 ORDER BY loclstName"
		End If
		
		Set rsCS = Server.CreateObject("ADODB.RecordSet")
			rsCS.Open sLocalSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText
			aAll = rsCS.GetRows
			closeobj(rsCS)

		getCtryStates = aAll			
	End Function
	
	'-------------------------------------------------------------------
	' Deletes the product from record
	'-------------------------------------------------------------------
	Sub setDeleteProd(sProdID)
		Dim sLocalSQL, rs
		sLocalSQL = "DELETE FROM sfProducts WHERE prodID ='" & sProdID & "'"
		
		Set rs = cnn.Execute(sLocalSQL)
		closeobj(rs)
	End Sub
	
	'-------------------------------------------------------------------
	' setDeleteCategory deletes the category from the db
	'-------------------------------------------------------------------
	Sub setDeleteCategory(iCatID)
		Dim sLocalSQL, rs
		sLocalSQL = "DELETE FROM sfCategories WHERE catID = " & iCatID
		
		Set rs = cnn.Execute(sLocalSQL)
		closeobj(rs)
	End Sub
	
	'-------------------------------------------------------------------
	'Added by Daniel From incGeneral
	'-------------------------------------------------------------------
	
	'-------------------------------------------------------------------
	' Function that takes a sql statement and executes on a table
	'-------------------------------------------------------------------
	Function RestrictedOpenTable(sSQL,cnn)
		' Variable Declarations
		Dim recSet
			
		Set recSet = Server.CreateObject("ADODB.RecordSet")
		recSet.Open sSQL, cnn, adOpenStatic, adLockOptimistic, adCmdText

		Set RestrictedOpenTable = recSet
	End Function
	
	'-------------------------------------------------------------------
	' Function that opens a full table
	'-------------------------------------------------------------------
	Function OpenTable(tableName, cnn)
		' Variable Declarations
		Dim recSet
		Dim sList
		Dim intId

		' Object Creation
		Set recSet = Server.CreateObject("ADODB.RecordSet")
		recSet.CursorLocation = adUseClient
		recSet.Open tableName, cnn, adOpenForwardOnly, adLockOptimistic, adCmdTable
				
		Set OpenTable = recSet
	End Function	
	
	'-------------------------------------------------------------------
	' Function to get list of payment processors and associated values	
	'-------------------------------------------------------------------
	Function getTransactionMethods
		Dim sLocalSQL, rsTransMethod
		sLocalSQL = "SELECT trnsmthdID,trnsmthdName,trnsmthdServerPath,trnsmthdLogin,trnsmthdPasswd	FROM sfTransactionMethods WHERE NOT trnsmthdName = 'InternetCash'"
		
		Set rsTransMethod = Server.CreateObject("ADODB.RecordSet")
		rsTransMethod.Open sLocalSQL, cnn, adOpenForwardOnly,adLockReadOnly, adCmdText 				
		
		getTransactionMethods = rsTransMethod.GetRows		
		closeobj(rsTransMethod)				
	End Function
	
		'-------------------------------------------------------------------
	' Function to retrieve the design list
	'-------------------------------------------------------------------	
	Function getDesignList(iChoice,iDesignId)
		' Variable Declarations
		Dim sSQL
		
		Select Case iChoice
			Case 1
				sSQL = "SELECT * FROM sfDesign WHERE dsgnID = "& iDesignId 
				Set getDesignList = RestrictedOpenTable(sSQL,cnn)
			Case 2
				sSQL = "SELECT dsgnID, dsgnName, dsgnDescription,dsgnIsActive FROM sfDesign ORDER BY dsgnName ASC"
				Set getDesignList = RestrictedOpenTable(sSQL,cnn)			
		End Select			
	End Function
	
	Function getList(iChoice,sTable,sColumn)
		' Variable Declarations
		Dim sSQL
		
		Select Case iChoice
			Case 1
				sSQL = "SELECT "& sColumn &" FROM "& sTable &" WHERE "& sColumn &" is not null"
				Set getList = RestrictedOpenTable(sSQL,cnn)
			Case 2
				sSQL = "SELECT slctvalLCID, slctvalLCIDLabel FROM sfSelectValues WHERE " _
					& "slctvalLCID is not null AND slctvalLCIDLabel is NOT null ORDER BY slctvalLCIDLabel"
				Set getList = RestrictedOpenTable(sSQL,cnn)
		End Select
	End Function
	
	Sub updateTable(iChoice, sTable)
		Dim sSQL
		Dim recSet
		
		Select Case iChoice
			Case 1
				Set recSet = OpenTable(sTable,cnn)
				Do While NOT recSet.EOF
					recSet.Fields("dsgnIsActive") = 0
					recSet.Update
				recSet.MoveNext					
				Loop
		End Select				
		
	End Sub
	
%>









