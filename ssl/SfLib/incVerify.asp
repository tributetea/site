<%
'@BEGINVERSIONINFO

'@APPVERSION: 50.4011.0.4

'@FILENAME: incverify.asp
	 



'@DESCRIPTION: Verify the order information

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde  Incorporated is expressly prohibited.
'
'(c) Copyright 2000,2001 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO
'------------------------------------------------------------------
' Checks PhoneFax type
' Returns 0 for recorded, 1 for non-recorded
'------------------------------------------------------------------

Dim Pkg


Function CheckPaymentMethod(sPayMethod)
Dim sLocalSQL, rsTransType, sReturn, sTempTransType

	sLocalSQL = "SELECT DISTINCT transtype FROM sfTransactionTypes WHERE transType='" & sPayMethod & "' and transIsActive = 1"
	
	Set rsTransType = Server.CreateObject("ADODB.RecordSet")
		rsTransType.Open sLocalSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText
		
		If rsTransType.EOF or rsTransType.BOF Then
			sReturn=0
		Else	
			sReturn=1
		End If	
		closeObj(rsTransType)
	CheckPayMentMethod = sReturn	
End Function

Function CheckPhoneFax
	Dim sLocalSQL, rsTrans, bType, Recorded
	
	sLocalSQL = "SELECT transName FROM sfTransactionTypes WHERE transType = 'PhoneFax' AND transIsActive = 1"
	Set rsTrans = Server.CreateObject("ADODB.RecordSet")
	rsTrans.Open sLocalSQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
	
	If rsTrans.EOF Or rsTrans.BOF Then
		bType = 0
	Else
	   'Response.write  rsTrans.Fields("transName") & "     "
		If rsTrans.Fields("transName") = "Recorded" Then
			bType = 0
		Else
			bType = 1
		End If		
	End If
	closeObj(rsTrans)
	'Response.write "TEST@   " & bType
	CheckPhoneFax = bType
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
			If rsActive.Fields(sPrefix & "ID") <> cLNG(iID) Then
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

'-------------------------------------------------------
' Add new customer's info in sfCustomers
' Returns ID of inserted row
'-------------------------------------------------------
Function getNewCustomer(sCustEmail,sPassword,sCustFirstName,sCustMiddleInitial,sCustLastName,sCustCompany,sCustAddress1,sCustAddress2,sCustCity,sCustState,sCustZip,sCustCountry,sCustPhone,sCustFax,bCustSubscribed)			
Dim	sLocalSQl, rsUpdate, iKeyID, bookMark

	Set rsUpdate = Server.CreateObject("ADODB.RecordSet")
		rsUpdate.CursorLocation = adUseClient
		rsUpdate.Open "sfCustomers ORDER BY custID",cnn,adOpenDynamic,adLockOptimistic,adCmdTable
		rsUpdate.AddNew
		rsUpdate.Fields("custFirstName")		= trim(sCustFirstName)
		rsUpdate.Fields("custMiddleInitial")	= trim(sCustMiddleInitial)
		rsUpdate.Fields("custLastName")			= trim(sCustLastName)
		rsUpdate.Fields("custCompany")			= trim(sCustCompany)
		rsUpdate.Fields("custAddr1")			= trim(sCustAddress1)
		rsUpdate.Fields("custAddr2")			= trim(sCustAddress2)
		rsUpdate.Fields("custCity")				= trim(sCustCity)
		rsUpdate.Fields("custState")			= trim(sCustState)
		rsUpdate.Fields("custZip")				= trim(sCustZip)
		rsUpdate.Fields("custCountry")			= trim(sCustCountry)
		rsUpdate.Fields("custPhone")			= trim(sCustPhone)
		rsUpdate.Fields("custFax")				= trim(sCustFax)	
		rsUpdate.Fields("custPasswd")			= trim(sPassword)
		rsUpdate.Fields("custEmail")			= trim(sCustEmail)
		rsUpdate.Fields("custTimesAccessed")	= 1
		rsUpdate.Fields("custLastAccess")		= Date()
		rsUpdate.Fields("custIsSubscribed")   = bCustSubscribed
		
		rsUpdate.Update		
		'bookMark = rsUpdate.AbsolutePosition 
		'rsUpdate.Requery 
		'rsUpdate.AbsolutePosition = bookMark			
		iKeyID  = rsUpdate.Fields("custID")
		closeObj(rsUpdate)	
		
		getNewCustomer = iKeyID
End Function

'----------------------------------------------------------
' Write into sfCShpAddresses
' Returns ID written to
'----------------------------------------------------------
Function setShipping(sShipFirstName,sShipMiddleInitial,sShipLastName,sShipCompany,sShipAddress1,sShipAddress2,sShipCity,sShipState,sShipZip,sShipCountry,sShipPhone,sShipFax,sShipEmail)
	Dim rsShipping, sLocalSQL, iShipID, bookMark
	
	sLocalSQL = "SELECT cshpaddrIsActive,cshpaddrID FROM sfCShipAddresses WHERE cshpaddrCustID = " & Request.Cookies("sfCustomer")("custID") 
	
	' Set the old ship addresses to 0
	Set rsShipping = Server.CreateObject("ADODB.RecordSet")
		rsShipping.Open sLocalSQL, cnn, adOpenDynamic,adLockOptimistic, adCmdText
		
		If NOT rsShipping.EOF Then
			Do While NOT rsShipping.EOF
					rsShipping.Fields("cshpaddrIsActive") = 0
					rsShipping.Update
					rsShipping.MoveNext
			Loop			
		End If
		rsShipping.Close
		
		rsShipping.CursorLocation = adUseClient
		rsShipping.Open "sfCShipAddresses", cnn, adOpenDynamic,adLockOptimistic, adCmdTable
		rsShipping.AddNew
			rsShipping.Fields("cshpaddrCustID")				= trim(Request.Cookies("sfCustomer")("custID"))
			rsShipping.Fields("cshpaddrShipFirstName")		= trim(sShipFirstName)
			rsShipping.Fields("cshpaddrShipMiddleInitial")	= trim(sShipMiddleInitial)
			rsShipping.Fields("cshpaddrShipLastName")			= trim(sShipLastName)
			rsShipping.Fields("cshpaddrShipCompany")			= trim(sShipCompany)
			rsShipping.Fields("cshpaddrShipAddr1")			= trim(sShipAddress1)
			rsShipping.Fields("cshpaddrShipAddr2")			= trim(sShipAddress2)
			rsShipping.Fields("cshpaddrShipCity")				= trim(sShipCity)
			rsShipping.Fields("cshpaddrShipState")			= trim(sShipState)
			rsShipping.Fields("cshpaddrShipZip")				= trim(sShipZip)
			rsShipping.Fields("cshpaddrShipCountry")			= trim(sShipCountry)
			rsShipping.Fields("cshpaddrShipPhone")			= trim(sShipPhone)
			rsShipping.Fields("cshpaddrShipFax")				= trim(sShipFax)
			rsShipping.Fields("cshpaddrShipEmail")			= trim(sShipEmail)
			rsShipping.Fields("cshpaddrIsActive")				= 1	
		rsShipping.Update
				
		iShipID = rsShipping.Fields("cshpaddrID")
		closeobj(rsShipping)
		setShipping = iShipID
End Function
'--------------------------------------------------------------------
' parses through shipping form
' Returns the id if a match is found
' 0 is returned if nothing has been entered
'--------------------------------------------------------------------
Function CheckShippingChange(sShipFirstName,sShipMiddleInitial,sShipLastName,sShipCompany,sShipAddress1,sShipAddress2,sShipCity,sShipState,sShipZip,sShipCountry,sShipPhone,sShipFax,sShipEmail)

Dim sLocalSQL, rsShip, aShipArray, iRow, bNoMatch, iMatchID

	If sShipFirstName = "" And sShipLastName = "" And sShipAddress1 = "" And sShipCity = "" And sShipCountry = "" Then
		CheckShippingChange = 0
		Exit Function
	End If		

	sLocalSQL = "SELECT cshpaddrShipFirstName,cshpaddrShipMiddleInitial,cshpaddrShipLastName,cshpaddrShipCompany,cshpaddrShipAddr1,"_
				 & "cshpaddrShipAddr2,cshpaddrShipCity,cshpaddrShipState,cshpaddrShipZip,cshpaddrShipCountry,cshpaddrShipPhone,"_
				 & "cshpaddrShipFax,cshpaddrShipEmail,cshpaddrID FROM sfCShipAddresses WHERE cshpaddrCustID = " & Request.Cookies("sfCustomer")("custID")
	
	Set rsShip = Server.CreateObject("ADODB.RecordSet")
	rsShip.Open sLocalSQL, cnn, adOpenKeyset, adLockOptimistic, adCmdText
	' Assume no change at beginning	
	iMatchID = 0 
	
	If NOT rsShip.EOF Then
			aShipArray = rsShip.GetRows
	
			iRow = 0
			If vDebug = 1 Then 	
				Response.Write "<br>CheckShippingChange SQL : " & sLocalSQL		
				Response.Write "<br>" & (Trim((aShipArray(0,iRow)) <> Trim(sShipFirstName)))
				Response.Write "<br>" & (Trim(aShipArray(1,iRow)) <> Trim(sShipMiddleInitial))
				Response.Write "<br>" & (Trim(aShipArray(2,iRow)) <> Trim(sShipLastName)) 
				Response.Write "<br>" & (Trim(aShipArray(3,iRow)) <> Trim(sShipCompany))
				Response.Write "<br>" & (Trim(aShipArray(4,iRow)) <> Trim(sShipAddress1))
				Response.Write "<br>" & (Trim(aShipArray(5,iRow)) <> Trim(sShipAddress2)) 
				Response.Write "<br>" & (Trim(aShipArray(6,iRow)) <> Trim(sShipCity))
				Response.Write "<br>" & (Trim(aShipArray(7,iRow)) <> Trim(sShipState)) 
				Response.WRITE "<br>" & (Trim(aShipArray(8,iRow)) <> Trim(sShipZip)) 
				Response.Write "<br>" & (Trim(aShipArray(9,iRow)) <> Trim(sShipCountry)) 
				Response.Write "<br>" & (Trim(aShipArray(10,iRow)) <> Trim(sShipPhone))
				Response.Write "<br>" & (Trim(aShipArray(11,iRow)) <> Trim(sShipFax))  
				Response.Write "<br>" & (Trim(aShipArray(12,iRow)) <> Trim(sShipEmail))
			End If
			
	
			If rsShip.RecordCount > 0 Then
				For iRow = 0 to rsShip.RecordCount - 1			
						' Debug
						If vDebug = 1 Then 	Response.Write "<br>Rows: " & rsShip.RecordCount
							
						If (Trim((aShipArray(0,iRow)) <> Trim(sShipFirstName)) Or (Trim(aShipArray(1,iRow)) <> Trim(sShipMiddleInitial)) Or (Trim(aShipArray(2,iRow)) <> Trim(sShipLastName)) Or (Trim(aShipArray(3,iRow)) <> Trim(sShipCompany)) Or (Trim(aShipArray(4,iRow)) <> Trim(sShipAddress1)) Or (Trim(aShipArray(5,iRow)) <> Trim(sShipAddress2)) Or  (Trim(aShipArray(6,iRow)) <> Trim(sShipCity)) Or (Trim(aShipArray(7,iRow)) <> Trim(sShipState)) Or (Trim(aShipArray(8,iRow)) <> Trim(sShipZip)) Or (Trim(aShipArray(9,iRow)) <> Trim(sShipCountry)) Or (Trim(aShipArray(10,iRow)) <> Trim(sShipPhone)) Or (Trim(aShipArray(11,iRow)) <> Trim(sShipFax)) Or (Trim(aShipArray(12,iRow)) <> Trim(sShipEmail))) Then
							If vDebug = 1 Then Response.Write "<br>No Match"		
							iMatchID = 0	
						Else
							If vDebug = 1 Then Response.Write "<br>Match : " & aShipArray(13,iRow)
							iMatchID = aShipArray(13,iRow)	
							
							Call SetActive("cshpaddr",iMatchID)
							closeobj(rsShip)
							CheckShippingChange = iMatchID
							Exit Function
						End If
				Next 
			End If	
	End If
	closeobj(rsShip)
	CheckShippingChange = iMatchID	
End Function

'------------------------------------------------------------------
' Returns shipping amount
' Returns a string
'------------------------------------------------------------------
Function GetShipping(iTotalPur, iPremiumShipping, aCheck, dCity,dState,dZip, dCountry, sTotalPrice,sType)'JF

	Dim SQL, sShipCode, sProdID, sShipping, oCountry, oZip, iLength, iWidth, iHeight, oCity, oState  'JF
	Dim iQuantity, iWeight, iShipType, iShipMin, iSpcShipAmt,uspsUsername,uspsPassword, CanadaPostRefNum  'JF added for Canada
	Dim rsAdmin, rsProdShipping, rsShipping, ups, arrCheck, sCheck, sErrMsg, obj, posit
    Dim sFreeship, ParseUsername, ParseLogin, upsUsername, upsPassword,TotaL_with_Attributes
    Dim boQty,shpQty,allQty 'SFAE
    dim noShip
'+jf 8/23/01
	dim oLCID
'-jf 8/23/01

	Set rsAdmin = Server.CreateObject("ADODB.RecordSet")
	'rsAdmin.Open "sfAdmin", cnn, adOpenForwardOnly, adLockReadOnly, adCmdTable
	rsAdmin.Open "sfAdmin", cnn, adOpenStatic, adLockReadOnly , adCmdTable 
	'Set UPS & USPS Variables
	iShipMin 		= rsAdmin.Fields("adminShipMin") 
	iSpcShipAmt 	= rsAdmin.Fields("adminSpcShipAmt")
	
	sShipCode 		= Request.Form("Shipping")
'if the order from didn't have the shipping combo box +JF
if trim(sShipCode) ="" then
GetShipping="0"
exit function
end if
'+jf
   posit =instr(sShipCode,",")
   If posit > 0 Then
     sShipCode = left(trim(sShipCode),posit-1) 'FreeShipping''''''''''''''''''''''
   End If
    
	oCountry 		= rsAdmin.Fields("adminOriginCountry")
	oZip 			= rsAdmin.Fields("adminOriginZip")

	oCity 			= "x"  'JF Seems to work with bogus City as long as Zip is correct, no state in DB
	oState 			= "x"  'JF Seems to work with bogus State as long as Zip is correct, no state in DB
'+jf 8/23/01
	oLCID=rsAdmin.Fields("AdminLCID")
'-jf 8/23/01

	If inStr(rsAdmin.Fields("adminUsPsUserName"),",") Then
	
	ParseUsername = Split(rsAdmin.Fields("adminUsPsUserName"),",")
	
	uspsUsername = ParseUsername(0)
	upsUsername = ParseUsername(1)
	'+JF added for Canada
	if inStr(mid(rsAdmin.Fields("adminUsPsUserName"),len(uspsUsername)+2),",") Then
		CanadaPostRefNum = ParseUsername(2)
	end if
	'-JF
	Else
	uspsUsername = rsAdmin.Fields("adminUsPsUserName")
	End If

	If InStr(RsAdmin.Fields("adminUsPsPassword"),",") Then
     	ParseLogin = Split(RsAdmin.Fields("adminUsPsPassword"),",")
       	uspsPassword	= ParseLogin(0)
       	upsPassword = ParseLogin(1)
 	Else
    	uspsPassword = RsAdmin.Fields("adminUsPsPassword")
    End If
    If Application("AppName")= "StoreFrontAE" Then
		If NOT isnull(RsAdmin.Fields("ltlUN"))then  
		  dim ltlID,ltlEmail
		  ltlID = RsAdmin.Fields("ltlUN")
		  ltlEmail = RsAdmin.Fields("ltlEmail")
		end if   
    End If 
    
   	If iTotalPur = 0 and posit > 0 Then
    	getShipping = "0"
    	Exit Function
	elseIf ((Trim(rsAdmin.Fields("adminFreeShippingIsActive")) = "1") AND (cDbl(rsAdmin.Fields("adminFreeShippingAmount")) <= cDbl(sTotalPrice))) Then
	  if posit > 0 then	
		getShipping = "0"'
		sShipping = "0"
	    Exit Function
	  end if  
	end if
  
	'Collect UPS Error Message
	arrCheck 	= split(aCheck, "|")
	If aCheck 	<> "" Then
		sCheck = arrCheck(0)
		sErrMsg = arrCheck(1)
	Else
		sCheck = ""
	End If
	
	'if UPS Failed, use secondary Shipping method
	If sCheck <> "FAIL" Then
		iShipType = rsAdmin.Fields("adminShipType")
	Else
		Response.Write "<font face=verdana size=2><b><center>Carrier Based Shipping has failed and the secondary Shipping Method is being used<br>Error Description: " & sErrMsg & "</center></b></font>"
		iShipType = rsAdmin.Fields("adminShipType2")
		sShipMethodName = "	Regular Shipping" 
		'Guard against an infanite loop
		If iShipType = 2 Then iShipType = 1  'changes it to valuebased shipping
	End If

	Set rsShipping = Server.CreateObject("ADODB.RecordSet")

	If iShip > 0 Then
	'Select Shipping Type
	Select Case iShipType
		'Valuebased Shipping
		Case 1 
		If iTotalPur > 0 Then
		'TotaL_with_Attributes = iTotalPur 'Upadating fix to #286
		
								Set rsProdShipping = Server.CreateObject("ADODB.RecordSet")
				SQL = "SELECT * FROM sfTmpOrderDetails WHERE odrdttmpShipping >= 1 AND odrdttmpSessionID = " & Session("SessionID")
				
				IF Application("AppName") = "StoreFrontAE" Then 'SFAE
					SQL = "Select * FROM sfTmpOrderDetails as A LEFT JOIN sfTmpOrderDetailsAE as B ON A.odrdttmpID = B.odrdttmpaeID "
					SQL = SQL & "WHERE odrdttmpShipping >= 1 AND odrdttmpSessionID = " & Session("SessionID")
				End IF 
				iTotalPur=0
				If vDebug = 1 Then Response.Write SQL & "<br><br>"
				rsShipping.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
				'Total product shipping
				Do While NOT rsShipping.EOF
					sProdID		= rsShipping.Fields("odrdttmpProductID")
					iQuantity	= rsShipping.Fields("odrdttmpQuantity")
					
					If Application("AppName") = "StoreFrontAE" Then 'SFAE
						allQty = rsShipping.Fields("odrdttmpQuantity")
						boQty =  rsShipping.Fields("odrdttmpBackOrderQTY")
						shpQty = allQty - boQty
						If sType = "BackOrder" Then iQuantity = boQty
						If sType = "Shipped" Then iQuantity = shpQty
						'Response.Write "<BR> boqty:" & boqty
						'Response.Write "<BR> shpqty:" & shpqty
						'Response.Write "<BR> allqty:" & allqty		
						'if iQuantity <= 0 then 
						'	GetShipping = 0
						'	Exit function
						'end if	
				   End If			
					If iQuantity > 0 Then 
					SQL = "SELECT ProdPrice,ProdSalePrice,ProdSaleIsActive,ProdShipIsActive FROM sfProducts WHERE prodShipIsActive = 1 AND prodID = '" & sProdID & "'"
					rsProdShipping.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText 
					
					If Not rsProdShipping.EOF Then 
										
						if rsProdShipping("ProdShipIsActive") then

							if rsProdShipping("ProdSaleIsActive") then
								
								iTotalPur=iTotalPur + rsProdShipping("ProdSalePrice")*iQuantity
							else
								iTotalPur=iTotalPur + rsProdShipping("ProdPrice")*iQuantity
							end if
			                    '#286 next line
			                    iTotalpur = getPriceWithAtt(iTotalpur, rsShipping.Fields("odrdttmpID"))
						End If
					End If
					rsProdShipping.Close 
					end if
					rsShipping.MoveNext 
				Loop
rsShipping.close
set rsProdShipping = nothing

'		If TotaL_with_Attributes >  iTotalPur then 
'		  iTotalPur = TotaL_with_Attributes '#286
'		end if	
			if iTotalPur > 0 Then 
			SQL = "SELECT valShpPurTotal, valShpAmt FROM sfValueShipping WHERE valShpPurTotal <= " & FormatNumber(iTotalPur,2) & " ORDER BY valShpPurTotal DESC" '
			
			If vDebug = 1 Then Response.Write SQL & "<br><br>"
		
			rsShipping.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
			sShipping = rsShipping.Fields("valShpAmt")
			Else
			sShipping = 0
			End If
			Else
			sShipping = 0
			End If
'-----------------------------------------------------------------			
		'Carrier shipping
		Case 2 
			Set rsProdShipping = Server.CreateObject("ADODB.RecordSet")
			 Dim totshpQty
                        Dim totBOQty
			'Gather Product Information
			
			SQL = "SELECT odrdttmpProductID, odrdttmpQuantity FROM sfTmpOrderDetails WHERE odrdttmpShipping >= 1 AND odrdttmpSessionID = " & Session("SessionID")
			
			IF Application("AppName") = "StoreFrontAE" Then 'SFAE
				SQL = "Select * FROM sfTmpOrderDetails as A LEFT JOIN sfTmpOrderDetailsAE as B ON A.odrdttmpID = B.odrdttmpaeID "
				SQL = SQL & "WHERE odrdttmpShipping >= 1 AND odrdttmpSessionID = " & Session("SessionID")
			End IF 
			
			If vDebug = 1 Then Response.Write SQL & "<br><br>"
			rsShipping.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
			
			'Total Product weights, Get largest Length, width, height
			iLength = 0
			iWidth = 0
			iHeight = 0
			noShip=1
			Do While NOT rsShipping.EOF
				sProdID = rsShipping.Fields("odrdttmpProductID")
				iQuantity = rsShipping.Fields("odrdttmpQuantity")
				
				If Application("AppName") = "StoreFrontAE" Then 'SFAE
						allQty = rsShipping.Fields("odrdttmpQuantity")
						boQty =  rsShipping.Fields("odrdttmpBackOrderQTY")
						totBOQty = totBOQty + boQty
                        shpQty = allqty - boQty
                        totshpQty = totshpQty + shpQty
						If sType = "BackOrder" Then iQuantity = boQty
						If sType = "Shipped" Then iQuantity = shpQty
						'Response.Write "<BR> boqty:" & boqty
						'Response.Write "<BR> shpqty:" & shpqty
						'Response.Write "<BR> allqty:" & allqty	
'						if iQuantity <= 0 then 
'							GetShipping = 0
'							Exit function
'						end if		
				End If
			If iQuantity > 0 Then			   									
				SQL = "SELECT prodWeight, prodLength, prodWidth, prodHeight FROM sfProducts WHERE prodShipIsActive = 1 AND prodID = '" & sProdID & "'"
				If vDebug = 1 Then Response.Write SQL & "<br><br>"
				rsProdShipping.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText 
				If Not rsProdShipping.EOF Then 
				noShip=0
					If rsProdShipping.Fields("prodWeight") <> "" AND Not IsNull(rsProdShipping.Fields("prodWeight")) Then
						iWeight = CDbl(iWeight) + (CDbl(rsProdShipping.Fields("prodWeight")) * CDbl(iQuantity))
					Else 
						iWeight = CDbl(iWeight)	
					End If	
					
					If (Not IsNull(rsProdShipping.Fields("prodLength"))) AND rsProdShipping.Fields("prodLength") <> "" Then
						if rsProdShipping.Fields("prodLength") > iLength then	
						iLength = rsProdShipping.Fields("prodLength")
						end if
					End If
					If 	(Not IsNull(rsProdShipping.Fields("prodWidth"))) AND rsProdShipping.Fields("prodWidth") <> "" Then
						if rsProdShipping.Fields("prodLength") > iWidth then	
						iWidth =rsProdShipping.Fields("prodWidth")
						end if
					End If
					If 	(Not IsNull(rsProdShipping.Fields("prodHeight"))) AND rsProdShipping.Fields("prodHeight") <> "" Then
						iHeight = iHeight + (rsProdShipping.Fields("prodHeight") * CDbl(iQuantity))
					End If	
				End If
				rsProdShipping.Close 
			end if	
				rsShipping.MoveNext 
			Loop
			rsShipping.Close
			
						if noShip=1 then 'JF
			iShipType = rsAdmin.Fields("adminShipType2")
			'sShipMethodName = "	Regular Shipping" 
			getShipping=0
				exit function
					closeObj(upsObj)
					closeObj(rsAdmin)
					closeObj(rsShipping)
					closeObj(rsProdShipping)

			end if

			'?
			If sShipping <> "FAIL" Then	
				Dim upsObj,DHLObj, canadapostObj, rsCountry, dCountryFull 'JF added for Canada & DHL

				SQL = "SELECT shipMethod,shipCode, shipRates FROM sfShipping WHERE shipID = " & sShipCode
				If vDebug = 1 Then Response.Write SQL & "<br><br>"
				rsShipping.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
				
				if trim(rsShipping.Fields("shipCode"))= "FedEx" or instr(rsShipping.Fields("shipMethod"),"UPS") then
					'Set the aceptable weight range of UPS
					'If iWeight <= 0 Then iWeight = 1
					If iWeight > 150 Then 
						
						Pkg = FormatNumber(iWeight/150,2)
						iWeight = 150
					End If	
				
				end if
				'On Error Resume Next				
'+jf 8/23/01
				On Error Resume Next				
'-jf 8/23/01

				If Trim(rsShipping.Fields("shipCode")) = "FedEx" Then
						Set obj = Server.CreateObject("SFServer.FedEx")
						obj.setFunc 
						obj.setScreen 
						obj.setOriginZip oZip
						obj.setOriginCountryCode "US"
						obj.setDestZip dZip
						obj.setDestCountyCode dCountry
						'---------------------------------------------------------------------------------
						'Issue #255 
						if iWeight > 0 and iWeight <1 then 
							iWeight=1
						end if
						'---------------------------------------------------------------------------------
						obj.setWeight iWeight
						obj.queryFedEx 
						
						sShipping = obj.getTotalCharge 
					
						Set obj = Nothing 
						If Err.number <> 0 Then
							'+jf 8/23/01
							If Err.number = 424 or Err.number=-2147467259 Then 
							'-jf 8/23/01	
								sShipping = "The FedEx component is not properly installed."
							Else
								sShipping = Err.number & " " & Err.description 
							End If
						End If
				'+jf 8/23/01
					On Error goto 0					
				'-jf 8/23/01

				ElseIf Instr(rsShipping.Fields("shipCode"), "USPS") <> 0 Then
						dCountry = ucase(dCountry)
					    If rsShipping.Fields("shipCode") = "USPSParcel" Then
							sShipping =  getDomesticUsPsShipping("Parcel",dZip,"",iLength,iWidth,iHeight,sProdID,1,uspsUsername,uspsPassword,oCountry,oZip,iWeight,dCountry) 'scontainer = "none"
						ElseIf rsShipping.Fields("shipCode") = "USPSExpress" Then
							sShipping =  getDomesticUsPsShipping("Express",dZip,"",iLength,iWidth,iHeight,sProdID,1,uspsUsername,uspsPassword,oCountry,oZip,iWeight,dCountry) 'scontainer = "none"
						ElseIf rsShipping.Fields("shipCode") = "USPSPriority" Then
							sShipping = getDomesticUsPsShipping("Priority",dZip,"",iLength,iWidth,iHeight,sProdID,1,uspsUsername,uspsPassword,oCountry,oZip,iWeight,dCountry) 'scontainer = "none"
						ElseIf rsShipping.Fields("shipCode") = "USPSInternational" Then
						  SQL = "SELECT loclctryName FROM sfLocalesCountry WHERE loclctryAbbreviation = '" & dCountry & "'"
						  Set rsCountry = Server.CreateObject("ADODB.RecordSet")
						  rsCountry.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
						  sShipping = getinternationalusps(trim(rsCountry("loclctryName")),"other packages",iLength,iWidth,iHeight,sProdID,1,uspsUsername,uspsPassword,oCountry,oZip,iWeight)
						  closeObj(rsCountry)
					   End If
 '==============================================					   
'+JF for Canada
				ElseIf Instr(rsShipping.Fields("shipMethod"), "Canada Post") <> 0 Then
'+jf 8/23/01					
					sShipping=getCanadaPostRates(rsShipping.Fields("shipCode"),dZip,dCity,dState,iLength,iWidth,iHeight,CanadaPostRefNum,iWeight,dCountry,oCountry,oLCID)
'-jf 8/23/01					
					
'-JF
'==============================================

'==============================================
'+JF for DHL
				ElseIf Instr(rsShipping.Fields("shipCode"), "DHL") <> 0 Then
				  	'If NOT(upsUsername <> "" AND upsPassword <> "") Then
				  	'	sShipping = "Not Registered for UPS Rate Service"
				  	'Else
				  	iWeight=CalculateDHLDimWeight(iWeight,iLength,iWidth,iHeight,dCountry,oCountry)
				 	Set DHLObj = Server.CreateObject("sfServer.DHL")
					DHLObj.setMH_1_MsgTyp "BkRateReq"
					DHLObj.setMH_1_SvcFunc "BookRate"
					DHLObj.setMH_1_Svc "rating"
					DHLObj.setMH_1_ClntRef "rateinquirymsg"
					DHLObj.setMH_1_MsgTm "1998-10-09 18:00:16.000+00:00"
					DHLObj.setMH_1_OrgnNm "K111111"
					DHLObj.setMH_1_MsgVsn "01"
					DHLObj.setMH_1_MsgRvsn "01"
					DHLObj.setShpShp_1_WghtUnit "L"
					DHLObj.setShpShp_1_DimUnit "I"
					DHLObj.setShpShp_1_Pcs "1"
					DHLObj.setShpr_1_City oCity
					DHLObj.setShpr_1_DvsnNm oState
					DHLObj.setShpr_1_PostCd oZip
					DHLObj.setShpr_1_CtryCd oCountry
					DHLObj.setCnsg_1_City dCity
					DHLObj.setCnsg_1_DvsnNm dState
					DHLObj.setCnsg_1_PostCd dZip
					DHLObj.setCnsg_1_CtryCd dCountry
					DHLObj.setproduct "non_dutiable"
					'---------------------------------------------------------------------------------
					'Issue #255 
					if iWeight > 0 and iWeight <0.1 then 
						iWeight=0.1
					end if
					'---------------------------------------------------------------------------------
					DHLObj.setweight iWeight
					
					dim tempResultStr
					
					DHLObj.QueryDHL	
					tempResultStr = DHLObj.getTotalCharge
			'+jf 8/23/01		
					sShipping=trim(parseTempString(tempResultStr))
					if MyNumeric(sShipping) then
					'do nothing
					else
					sShipping="DHL service not available for this order."
					end if
					'Ups Error Handling
					If sShipping = "" Then
						sShipping = DHLObj.ErrorMessage
					End If
					set DHLobj =nothing
					If Err.number = 424 Then 
						sShipping = "The DHL component is not properly installed."
					ElseIf sShipping = "" Then
						sShipping = "The DHL component is not properly installed."
					End If
			'-jf 8/23/01
'-JF			
'+++++++++++++++++++++++++++++++++++++++++++++++++++
				ElseIf rsShipping.Fields("shipCode")= "LTL" then
				
				dim iClass,sContainer, arrLtl, bLTL
				iClass = 50
				sContainer ="Box"
				bLTL =true
				
				sShipping = get_ltl(dZip, sContainer, iLength, iWidth, iHeight, sProdId, 1, ltlid, ltlEmail, oZip, iWeight, sType, totBOQty, totshpQty, C_BGCOLOR2, C_BGCOLOR4, C_BGCOLOR5, C_BKGRND2, C_BKGRND4, C_BKGRND5, C_FONTFACE2, C_FONTFACE4, C_FONTFACE5, C_FONTCOLOR2, C_FONTCOLOR4, C_FONTCOLOR5, C_FONTSIZE4, C_FONTSIZE5, rsShipping.Fields("shiprates"), iTotalpur, iPremiumShipping, aCheck, dCity, dState, dCountry, sTotalPrice, iship, sShipmethodname)
				
				
				
				  if instr(sShipping,"|")> 0  then
                    'Response.Write sShipping                				   
				    arrLtl = split(sShipping,"|")
				    sShipping = arrltl(0)
				    
				    if isnumeric(sShipping) then
				     sShipping = CDbl(sShipping)
				    'remove the Doller sign
				    end if
				  end if  
				
				Else
				  	
				  	If NOT(upsUsername <> "" AND upsPassword <> "") Then
				  		sShipping = "Not Registered for UPS Rate Service"
				  	Else
				 	Set upsObj = Server.CreateObject("sfServer.UPS")
					upsObj.setActionCode "3"
					upsObj.setServiceLevelCode Trim(rsShipping.Fields("shipCode"))
					upsObj.setRateChart "Regular Daily Pickup"
					upsObj.setShipperPostalCode oZip
					upsObj.setConsigneePostalCode dZip
					upsObj.setConsigneeCountry dCountry
					'---------------------------------------------------------------------------------
					'Issue #255 
					if iWeight > 0 and iWeight <0.01 then 
						iWeight=0.01
					end if
					'---------------------------------------------------------------------------------
					upsObj.setPackageActualWeight iWeight
					upsObj.setResidentialInd "0"
					upsObj.setPackagingType "00"
					upsObj.setLength iLength
					upsObj.setWidth iWidth
					upsObj.setHeight iHeight
					upsObj.QueryUPS					
					sShipping = upsObj.getTotalCharge
					'Response.write "<br>Weight:" & iWeight
					'Ups Error Handling
					
					If sShipping = "" Then
						sShipping = upsObj.ErrorMessage
					End If
					If Err.number = 424 Then 
						sShipping = "The UPS component is not properly installed."
					ElseIf sShipping = "" Then
						sShipping = Err.number & " " & Err.description  
					End If 
				End If
				'On Error GoTo 0	

			End If 
			End If
			sShipping = sShipping
			'+jf 8/23/01
			if trim(sShipping)="" then
			
			sShipping="Unable to connect to " & rsShipping.Fields("shipCode") & " shipping"
			end if
			'-jf 8/23/01

			'Response.Write "hello:" & sShipping & ":olleh"
			if isnumeric(sShipping) then
			 If MyNumeric(sShipping) Then
			   	sShipping = CDbl(sShipping) * rsShipping.Fields("ShipRates")
			
			 Else
			  sShipping = getShipping(iTotalPur, iPremiumShipping, "FAIL" & "|" & sShipping, dCity, dState, dZip, dCountry, sTotalPrice,sType)
			  bLTL=False 'JF once it has gone to secondary shipping, make sure it doesn't still think it is LTL
			 end if
			else	
				sShipping = getShipping(iTotalPur, iPremiumShipping, "FAIL" & "|" & sShipping, dCity, dState, dZip, dCountry, sTotalPrice,sType)
				bLTL=False 'JF once it has gone to secondary shipping, make sure it doesn't still think it is LTL
			 End If 
			
			Set upsObj = nothing				

 

'-----------------------------------------------------------------			
		'Product Based
		Case 3 
			Set rsProdShipping = Server.CreateObject("ADODB.RecordSet")
				SQL = "SELECT * FROM sfTmpOrderDetails WHERE odrdttmpShipping >= 1 AND odrdttmpSessionID = " & Session("SessionID")
				
				IF Application("AppName") = "StoreFrontAE" Then 'SFAE
					SQL = "Select * FROM sfTmpOrderDetails as A LEFT JOIN sfTmpOrderDetailsAE as B ON A.odrdttmpID = B.odrdttmpaeID "
					SQL = SQL & "WHERE odrdttmpShipping >= 1 AND odrdttmpSessionID = " & Session("SessionID")
				End IF 
				
				If vDebug = 1 Then Response.Write SQL & "<br><br>"
				rsShipping.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
				'Total product shipping
				Do While NOT rsShipping.EOF
					sProdID		= rsShipping.Fields("odrdttmpProductID")
					iQuantity	= rsShipping.Fields("odrdttmpQuantity")
					
					If Application("AppName") = "StoreFrontAE" Then 'SFAE
						allQty = rsShipping.Fields("odrdttmpQuantity")
						boQty =  rsShipping.Fields("odrdttmpBackOrderQTY")
						shpQty = allQty - boQty
						If sType = "BackOrder" Then iQuantity = boQty
						If sType = "Shipped" Then iQuantity = shpQty
						'Response.Write "<BR> boqty:" & boqty
						'Response.Write "<BR> shpqty:" & shpqty
						'Response.Write "<BR> allqty:" & allqty		
						if iQuantity <= 0 then 
							GetShipping = 0
							Exit function
						end if	
				   End If			
					
					SQL = "SELECT prodShip FROM sfProducts WHERE prodShipIsActive = 1 AND prodID = '" & sProdID & "'"
					rsProdShipping.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText 
					
					If Not rsProdShipping.EOF Then 
						sShipping = CDbl(sShipping) + (CDbl(rsProdShipping.Fields("prodShip")) * CDbl(iQuantity))
					End If
						
					rsProdShipping.Close 
					rsShipping.MoveNext 
				Loop
		End Select

	End If	
	'Default Minimum Shipping and premium shipping
	If iShipMin = "" Then iShipMin = 0
	If iSpcShipAmt = "" Then iSpcShipAmt = 0
	
	'Add in Premium shipping, check minimum shipping, apply shipping sale
	If iPremiumShipping = "1" Then 								
		sShipping = CDbl(iSpcShipAmt) + CDbl(sShipping)
	End If	
	
	If isnumeric(sShipping) then '#313
    	 If CDbl(sShipping) < CDbl(iShipMin) Then sShipping = iShipMin
	End If
	If Not IsNull(Pkg) AND Pkg > 0 Then
		sShipping = sShipping * Pkg
	End If
	If bLtl = true Then
	   getShipping = "@" & sShipping & "|" & arrLtl(1)
	    session("sltl") = "@" & sShipping & "|" & arrLtl(1)
	    
	else

	 getShipping = formatNumber(sShipping, 2)  
	End if   
	
	'End If
	closeObj(upsObj)
	closeObj(rsAdmin)
	closeObj(rsShipping)
	closeObj(rsProdShipping)
	'
	 
End Function
Function getPriceWithAtt(iTotalPrice, DetID)
Dim sql
Dim tmpAmt
Dim rsProdShipping
Dim rsProdShipping2
Set rsProdShipping = Server.CreateObject("ADODB.RecordSet")
Set rsProdShipping2 = Server.CreateObject("ADODB.RecordSet")
tmpAmt = iTotalPrice
   
   sql = "SELECT * from sfTmpOrderAttributes where odrattrtmpOrderDetailID = " & DetID
      rsProdShipping.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
        Do While Not rsProdShipping.EOF
           sql = "SELECT * from sfAttributeDetail where attrdtID= " & rsProdShipping("odrattrtmpAttrID")
            rsProdShipping2.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
            If rsProdShipping2("attrdttype") = 1 Then
            tmpAmt = tmpAmt + rsProdShipping2("attrdtPrice")
            ElseIf rsProdShipping2("attrdttype") = 2 Then
            tmpAmt = tmpAmt - rsProdShipping2("attrdtPrice")
            End If
    rsProdShipping2.Close
    rsProdShipping.MoveNext
        Loop
         getPriceWithAtt = tmpAmt
         rsProdShipping.Close
         Set rsProdShipping = Nothing
         Set rsProdShipping2 = Nothing
End Function

Function CalculateDHLDimWeight(ActualWeight,L,W,H,DestCountry,OrigCountry) 'jf - dhl
	dim dimWeight

	if trim(DestCountry)=trim(OrigCountry) then
		dimWeight=(L*W*H)/194
	else
		dimWeight=(L*W*H)/166
	end if

	if dimWeight>ActualWeight then
		CalculateDHLDimWeight=dimWeight
	else
		CalculateDHLDimWeight=ActualWeight
	end if
end function

Function parseTempString(stringtoparce) 'jf - dhl
	if instr(stringtoparce,"Total Charge") then
		stringtoparce=mid(stringtoparce,instr(stringtoparce,"Total Charge")+12)
		if instr(stringtoparce,"$") then
			stringtoparce=mid(stringtoparce,instr(stringtoparce,"$"))
		end if
		if instr(stringtoparce,".") then
			stringtoparce=mid(stringtoparce,2,instr(stringtoparce,".")+1)
		end if
	end if				
	parseTempString=stringtoparce	
end Function
'-----------------------------------------------------------------------
' Generates a unique password
'-----------------------------------------------------------------------
Function generatePassword
	Dim sPassword,Random_Number_Min,Random_Number_Max
  	
  	Randomize
	Random_Number_Min = 10000000
	Random_Number_Max = 99999999

	sPassword = Int(((Random_Number_Max-Random_Number_Min+1) * Rnd) + Random_Number_Min)
	generatePassword = sPassword
End Function

'-----------------------------------------------------------------------
'Stores Shipping in temp orders table 
'-----------------------------------------------------------------------
Sub putShipping(iShipping)
      Dim SQLText, myID, adExecuteNoRecords
	  	  adExecuteNoRecords = &H00000080
	  	  myID = Session("SessionID")
      	  SQLText = "UPDATE sfTmpOrderDetails" & _
          		        " SET odrdttmpShipping = " & cdbl(iShipping) & _
                	        " WHERE (odrdttmpSessionID = " & myID & ")"
         cnn.execute SQLText, ,adCmdText + adExecuteNoRecords
End Sub

'-----------------------------------------------------------------------
'Retrive CustPassword
'-----------------------------------------------------------------------
Function getPassword(iCustID)
	Dim sSQL, sPassword, rsPass
	sPassword = ""
	If iCustID <> "" Then
		Set rsPass = Server.CreateObject("ADODB.Recordset")
			sSQL = "SELECT custPasswd FROM sfCustomers WHERE custID = " & iCustID
			rsPass.Open sSQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
		If Not (rsPass.EOF And rsPass.BOF) Then	sPassword = rsPass.Fields("custPasswd")
			getPassword = trim(sPassword)
	Else
		getPassword = ""
	End If 
	closeObj(rsPass)
End Function

'-----------------------------------------------------------------------
'getDomesticUsPsShipping
'returns: 
'	1.No Account = User does not have a user name and password with usps
'	2.Not Domestic = Not a domestic ship
'	3.TooBig = Total girth is over 108 inches 
'	4.TooHeavy = package is over 70 pounds
'	5.NoQuantity = No items
'	6.NoDiminsions = no diminsions set for product 
'	7.price = no errors
'	8. ParcelServeOnly = Package is to big for Express or Priority
'
'Inputs:
'sService = "Express"  "Parcel" "Priority" ''' 
'dzip = Destination Zip code
'sContainer = "None" unless package is shipped in a isps container
'-----------------------------------------------------------------------
Function getDomesticUsPsShipping(sService,dZip,sContainer,iLength,iWidth,iHeight,sProdID,iQuantity,uspsUsername,uspsPassword,oCountry,oZip,iWeight,dCountry)
	Dim strSize,sLen,uPounds,uOunces,sMachine,httpconn,xmlDoc,srvr,Port,Path,query,msg,contentType,proxyServer,proxyPort,sRespcode,resp,RequestLevel,PackageLevel,PackageElementLevel,t

	  oCountry =trim(lcase(oCountry))
	  If uspsPassword = "" or uspsUsername = "" Then
	    getDomesticUsPsShipping = "Not Registered for USPS Rate Service"
	    Exit Function	  
	  Elseif oCountry <> "us" And oCountry <> "ca" And oCountry <> "mx" Then
   	    getDomesticUsPsShipping ="Domestic Rates only available for destinations within the United States"
	    Exit Function
	  elseif lcase(dCountry) <> "us" then   
	     getDomesticUsPsShipping ="Domestic Rates only available for destinations within the United States"
	    Exit Function	  
	  End If   
	  
			'Gather Product Information
			
			
			
				If iQuantity < 1 Then
				  getDomesticUsPsShipping ="NoQuantity"
	              Exit Function
				End If	
				
				
		If iHeight = 0 Then iHeight = 1
		If iWidth = 0 Then  iWidth = 1
		If iLength = 0 Then iLength = 1
			
	    strSize= getsize(iWidth, iLength,iHeight)
	    If strSize = "TooBig" Then
		    getDomesticUsPsShipping = "TooBig"
		    Exit Function
		End If	
	  	
'	if iWeight < 1.0 then
'	 iWeight = 1.0
'	end if 
				If iWeight > 70 Then 
				'Dim Pkg
				Pkg = FormatNumber(iWeight/70,2)
				iWeight = 70
			End If	

	On Error Resume Next
'		upounds = Int(iWeight)
'		uOunces = iWeight mod Int(iWeight)
		
		'---------------------------------------------------------------------------------
		'Issue #255 
		if iWeight <> 0 then 
			upounds = Int(iWeight)
			uOunces = int((iWeight - uPounds)*16)
		
			'will break if both are 0 which will 
			'happen if weight is so little that it rounded down to 0
			if uOunces=0 and upounds=0 then
			uOunces=1
			end if
		else 'let it fail if there is a true weight of zero
			upounds = 0
			uOunces = 0
		end if
		'---------------------------------------------------------------------------------
		If sContainer = "" then 
			sContainer = "None"
		End If  
	Select case sService
		Case "Parcel"
			sMachine=""
			If iHeight > 2.99 And iHeight < 17.01 Then
				sMachine ="True"
			Else
				sMachine = "False"
			End If  
			If sMachine <> "False" Then
				If iLength > 5.99 And iLength < 34.01 Then
					sMachine ="True"
				Else
					sMachine = "False"
				End If  
			End If 
			
			If sMachine <> "False" Then 
				If iWidth > .24 And iWidth < 17.01 Then
					sMachine ="True"
				Else
					sMachine = "False"
				End If
			End if
			If sMachine <> "False" Then   
				If iWeight > .07 And iWeight < 35.01 Then
					sMachine ="True"
				Else
					sMachine = "False"
				End If
			End If 
      Case "Express"  
			sMachine =""
	  Case "Priority"
			sMachine ="" 
	End Select
Set httpconn = server.createobject("httpcom.chttpcom")		
set xmlDoc = Server.CreateObject("MSXML.DOMDocument")

	Set RequestLevel = xmlDoc.createElement("RateRequest") 
	Set RequestLevel = xmlDoc.createElement("RateRequest")
		RequestLevel.setAttribute "USERID", uspsUsername
		RequestLevel.setAttribute "PASSWORD", uspsPassword
	Set PackageLevel = xmlDoc.createElement("Package")
		PackageLevel.setAttribute "ID", sProdID
	Set PackageElementLevel = xmlDoc.createElement("Service")
	Set t = xmlDoc.createTextNode(sService)                  
		PackageElementLevel.appendChild (t)
	Call PackageLevel.appendChild(PackageElementLevel)
	Set PackageElementLevel = xmlDoc.createElement("ZipOrigination")
		Set t = xmlDoc.createTextNode(ozip)                            
		PackageElementLevel.appendChild (t)
	Call PackageLevel.appendChild(PackageElementLevel)
	Set PackageElementLevel = xmlDoc.createElement("ZipDestination")
	Set t = xmlDoc.createTextNode(dzip)                               
		PackageElementLevel.appendChild (t)
	Call PackageLevel.appendChild(PackageElementLevel)
	Set PackageElementLevel = xmlDoc.createElement("Pounds")
	Set t = xmlDoc.createTextNode(uPounds)                                   
		PackageElementLevel.appendChild (t)
	Call PackageLevel.appendChild(PackageElementLevel)
	Set PackageElementLevel = xmlDoc.createElement("Ounces")
	Set t = xmlDoc.createTextNode(uOunces)    ')                             
		PackageElementLevel.appendChild (t)
	Call PackageLevel.appendChild(PackageElementLevel)
	Set PackageElementLevel = xmlDoc.createElement("Container")
	Set t = xmlDoc.createTextNode(sContainer)                      
		PackageElementLevel.appendChild (t)
	Call PackageLevel.appendChild(PackageElementLevel)
	Set PackageElementLevel = xmlDoc.createElement("Size")
	Set t = xmlDoc.createTextNode(strSize)                    
		PackageElementLevel.appendChild (t)
	Call PackageLevel.appendChild(PackageElementLevel)
	Set PackageElementLevel = xmlDoc.createElement("Machinable")
	Set t = xmlDoc.createTextNode(sMachine)                   
		PackageElementLevel.appendChild (t)
	Call PackageLevel.appendChild(PackageElementLevel)
	Call RequestLevel.appendChild(PackageLevel)

	Call xmlDoc.appendChild(RequestLevel)
	'Call xmlDoc.save(Server.MapPath("/Sendme.xml"))

	srvr		= "Production.shippingapis.com" 'You must modify the Server name
	Port		= 80
	Path		= "/ShippingAPI.dll" 'You must modify the Path name
	query		= ""
	msg			= "API=Rate&XML=" & URL_Encode(xmlDoc.xml)
	contentType =  "application/x-www-form-urlencoded"
	proxyServer = "" 'You must modify proxy server or leave
	proxyPort	= 0 'Modify the proxy port, if necessary. Leave blank if no

	resp = ""
	on error resume next
	Err.Number = 0
	sRespcode = httpConn.GetResponse(srvr, Port, Path, query, msg, contentType, proxyServer, proxyPort, resp)
	If Err.Number <> 0 Then
		
		If Err.number = 424 Then
			getDomesticUsPsShipping = "The USPS component is not properly installed."
		Else
			getDomesticUsPsShipping = Err.description 
		End If
		Set xmlDoc = Nothing
		Set httpConn = Nothing

		Set xmlDoc = Nothing
		Set httpConn = Nothing
		closeObj(rsAdmin)
		closeObj(rsShipping)
		closeObj(rsProdShipping)
		Exit Function
	Else
		getDomesticUsPsShipping = GetPrice(resp)
		
	End If
		Set xmlDoc = Nothing
		Set httpConn = Nothing
		Set xmlDoc = Nothing
		Set httpConn = Nothing
	
End Function
'jf CanadaPost
'+jf 8/23/01
Function getCanadaPostRates(sService,dZip,dCity,dState,iLength,iWidth,iHeight,CanadaPostID,iWeight,dCountry,oCountry,oLCID)
'-jf 8/23/01
	Dim strSize,sLen,uPounds,uOunces,sMachine,httpconn,xmlDoc,srvr,Port,Path,query,msg,contentType,proxyServer,proxyPort,sRespcode,resp,RequestLevel,PackageLevel,PackageElementLevel,t
	Dim xmlRequest
	dim objParcel
	'Dim Pkg
	dim HttpObj
	dim DomDoc1
	dim strXML
	dim ServiceCode
	dim NumOfServices
	dim	aObjDOMNodeList
	dim counter
	dim shipAmt
'+jf 8/23/01	
	dim oLang


	oCountry =trim(lcase(oCountry))
	If CanadaPostID = "" Then
	  getCanadaPostRates = "Not Registered for Canada Post Service"
	  Exit Function	  
'	Elseif oCountry <> "ca" Then
'  	  getCanadaPostRates ="Rates only available for shipments originating within Canada"
'	  Exit Function
	End If   
	'convert to centimeters		
	iHeight=iHeight * 2.54
	iWidth=iWidth * 2.54
	iLength=iLength * 2.54

	iWeight=iWeight * 0.453592 'Convert from LB to KG
	
'-jf 8/23/01		
	If iHeight = 0 Then iHeight = 1
	If iWidth = 0 Then  iWidth = 1
	If iLength = 0 Then iLength = 1
	
	'if iWeight < 1.0 then
	' iWeight = 1.0
	'end if 
	
	If iWeight > 29 Then 
		Pkg = FormatNumber(iWeight/29,2)
		iWeight = 29
	End If	

	
	'+jf 8/23/01	
	if oLCID=3084 then
		oLang="fr"
	else
		oLang="en"
	end if
	xmlRequest = " <?xml version=""1.0"" ?> " & _
					"<eparcel>    " & _
					"<language> " & oLang & " </language>    " & _
					"<ratesAndServicesRequest>  " & _
					"  <merchantCPCID> " & CanadaPostID & " </merchantCPCID>   " & _
					"  <lineItems>  " & _
					"    <item>  " & _
					"      <quantity> 1 </quantity>  " & _
					"      <weight>  " & iWeight & "  </weight>  " & _
					"      <length>  " & iLength & "  </length>                  " & _
					"      <width>  " & iWidth & "  </width>  " & _
					"      <height>  " & iHeight & "  </height>  " & _
					"      <description> Package to Ship </description>  " & _
					"    </item>           " & _
					"  </lineItems>          " & _
					"  <city>  " & dCity & " </city>          " & _
					"  <provOrState>  " & dState & " </provOrState>  " & _
					"  <country> " & dCountry & " </country>  " & _
					"  <postalCode>  " & dZip & " </postalCode>  " & _
					"</ratesAndServicesRequest>  " & _
					"</eparcel>"
					
	on error resume next
'-jf 8/23/01					
	Set HttpObj = Server.CreateObject("MSXML2.ServerXMLHTTP")
	Set DomDoc1 = Server.CreateObject("MSXML2.DomDocument")

	HttpObj.open "POST", "http://216.191.36.73:30000"
	HttpObj.send xmlRequest
   
	strXML = HttpObj.responseText
'+JF 8/23/01
    'Response.Write ":" & Err.Number & ":"
    'Response.end
    If Err.Number <> 0 Then
		If Err.number = -2147467259 Then
			getCanadaPostRates = "The Canada Post component is not properly installed."
		Else
			getCanadaPostRates = Err.description 
		End If

		Exit Function
	End if
	on error goto 0
'-JF 8/23/01

	DomDoc1.loadXML strXML
	While DomDoc1.readyState <> 4
    'loop through to make sure DomDoc1 is loaded
    Wend
   
   
   
   If DomDoc1.parseError.errorCode <> 0 Then
	shipAmt = "Could Not Load Data to Parse.<br>" & "Description: " & _
	DomDoc1.parseError.reason & _
	"<br>Line: " & DomDoc1.parseError.Line
   else
   
   
	DomDoc1.async = False
'+jf 8/23/01
	Set aObjDOMNodeList = DomDoc1.selectSingleNode("//statusCode")
	if aObjDOMNodeList.text <> "1" then
		
		Set aObjDOMNodeList =nothing
		Set aObjDOMNodeList = DomDoc1.selectSingleNode("//statusMessage")
		
		getCanadaPostRates = aObjDOMNodeList.text
	    Exit Function	  
	end if

	
		Set aObjDOMNodeList =nothing
		'-jf 8/23/01	
	Set aObjDOMNodeList = DomDoc1.selectNodes("//product")
	
	NumOfServices=aObjDOMNodeList.length
	Set aObjDOMNodeList =nothing
'	Response.Write NumOfServices

'	on error resume next
	counter=0
	do while counter<NumOfServices
		'Response.Write counter
		Set aObjDOMNodeList = DomDoc1.selectNodes("//product").Item(counter).Attributes(0)

		ServiceCode=aObjDOMNodeList.text
		Set aObjDOMNodeList =nothing
		if ServiceCode=sService then
			Set aObjDOMNodeList = DomDoc1.selectNodes("//product").Item(counter).childNodes(1)
			shipAmt=aObjDOMNodeList.text
			counter=NumOfServices
			Set aObjDOMNodeList =nothing
		end if
		counter=counter+1
	loop
'	response.end
'
	if trim(shipAmt)="" then
		shipAmt="Selected Canada Post service not available for your order."
	end if
End If
getCanadaPostRates=shipAmt
End Function



'-----------------------------------------------------
' URL Encode string
'-----------------------------------------------------
Function URL_Encode(s)
 Dim URL_SAFE_CHARS, outStr, i, c, res    
 URL_SAFE_CHARS = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890$-_.+!*'(),"

    outStr = ""
    For i = 1 To Len(s)
        c = Right(Left(s, i), 1) 'set c equal to ith char
        res = InStr(URL_SAFE_CHARS, c)
        If res = 0 Then
            'escape unsafe characters as %xx where xx is the
            'hexadecimal value of the character's ASCII code.
            outStr = outStr & "%" & Right("00" & CStr(Hex(Asc(c))), 2)
        Else
            outStr = outStr & c
        End If
    Next
    
    URL_Encode = outStr
    
End Function


'--------------------------------------------------------------
' GetSize function
'--------------------------------------------------------------
Function getSize(iWidth,iLength,iHeight)
	Dim iGirth, totGirth
	iGirth = iWidth + iHeight + iWidth + iHeight
	totGirth =iGirth + iLength
	'max total for girth + lenght is 108
	
	If totGirth > 108 Then 
		getSize = "OverSize"
		Exit Function
	ElseIf totGirth < 84 Then
		getSize = "Regular"
		Exit function
	Elseif totGirth >  84 And totGirth < 108 Then
       getSize = "Large"
       Exit Function
	Elseif totGirth > 108  And totGirth < 130 Then
		getSize = "OverSize"
		Exit Function
	Elseif totGirth >  84 And totGirth < 108 Then
	    getSize = "TooBig"
		Exit Function  	
	End If
End Function
function getClass(iWidth,iLength,iHeight)
	Dim iGirth, totGirth
	iGirth = iWidth + iHeight + iWidth + iHeight
	 totGirth =iGirth + iLength
     getClass = totGirth
end Function     
'-----------------------------------------------------
' GetPrice
'-----------------------------------------------------
Function GetPrice(vXml)

On error resume next
Dim sNew,posit

	posit = instr(vxml,"<Postage>")
	If posit > 0 then
		snew = mid(vxml,posit,len(vxml)- posit )	
		posit = instr(snew,"</Postage>")
		snew = mid(snew,10,posit-10)
		getprice = snew 
	Else  
		getprice = vxml
	End if
End Function

'-------------------------------------------------------
' Update saved cart customers' info in sfCustomers
'-------------------------------------------------------
Sub setUpdateSvdCustomer(sNewEmail,sPassword,sFirstName,sMiddleInitial,sLastName,sCompany,sAddress1,sAddress2,sCity,sState,sZip,sCountry,sPhone,sFax,bSubscribed,iCustID)
	Dim	sLocalSQl, rsUpdate, iOldNum
	
	sLocalSQL = "Select custFirstName, custMiddleInitial, custLastName, custCompany, custAddr1, custAddr2, custCity, custState, custZip, custCountry, "_
				& "custPhone, custFax, custTimesAccessed, custLastAccess, custEmail, custPasswd, custIsSubscribed FROM sfCustomers WHERE custID = " & iCustID
	
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
				rsUpdate.Fields("custPasswd")			= sPassword
				If sNewEmail <> "" Then
					rsUpdate.Fields("custEmail")		= sNewEmail
				End If		
				If bSubscribed = "" Then
					rsUpdate.Fields("custIsSubscribed") = 0
				Else
					rsUpdate.Fields("custIsSubscribed") = 1
				End If	
				rsUpdate.Update		
		End If
		closeObj(rsUpdate)	
End Sub

'---------------------------------------------------------------
'INPUTS
'sMailTypes:
'	The following are valid international mail types:
'	“letters or letter packages”
'	“other packages”
'	“postcards or aerogrammes”
'	“regular printed matter”
'	“books or sheet music”
'	“publishers’ periodicals”
'	“matter for the blind”
'dCountry: = destination Country
'---------------------------------------------------------------
Function getInterNationalUsPS(dCountry,sMailType,iLength,iWidth,iHeight,sProdID,iQuantity,uspsUsername,uspsPassword,oCountry,oZip,iWeight)
	Dim	sLen,uPounds,uOunces,httpconn,xmlDoc,srvr,Port,Path,query,msg,contentType,proxyServer,proxyPort,sRespcode,resp,RequestLevel,PackageLevel,PackageElementLevel,t
On error resume next
dim RateLevel,RateElementLevel	


		oCountry =trim(lcase(oCountry))
	    If uspsPassword = "" or uspsUsername = "" Then
			getInterNationalUsPs ="No Account"
			Exit function
		End if
       If Lcase(dCountry)="united states" then
	     getInterNationalUsPs = "Destination Country is not International"
         Exit function
       end if
	   If iQuantity < 1 Then
		  getInterNationalUsPs = "NoQuantity"
	      Exit function
       End if
  							
			         

	 'if iWeight < 1.0 then
	 '  iWeight = 1.0
	 'end if  
	
	On error resume next
'	 upounds = Int(iWeight)
'	 uOunces = iWeight mod Int(iWeight)
		'---------------------------------------------------------------------------------
		'Issue #255 
		if iWeight <> 0 then 
			upounds = Int(iWeight)
			uOunces = int((iWeight - uPounds)*16)
		
			'will break if both are 0 which will 
			'happen if weight is so little that it rounded down to 0
			if uOunces=0 and upounds=0 then
			uOunces=1
			end if
		else 'let it fail if there is a true weight of zero
			upounds = 0
			uOunces = 0
		end if
		'---------------------------------------------------------------------------------	 	 
    If uPounds > 70 Then
      getInterNationalUsPs = "TooHeavy"
     Exit Function
    End If    
      
	Set xmlDoc = Server.CreateObject("MSXML.DOMDocument")                     
	Set RequestLevel = xmlDoc.createElement("IntlRateRequest")  
	Set RequestLevel = xmlDoc.createElement("IntlRateRequest")
		RequestLevel.setAttribute "USERID", uspsUsername
		RequestLevel.setAttribute "PASSWORD", uspsPassword
	Set RateLevel = xmlDoc.createElement("Package")
		RateLevel.setAttribute "ID", sProdId
	Set RateElementLevel = xmlDoc.createElement("Pounds")
	Set t = xmlDoc.createTextNode(uPounds)
		RateElementLevel.appendChild (t)
		Call RateLevel.appendChild(RateElementLevel)
	Set RateElementLevel = xmlDoc.createElement("Ounces")
	Set t = xmlDoc.createTextNode(uOunces)
		RateElementLevel.appendChild (t)
		Call RateLevel.appendChild(RateElementLevel)
	Set RateElementLevel = xmlDoc.createElement("MailType")
	Set t = xmlDoc.createTextNode(sMailType)
		RateElementLevel.appendChild (t)
	Call RateLevel.appendChild(RateElementLevel)
	Set RateElementLevel = xmlDoc.createElement("Country")
	Set t = xmlDoc.createTextNode(dCountry)
		RateElementLevel.appendChild (t)
	Call RateLevel.appendChild(RateElementLevel)
	Call RequestLevel.appendChild(RateLevel)
	Call xmlDoc.appendChild(RequestLevel)
	
	Set httpconn = server.createobject("httpcom.chttpcom")	

		srvr = "Production.shippingapis.com" 'You must modify the Server name
		Port = 80
		Path = "/ShippingAPI.dll" 'You must modify the Path name
		query = "" '"?API=Rate&XML=" & xmlDoc.xml 'xmldoc.xmloc.xml
		msg = "API=IntlRate&XML=" & URL_Encode(xmlDoc.xml)
		contentType =  "application/x-www-form-urlencoded"
		proxyServer = "" 'You must modify proxy server or leave
		proxyPort = 0 'Modify the proxy port, if necessary. Leave blank if no proxy

	resp = ""

	On Error Resume Next
		Err.Number = 0
		sRespcode = httpConn.GetResponse(srvr, Port, Path, query, msg, contentType, proxyServer, proxyPort, resp)
	If Err.Number <> 0 Then
	   
   	   '+JF 8/23/01
	    If Err.number = 424 Then
			getInterNationalUsPS = "The USPS component is not properly installed."
		Else
			getInterNationalUsPS = Err.description 
		End If
		'-JF 8/23/01

	    Set xmlDoc = Nothing
	    Set httpConn = Nothing
	    Set xmlDoc = Nothing
	    Set httpConn = Nothing
	    closeObj(rsAdmin)
		closeObj(rsShipping)
		closeObj(rsProdShipping)
		Exit Function
Else
Dim strResult

   strResult = GetInterPrice(resp)
   
  if isnumeric(strResult)then
   getInterNationalUsPs = strResult
   
  else
   getInterNationalUsPs = getPrice(resp)
  end if  
End If
    Set xmlDoc = Nothing
    Set httpConn = Nothing
    Set xmlDoc = Nothing
    Set httpConn = Nothing
End Function

Function MyNumeric(iNum) 
 
 Dim sNumbers, i, c, res    
 sNumbers = "1234567890-.,"
    For i = 1 To Len(iNum)
        c = Right(Left(iNum, i), 1) 'set c equal to ith char
        res = InStr(sNumbers, c)
        If res = 0 Then
			MyNumeric = False
			Exit Function
        End If
    Next
    MyNumeric = True	
End Function
Function GetInterPrice(vXml)


On error resume next
Dim sNew,posit
dim sSearch

  

sSearch = "<Service ID="& chr(34) & "2" & chr(34) & ">"
posit = instr(vxml,sSearch)
  If posit > 0 then
    snew = mid(vxml,posit+100,len(vxml)- posit )	
	posit = instr(snew,"</Postage>")
	snew = mid(snew,10,posit-10)
	posit = instr(snew,"e>")
	snew = mid(snew,posit + 2,len(snew)-posit -1 )
	GetInterPrice = snew 
 Else  
	GetInterPrice = vxml
 End if
End Function




Function get_ltl(dZip, sContainer, iLength, iWidth, iHeight, sProdId, iQuantity, ltlUN, ltlEmail, oZip, iWeight, sType, boQty, shpQty, C_BGCOLOR2, C_BGCOLOR4, C_BGCOLOR5, C_BKGRND2, C_BKGRND4, C_BKGRND5, C_FONTFACE2, C_FONTFACE4, C_FONTFACE5, C_FONTCOLOR2, C_FONTCOLOR4, C_FONTCOLOR5, C_FONTSIZE4, C_FONTSIZE5, bShiprates, iTotalpur, iPremiumShipping, aCheck, dCity, dState, dCountry, sTotalPrice, iship, sShipmethodname)
dim sReturn,iClass
dim objFQRating
dim i,dblLtlPRice, ltlIndex
'check user
 Dim arrltl()
    Dim arrLTL2()
    Dim arrltla()
    Dim arrLTLa2()
    Dim arrltlb()
    Dim arrLTLb2()
    Dim arrltlc()
    Dim arrLTLc2()
    Dim TempRate
    Dim TempBillRate
    Dim TempBORate 
    Dim ii
    Dim f

if trim(ltlUn) = "" or trim(ltlEmail) = "" then
 get_LTL = "Not Registered for LTL Carriers Rate Service"
 exit function
end if  



on error resume next
'Create The Rating Object
Set objFQRating = CreateObject("FQRating.cFQRating")



 if err.number <> 0 then

  get_LTL = "LTL component is not properly installed."
 exit function
end if

 iClass = 50

'Populate the Email/Password Properties to Log-In
'set index for backorders...

if not isnull(Session("LTLIndex")) or Session("LTLIndex") <> "" then

 ltlIndex = cint(Session("LTLIndex")) 
 if ltlIndex < 1 then ltlIndex = 1
elseif ltlIndex < 1 or not isnumeric(ltlIndex) then 

 ltlIndex = 1
end if


if Session("SpecialBilling") =1 AND sType="All" THEN			

If boQty <> 0 And Trim(boQty) <> "" Then
		  Session("BackOrderPrices") = ""
                   Session("BackOrderCarriers") = ""
                   Session("BackOrderOptionIDs") = ""
                   Session("BackOrderTransits") = ""
       Session("backordershipping") = getShipping(iTotalpur, iPremiumShipping, aCheck, dCity, dState, dZip, dCountry, sTotalPrice, "BackOrder", iship, sShipmethodname, C_BGCOLOR2, C_BGCOLOR4, C_BGCOLOR5, C_BKGRND2, C_BKGRND4, C_BKGRND5, C_FONTFACE2, C_FONTFACE4, C_FONTFACE5, C_FONTCOLOR2, C_FONTCOLOR4, C_FONTCOLOR5, C_FONTSIZE4, C_FONTSIZE5)
End If
If shpQty <> 0 And Trim(boQty) <> "" Then
 Session("ShippedPrices") = ""
                   Session("ShippedCarriers") = ""
                   Session("ShippedOptionIDs") = ""
                   Session("ShippedTransits") = ""
             Session("BillShipping") = getShipping(iTotalpur, iPremiumShipping, aCheck, dCity, dState, dZip, dCountry, sTotalPrice, "Shipped", iship, sShipmethodname, C_BGCOLOR2, C_BGCOLOR4, C_BGCOLOR5, C_BKGRND2, C_BKGRND4, C_BKGRND5, C_FONTFACE2, C_FONTFACE4, C_FONTFACE5, C_FONTCOLOR2, C_FONTCOLOR4, C_FONTCOLOR5, C_FONTSIZE4, C_FONTSIZE5)
      
End If
end if		
If Session("specialbilling") = 1 And LCase(sType) = "all" Then
'do nothing
Else

'NOTE: TOTALLY BOGUS RESULTS WHEN USING THE TEST ACCOUNT''''''''''''''''''
objfqrating.Email = ltlEmail  '"xmltest@freightquote.com"
objfqrating.Password =  ltlUN  ' "xml"

'Set the Origin/Destination Zip Codes
	objFQRating.oaddress.zip =ozip
	objFQRating.daddress.zip =dzip

	'Populate the required Product Properties
	objFQRating.FQProds.Class1 =iClass
	objFQRating.FQProds.Description1 = sprodid
	objFQRating.FQProds.PackageType1 = sContainer
	objFQRating.FQProds.Pieces1 = iQuantity
	'---------------------------------------------------------------------------------
	'Issue #255 
	if iWeight > 0 and iWeight <1 then 
		iWeight=1
	end if
	'---------------------------------------------------------------------------------

	objFQRating.FQProds.Weight1 = iWeight
	objFQRating.BILLTO = "SITE"

	'Run Get Quote to get the quote
    objFQRating.GetQuote
End If


If Session("specialbilling") = 1 And LCase(sType) = "all" Then
'go through this to combine billing and backorder
    
If Trim(Session("BackOrderPrices")) <> "" And Trim(Session("ShippedPrices")) <> "" Then
    arrltl = Split(Session("BackOrderPrices"), "|")
    arrLTL2 = Split(Session("ShippedPrices"), "|")
    arrltla = Split(Session("BackOrderCarriers"), "|")
    arrLTLa2 = Split(Session("ShippedCarriers"), "|")
    arrltlb = Split(Session("BackOrderOptionIDs"), "|")
    arrLTLb2 = Split(Session("ShippedOptionIDs"), "|")
    arrltlc = Split(Session("BackOrderTransits"), "|")
    arrLTLc2 = Split(Session("ShippedTransits"), "|")
    Dim tempcount
    
    
        'close existing tables
        'sReturn = sReturn & "</table></td></tr></table></td></tr>" & vbCrLf
        sReturn = sReturn & "</table></td></tr>" & vbCrLf
        sReturn = sReturn & "<script>" & Chr(13) & vbCrLf
        '''''
        'build java
        sReturn = sReturn & "<!--" & Chr(13) & vbCrLf
        sReturn = sReturn & "function selectcarrier(chkindex,pQuoteID,pOptionID,sLTLCarrier) {" & vbCrLf
        sReturn = sReturn & "var ichkcount =" & objfqRating.FQResults.Count & ";" & vbCrLf
        sReturn = sReturn & "var e;" & vbCrLf
       ' sReturn = sReturn & "for (var i = 0; i < ichkcount - 1; i++) {" & vbCrlf
    '   sReturn = sReturn & "  document.frmLTL.Carrier[i].checked =false; " & vbCrlf
     '   sReturn = sReturn & "  }" & vbCrlf
      '  sReturn = sReturn & "  document.frmLTL.Carrier[chkindex -1 ].checked = true; " & vbCrlf
         sReturn = sReturn & "document.frmLTL.action = " & Chr(34) & "verify.asp?OptionID=" & Chr(34) & " + chkindex ;" & vbCrLf
        sReturn = sReturn & "document.frmLTL.OptionID.value = chkindex;" & vbCrLf
        sReturn = sReturn & "document.frmLTL.ltlrate.value = pOptionID;" & vbCrLf
        sReturn = sReturn & "document.frmLTL.ltlcarrier.value = sLTLCarrier;" & vbCrLf
        'sReturn = sReturn & "alert(sLTLCarrier);"  & vbCrlf
        sReturn = sReturn & "document.frmLTL.submit();" & vbCrLf
        sReturn = sReturn & "}" & Chr(13) & vbCrLf
        sReturn = sReturn & "//-->" & Chr(13) & vbCrLf
        sReturn = sReturn & "</script>" & vbCrLf
        'build form and new tables
        sReturn = sReturn & "<form method=post name=frmLTL action=verify.asp>"
        sReturn = sReturn & "<tr>" & vbCrLf
        sReturn = sReturn & "<td width=100% align=center class='tdContent2'>" & vbCrLf
        sReturn = sReturn & "<table border=0 width=100% cellspacing=0 cellpadding=0>" & vbCrLf
        sReturn = sReturn & "<tr>" & vbCrLf
        sReturn = sReturn & "<td width=100% align=left class='tdMiddleTopBanner'><font class='Middle_Top_Banner_Small'><B>Select a shipping Option:</font></B></td>" & vbCrLf
        sReturn = sReturn & "</tr>" & vbCrLf
        sReturn = sReturn & "<tr>" & vbCrLf
        sReturn = sReturn & "<td>" & vbCrLf
        sReturn = sReturn & "<table border=0 width=100% cellspacing=0 cellpadding=0>" & vbCrLf
        sReturn = sReturn & "<tr>" & vbCrLf
        sReturn = sReturn & "<td width=10% align=center class='tdContentBar' ><B>Select</B></td>" & vbCrLf
        sReturn = sReturn & "<td width=40% align=center class='tdContentBar' ><B>Method</B></td>" & vbCrLf
      If iConverion = 1 Then
          sReturn = sReturn & "<td width=10% align=center class='tdContentBar' ><B>Transit Time (days)</B></td>" & vbCrLf
          sReturn = sReturn & "<td width=40% align=center class='tdContentBar' ><B>Rate</B></td>" & vbCrLf
      Else
          sReturn = sReturn & "<td width=25% align=center class='tdContentBar' ><B>Transit Time (days)</B></td>" & vbCrLf
          sReturn = sReturn & "<td width=25% align=center class='tdContentBar' ><B>Rate</B></td>" & vbCrLf
       
      End If
        sReturn = sReturn & "</tr>" & vbCrLf

    'build results table
    tempcount = 0
     For i = 1 To UBound(arrltl)
        For ii = 1 To UBound(arrLTL2)
               If Trim(arrltla(i)) = Trim(arrLTLa2(ii)) Then
               tempcount = tempcount + 1

               sReturn = sReturn & "<tr>" & vbCrLf






               If i = ltlIndex Then
                 sShipmethodname = arrltla(i)
'                    arrltl = Split(Session("BackOrderPrices"), "|")
'                    arrLTL2 = Split(Session("ShippedPrices"), "|")

                     If Trim(Session("BackOrderPrices")) <> "" And Trim(Session("ShippedPrices")) <> "" Then
                        TempRate = "$" & CDBL(Mid(arrltl(i), 1)) + CDBL(Mid(arrLTL2(i), 1))
                     ElseIf Trim(Session("BackOrderPrices")) <> "" Then
                        TempRate = arrltl(i)
                     ElseIf Trim(Session("ShippedPrices")) <> "" Then
                        TempRate = arrLTL2(i)
                     End If
                     dblLTLPrice = TempRate * CDBL(bShiprates)
                     'sReturn = sReturn & "<td width=5% align=center><input type=radio name=Carrier checked =true onClick=selectcarrier(" & i & "," & objFQRating.FQResults.item(1).OptionID & "," & chr(34) & TempRate & chr(34) & "," & chr(34) & objFQRating.FQResults.item(i).Carrier & chr(34) & ")></td>" & vbCrlf
                     sReturn = sReturn & "<td width=5% align=center  class='tdContent' ><input type=radio name=Carrier value=" & Chr(34) & arrltla(i) & Chr(34) & " checked =true onClick=selectcarrier(" & i & "," & arrltlb(i) & "," & Chr(34) & TempRate & Chr(34) & ",this.value)></td>" & vbCrLf
                  









             Else

                   arrltl = Split(Session("BackOrderPrices"), "|")
                   arrLTL2 = Split(Session("ShippedPrices"), "|")
                    If Trim(Session("BackOrderPrices")) <> "" And Trim(Session("ShippedPrices")) <> "" Then
                     TempRate = "$" & CDBL(Mid(arrltl(i), 1)) + CDBL(Mid(arrLTL2(i), 1))
                    ElseIf Trim(Session("BackOrderPrices")) <> "" Then
                     TempRate = arrltl(i)
                    ElseIf Trim(Session("ShippedPrices")) <> "" Then
                     TempRate = arrLTL2(i)
                    End If
                sReturn = sReturn & "<td width=5% align=center class='tdContent' ><input type=radio name=Carrier value=" & Chr(34) & arrltla(i) & Chr(34) & " onClick=selectcarrier(" & i & "," & arrltlb(i) & "," & Chr(34) & TempRate & Chr(34) & ",this.value)></td>" & vbCrLf
             End If
             sReturn = sReturn & "<td width=45% align=center class='tdContent' >" & arrltla(i) & "</td>" & vbCrLf
             sReturn = sReturn & "<td width=2% align=center class='tdContent' >" & arrltlc(i) & "</td>" & vbCrLf
             If iConverion = 1 Then
               sReturn = sReturn & "<td width=45% align=center class='tdContent' >" & "<script> document.write(""" & TempRate & " = ("" + OANDAconvert(" & CDBL(TempRate) & ", " & Chr(34) & CurrencyISO & Chr(34) & ") + "")"");</script></td>" & vbCrLf
             Else
               sReturn = sReturn & "<td width=33% align=center  class='tdContent' >" & "<B>" & TempRate & "</B></td>" & vbCrLf
             End If
            sReturn = sReturn & "</tr>" & vbCrLf
            End If
       Next
    Next
      ' get request form variable for re-submit
      
     ' ERR.Clear
     ' On Error Resume Next
     
      For f = 1 To Request.Form.Count
        sReturn = sReturn & "<input type=hidden name= " & Request.Form.Key(f) & " value= " & Request.Form.Item(f) & ">"
      Next
    'close all tables
    sReturn = sReturn & "</table>" & vbCrLf
    sReturn = sReturn & "</td>"
    sReturn = sReturn & "</tr>"
    sReturn = sReturn & "</table>" & vbCrLf
    
    'attach hidden variables to form
    sReturn = sReturn & "<input type=hidden name=QuoteID>" & vbCrLf
    sReturn = sReturn & "<input type=hidden name=OptionID>" & vbCrLf
    sReturn = sReturn & "<input type=hidden name=FromVerify value =" & Chr(34) & "1" & Chr(34) & ">" & vbCrLf
    sReturn = sReturn & "<input type=hidden name=ltlrate>" & vbCrLf
    sReturn = sReturn & "<input type=hidden name=ltlcarrier>" & vbCrLf
    sReturn = sReturn & "</td>"
    sReturn = sReturn & "</tr>"
    sReturn = sReturn & "</form>" & vbCrLf
    
    'build java for re-submit
    sReturn = sReturn & "<script>" & Chr(13) & vbCrLf
    sReturn = sReturn & "<!--" & Chr(13) & vbCrLf
    sReturn = sReturn & "function resetme(chkID) {" & vbCrLf
'    sReturn = sReturn & "alert(chkID);" & vbCrlf
    sReturn = sReturn & "var ichkcount =" & tempcount & ";" & vbCrLf
    sReturn = sReturn & "for (var i = 0; i < ichkcount - 1; i++) {" & vbCrLf
    sReturn = sReturn & "  document.frmLTL.Carrier[i].checked =false; " & vbCrLf
    sReturn = sReturn & "  }" & vbCrLf
    sReturn = sReturn & "document.frmLTL.Carrier[chkID -1].checked =true;  " & vbCrLf
    sReturn = sReturn & "}" & Chr(13) & vbCrLf
    sReturn = sReturn & "//-->" & Chr(13) & vbCrLf
    sReturn = sReturn & "</script>" & vbCrLf
    'reopen old tables
    sReturn = sReturn & "<td width=100% align=center class='tdContent2'>" & vbCrLf
    sReturn = sReturn & "<table border= 0 width=100% cellspacing=0 cellpadding=0>" & vbCrLf
    sReturn = sReturn & "<tr><td class='tdContent2'>" & vbCrLf
    sReturn = sReturn & " <table border=0 width=100% cellspacing=0 cellpadding=2 class='tdContent2'>" & vbCrLf




    get_ltl = CDBL(dblLTLPrice) & "|" & sReturn
    Session("sLTL") = FormatCurrency(dblLTLPrice) & "|" & sReturn
    
Else

    sReturn = sReturn & "No Carriers Found." & vbCrLf
'    sReturn = sReturn & objfqRating.XMLQuoteResponse & vbCrLf
    get_ltl = sReturn
End If


Else

If objfqRating.FQResults.Count > 0 Then
        'close existing tables
        'sReturn = sReturn & "</table></td></tr></table></td></tr>" & vbCrLf
        sReturn = sReturn & "</table></td></tr>" & vbCrLf
        sReturn = sReturn & "<script>" & Chr(13) & vbCrLf
        '''''
        'build java
        sReturn = sReturn & "<!--" & Chr(13) & vbCrLf
        sReturn = sReturn & "function selectcarrier(chkindex,pQuoteID,pOptionID,sLTLCarrier) {" & vbCrLf
        sReturn = sReturn & "var ichkcount =" & objfqRating.FQResults.Count & ";" & vbCrLf
        sReturn = sReturn & "var e;" & vbCrLf
       ' sReturn = sReturn & "for (var i = 0; i < ichkcount - 1; i++) {" & vbCrlf
    '   sReturn = sReturn & "  document.frmLTL.Carrier[i].checked =false; " & vbCrlf
     '   sReturn = sReturn & "  }" & vbCrlf
      '  sReturn = sReturn & "  document.frmLTL.Carrier[chkindex -1 ].checked = true; " & vbCrlf
         sReturn = sReturn & "document.frmLTL.action = " & Chr(34) & "verify.asp?OptionID=" & Chr(34) & " + chkindex ;" & vbCrLf
        sReturn = sReturn & "document.frmLTL.OptionID.value = chkindex;" & vbCrLf
        sReturn = sReturn & "document.frmLTL.ltlrate.value = pOptionID;" & vbCrLf
        sReturn = sReturn & "document.frmLTL.ltlcarrier.value = sLTLCarrier;" & vbCrLf
        'sReturn = sReturn & "alert(sLTLCarrier);"  & vbCrlf
        sReturn = sReturn & "document.frmLTL.submit();" & vbCrLf
        sReturn = sReturn & "}" & Chr(13) & vbCrLf
        sReturn = sReturn & "//-->" & Chr(13) & vbCrLf
        sReturn = sReturn & "</script>" & vbCrLf
        
        'build form and new tables
        sReturn = sReturn & "<form method=post name=frmLTL action=verify.asp>"
        sReturn = sReturn & "<tr>" & vbCrLf
        sReturn = sReturn & "<td width=100% align=center class='tdContent2' >" & vbCrLf
        sReturn = sReturn & "<table border=0 width=100% cellspacing=0 cellpadding=0>" & vbCrLf
        sReturn = sReturn & "<tr>" & vbCrLf
        sReturn = sReturn & "<td width=100% align=left class='tdContentBar'><font class='Middle_Top_Banner_Small'><B>Select a shipping Option:</font></B><br>&nbsp;</td>" & vbCrLf
        sReturn = sReturn & "</tr>" & vbCrLf
        sReturn = sReturn & "<tr>" & vbCrLf
        sReturn = sReturn & "<td>" & vbCrLf
        sReturn = sReturn & "<table border=0 width=100% cellspacing=0 cellpadding=0>" & vbCrLf
        sReturn = sReturn & "<tr>" & vbCrLf
        sReturn = sReturn & "<td width=10% align=center class='tdContentBar' ><B>Select</B></td>" & vbCrLf
        sReturn = sReturn & "<td width=40% align=center class='tdContentBar' ><B>Method</B></td>" & vbCrLf
      If iConverion = 1 Then
          sReturn = sReturn & "<td width=10% align=center class='tdContentBar' ><B>Transit Time (days)</B></td>" & vbCrLf
          sReturn = sReturn & "<td width=40% align=center class='tdContentBar' ><B>Rate</B></td>" & vbCrLf
      Else
          sReturn = sReturn & "<td width=25% align=center class='tdContentBar' ><B>Transit Time (days)</B></td>" & vbCrLf
          sReturn = sReturn & "<td width=25% align=center class='tdContentBar' ><B>Rate</B></td>" & vbCrLf
       
      End If
        sReturn = sReturn & "</tr>" & vbCrLf

    'build results table
    
     For i = 1 To objfqRating.FQResults.Count
               
               sReturn = sReturn & "<tr>" & vbCrLf
               If sType = "BackOrder" Then
                   Session("BackOrderPrices") = Session("BackOrderPrices") & "|" & objfqRating.FQResults.Item(i).Rate
                   Session("BackOrderCarriers") = Session("BackOrderCarriers") & "|" & objfqRating.FQResults.Item(i).Carrier
                   Session("BackOrderOptionIDs") = Session("BackOrderOptionIDs") & "|" & objfqRating.FQResults.Item(i).OptionID
                   Session("BackOrderTransits") = Session("BackOrderTransits") & "|" & objfqRating.FQResults.Item(i).Transit
               ElseIf sType = "Shipped" Then
                   
                   Session("ShippedPrices") = Session("ShippedPrices") & "|" & objfqRating.FQResults.Item(i).Rate
                   Session("ShippedCarriers") = Session("ShippedCarriers") & "|" & objfqRating.FQResults.Item(i).Carrier
                   Session("ShippedOptionIDs") = Session("ShippedOptionIDs") & "|" & objfqRating.FQResults.Item(i).OptionID
                   Session("ShippedTransits") = Session("ShippedTransits") & "|" & objfqRating.FQResults.Item(i).Transit
               End If
               If i = ltlIndex Then
                 sShipmethodname = objfqRating.FQResults.Item(i).Carrier
 
                     dblLTLPrice = objfqRating.FQResults.Item(i).Rate
                     dblLTLPrice = dblLTLPrice * CDBL(bShiprates)
                     dblLTLPrice = FormatCurrency(CStr(dblLTLPrice))
                     TempRate = objfqRating.FQResults.Item(i).Rate
                     TempRate = TempRate * CDBL(bShiprates)
                     TempRate = FormatCurrency(TempRate)
                     sReturn = sReturn & "<td class='tdContent' width=5% align=center> <input type=radio name=Carrier value=" & Chr(34) & objfqRating.FQResults.Item(i).Carrier & Chr(34) & " checked =true onClick=selectcarrier(" & i & "," & objfqRating.FQResults.Item(i).OptionID & "," & Chr(34) & TempRate & Chr(34) & ",this.value)></td>" & vbCrLf
   '              End If
             Else
    
                   TempRate = objfqRating.FQResults.Item(i).Rate
                   TempRate = CDBL(TempRate) * CDBL(bShiprates)
                   TempRate = FormatCurrency(TempRate)
    '             End If
                sReturn = sReturn & "<td class='tdContent' width=5% align=center><input type=radio name=Carrier value=" & Chr(34) & objfqRating.FQResults.Item(i).Carrier & Chr(34) & " onClick=selectcarrier(" & i & "," & objfqRating.FQResults.Item(i).OptionID & "," & Chr(34) & TempRate & Chr(34) & ",this.value)></td>" & vbCrLf
             End If
             sReturn = sReturn & "<td class='tdContent' width=45% align=center>" & objfqRating.FQResults.Item(i).Carrier & "</td>" & vbCrLf
             sReturn = sReturn & "<td class='tdContent' width=2% align=center>" & objfqRating.FQResults.Item(i).Transit & "</td>" & vbCrLf
             
             If iConverion = 1 Then
                        
             sReturn=sReturn & "<td class='td Content' width=45% align=center>" & "<script> document.write(""" & TempRate & " = ("" + OANDAconvert(" & cDbl(TempRate) & ", """ & CurrencyISO & """) + "")"");</script></td>" & vbCrLf
             '  sReturn = sReturn & "<td class='tdContent' width=45% align=center>HHH<font class='ContentBar_Small'>" & "<script> document.write(""" & TempRate & " = ("" + OANDAconvert(" & cDbl(TempRate) & ", " & Chr(34) & CurrencyISO & Chr(34) & ") + "")"");</script></font></i></td>" & vbCrLf
             
             Else
             
               sReturn = sReturn & "<td class='tdContent' width=33% align=center>" & "<B>" & TempRate & "</B></td>" & vbCrLf
             End If
            sReturn = sReturn & "</tr>" & vbCrLf
            
    Next
      ' get request form variable for re-submit
      
     ' ERR.Clear
     ' On Error Resume Next
      For f = 1 To Request.Form.Count
        sReturn = sReturn & "<input type=hidden name= " & Request.Form.Key(f) & " value= " & Request.Form.Item(f) & ">"
      Next
      
    'close all tables
    sReturn = sReturn & "</table>" & vbCrLf
    sReturn = sReturn & "</td>"
    sReturn = sReturn & "</tr>"
    sReturn = sReturn & "</table>" & vbCrLf
    
    'attach hidden variables to form
    sReturn = sReturn & "<input type=hidden name=QuoteID>" & vbCrLf
    sReturn = sReturn & "<input type=hidden name=OptionID>" & vbCrLf
    sReturn = sReturn & "<input type=hidden name=FromVerify value =" & Chr(34) & "1" & Chr(34) & ">" & vbCrLf
    sReturn = sReturn & "<input type=hidden name=ltlrate>" & vbCrLf
    sReturn = sReturn & "<input type=hidden name=ltlcarrier>" & vbCrLf
    sReturn = sReturn & "</td>"
    sReturn = sReturn & "</tr>"
    sReturn = sReturn & "</form>" & vbCrLf
    'build java for re-submit
    sReturn = sReturn & "<script>" & Chr(13) & vbCrLf
    sReturn = sReturn & "<!--" & Chr(13) & vbCrLf
    sReturn = sReturn & "function resetme(chkID) {" & vbCrLf
'    sReturn = sReturn & "alert(chkID);" & vbCrlf
    sReturn = sReturn & "var ichkcount =" & objfqRating.FQResults.Count & ";" & vbCrLf
    sReturn = sReturn & "for (var i = 0; i < ichkcount - 1; i++) {" & vbCrLf
    sReturn = sReturn & "  document.frmLTL.Carrier[i].checked =false; " & vbCrLf
    sReturn = sReturn & "  }" & vbCrLf
    sReturn = sReturn & "document.frmLTL.Carrier[chkID -1].checked =true;  " & vbCrLf
    sReturn = sReturn & "}" & Chr(13) & vbCrLf
    sReturn = sReturn & "//-->" & Chr(13) & vbCrLf
    sReturn = sReturn & "</script>" & vbCrLf
    'reopen old tables
    sReturn = sReturn & "<td width=100% align=center  class='tdContent2'>" & vbCrLf
    sReturn = sReturn & "<table border= 0 width=100% cellspacing=0 cellpadding=0>" & vbCrLf
    sReturn = sReturn & "<tr><td class='tdContent2'>" & vbCrLf
    sReturn = sReturn & " <table border=0 width=100% cellspacing=0 cellpadding=2 class='tdContent2'>" & vbCrLf




    get_ltl = cDbl(dblLTLPrice) & "|" & sReturn
    Session("sLTL") = FormatCurrency(dblLTLPrice) & "|" & sReturn
    'Response.Write "S"
Else

    sReturn = sReturn & "No Carriers Found." & vbCrLf
    sReturn = sReturn & objfqRating.XMLQuoteResponse & vbCrLf
    get_ltl = sReturn
End If
End If
Set objfqRating = Nothing
End Function
%>








