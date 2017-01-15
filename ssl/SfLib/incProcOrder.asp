<%

'@BEGINVERSIONINFO

'@APPVERSION: 50.4014.0.3

'@FILENAME: incProcOrder.asp
	 


'@DESCRIPTION: Process the customers order

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws  and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000,2001 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO

'Modified 10/23/01 
'Storefront Ref#'s: 147 'JF
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

'-------------------------------------------------------------------
' Function to retrieve customer shipping row
'-------------------------------------------------------------------
Function getCustomerShippingRow(iCustID)
	Set getCustomerShippingRow = getRowActive("sfCShipAddresses","cshpaddrCustID","cshpaddrIsActive",iCustID,cnn)
End Function

'-------------------------------------------------------------------
' Function to retrieve customer creditcard row
'-------------------------------------------------------------------
Function getCustomerCardRow(iCustID)
	Set getCustomerCardRow = getRowActive("sfCPayments","payCustId","payIsActive",iCustID,cnn)
End Function


'-------------------------------------------------------------------
' Function to retrieve customer row
'-------------------------------------------------------------------
Function getCustomerRow(iCustID)
	Set getCustomerRow = getRow("sfCustomers","custID",iCustID,cnn)
End Function

'-------------------------------------------------------------------
' Subroutine setUpdateSavedCartCustID
'-------------------------------------------------------------------
Sub setUpdateSavedCartCustID(iCustID,iDeletedCustID)
	Dim sSQL, rsTmpCust
	sSQL = "Select odrdtsvdCustID FROM sfSavedOrderDetails WHERE odrdtsvdCustID=" & makeInputSafe(iDeletedCustID)
	Set rsTmpCust = Server.CreateObject("ADODB.RecordSet")		
		rsTmpCust.Open sSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText					
			Do While NOT rsTmpCust.EOF
					rsTmpCust.Fields("odrdtsvdCustID")	= makeInputSafe(trim(iCustID))
					rsTmpCust.Update	
					rsTmpCust.MoveNext
			Loop
		closeobj(rsTmpCust)		
End Sub

'--------------------------------------------------------------------
' Function to update sessionid
'--------------------------------------------------------------------
Sub setUpdateTmpOrdersSessionID(OldSessionID,NewSessionID)
Dim sSQL, rsTmp 
	sSQL = "SELECT odrdttmpSessionID FROM sfTmpOrderDetails WHERE odrdttmpSessionID=" & OldSessionID
		Set rsTmp = Server.CreateObject("ADODB.RecordSet")		
		rsTmp.Open sSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText					
			Do While NOT rsTmp.EOF
					rsTmp.Fields("odrdttmpSessionID")	= makeInputSafe(trim(NewSessionID))
					rsTmp.Update	
					rsTmp.MoveNext
			Loop
		closeobj(rsTmp)		
End Sub

'-----------------------------------------------------------------------
' Deletes saved customer row
'-----------------------------------------------------------------------
Sub DeleteCustRow(iCustID)
	Dim rsDelete, sSQL
	
	sSQL = "DELETE FROM sfCustomers WHERE custID= " & makeInputSafe(iCustID) & " AND custFirstName = 'SavedCartCustomer'"
	Set rsDelete = cnn.Execute(sSQL)
	closeObj(rsDelete)
End Sub

'--------------------------------------------------------------------
' Function : getShippingList
' This returns the list for shipping options in HTML format for dropdown box.
'--------------------------------------------------------------------	
Function getShippingList(blnFree)
	Dim sShipList, rsShipList,iPrmShip,rsShipMethod, sLocalSQL, iShipMethod, iCounter
	
	sLocalSQL = "SELECT adminShipType, adminPrmShipIsActive FROM sfAdmin"	
	
	Set rsShipMethod = Server.CreateObject("ADODB.RecordSet")
	rsShipMethod.Open sLocalSQL, cnn, adOpenDynamic, adLockOptimistic, adCmdText
	
		iShipMethod = rsShipMethod.Fields("adminShipType")
		iPrmShip    = rsShipMethod.Fields("adminPrmShipIsActive")	

		If iShipMethod = 1 Then
				if blnFree = true then
					sShipList = "<option value=""3,400"" style=""WEIGHT: bold;COLOR: red"">Free Shipping</option>"
				Else
					sShipList =  "<option value=""0"">Regular Shipping</option>"
				End If
				
				If iPrmShip = 1 Then
						sShipList = sShipList & "<option value=""1"">Premium Shipping</option>"
				End If
							
		ElseIf iShipMethod = 2 Then		
			sLocalSQL = "SELECT shipID, shipMethod FROM sfShipping WHERE shipIsActive = 1"	
			
			Set rsShipList = Server.CreateObject("ADODB.RecordSet")
			if blnFree = true then
			    dim sSql
'			
			       sSql = "SELECT * FROM sfShipping WHERE shipMethod = 'Free Shipping' "
			       rsShipList.Open sSQL, cnn, adOpenKeyset, adLockOptimistic, adCmdText

			 sShipList = sShipList &"<option value=" & chr(34) & rsShipList.Fields("shipID")& ",400" & chr(34) & "style=""WEIGHT: bold;COLOR: red"">Free Shipping</option>" 
			    rsShipList.Close 
			end if  
			rsShipList.Open sLocalSQL, cnn, adOpenKeyset, adLockOptimistic, adCmdText
						
			If rsShipList.EOF Or rsShipList.BOF Then
			   sShipList = ""
			Else					
				For iCounter = 1 to rsShipList.RecordCount
					sShipList = sShipList & "<option value=""" & Trim(rsShipList.Fields("shipID"))& """>" & Trim(rsShipList.Fields("shipMethod")) & "</option>"
					rsShipList.MoveNext
				Next	
			End If
					
		ElseIf iShipMethod = 3 Then
				if blnFree = true then
					sShipList = "<option value=""3,400"" style=""WEIGHT: bold;COLOR: red"">Free Shipping</option>"
				Else
					sShipList =  "<option value=""0"">Regular Shipping</option>"
				End If
				
				If iPrmShip = 1 Then
						sShipList = sShipList & "<option value=""1"">Premium Shipping</option>"
				End If
				
		End If

	closeobj(rsShipList)
	closeObj(rsShipMethod)
	getShippingList = sShipList
End Function


'--------------------------------------------------------------------
' Function : getStateList
' This generates a dropdown list option 
'--------------------------------------------------------------------	
Function getStateList()
	Dim rsSList, SQL, txtList
	Set rsSList = Server.CreateObject("ADODB.Recordset")
				
	SQL = "SELECT loclstAbbreviation, loclstName FROM sfLocalesState WHERE loclstLocaleIsActive=1 ORDER BY loclstName"
	rsSList.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
	txtList = ""
			
	Do While Not rsSList.EOF 
		txtList = txtList & "<option value = """ & trim(rsSList.Fields("loclstAbbreviation")) & """>" & trim(rsSList.Fields("loclstName")) & "</option>"
		rsSList.MoveNext 
	Loop
	'#541
txtList = txtList & "<option value =' '> </option>"	
	rsSList.Close 
	Set rsSList = Nothing
	getStateList = txtList
End Function

'--------------------------------------------------------------------
' Function : getCountryList
' This generates a dropdown list option for countries 
'--------------------------------------------------------------------	
Function getCountryList()
	Dim rsCList, SQL, txtList
	Set rsCList = Server.CreateObject("ADODB.Recordset")

	SQL = "SELECT loclctryAbbreviation, loclctryName FROM sfLocalesCountry WHERE loclctryLocalIsActive=1 ORDER BY loclctryName"
	rsCList.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
		
	txtList = ""
			
	Do While Not rsCList.EOF 
		txtList = txtList & "<option value = """ & trim(rsCList.Fields("loclctryAbbreviation")) & """>" & trim(rsCList.Fields("loclctryName")) & "</option>"
		rsCList.MoveNext 
	Loop
	rsCList.Close 
	'#541
	txtList = txtList & "<option value =' '> </option>"			
	Set rsCList = Nothing
	getCountryList = txtList 		
End Function
%>








