<% 
'@BEGINVERSIONINFO

'@APPVERSION: 50.4011.0.4

'@FILENAME: mail.asp
	


'

'@DESCRIPTION: Creates Notification eMail

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws as an unpublished work, and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO

' #259 - MS

On Error Resume Next
Sub createMail(sType,sInformation)
'If vDebug  <> 1 Then  On Error Resume Next
	Dim iRow, sBody, sPrimary, sSecondary, sSubject, sMessage, sMailMethod, sMailServer, _ 
		rsAdmin,sMathSign, Mailer, Mailer2, sCustEmail, arrInfo, sMBody, sCBody, sPath, sHomePath
		
	Set rsAdmin = Server.CreateObject("ADODB.RecordSet")
	rsAdmin.Open "sfAdmin", cnn, adOpenForwardOnly , adLockReadOnly, adCmdTable
	
	sPrimary = rsAdmin.Fields("adminPrimaryEmail")
	sSecondary = rsAdmin.Fields("adminSecondaryEmail")
	sSubject = rsAdmin.Fields("adminEmailSubject")
	sMessage = rsAdmin.Fields("adminEmailMessage")
	sMailMethod = rsAdmin.Fields("adminMailMethod")
	sMailServer = rsAdmin.Fields("adminMailServer")
	sPath = rsAdmin.Fields("adminSSLPath")
	sHomePath = rsAdmin.Fields("adminDomainName")
	closeObj(rsAdmin)
	
	If Mid(sHomePath, len(sHomePath)-1, 1) <> "/" Then
		sHomePath = sHomePath & "/"
	End If
	
If sType = "Confirm" Then
	dim strCustName
	sCustEmail = sInformation
	strCustName =  sCustFirstName & " " & sCustMiddleInitial & " " & sCustLastName
	'Build basic Email Body Info
	
	sBody = VbCrLf & "-----------------" & VbCrLf & "Sold To" & VbCrLf & "-----------------" & VbCrLf  
	sBody = sBody & "" & strCustName & VbCrLf
	sBody = sBody & sCustCompany & vbCrLf 'issue #259
	sBody = sBody & sCustAddress1 &  VbCrLf
	if trim(sCustAddress2) <> "" then
     sBody = sBody & sCustAddress2 &  VbCrLf
    end if
    sBody = sBody & sCustCity & ", " & sCustState & " " & sCustZip & VbCrLf
    sBody = sBody & sCustCountryName &  VbCrLf  
    sBody = sBody &  sCustPhone &  VbCrLf
    sBody = sBody & "Fax Number: " & sCustFax &  VbCrLf  
    sBody = sBody & "Email Address: " & sCustEmail &  VbCrLf  
    sBody = sBody & "Payment Method: " & sPaymentMethod &  VbCrLf  
    sBody = sBody & VbCrLf & "-----------------" & VbCrLf & "Shipped To" & VbCrLf & "-----------------" & VbCrLf  
    strCustName =  sShipCustFirstName & " " & sShipCustMiddleInitial & " " & sShipCustLastName
	
    sBody = sBody & strCustName & VbCrLf
		sBody = sBody & sShipCustCompany & vbCrLf 'issue #259
    sBody = sBody &  sShipCustAddress1 &  VbCrLf
	if trim(sShipCustAddress2) <> "" then
     sBody = sBody  & sShipCustAddress2 &  VbCrLf
    end if
    sBody = sBody & sShipCustCity & ", " & sShipCustState & " " & sShipCustZip & VbCrLf
    sBody = sBody &  sShipCustCountryName &  VbCrLf  
    sBody = sBody &  sShipCustPhone &  VbCrLf
    sBody = sBody &  sCustFax &  VbCrLf  
    sBody = sBody &  sShipCustEmail &  VbCrLf  
	sBody = sBody & VbCrLf & "-----------------" & VbCrLf & "Purchase Summary" & VbCrLf & "-----------------" & VbCrLf  
	sBody = sBody & "Order ID: " & iOrderID & VbCrLf	
	iRow = 0
	For iCounter = 0 To iProductCounter - 1	
		sBody = sBody & VbCrLf & "Item " & iCounter+1 & VbCrLf
		sBody = sBody & "Product ID: " & aAllProd(0,iCounter)(4) & VbCrLf 'SFUPDATE
		sBody = sBody & "Product Name: " & aAllProd(0,iCounter)(0) & VbCrLf 
		
		' aAllProd(0,iCounter)(2) is the attrNum indicator. This is in place of a null tester
		' aAllProd is a two 2d array of arrays, aAllProd(attributeArray, productArray)	
		 If (aAllProd(0, iCounter)(2) > 0) Then
         sBody = sBody & "Attributes: "
             For iRow = 1 To aAllProd(0, iCounter)(2) - 1
                If aAllProd(iRow, iCounter)(2) = 1 Then
                   sMathSign = "+ "
                ElseIf aAllProd(iRow, iCounter)(2) = 2 Then
                   sMathSign = "- "
                End If
                sBody = sBody & aAllProd(iRow, iCounter)(0) & ", "
                If aAllProd(iRow, iCounter)(1) <> "0" AND isnumeric(aAllProd(iRow, iCounter)(1)) Then
                    sBody = sBody & sMathSign & FormatCurrency(aAllProd(iRow, iCounter)(1)) & vbCrLf
                Else
                    sBody = sBody & vbCrLf
                End If
             Next
                If aAllProd(iRow, iCounter)(2) = 1 Then
                   sMathSign = "+ "
                ElseIf aAllProd(iRow, iCounter)(2) = 2 Then
                   sMathSign = "- "
                End If
             sBody = sBody & aAllProd(iRow, iCounter)(0) & ", "
              If aAllProd(iRow, iCounter)(1) <> "0" Then
                sBody = sBody & sMathSign & FormatCurrency(aAllProd(iRow, iCounter)(1)) & vbCrLf
              Else
                    sBody = sBody & vbCrLf
              End If
        End If   	
		
		IF Application("AppName") = "StoreFrontAE" then 'SFAE
			sBody = sBody & "Product Price: " & FormatCurrency(aProdInfoAE(iCounter,1)) & VbCrLf
			If aProdInfoAE(iCounter,2) > 0 Then
'				sBody = sBody & "Quantity: " & aAllProd(0,iCounter)(3) & VbCrLf
'				sBody = sBody & "BackOrdered Qty: " &aProdInfoAE(iCounter,2) & " (of above qty)" & VbCrLf
				sBody = sBody & "Quantity: " & aAllProd(0,iCounter)(3) & " (" & aProdInfoAE(iCounter,2) & " on backorder)" & VbCrLf
			Else
				sBody = sBody & "Quantity: " & aAllProd(0,iCounter)(3) & VbCrLf
			End If
			If aProdInfoAE(iCounter,3) > 0 Then
				sBody = sBody & "Gift Wrap Price: " & FormatCurrency(aProdInfoAE(iCounter,4)) & VbCrLf
				sBody = sBody & "Gift Wrap Qty: " & aProdInfoAE(iCounter,3) & VbCrLf
			End If
		ENd If
		
		IF Application("AppName") = "StoreFront" then 'SFUPDATE
			sBody = sBody & "Product Price: " & FormatCurrency(getGlobalSalePrice(aAllProd(0,iCounter)(5))) & VbCrLf
			sBody = sBody & "Quantity: " & aAllProd(0,iCounter)(3) & VbCrLf
		ENd If
		'sBody = sBody & "Quantity: " & aAllProd(0,iCounter)(3) & VbCrLf 'SFUPDATE
      	'sBody = sBody & "Product ID: " & aAllProd(0,iCounter)(4) & VbCrLf 'SFUPDATE
	Next
	

	IF Application("AppName") = "StoreFrontAE" then 'SFAE
		If Session("CouponDiscountPercent") + Session("CouponDiscountAmount") > 0  Then
			sBody = sBody & vbcrlf & "Coupon Discount: -" & FormatCurrency(Session("CouponDiscountPercent") + Session("CouponDiscountAmount"))
		End If		
		If Session("StoreWideDiscount") > 0 then
			sBody = sBody & vbcrlf & "Store Wide Discount: -" & FormatCurrency(Session("StoreWideDiscount"))
		End If
	End If
	
	sBody = sBody & VbCrLf & "SubTotal:    " & FormatCurrency(sTotalPrice) & VbCrLf & "Shipping:    " & FormatCurrency(sShipping) & "(" & sShipMethodName & ")" & VbCrLf & "Handling:    " & FormatCurrency(sHandling) _
		& VbCrLf & "State Tax:   " & FormatCurrency(iSTax) & VbCrLf & "Country Tax: " & FormatCurrency(iCTax) & VbCrLf & "Grand Total: " & FormatCurrency(sGrandTotal)
		
	
	IF Application("AppName") = "StoreFrontAE" then 'SFAE
		If  Session("SpecialBilling") <> 0 then
			sBody = sBody & vbcrlf & "Billed Amount:" & FormatCurrency(Session("BillAmount"))
			sBody = sBody & vbcrlf & "Remaining Amount:" & FormatCurrency(Session("BackOrderAmount"))
		End if
	End IF
 	If iShipType = 2 Then
		sBody = sBody & vbcrlf & "NOTICE: The shipping costs as shown above do not necessarily represent the carrier's published rates and may include additional charges levied by the merchant."
	End If

	sBody = sBody & VbCrLf  & VbCrLf & "Special Instructions: " &  sShipInstructions 
	
		
	If vDebug = 1 Then Response.Write "<P><B> BODY:</b> " & sBody
	
	'Merchant Email Body
	sPath = LCase(sPath)
	sPath = Replace(sPath, "/process_order.asp", "")
	sPath = sPath & "/admin/sfReports1.asp?OrderID=" & iOrderID
	sMBody = sBody & VbCrLf & "Retrieve Order:" & VbCrLf & sPath
	'Customer Email Body
	sCBody = sBody & VbCrLf & "User Name: " & sCustEmail & VbCrLf & "Password: " & sMailPassword & VbCrLf & "Use this information next time you order for quick access to your Customer Information."
Elseif sType = "FPWD" Then
	' Build Body
	arrInfo = split(sInformation, "|")
	sCustEmail = arrInfo(0)
	sBody = "Here is the password you requested. Please use it for login to " & sHomePath & VbCrLf
	sBody = sBody & "Your password for the e-mail account : " & sCustEmail & " is : " & arrInfo(1)
	sSubject = "Requested Password"
ElseIf sType = "PromoMail" Then
	arrInfo = split(sInformation, "|")
	sCustEmail = arrInfo(0)
	sSubject = arrInfo(1)
	sMessage = arrInfo(2)
	sCBody = "To Remove yourself from the mailing list please go to this link:" & VbCrLf _ 
			& sHomePath & "unsubscribe.asp?email=" & arrInfo(0)
	sSecondary = "" 
ElseIf sType= "EmailFriend" Then
	arrInfo = split(sInformation, "|")
	'dim i
	
	'for i = 0 to ubound(arrinfo)
	'Response.Write i & ":  "  &  arrinfo(i) & "<BR>"
	'next
	sCustEmail = arrInfo(0)
	sPrimary = arrInfo(1)
	sSecondary = arrInfo(5)
	sSubject = arrInfo(4)
	sMessage = arrInfo(2) & vbcrlf & sHomePath & "detail.asp?product_id=" & server.urlencode((arrInfo(3)))
ElseIf sType ="EmailWishList" Then 'SFAE
	arrInfo = split(sInformation, "|")
	sCustEmail = arrInfo(0)
	sPrimary = arrInfo(1)
	sSecondary = arrInfo(5)
	sSubject = arrInfo(4)
	sMessage = arrInfo(2)
	
ElseIf sType ="InvenNotification" Then 'SFAE
	arrInfo = split(sInformation,"|")
	sPrimary = sPrimary 
	sCustEmail = sPrimary 
	sSecondary = sSecondary
	sSubject = arrInfo(0)
	sMessage = ""
	sMBody = arrInfo(1)
	
End If




If sMailMethod = "Simple Mail" Then
		Set Mailer = Server.CreateObject("SimpleMail.smtp.1")
		Mailer.OpenConnection sMailServer
		
		If sType ="FPWD" Then
			Mailer.SendMail sCustEmail, sSubject, sBody
		Else
			If sType <> "PromoMail"  And sType <> "EmailFriend" And sType <> "EmailWishList" Then 
			Mailer.SendMail sPrimary, sCustEmail, sSubject, sMBody
			end if
			If sType <> "PromoMail" And sType <> "EmailFriend" And sType <> "EmailWishList" and sSecondary <> "" Then 
			Mailer.SendMail sSecondary, sCustEmail, sSubject, sMBody
			end if
			If sType = "InvenNotification" Then 
				Set Mailer = nothing
				Exit sub 'SFAE
			End IF
			
			Mailer.SendMail sCustEmail, sPrimary, sSubject, sMessage & VbCrLf & sCBody
			'#338
			If sSecondary <> "" And Not IsNull(sSecondary) Then
                       If LCase(sType) <> "confirm" Then
                         Mailer.SendMail sSecondary, sPrimary, sSubject, sMessage & VbCrLf & sCBody				
                       
                      End If
             End If
			
		End If		
		
		Mailer.CloseConnection
		Set Mailer = nothing
ElseIf sMailMethod = "CDOSYS" Then	

		Dim Message 'As New CDO.Message
		Dim Configuration 'As New CDO.Configuration
		Dim Fields 'As ADODB.Fields
		Set Message = Server.CreateObject("CDO.Message")
		Set Configuration = Server.CreateObject("CDO.Configuration")
		Dim Item
		'cdoSendUsingMethod, cdoSMTPServerPort, cdoSMTPServer, cdoSMTPConnectionTimeout, cdoSMTPAuthenticate, cdoURLProxyServer, cdoURLProxyBypass, cdoURLGetLatestVersion
		Set Fields = Configuration.Fields
			With Fields
			    Fields(cdoSendUsingMethod) = 2
			    Fields(cdoSMTPServerPort) = 25
			    Fields(cdoSMTPServer) = sMailServer
			    Fields(cdoSMTPConnectionTimeout) = 20
 				Fields(cdoSMTPAuthenticate)      = 0
  				Fields(cdoURLProxyServer)        = "server:80"
  				Fields(cdoURLProxyBypass)        = "<local>"
  				Fields(cdoURLGetLatestVersion)   = True
			    Fields(cdoSendUserName)         = ""
  				Fields(cdoSendPassword)         = ""
			    .Update
			End With
			With Message
			         Set .Configuration = Configuration
			       If sType = "FPWD" Then
			         .To = sCustEmail
			         .From = sPrimary
			         .Subject = sSubject
			         .TextBody = sBody
			         .send
			       Else
			          If sType <> "PromoMail" And sType <> "EmailFriend" Then
			             .To = sPrimary
			             If sSecondary <> "" And Not IsNull(sSecondary) Then
			              .CC = sSecondary
			             End If
			            .From = sCustEmail
			            .Subject = sSubject
			            .TextBody = sMBody
			            .send
			            If sType = "InvenNotification" Then
			                 Exit Sub 'SFAE
			            End If
			          End If
			          
			      If sSecondary <> "" And Not IsNull(sSecondary) Then
                       If LCase(sType) <> "confirm" Then
                         .CC = sSecondary
                       Else
                         .CC = ""
                      End If
                  End If
			       .To = sCustEmail
			       .From = sPrimary
			       .Subject = sSubject
			       .TextBody = sMessage & vbCrLf & sCBody
			       .send
			    End If
			 End With


		Set Message = Nothing
		Set Configuration = Nothing
		Set Fields = Nothing
					
ElseIf sMailMethod = "SimpleMail 2.0" Then

		Set Mailer = Server.CreateObject("SimpleMail.smtp")
		Mailer.OpenConnection sMailServer
		If sType = "FPWD" Then
			Mailer.SendMail sCustEmail,sPrimary,"",sSubject,sBody '#396
		Else
			If sType <> "PromoMail" And sType <> "EmailFriend" And sType <> "EmailWishList" Then 
				Mailer.SendMail sPrimary, sCustEmail, sSecondary, sSubject, sMBody
			End If
			If sType = "InvenNotification" Then 
				Set Mailer = nothing
				Exit sub 'SFAE
			End IF
			'#338
			 If sSecondary <> "" And Not IsNull(sSecondary) Then
                If LCase(sType) <> "confirm" Then
           			Mailer.SendMail sCustEmail, sPrimary, sSecondary, sSubject, sMessage & VbCrLf & sCBody
				else
					Mailer.SendMail sCustEmail, sPrimary, "", sSubject, sMessage & VbCrLf & sCBody			
				End If
             End If

			'Mailer.SendMail sCustEmail, sPrimary, sSecondary, sSubject, sMessage & VbCrLf & sCBody
		End If
	
		Mailer.CloseConnection
		Set Mailer = nothing
		
ElseIf sMailMethod = "ASP Mail" Then
		Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		Mailer.QMessage = TRUE
		If sType="FPWD" Then
			Mailer.RemoteHost = sMailServer
			Mailer.AddRecipient sCustEmail, sCustEMail
			Mailer.FromAddress = sPrimary
			Mailer.FromName = sPrimary
			Mailer.Subject = sSubject
			Mailer.BodyText = sBody
			Mailer.SendMail
		Else		
			If sType <> "PromoMail" And sType <> "EmailFriend" And sType <> "EmailWishList" Then 
				Mailer.RemoteHost = sMailServer
				Mailer.AddRecipient sPrimary, sPrimary
				Mailer.AddRecipient sSecondary, sSecondary
				Mailer.FromAddress = sCustEmail
				Mailer.FromName = sCustEmail
				Mailer.Subject = sSubject
				Mailer.BodyText = sMBody
				Mailer.SendMail
				Mailer.ClearBodyText
				Mailer.ClearRecipients
				If sType = "InvenNotification" Then 
					Set Mailer = nothing
					Exit sub 'SFAE
				End IF
			
			End If
			Mailer.RemoteHost = sMailServer
			Mailer.AddRecipient sCustEmail, sCustEmail
			If sSecondary <> "" And Not IsNull(sSecondary) Then
                       If LCase(sType) <> "confirm" Then
						Mailer.AddRecipient sSecondary, sSecondary
                      End If
             End If
			
			Mailer.FromAddress = sPrimary
			Mailer.FromName = sPrimary
			Mailer.Subject = sSubject
			Mailer.BodyText = sMessage & VbCrLf & sCBody
			Mailer.SendMail
			Mailer.ClearBodyText
			Mailer.ClearRecipients
		End If
		Set Mailer = Nothing

ElseIf sMailMethod = "CDONTS Mail" Then
		Set Mailer = Server.CreateObject("CDONTS.NewMail")
		Mailer.MailFormat = 0
		
		'#396
		
		If sType = "FPWD" Then
			Mailer.To = sCustEmail		
			Mailer.From = sPrimary
			Mailer.Subject = sSubject
			Mailer.Body = sBody
			Mailer.Send			
			Set Mailer = Nothing
		Else
		
			If sType <> "PromoMail" And sType <> "EmailFriend" And sType <> "EmailWishList" Then 
				Mailer.To = sPrimary
				If sSecondary <> "" AND Not IsNull(sSecondary) Then
				Mailer.Cc = sSecondary
				End If
				
				Mailer.From = sCustEmail
				Mailer.Subject = sSubject
				Mailer.Body = sMBody
				Mailer.Send					
				Set Mailer = Nothing
				
				If sType = "InvenNotification" Then 
					Set Mailer = nothing
					Exit sub 'SFAE
				End IF
			
				
			End If
			
		
			Set Mailer2 = Server.CreateObject("CDONTS.NewMail")
			Mailer2.MailFormat = 0
			 If sSecondary <> "" And Not IsNull(sSecondary) Then
                       If LCase(sType) <> "confirm" Then
                         Mailer2.CC = sSecondary
                       Else
                         Mailer2.CC = ""
                      End If
             End If
		
		
			Mailer2.To = sCustEmail
			
			Mailer2.From = sPrimary
			Mailer2.Subject = sSubject
			Mailer2.Body = sMessage & VbCrLf & sCBody
			Mailer2.Send
			Set Mailer2 = nothing
		End if
		Set Mailer = nothing
		
ElseIf sMailMethod= "AB Mail" Then
		Set Mailer = Server.CreateObject("ABMailer.Mailman")
		Mailer.Clear
		
		If sType = "FPWD" Then
			Mailer.SendTo = sCustEmail
			Mailer.ReplyTo = sPrimary
			Mailer.FromAddress=sPrimary
			Mailer.MailSubject = sSubject			
			Mailer.MailDate = ""
			Mailer.ServerAddr = sMailServer
			Mailer.MailMessage = sBody
			Mailer.SendMail			
		Else
			If sType <> "PromoMail" And sType <> "EmailFriend" And sType <> "EmailWishList" Then 
				Mailer.SendTo = sPrimary
				Mailer.ReplyTo = sCustEmail
				Mailer.FromAddress=sCustEmail
				Mailer.MailSubject = sSubject
				Mailer.SendCc = sSecondary
				Mailer.MailDate = ""
				Mailer.ServerAddr = sMailServer
				Mailer.MailMessage = sMBody
				Mailer.SendMail
				If sType = "InvenNotification" Then 
					Set Mailer = nothing
					Exit sub 'SFAE
				End IF
			End If
			
			Mailer.Clear
			Mailer.SendTo = sCustEmail

			 If sSecondary <> "" And Not IsNull(sSecondary) Then
                       If LCase(sType) <> "confirm" Then
                         			Mailer.SendCc = sSecondary
                       Else
                         			Mailer.SendCc = ""
                      End If
             End If

			
			Mailer.ReplyTo = sPrimary
			Mailer.FromAddress=sPrimary
			Mailer.MailSubject = sSubject
			Mailer.MailDate = ""
			Mailer.ServerAddr = sMailServer
			Mailer.MailMessage = sMessage & VbCrLf & sCBody
			Mailer.SendMail
		End If
		Set Mailer = nothing

ElseIf sMailMethod = "Bamboo Mail" Then
		Set Mailer = Server.CreateObject("Bamboo.SMTP")
		Mailer.Server = sMailServer
		
		If sType = "FPWD" Then 
			Mailer.RCPT = sCustEmail
			Mailer.From = sPrimary
			Mailer.Fromname = sPrimary
			Mailer.Subject = sSubject
			Mailer.Message =  sBody
			Mailer.Send
			
		Else
		If sType <> "PromoMail" And sType <> "EmailFriend" And sType <> "EmailWishList" Then 
				Mailer.RCPT = sPrimary
				Mailer.From = sCustEmail
				Mailer.FromName = sCustEmail
				Mailer.Subject = sSubject
				Mailer.Message = sMessage & VbCrLf & sMBody
				Mailer.Send
				If sType = "InvenNotification" Then 
					Set Mailer = nothing
					Exit sub 'SFAE
				End IF
			
			End If
			Set Mailer2 = Server.CreateObject("Bamboo.SMTP")
			Mailer2.Server = sMailServer
			Mailer2.RCPT = sCustEmail
			Mailer2.From = sPrimary
			Mailer2.FromName = sPrimary
			Mailer2.Subject = sSubject
			Mailer2.Message = sMessage & VbCrLf & sCBody
			Mailer2.Send
			Set Mailer2 = nothing
		End If
		Set Mailer = nothing
		
ElseIf sMailMethod = "J Mail" Then
		Set Mailer = Server.CreateObject("JMail.SMTPMail")
		
		If sType = "FPWD" Then
			Mailer.ServerAddress = sMailServer
			Mailer.Sender = sPrimary
			Mailer.SenderName = sPrimary
			Mailer.AddRecipientEx sCustEmail, sCustEmail	'#396		
			Mailer.Subject = sSubject
			Mailer.Body = sBody
			Mailer.Execute
		Else
			If sType <> "PromoMail" And sType <> "EmailFriend" And sType <> "EmailWishList" Then 
				Mailer.ServerAddress = sMailServer
				Mailer.Sender = sCustEmail
				Mailer.SenderName = sCustEmail
				Mailer.AddRecipientEx sPrimary, sPrimary
				Mailer.AddRecipientEx sSecondary, sSecondary
				Mailer.Subject = sSubject
				Mailer.Body = sMBody
				Mailer.Execute
				If sType = "InvenNotification" Then 
					Set Mailer = nothing
					Exit sub 'SFAE
				End IF
			
			End If

			Mailer.ServerAddress = sMailServer
			Mailer.Sender = sPrimary
			Mailer.SenderName = sPrimary
			Mailer.AddRecipientEx sCustEmail, sCustEmail
			'#338
			If sSecondary <> "" And Not IsNull(sSecondary) Then
                If LCase(sType) <> "confirm" Then
					Mailer.AddRecipientEx sSecondary, sSecondary
				end if
			end if
			Mailer.Subject = sSubject
			Mailer.Body = sMessage & VbCrLf & sCBody
			Mailer.Execute
		End If
		Set Mailer = nothing
			
ElseIf sMailMethod = "OCX Mail" Then
		Set Mailer = Server.CreateObject("ASPMail.ASPMailCtrl.1")
	
		If sType = "FPWD" Then
			Mailer.SendMail sMailServer, sCustEmail, sPrimary, sSubject, sBody			
		Else
			If sType <> "PromoMail" And sType <> "EmailFriend" And sType <> "EmailWishList" Then 
			Mailer.SendMail sMailServer, sPrimary, sCustEmail, sSubject, sMBody
			end if
			If sType <> "PromoMail" And sType <> "EmailFriend" And sType <> "EmailWishList" and sSecondary <> "" Then 
			Mailer.SendMail sMailServer, sSecondary, sCustEmail, sSubject, sMBody
			end if
			If sType = "InvenNotification" Then 
				Set Mailer = nothing
				Exit sub 'SFAE
			End IF
			Mailer.SendMail sMailServer, sCustEmail, sPrimary, sSubject, sMessage & VbCrLf & sCBody
			If  sSecondary <> "" Then 
			 If LCase(sType) <> "confirm" Then
			Mailer.SendMail sMailServer, sSecondary, sPrimary, sSubject, sMessage & VbCrLf & sCBody
			end if
			end if
		End If
		
		Set Mailer = nothing
		
ElseIf sMailMethod = "No Mail" Then
End If
		
End Sub
%>







