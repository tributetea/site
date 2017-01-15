<%@ Language=VBScript %>
<%	option explicit 
	Response.Buffer = True
	%>
<!--#include file="SFLib/incDesign.asp"-->
<!--#include file="SFLib/db.conn.open.asp"-->
<!--#include file="SFLib/adovbs.inc"-->
<!--#include file ="SFLib/incLogin.asp"-->
<!--#include file ="SFLib/incGeneral.asp"-->
<!--#include file="SFLib/mail.asp"-->
<!--#include file="error_trap.asp"-->
<%
	Const vDebug = 0

'@BEGINVERSIONINFO

'@APPVERSION: 50.4011.0.2

'@FILENAME: login.asp
 

'

'@DESCRIPTION: Handles customer login

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000, 2001 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO
%>
<%
Dim sFirstTime, sReturn, sFormActionType, sThanks, iCustID, sOutput, sChangeUser, sCustomerForm, sSubmitAction
Dim iCookieID, sForgotPwd, sSendEmail, iAuthenticate, bSvdCartCustomer, bEmailAddress, sNewAccount, sChangeCart, sChange, sDirection, sSvdCartCustEmail, sPaymentMethod
Dim sLoggedIn, sEmail, sPassword, sLogin

'------------------------------------------------------------
' For savecart
'------------------------------------------------------------

' Check which button has been depressed
sFirstTime	  	= Trim(Request.Form("SignUp.x"))
sLoggedIn	  	= Trim(Request.Cookies("sfOrder")("SessionID"))
sReturn	  	= Trim(Request.Form("Return.x"))
sChangeUser 	= Trim(Request.Form("ChangeUser.x"))
sForgotPwd	  	= Trim(Request.QueryString("FPWD")) ' Request.Form("FPWD.x")
sSendEmail	  	= Trim(Request.Form("SendEmail.x"))
sNewAccount  	= Trim(Request.QueryString("New")) '.Form("New.x")
If sNewAccount = "" Then
	sNewAccount = Trim(Request.QueryString("Type"))
End If
sChangeCart = Trim(Request.Form("ChangeCart.x"))
sChange		= Trim(Request.Form("Change.x"))
sLogin = Request.Cookies("sfThanks")("PreviousAction")
iCustID = Request.Cookies("sfCustomer")("custID")

 ' Get payment method and write to a cookie
If Trim(Request.Form("PaymentMethod")) <> "" Then
    sPaymentMethod = Trim(Request.Form("paymentmethod"))
    Response.Cookies("sfOrder").Expires = Date() + 1
    Response.Cookies("sfOrder")("PaymentMethod") = sPaymentMethod    
End If


' For people already logged in 
If (sLoggedIn = Session("SessionID") AND iCustID <> "" And sChangeCart = "" And sChange = "" AND sForgotPwd = "") Then
	
	' Possibility - get a new account 
	If sNewAccount <> "" Then	   
	   sFormActionType = "NewAccount"
	
	' Through new account, get them a new account   
	ElseIf sFirstTime <> "" Then	
		sFormActionType = "FirstTime"
	
	' If there is email and password, authenticate
	ElseIf Trim(Request.Form("Email")) <> "" And Trim(Request.Form("Passwd")) <> "" Then
		sEmail	  = Trim(Request.Form("Email"))
		sPassword = Trim(Request.Form("Passwd"))
	
		iAuthenticate = customerAuth(sEmail,sPassword,"loose")
							
		If iAuthenticate <> "" AND iAuthenticate > 0 Then
				' Check if there is a custID
				If Trim(Request.Cookies("sfCustomer")("custID")) = "" Then
					sFormActionType = "Returning"	
				Else	
					If Request.Cookies("sfCustomer")("custID") <> iAuthenticate Then
					   Response.Cookies("sfCustomer")("custID") = iAuthenticate
					   Response.Cookies("sfCustomer").Expires = Date() + 730
					End If		
					' Redirect to proc_order
					  Response.Redirect(C_SecurePath)
				' End custId if		
				End If 		  
		  Else
				sFormActionType = "FirstTime"
		  ' End Authenticate If
		  End If		 
		   
		Else
			If Request.Cookies("sfCustomer")("custID") <> "" Then
				 Response.Redirect(C_SecurePath)
			Else
				sFormActionType	= "Returning"	
			End If		 	
				
		' End auth if		
		End If			
	
Else 
	If sFirstTime <> "" Then
		sFormActionType = "FirstTime"
	ElseIf sReturn <> "" Then
		sFormActionType = "Returning"	
	ElseIf sForgotPwd <> "" Then
		sFormActionType = "ForgotPwd"
	ElseIf sSendEmail <> "" Then
		sFormActionType = "SendEmail"	
	ElseIf sNewAccount <> "" Then
		sFormActionType = "NewAccount"	
	ElseIf sChangeCart <> "" Then
		sFormActionType = "ChangeCart"	
	ElseIf sChange <> "" Then
		sFormActionType = "Change"	
	End If
End If


'-----------------------------------------------------------
' Cases for saved cart
'-----------------------------------------------------------	
	
	If vDebug = 1 Then Response.Write "<br>FormActionType: " & sFormActionType	
		
	' First time at login, no actions	
	If sFormActionType = "" And Request.Cookies("sfCustomer")("custID") = "" Then		
		sOutput = "General"
		
	ElseIf sFormActionType = "" And Request.Cookies("sfCustomer")("custID") <> "" Then		
		sOutput = "HasID"
 
	ElseIf sFormActionType = "NewAccount" Then		
		sOutput = "NewAccount"
	
	' Forgot password
	ElseIf sFormActionType = "ForgotPwd" Then		
		sOutput = "Email"		
	
	ElseIf sFormActionType = "ChangeCart" Then
		sOutput = "ChangeCart"
	
	ElseIf sFormActionType = "Change" Then
		sEmail = Trim(Request.Form("Email"))
		sPassword = Trim(Request.Form("Passwd"))

		' Authenticate
		iAuthenticate = customerAuth(sEmail,sPassword,"loose")
		
		If iAuthenticate <> "" AND iAuthenticate > 0 Then
			' Write New Cookie
			Response.Cookies("sfCustomer")("custID") = iAuthenticate
			Response.Cookies("sfCustomer").Expires = Date() + 730		
			
			' Associate sessionID with cookie
			Response.Cookies("sfOrder")("SessionID") = Session("SessionID")
			Response.Cookies("sfOrder").Expires = Date() + 1
			
			' Redirect to Svd Cart
			Response.Redirect("savecart.asp")
		Else 
			sOutput = "FailedAuthChange"
		End If		
	
	' Send Mail
	ElseIf sFormActionType = "SendEmail" Then		
		sEmail = Request.Form("Email")
		' Send email with password, returns a success or failure boolean
		bEmailAddress = SendPassword(sEmail)
		
		If bEmailAddress = 1 Then
			sOutPut = "SentEmail"
		Else
			sOutPut = "FailedEmail"
		End If		
		
	' First time login		
	ElseIf sFormActionType = "FirstTime" Then	
		 sEmail = Request.Form("Email")
		 sPassword = Request.Form("Passwd")
		
		' Check if email and password correspond to any customer on record	
			If customerAuth(sEmail,sPassword,"loose") > 0 Then
					
					' Get custID
					iCustID = customerAuth(sEmail,sPassword,"loose")	
				
					' Write a new cookie with found custID
					Response.Cookies("sfCustomer")("custID") = iCustID
					Response.Cookies("sfCustomer").Expires = Date() + 730	
				
					' Associate sessionID with cookie
					Response.Cookies("sfOrder")("SessionID") = Session("SessionID")
					Response.Cookies("sfOrder").Expires = Date() + 1				
			
						' For Customers from SaveCart, special case
						If sLogin <> "" Then
				
							' Update Saved Table with iCustID in Table
							Call UpdateCustID(iCustID)							
									If Request.Cookies("sfThanks")("PreviousAction") = "SaveCart" Then
										' Redirect to thanks page
										Response.Cookies("sfThanks").Expires = NOW()
										Response.Redirect "addproduct.asp?logedin=1" 
							
									ElseIf Request.Cookies("sfThanks")("PreviousAction") = "FromShopCart" Then
										' Delete from temp										
										Call setDeleteOrder("odrdttmp",Request.Cookies("sfThanks")("DeleteTmpOrderID"))
										Response.Cookies("sfThanks")("DeleteTmpOrderID") = ""
										Response.Cookies("sfThanks").Expires = NOW()										
										Response.Redirect "order.asp" 					
									End If	
							
						' End SaveCart New customers If
						End If 	

			ElseIf customerAuth(sEmail,sPassword,"loosest") > 0 Then			
				' Email match. Prompt for new email				  
				sOutput = "EmailMatch"
			
			Else		
				
				' Write to customer table, write a cookie, and get back the id							
				iCustID = getCustomerID(sEmail,sPassword)	
			   If vDebug = 1 Then Response.write "CustID:" & iCustID
			   
 						' Associate sessionID with cookie
						Response.Cookies("sfOrder")("SessionID") = Session("SessionID")
						Response.Cookies("sfOrder").Expires = Date() + 1
	
					    ' For Customers from SaveCart, special case
						If Request.Cookies("sfThanks")("PreviousAction") <> "" Then
				
							' Update Saved Table with iCustID in Table
							Call UpdateCustID(iCustID)							
							
							If Request.Cookies("sfThanks")("PreviousAction") = "SaveCart" Then							
								Response.Cookies("sfThanks").Expires = NOW()
								' Redirect to thanks page
								Response.Redirect "addproduct.asp?logedin=1" 
							ElseIf Request.Cookies("sfThanks")("PreviousAction") = "FromShopCart" Then
								' Delete from temp
								Call setDeleteOrder("odrdttmp",Request.Cookies("sfThanks")("DeleteTmpOrderID"))
								Response.Cookies("sfThanks")("DeleteTmpOrderID") = ""
								Response.Cookies("sfThanks").Expires = NOW()
								Response.Redirect "order.asp"								
							End If								
							
						' End SaveCart New customers If
						End If 	
			
				    ' Assume new customer checkout for other cases
					' Redirect to form to enter customer Info  
					Response.Redirect("order.asp")
					
		' End existing cookie if			
		  End If			

	ElseIf sFormActionType = "Returning" Then	
		sEmail = Trim(Request.Form("Email"))
		sPassword = Trim(Request.Form("Passwd"))
		
		' Authenticate customer
		iAuthenticate = customerAuth(sEmail,sPassword,"loose")
			
		If iAuthenticate <> "" AND iAuthenticate > 0 Then
		
			' Associate sessionID with cookie
			Response.Cookies("sfOrder").Expires = Date() + 1
			Response.Cookies("sfOrder")("SessionID") = Session("SessionID")				
			
				' Check if customer still has a cookie				
				iCustID = Request.Cookies("sfCustomer")("custID")				
			
				' Write to cookie if none exists for custID	
				If iCustID <> iAuthenticate Or iCustID = "" Then
					Response.Cookies("sfCustomer")("custID") = iAuthenticate					
					Response.Cookies("sfCustomer").Expires = Date() + 730
					iCustID = iAuthenticate
				End If			
				
						' For Customers from SaveCart, special case
						If Request.Cookies("sfThanks")("PreviousAction") <> "" Then
										
							' Update Saved Table with iCustID in Table
							Call UpdateCustID(iCustID)
		
							If Request.Cookies("sfThanks")("PreviousAction") = "SaveCart" Then
								' Redirect to thanks page
								Response.Cookies("sfThanks").Expires = NOW()
								Response.Redirect "addproduct.asp?logedin=1" 
							ElseIf Request.Cookies("sfThanks")("PreviousAction") = "FromShopCart" Then
								' Delete from temp
								Call setDeleteOrder("odrdttmp",Request.Cookies("sfThanks")("DeleteTmpOrderID"))
								Response.Cookies("sfThanks")("DeleteTmpOrderID") = ""
								Response.Cookies("sfThanks").Expires = NOW()
								Response.Redirect "order.asp"								
							End If	
							
						' End SaveCart New customers If
						End If 
			
				' Check if it is a savedcart customer
				bSvdCartCustomer = getSvdCartCustomer(iCustID,"boolean")
			
				If bSvdCartCustomer = 1 Then
					If vDebug = 1 Then Response.Write "<Br> Wish List Customer " & iCustID
					Response.Redirect("order.asp")					
				Else
					' Redirect to proc_order
					Response.Redirect("savecart.asp")
				End If 
				
		Else
			' Assume new person
			' Write to customer table, write a cookie, and get back the id							
				iCustID = getCustomerID(sEmail,sPassword)	
			   If vDebug = 1 Then Response.write "CustID:" & iCustID
			   
 						' Associate sessionID with cookie
						Response.Cookies("sfOrder")("SessionID") = Session("SessionID")
						Response.Cookies("sfOrder").Expires = Date() + 1
	
					    ' For Customers from SaveCart, special case
						If Request.Cookies("sfThanks")("PreviousAction") <> "" Then
				
							' Update Saved Table with iCustID in Table
							Call UpdateCustID(iCustID)							
							
							If Request.Cookies("sfThanks")("PreviousAction") = "SaveCart" Then							
								Response.Cookies("sfThanks").Expires = NOW()
								' Redirect to thanks page
								Response.Redirect "addproduct.asp?logedin=1" 
							ElseIf Request.Cookies("sfThanks")("PreviousAction") = "FromShopCart" Then
								' Delete from temp
								Call setDeleteOrder("odrdttmp",Request.Cookies("sfThanks")("DeleteTmpOrderID"))
								Response.Cookies("sfThanks")("DeleteTmpOrderID") = ""
								Response.Cookies("sfThanks").Expires = NOW()
								Response.Redirect "order.asp"								
							End If								
							
						' End SaveCart New customers If
						End If 	
						
					' If all else fails, just go to order.asp	
					Response.Redirect("order.asp")  
		End If	

	' End FormAction If	
	End If	
%>
<html>
<head>
<script language="javascript" src="SFLib/sfCheckErrors.js"></script>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
</script>
<title><%= C_STORENAME %>-StoreFront Web Store Login Page</title>



<!--Header Begin -->
<link rel="stylesheet" href="sfCSS.css" type="text/css">
</head>

<body link="<%= C_LINK %>" vlink="<%= C_VLINK %>" alink="<%= C_ALINK %>">

<table border="0" cellpadding="1" cellspacing="0"  class="tdbackgrnd" width="<%= C_WIDTH %>" align="center">
  <tr>
    <td>
      <table width="100%" border="0" cellspacing="1" cellpadding="3">
			 <tr>
          <td align="middle"  class="tdTopBanner"> 
            <%If C_BNRBKGRND = "" Then%>
            <%Else%>
            <img src="buttons/tt_blue.gif" border="0" width="275" height="36"> 
            <%End If%>
          </td>
			</tr>       
<!--Header End --> 
        <%
    '-----------------------------------------------------------
    ' Begin OutPut Block
    '-----------------------------------------------------------
    %>    
     
        <tr>
          <td   class="tdMiddleTopBanner">
  	        <%If sOutPut = "General" Then %>  
			Please Login
	        <%ElseIf sOutPut = "NewAccount" Then %>
			New Account 		
	        <%ElseIf sOutPut = "EmailMatch" Then %>	
			Matching Email Found			
	        <%ElseIf sOutPut = "HasID" Then %>  
			Returning Customer Login		
	        <%ElseIf sOutPut = "FailedAuth" Then %>	
			Failed Authentication
	        <%ElseIf sOutPut = "FailedAuthChange" Then %>	
			Change Cart Failed Authentication			
	        <%ElseIf sOutPut = "Email" Then %>	
			Send Password to Email Address
	        <%ElseIf sOutPut = "SentEmail" Then %>
			Email Sent to Address		
	        <%ElseIf sOutPut = "FailedEmail" Then %>
			No Email Found On Record 				
	        <%ElseIf sOutPut = "ChangeCart" and Application("AppName")= "StoreFrontAE" Then %>	
			Change Wish List	
			<%ElseIf sOutPut = "ChangeCart" and Application("AppName")<> "StoreFrontAE" Then %>	
			Change Saved Cart	
	        <%Else%>
	        Login with Email and Password
	        <%
	End If 
	
	If sOutPut = "FailedAuth" Then
		sSubmitAction = "" 
	Else
		sSubmitAction = "return sfCheck(this)"
	End If 
	%>	
		  </td>
		</tr>
		
		<% If sOutPut <> "Email" AND sOutPut <> "SentEmail"  AND sOutPut <> "FailedEmail" Then %>
		<tr>
		  <td class="tdBottomTopBanner">
		    <%If sOutPut = "NewAccount" Then %>  
			 Your Saved Order is your personal, private collection of items that you are thinking about purchasing. Put as many items in your Saved Order as you want. When you decide to buy them, simply add them to your Current Order.<P>To access an existing Saved Order, log in with the e-mail and password you chose when you first signed in. If you've forgotten your password, click on
            Forgot Password for help. To create a new Saved Order, select New Account. 
            <% Elseif sOutPut = "General" Then %>
            Your Saved Order is your personal, private collection of items that you are thinking about purchasing. Put as many items in your Saved Order as you want. When you decide to buy them, simply add them to your Current Order.
           	 To create a new Saved Order, select New Account.            
		    <%ElseIf sOutPut = "HasID" Then %>  
            Please login with the e-mail address and password you chose when you first signed in. If you've forgotten your password, click on
            Forgot Password for help
		    <%ElseIf sOutPut = "FailedAuth" Then %>		
			We're sorry, but your combination of e-mail address and password is not recognized. If you've forgotten your password, you can click on
            Forgot Password and an e-mail will be sent to the address on record. Alternatively, you can sign in as a new user.	
		    <%ElseIf sOutPut = "FailedAuthChange" Then %>		
			We're sorry, but your combination of e-mail address and password is not recognized. Please retype them.
		    <%ElseIf sOutPut = "Email" Then %>
			If you're an existing customer (you've made a purchase or saved something to your order previously), then we can retrieve your password if you type in the e-mail address on record.
		    <%ElseIf sOutPut = "ChangeCart" Then %>	
			Please enter the e-mail address and password corresponding to an existing saved
            order. Your current saved order can be accessed through the change
            order option on the saved order page.
		    <%ElseIf sOutPut = "EmailMatch" Then %>	
			A matching e-mail address has been found. Please either create a new account, or click on
            Forgot Password for more help
		    <%Else%>
		    Please choose a login with an e-mail account and a password. This will be used for future retrieval of billing and shipping records.
		    <%End If %>	
		    </td>
		  </tr>
		  <tr>
		    <td class="tdContent" align="middle"><br>
		   
			  <%If (sOutput <> "FailedAuth" and sOutPut <> "HasID" and sOutPut <> "ChangeCart" and sOutPut <> "FailedAuthChange") Then
			
				If Request.Cookies("sfCustomer")("custID") = "" Or sOutPut = "NewAccount" Or sOutput = "EmailMatch" Or sOutPut <> "SameUser" Then %>		   
				<form method="post" action="login.asp" name="login" onSubmit="<%=sSubmitAction%>">
		          <table border="0" width="75%">
		            <tr>
		              <td width="100%" align="middle" class="tdBottomTopBanner2">
		       	       
			            <% If sOutPut = "NewAccount" or sOutput = "EmailMatch" Then %>
					New Account 		
		                <% Else  %>
					First Time Here?      
		                <% End If %>
		        
		              </td>
		            </tr>
		            <tr>
		              <td width="100%" align="middle" class="tdContent">
		                <table border="0" width="100%">
		                  <tr>
		                    <td width="50%" align="right"><b>E-Mail Address:</b></td>
		                    <td width="50%">
                            <input type="text" name="Email" title="Email Address" style="<%= C_FORMDESIGN %>" maxlength="100">
		                    </tr>
		                    <tr>
		                      <td width="50%" align="right"><b>Password:</b></td>
		                      <td width="50%">
                              <input type="Password" name="Passwd" title="Password" style="<%= C_FORMDESIGN %>" maxlength="10"></td>
		                    </tr>
		                    <tr>
		                      <td width="100%" align="middle" colspan="2">
					            
                          <input type="image" src="buttons/signup.gif" width="92" height="22" border="0" name="SignUp" onsubmit="javascript:CheckLoginInput(this)">   
		                      </td>
		                    </tr>
		                  </table>
		                </td>
		              </tr>
		            </table>
		          </form>
		          <br>
		    
					<%If sOutput = "EmailMatch" Then %>
						<form method="post" action="login.asp" name="login" onSubmit="<%= sSubmitAction%>">
					       <table border="0" width="75%">
					        <tr>
					          <td width="100%" align="middle" class="tdBottomTopBanner2">Existing Member Login</td>
					        </tr>
					        <tr>
					          <td width="100%" align="middle" class="tdContent">
					            <table border="0" width="100%">
						           <tr>
						            <td width="50%" align="right"><b> E-Mail:</b></td>
						            <td width="50%">
                                    <input name="Email" type="text" title="Email Address" style="<%= C_FORMDESIGN %>" maxlength="100" SIZE="20"></td>
						            </tr>
						            <tr>
						              <td width="50%" align="right"><b>Password:</b></td>
						              <td width="50%">
                                      <input type="password" name="Passwd" title="Password" style="<%= C_FORMDESIGN %>" maxlength="10" SIZE="20"></td>
						            </tr>
						            <tr>
						              <td width="100%" align="middle" colspan="2">
									    
                          <input type="image" src="buttons/login.gif" name="Return" border="0" width="92" height="22">
						              </td>
						            </tr>
						          </table>
						      </td>
						      </tr>
						      </table>
						      <p>
						      <a href="login.asp?FPWD=True"><img src="<%= C_BTN17 %>" border="0"></a>
						      <!--<input type="image" border="0" src="<%= C_BTN17 %>" name="FPWD">-->
						      </form>
						  <%End If
						
						' End cookies check if
						End If%>
		    
		                  <%Else  %>
		                  <form method="post" action="login.asp" name="login" onSubmit="<%= sSubmitAction%>">
		                    <table border="0"  width="75%">
		                      <tr>
		                        <td width="100%" align="middle" class="tdBottomTopBanner2">
		                          <%If sOutPut = "ChangeCart" Then %>	
					Existing Cart Log In
				                  <%Else%>		
					Existing Member
		                          <%End If%>
		                        </td>
		                      </tr>
		                      <tr>
		                        <td width="100%" align="middle" class="tdContent">
		                          <table border="0" width="100%">
		                            <tr>
		                              <td width="50%" align="right"><b> E-Mail:</b></td>
		                              <td width="50%">
                                      <input name="Email" type="text" title="Email Address" style="<%= C_FORMDESIGN %>" maxlength="100"></td>
		                            </tr>
		                            <tr>
		                              <td width="50%" align="right"><b>Password:</b></td>
		                              <td width="50%">
                                      <input type="password" name="Passwd" title="Password" style="<%= C_FORMDESIGN %>" maxlength="10"></td>
		                            </tr>
		                            <tr>
		                              <td width="100%" align="middle" colspan="2">
		                                <%If sOutPut = "ChangeCart" or sOutPut = "FailedAuthChange" Then %>	
						                
                          <input Type="image" src="buttons/login.gif" name="Change" border="0" width="92" height="22">
		                                <%Else%>
						                
                          <input type="image" src="buttons/login.gif" name="Return" border="0" width="92" height="22">
					                    <%End If%>
		                              </td>
		                            </tr>
		                          </table>
		                        </td>
		                      </tr>
		                    </table>
		                    <p><a href="login.asp?FPWD=True"><img src="<%= C_BTN17 %>" border="0"></a>
		                    <!--<input type="image" border="0" src="<%= C_BTN17 %>" name="FPWD">-->
				            <%If sOutPut = "FailedAuth" or sOutPut = "HasID" Then %>	
					        <a href="login.asp?New=true"><img src="<%= C_BTN19 %>" border="0"></a>
		                    <!--<input type="image" src="<%= C_BTN19 %>" border="0" name="New">-->
		                    <%End If%>	
		                    </form>
		                    <%End If%>
		                  </td>
		          </tr>
        
                  <%  
    Else ' Send Email or print confirmation of sent email
    %>
    
 
	                <tr>
	                  <td class="tdContent" align="middle"><br>
		                <%
		If sOutput = "Email" Then
		%>	
		                <form method="post" action="login.asp" onSubmit="<%= sSubmitAction%>">
                          <table border="0"  width="75%">
                            <tr>
                              <td width="100%" align="middle" class="tdBottomTopBanner2">
			Please Type In Your E-Mail Address	
                              </td>
                            </tr>
                            <tr>
                              <td width="100%" align="middle" class="tdContent">
                                <table border="0" width="100%">
                                  <tr>
                                    <td width="50%" align="right"><b>E-Mail Address:</b></td>
                                    <td width="50%">
                                    <input type="text" name="Email" title="Email Address" style="<%= C_FORMDESIGN %>">
                                    </tr>
                                    <tr>
                                      <td width="100%" align="middle" colspan="2">
				                        
                          <input type="image" src="buttons/submit.gif" border="0" name="SendEmail" width="92" height="22">
                                      </td>         
                                    </tr>
                                  </table>
                                </td>
                              </tr>
                            </table>
                          </form>
                        </td>
                      </tr> 
        
		              <% ElseIf sOutPut = "SentEmail" Then %>
			          <table border="0"  width="75%">
                        <tr>
                          <td width="100%" align="middle" class="tdContent">
					An e-mail with the customer password has been sent to this e-mail:
					        <br><%=Request.Form("Email")%>
					        <br>
				          </td>
                        </tr>
                      </table>
                      <br>
            
             	
            
  		              <% ElseIf sOutPut = "FailedEmail" Then %>
			          <table border="0"  width="75%">
                        <tr>
                          <td width="100%" align="middle" class="tdContent">
					No record exists for a customer with the following e-mail: 
					        <br><%=Request.Form("Email")%>
					        <p>Would you like to <a href="login.asp?Type=NewAccount">sign in</a> as a new customer?
				            </td>
                          </tr>
                        </table>
                        <br>
            
                    
 	          	
	                    <%	
		' End sOutPut If 
		End If
    ' End Send Email If  
    End If
    '-----------------------------------------------------------
    ' End OutPut Block
    '-----------------------------------------------------------
    closeObj(cnn)
    %>
<!--Footer begin-->
                        <!--#include file="footer.txt"-->
                      </table>
                    </td>
                  </tr>
              </table>
            </body>
          </html>

<!--Footer End-->


