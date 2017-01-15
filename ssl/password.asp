<%@ Language=VBScript %>
<%	option explicit 
	Response.Buffer = True
	%>
	
<!--#include file="SFLib/db.conn.open.asp"-->
<!--#include file ="SFLib/incProcOrder.asp"-->
<!--#include file="error_trap.asp"-->
<!--#include file="SFLib/incGeneral.asp"-->
<!--#include file="SFLib/incDesign.asp"-->
<!--#include file="SFLib/incLogin.asp"-->
<!--#include file="SFLib/ADOVBS.inc"-->
	<%
	Const vDebug = 0
    
	'@BEGINVERSIONINFO

	'@APPVERSION: 50.4011.0.2

    '@FILENAME: password.asp
	 



	'@DESCRIPTION: Handles password Information

	'@STARTCOPYRIGHT
	'The contents of this file is protected under the United States
	'copyright laws and is confidential and proprietary to
	'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
	'expressed written permission of LaGarde, Incorporated is expressly prohibited.
	'
	'(c) Copyright 2000,2001 by LaGarde, Incorporated.  All rights reserved.
	'@ENDCOPYRIGHT

	'@ENDVERSIONINFO

Dim sStatus, sOutput, sEmail, sPasswd, sOldPassword, sNewPassword, iAuthenticate, bEmailAddress
Dim iFormLoop
sStatus = Trim(Request.QueryString("status"))
	
If sStatus = "fpwd" then
	sOutput = "FPWD"
ElseIf sStatus = "change" then
	sOutput="ChangePwd"
End If	
If Trim(Request.Form("SendEmail.x")) <> "" or (Request.Form.Count=1 and Trim(Request.Form("SubmitNewPWD.x"))="") Then
'above if is there to see if this is someone requesting a new password,
'if they hit return while in the e-mail box, the first condition is false,
'so I had to add the other conditions
	sEmail = Trim(Request.Form("Email"))
	sPasswd = Trim(Request.Form("Password"))

	bEmailAddress = SendPassword(sEmail)
			
	If bEmailAddress = 1 Then
		sOutPut = "EmailSent"
	Else
		sOutPut = "NoEmailMatch"
	End If	

	
ElseIf Trim(Request.Form("SubmitNewPWD.x")) <> "" Then
	sEmail = Trim(Request.Form("Email"))
	sOldPassword = Trim(Request.Form("OldPassword"))
	sNewPassword = Trim(Request.Form("NewPassword"))
	iAuthenticate = customerAuth(sEmail,sOldPassword,"loose")
	
	If  iAuthenticate > 0 Then			
		Call UpdatePassword(iAuthenticate,sNewPassword)
		Response.Cookies("sfOrder")("SessionID") = Session("SessionID")
		Response.Cookies("sfOrder").Expires = Date() + 1
		Response.Cookies("sfCustomer")("custID") = iAuthenticate
		Response.Cookies("sfCustomer").Expires = Date() + 1		
		sOutput = "ChangedPwd"
	Else
		sOutput = "FailedPwdUpdate"		
	End If	
End If


%>
<html>
<head>

<SCRIPT language = "javascript" src="../SFLib/sfCheckErrors.js"></SCRIPT>
<SCRIPT language = "javascript">
function checkpwd(WholeForm) {
	if(WholeForm.NewPassword.value != WholeForm.NewPassword2.value){
		alert('Your passwords do not match. Please retype them again.');
		WholeForm.NewPassword.value = "";
		WholeForm.NewPassword2.value= "";		
		WholeForm.NewPassword.focus();	
		return false;		
	}
	else {
		return true;
	}
}
</script>
<meta http-equiv="Pragma" content="no-cache">

<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title><%= C_STORENAME %>-SF Password Page</title>


<!--Header Begin -->
<link rel="stylesheet" href="sfCSS.css" type="text/css">
</head>

<body  link="<%= C_LINK %>" vlink="<%= C_VLINK %>" alink="<%= C_ALINK %>">

<table border="0" cellpadding="1" cellspacing="0" class="tdbackgrnd" width="<%= C_WIDTH %>" align="center">
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
<!--Header End -->
        
        <%
    '-----------------------------------------------------------
    ' Begin OutPut Block
    '-----------------------------------------------------------
    %>  
        <tr>
          <td  align="middle" class="tdMiddleTopBanner">
   		    <% If sOutput = "ChangePwd" Then %>
		Change Login/Password
		    <% ElseIf sOutPut = "FailedPwdUpdate" Then%>
		Failed Password Update
		    <% ElseIf sOutput = "EmailSent" Then %>
		E-Mail With Password Sent
		    <% ElseIf sOutput = "NoEmailMatch" Then %>
		No Matching E-Mail Found	
		    <% ElseIf sOutput = "FPWD" Then %>
		Forgotten Password Help		
		    <% End If %>
	        </td>
	        </tr>	
            <tr>
	          
          <td class="tdBottomTopBanner2" align="left"> 
            <% If sOutput = "ChangePwd" Then %>
            Please enter your e-mail and current password followed by your new 
            password. 
            <% ElseIf sOutPut = "FailedPwdUpdate" Then%>
            The password you entered for this e-mail is not valid. 
            <% ElseIf sOutput = "EmailSent" Then %>
            A message has been sent to the following address. 
            <% ElseIf sOutput = "NoEmailMatch" Then %>
            No e-mail address was found to match. 
            <% ElseIf sOutput = "FPWD" Then %>
            Please enter your e-mail address and your password will be forwarded 
            to you immediately. 
            <% End If %>
          </td>
	            </tr>	
	
	            <% If sOutput="ChangePwd" Or sOutput="FailedPwdUpdate" Then %>
		        <tr>
		          <td class="tdContent" align="middle"><br>
			        <form name="changepwd" method="post" onSubmit="return checkpwd(changepwd)">
				      <table border="0" width="75%">
				
					    <% If sOutput="FailedPwdUpdate" Then %>
					    <tr>
					      <td width="100%" align="middle" class="tdBottomTopBanner2">
					Your current password and e-mail did not match with our records. Please try again.
				          </td>
					    </tr>
					    <% End If %>

					    <tr>
					      <td width="100%" align="middle" class="tdContent">
					        <table border="0" width="100%">
						      <tr>
							    <td width="50%" align="right"><b>E-Mail Address:</b></td>
							    <td width="50%"><input type="text" name="Email" title="E-Mail Address" style="<%= C_FORMDESIGN %>"> 
						        </tr>
						        <tr>
							      <td width="50%" align="right"><b>Current Password:</b></td>
							      <td width="50%"><input type="password" name="OldPassword" title="Current Password" style="<%= C_FORMDESIGN %>">
						          </tr>
						          <tr>
							        <td width="50%" align="right"><b>New Password:</b></td>
							        <td width="50%"><input type="password" name="NewPassword" title="New Password" style="<%= C_FORMDESIGN %>">
						            </tr>
						            <tr>
							          <td width="50%" align="right"><b>New Password Confirmation:</b></td>
							          <td width="50%"><input type="password" name="NewPassword2" title="New Password Confirmation" style="<%= C_FORMDESIGN %>">
						              </tr>
						              <tr>
							            <td width="100%" align="middle" colspan="2">
							              
                          <input type="image" src="buttons/submit.gif" border="0" name="SubmitNewPWD" width="92" height="22">
							            </td>         
						              </tr>
					                </table>
					              </td>
					            </tr>
				              </table>
			                </form>
                          </td>
                        </tr> 
	                    <% ElseIf sOutput="ChangedPwd" Then %>
		                <tr align="center">
		                  <td class="tdContent" align="center"><br>			
				            <table border="0" width="100%">
				              <tr>
					              <td width="100%" align="center" class="tdContent"><b>
						Your login has been updated. Please <a href="process_order.asp?Persist=1">return to checkout</a>.</b>
					              </td>
				                  </tr>
				                </table>
				                <br>
                              </td>
                            </tr> 
	
	                        <% ElseIf sOutput="FPWD" Then %>
		                    <tr>
		                      <td class="tdContent" align="center"><br>
			                    <form method="post" name=thisForm>
				                  <table border="0" width="100%">
				                    <tr>
					                  <td width="100%" align="center" class="tdBottomTopBanner2">
					Please Type In Your E-Mail Address	
				                      </td>
				                    </tr>
				                    <tr>
					                  <td width="100%" align="center" class="tdContent">
					                    <table border="0" width="100%">
						                  <tr>
							                <td width="50%" align="right"><b>E-Mail Address:</b></td>
							                <td width="50%"><input type="text" name="Email" title="Email Address" style="<%= C_FORMDESIGN %>">
						                    </tr>
						                    <tr>
							                  <td width="100%" align="center" colspan="2">
							                    
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
        
	                          <% ElseIf sOutput = "EmailSent" Then %>
		                      <tr align="center">			      			 
		                        <td>
			                      <br>
			                      <table border="0" width="100%" align="center">
                                    <tr align="center">
                                      
                <td width="100%" align="center" class="tdContent"> A message containing 
                  your password has been forwarded to the following e-mail address: 
                  <br>
                  <%= sEmail %>
					                    <br>
				                      </td>
                                    </tr>
                                  </table>
                                  <br>
                                </td>
                              </tr>    
                              <% ElseIf sOutPut = "NoEmailMatch" Then %>
                              <tr>
			                    <table border="0" width="100%">
                                  <tr>
                                    <td width="100%" align="middle" class="tdContent">
					No record exists for a customer with the following e-mail: 
					                  <br><%= sEmail %>
					                  <br>Please try again.
				                    </td>
                                  </tr>
                                </table>
                                <br>
            
                              </tr>   
                              <% 
	   End If 
	   closeObj(cnn)   
	%>   
                              <!--Footer begin-->
                <!--#include file="foot.txt"-->
              </table>
            </td>
          </tr>
        </table>
       </body>

    </html>
<!--Footer End-->    
