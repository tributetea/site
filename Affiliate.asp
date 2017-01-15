<%@ Language=VBScript %>
<%	
	option explicit 
	Response.Buffer = True
%>
<!--#include file="sfLib/incDesign.asp"-->
<!--#include file="SFLib/db.conn.open.asp"-->
<!--#include file="SFLib/adovbs.inc"-->
<!--#include file="SFLib/incGeneral.asp"--> 
<%
'@BEGINVERSIONINFO


'@APPVERSION: 50.4011.0.2

'@FILENAME: affiliate.asp
 

'

'@DESCRIPTION: Affiliate Partners Page

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000, 2001 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO

Dim sAction,rsAff,iBookMark,iAffId,sFilter,sName,rsRegAff,sAffiliateId,rsSales,sSQL,sTotalNet,sTotalSTax,sTotalCTax,sTotalShipping,sHandling,sGrandTotal,arrSales,i

If Request.Form("affSubmit.x") <> "" or Request.QueryString("AffID") <> "" Then
	If Request.Form("affSubmit.x") <> "" Then sFilter = "affName = '" & trim(Request.Form("txtAffId")) & "' AND affPassword = '" & trim(Request.Form("passAff")) & "'"
	
	IF  Session("AffLogin") ="ByPass" then
		If Request.QueryString("AffID") <> "" Then sFilter = "affID = " &  Request.QueryString("AffID") 
		Session("AffLogin") =""
	else
		If Request.QueryString("AffID") <> "" Then sFilter = "affID = " &  Request.QueryString("AffID") & " AND affPassword = '" & trim(Request.Form("passAff")) & "'"	
		Session("AffLogin") =""
	End IF
	
	Set rsAff = Server.CreateObject("ADODB.Recordset")
	rsAff.Open "sfAffiliates", cnn, adOpenForwardOnly, adLockReadOnly, adCmdTable
	rsAff.Filter = sFilter

	If Not (rsAff.EOF And rsAff.BOF) Then
		sAffiliateId = trim(rsAff.Fields("affName"))
		sSQL = "SELECT orderAmount, orderSTax, orderCTax, orderShippingAmount, orderHandling, orderGrandTotal, orderTradingPartner FROM sfOrders WHERE orderTradingPartner = '" & sAffiliateId & "'"
		Set rsSales = Server.CreateObject("ADODB.RecordSet")
		rsSales.Open sSQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
		
		If Not (rsSales.EOF and rsSales.BOF) Then arrSales = rsSales.GetRows 
		
		sTotalNet = 0 
		sTotalSTax = 0
		sTotalCTax = 0
		sTotalShipping = 0
		sHandling = 0 
		sGrandTotal = 0

		rsSales.Close 
		Set rsSales= Nothing

		If isArray(arrSales) Then
			For i=0 to uBound(arrSales, 2)
				If arrSales(0, i) <> "" Then sTotalNet = sTotalNet + cDbl(arrSales(0, i))
				If arrSales(1, i) <> "" Then sTotalSTax = sTotalSTax + cDbl(arrSales(1, i))
				If arrSales(2, i) <> "" Then sTotalCTax = sTotalCTax + cDbl(arrSales(2, i))
				If arrSales(3, i) <> "" Then sTotalShipping = sTotalShipping + cDbl(arrSales(3, i))
				If arrSales(4, i) <> "" Then sHandling = sHandling + cDbl(arrSales(4, i))
				If arrSales(5, i) <> "" Then sGrandTotal = sGrandTotal + cDbl(arrSales(5, i))
			Next
		End If 
			
		sAction = "Registered"
	Else
		sAction = "Not Found"
	End If
	closeObj(rsAff)
ElseIf Request.Form("newAffSubmit.x") <> "" Then
	Set rsAff = Server.CreateObject("ADODB.Recordset")
	rsAff.CursorLocation = adUseClient
	rsAff.Open "sfAffiliates ORDER BY affID", cnn, adOpenStatic, adLockOptimistic, adCmdTable
	
	sName = trim(Request.Form("txtName"))
	Do While Not rsAff.EOF
		If rsAff.Fields("affName") = sName Then sAction="Exists"
		rsAff.MoveNext
	Loop
	
	If sAction <> "Exists" Then
		rsAff.AddNew 
		rsAff.Fields("affName") = sName
		rsAff.Fields("affCompany") = trim(Request.Form("txtCompany"))
		rsAff.Fields("affAddress1") = trim(Request.Form("txtAdd1"))
		rsAff.Fields("affAddress2") = trim(Request.Form("txtAdd2"))
		rsAff.Fields("affCity") = trim(Request.Form("txtCity"))
		rsAff.Fields("affState") = trim(Request.Form("txtState"))
		rsAff.Fields("affZip") = trim(Request.Form("txtZip"))
		rsAff.Fields("affCountry") = trim(Request.Form("txtCountry"))
		rsAff.Fields("affPhone") = trim(Request.Form("txtPhone"))
		rsAff.Fields("affFAX") = trim(Request.Form("txtFax"))
		rsAff.Fields("affEmail") = trim(Request.Form("txtEmail"))
		rsAff.Fields("affHttpAddr") = trim(Request.Form("txtWeb"))
		rsAff.Fields("affPassword") = trim(Request.Form("Password"))
		rsAff.Update 
	
		iBookMark = rsAff.AbsolutePosition 
		rsAff.Requery 
		rsAff.AbsolutePosition = iBookMark
		iAffId = rsAff.Fields("affID")
		Session("AffLogin") = "ByPass"
	End If
		closeObj(rsAff)
	If sAction <> "Exists" Then Response.Redirect "Affiliate.asp?AffID=" & iAffId
ElseIf Request.QueryString("NewAccount") = "true" Then
	sAction = "New Account"
End If
%>

<html>

<head>
<script language="javascript" src="SFLib/sfCheckErrors.js"></script>
<meta http-equiv="Pragma" content="no-cache">

<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title><%= C_STORENAME %>-SF Affiliate Partners Page</title>


<!--Header Begin -->
<link rel="stylesheet" href="sfCSS.css" type="text/css">
</head>
<body bgproperties="fixed"  link="<%= C_LINK %>" vlink="<%= C_VLINK %>" alink="<%= C_ALINK %>">
<table border="0" cellpadding="1" cellspacing="0"  class="tdbackgrnd" width="<%= C_WIDTH %>" align="center">
  <tr>
    <td>
	  <table width="100%" border="0" cellspacing="1" cellpadding="3">
        <tr>
          <td align="middle"  class="tdTopBanner"><%If C_BNRBKGRND = "" Then%><%= C_STORENAME %><%Else%><img src="<%= C_BNRBKGRND %>" border="0"><%End If%></td>
        </tr>
<!--Header End -->
        <tr>
	      <td align="middle"   class="tdMiddleTopBanner">Affiliate Registration</td>        
        </tr>
        <tr>
          <td class="tdBottomTopBanner"> Enter your affiliate ID and password to view your referred sales.  If you would like to establish a new affiliate account, click <B>New Account</B> to begin.</td>
        </tr>
        <tr>
          <td align="center" class="tdContent2"  width="100%" nowrap>
       	    <% If sAction = "Registered" Then %>
       			<table border=0>
       			  <tr>
       			    <td align="left" colspan=2 width="100%">Please use this at the end of your link to this shopping
                        order.  Be sure the link is to an ASP page in this site.<br><br>
       			      <center>REFERER=<%= sAffiliateId %></center><br>
       			Example http://www.thisdomain.com/search.asp?REFERER=<%= sAffiliateId %><br><br></td>
       			  </tr>
       			  <tr>
       			    <td width="100%" colspan=2><hr width="100%"></td>
       			  </tr>
       			  <tr>
       			    <td align="Center" colspan=2 width="100%" class="Content_Small"><b>Sales Invoice</b></td>
       			  </tr>
       			  <tr>
       			    <td align="right" width="80%">Net Sales:</td><td width="20%" align="right"><%= FormatCurrency(sTotalNet) %></td>
       			    </tr>
       			    <tr>
       			      <td align="right" width="80%">State/Province Tax:</td><td width="20%" align="right"><%= FormatCurrency(sTotalSTax) %></td>
       			      </tr>
       			      <tr>
       			        <td align="right" width="80%">Country Tax:</td><td width="20%" align="right"><%= FormatCurrency(sTotalCTax) %></td>
       			        </tr>
       			        <tr>
       			          <td align="right" width="80%">Handling:</td><td width="20%" align="right"><%= FormatCurrency(sHandling) %></td>
       			          </tr>
       			          <tr>
       			            <td align="right" width="80%">Shipping:</td><td width="20%" align="right"><%= FormatCurrency(sTotalShipping) %></td>
       			            </tr>
       			            <tr>
       			              <td align="right" width="80%">Grand Total:</td><td width="20%" align="right"><%= FormatCurrency(sGrandTotal) %></td>
       			              </tr>
       			            </table>
       	                    <% ElseIf sAction = "New Account" or sAction = "Exists" Then %>
       	                    <form method="post" action="Affiliate.asp" onSubmit="this.Password.password=true;this.txtAdd2.optional=true;this.txtPhone.phoneNumber=true;this.txtEmail.eMail=true;this.txtFax.optional=true;return sfCheck(this)" name=frmNewAff>
       	                      <table width="100%">
       		                    <% If sAction = "Exists" Then %>
       		                    <tr>
       		                      <td colspan=2 align=center><font class="Error">We are sorry but your <b>"Affiliate ID"</b> already exists, please enter a different ID value.</font></td>
       		                    </tr>
       		                    <% End If %>
       	                        <tr>
	                              <td align="left" width="20%" nowrap>Affiliate ID:</td>
	                              <td align="left" width="80%" nowrap><input type="text" style="<%= C_FORMDESIGN %>" size="15" name="txtName" value="<%= Request.Form("txtName")%>"></td>
	                            </tr>
	                            <tr>
	                              <td align="left" width="20%" nowrap>Company:</td>
	                              <td align="left" width="80%" nowrap><input type="text" style="<%= C_FORMDESIGN %>" size="20" name="txtCompany" value="<%= Request.Form("txtCompany")%>"></td>
	                            </tr>
	                            <tr>
	                              <td align="left" width="20%" nowrap>Address Line 1:</td>
	                              <td align="left" width="80%" nowrap><input type="text" style="<%= C_FORMDESIGN %>" size="25" name="txtAdd1" value="<%= Request.Form("txtAdd1")%>"></td>
	                            </tr>
	                            <tr>
	                              <td align="left" width="20%" nowrap>Address Line 2:</td>
	                              <td align="left" width="80%" nowrap><input type="text" style="<%= C_FORMDESIGN %>" size="25" name="txtAdd2" value="<%= Request.Form("txtAdd2")%>"></td>
	                            </tr>
	                            <tr>
	                              <td align="left" width="20%" nowrap>City:</td>
	                              <td align="left" width="80%" nowrap><input type="text" style="<%= C_FORMDESIGN %>" size="15" name="txtCity" value="<%= Request.Form("txtCity")%>"></td>
	                            </tr>
	                            <tr>
	                              <td align="left" width="20%" nowrap>State:</td>
	                              <td align="left" width="80%" nowrap><input type="text" style="<%= C_FORMDESIGN %>" size="15" name="txtState" value="<%= Request.Form("txtState")%>"></td>
	                            </tr>
	                            <tr>
	                              <td align="left" width="20%" nowrap>Zip/Postal Code:</td>
	                              <td align="left" width="80%" nowrap><input type="text" style="<%= C_FORMDESIGN %>" size="10" name="txtZip" value="<%= Request.Form("txtZip")%>"></td>
	                            </tr>
	                            <tr>
	                              <td align="left" width="20%" nowrap>Country:</td>
	                              <td align="left" width="80%" nowrap><input type="text" style="<%= C_FORMDESIGN %>" size="20" name="txtCountry" value="<%= Request.Form("txtCountry")%>"></td>
	                            </tr>
	                            <tr>
	                              <td align="left" width="20%" nowrap>Phone Number:</td>
	                              <td align="left" width="80%" nowrap><input type="text" style="<%= C_FORMDESIGN %>" size="15" name="txtPhone" value="<%= Request.Form("txtPhone")%>"></td>
	                            </tr>
	                            <tr>
	                              <td align="left" width="20%" nowrap>Fax Number:</td>
	                              <td align="left" width="80%" nowrap><input type="text" style="<%= C_FORMDESIGN %>" size="15" name="txtFax" value="<%= Request.Form("txtFax")%>"></td>
	                            </tr>
	                            <tr>
	                              <td align="left" width="20%" nowrap>Email:</td>
	                              <td align="left" width="80%" nowrap><input type="text" style="<%= C_FORMDESIGN %>" size="20" name="txtEmail" value="<%= Request.Form("txtEmail")%>"></td>
	                            </tr>
	                            <tr>
	                              <td align="left" width="20%" nowrap>Web Site:</td>
	                              <td align="left" width="80%" nowrap><input type="text" style="<%= C_FORMDESIGN %>" size="25" name="txtWeb" value="<%= Request.Form("txtWeb")%>"></td>
	                            </tr>
	                            <tr>
	                              <td align="left" width="20%" nowrap>Password:</td>
	                              <td align="left" width="80%" nowrap><input type="password" style="<%= C_FORMDESIGN %>" size="15" name="Password"></td>
	                            </tr>
	                            <tr>
	                              <td align="left" width="20%" nowrap>Confirm Password:</td>
	                              <td align="left" width="80%" nowrap><input type="password" style="<%= C_FORMDESIGN %>" size="15" name="Password2"></td>
	                            </tr>
	                            <tr>
	                              <td align="center" width="100%" colspan=2><input type="image" src="<%= C_BTN18 %>" border="0" name="newAffSubmit"></td>
	                            </tr>
       	                      </table>
       	                    </form>
       	                    <% Else %>
       	                    <form method="post" action="Affiliate.asp" onSubmit="return sfCheck(this)" name=frmAff>
                              <table border="0" width="75%">
                                <tr>
                                  <td width="100%" align="middle" class="tdBottomTopBanner2">
			                        <%If sAction = "Not Found" Then%>
				                    Your Affiliate ID and Password did not match a registered user, please create a new account or try again.
			                        <%Else%>
				                    Please Type In Your Affiliate ID and Password or Click on New Account.
			                        <%End If%>
                                    </td>
                                  </tr>
                                  <tr>
                                    <td width="100%" align="center" class="tdContent">
                                      <table border="0" width="100%">
                                        <tr>
                                          <td width="50%" align="right"><b>Affiliate ID:</b></td>
                                          <td width="50%">
                                          <input type="text" name="txtAffId" title="Affiliate ID">
                                          </tr>
                                          <tr>
                                            <td width="50%" align="right"><b>Password:</b></td>
                                            <td width="50%">
                                            <input type="password" name="passAff" title="Password">
                                            </tr>
                                            <tr>
                                              <td width="100%" align="middle" colspan="2">
				                                <input type="image" src="<%= C_BTN18 %>" border="0" name="affSubmit"><a href="Affiliate.asp?NewAccount=true"><img src="<%= C_BTN19 %>" border="0" name="affNewAccount"></a>
                                              </td>         
                                            </tr>
                                          </table>
                                        </td>
                                      </tr>
                                    </table>
                                  </form> 
                                  <% 
        End If 
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



