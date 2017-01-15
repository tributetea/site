<%@ Language=VBScript %>
<%	option explicit 
	Response.Buffer = True
%>
<!--#include file="SFLib/incDesign.asp"-->
<!--#include file="SFLib/db.conn.open.asp"-->
<!--#include file="SFLib/adovbs.inc"-->
<!--#include file="SFLib/mail.asp"-->
<!--#include file="SFLib/incGeneral.asp"-->
<!--#include file="SFLIB/incAE.asp"-->
<%		
	'@BEGINVERSIONINFO


	'@APPVERSION: 50.4011.0.2

	'@FILENAME: emailfriend.asp
 



	'@DESCRIPTION: Send Email of product to a friend

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
<HTML>

<HEAD>
<SCRIPT language=javascript src="SFLib/sfCheckErrors.js"></SCRIPT>

<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title><%= C_STORENAME %>-SF Email a Friend Popup Window</title>




<%
Dim sClosing
sClosing = ""
Dim sMessage


	
If CheckLogin = 0 then 'AE
	js "window.close"
	Response.End 
End if	

sMessage = MakeWishList

Function  MakeWishList
Dim sql,rst
dim sProdLine	    
dim i ,rsAdmin
Dim sPrimary, sSecondary, sHomePath
Dim sMessage	
Dim iCustID,  iSessionID	
dim sAttName

		    
	' Request Cookie for custID 
	iCustID		= Trim(Request.Cookies("sfCustomer")("custID"))
	iSessionID	= Trim(Request.Cookies("sfOrder")("SessionID"))
	    
	If iCustID = "" or iSessionID <> Session("SessionID") Then 
		MakeWishList = ""
		Exit Function
	End If			
	

	Set rsAdmin = Server.CreateObject("ADODB.RecordSet")
	rsAdmin.Open "sfAdmin", cnn, adOpenForwardOnly , adLockReadOnly, adCmdTable
	sPrimary = rsAdmin.Fields("adminPrimaryEmail")
	sHomePath = rsAdmin.Fields("adminDomainName")
	closeObj(rsAdmin)

	If Mid(sHomePath, len(sHomePath)-1, 1) <> "/" Then
		sHomePath = sHomePath & "/"
	End If
	
	sMessage = ""
	sMessage =	vbCRLF & vbCrLF  & "---------------------------------------------------"
	sMessage = sMessage & vbCrLF & "    My Wish List at " & C_STORENAME 
	sMessage = sMessage & vbCrLF & "---------------------------------------------------" & vbCrLF & VBCRLF
	
	
	'read from sfSvdOrderDetails
	sql = "SELECT * FROM sfSavedOrderDetails WHERE odrdtsvdCustID=" & iCustID
	
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText
	If rst.recordcount > 0 then
		rst.MoveFirst
		For i = 1 to rst.RecordCount 
			'sProdLine = vbCRLF & vbCRLF & "ProductId: " & rst("odrdtsvdProductId") 
			sAttName=GetAttName(rst("odrdtsvdID"),"svd")
			
			sProdLine =vbCRLF  & "Product: "
			sProdLine = sProdLine &  GetProductName(rst("odrdtsvdProductID"))
			If sAttName <> "" then 
				sProdLine = sProdLine & vbCRLF & "Details: " & sAttName
			End if
			sProdLine = sProdLine & vbCrLF & "Quantity:" & rst("odrdtsvdQuantity")
			sProdLine = sProdLine & vbCrLF & sHomePath & "detail.asp?product_id=" & server.urlEncode(rst("odrdtsvdProductID"))
			sProdLine = sProdLine & vbCrLF
			sMessage = sMessage & sProdLine
			
			If not rst.EOF then rst.MoveNext 
		Next 
		
	End If
	closeobj(rst)
	
	MakeWishList = sMessage
	

End Function

		

If Request.Form("btnEmailSubmit.x") <> "" Then
	Dim sInformation
		
	sInformation = Request.Form("txtFriend") & _
	 "|" & Request.Form("txtEmail") & _
	 "|" & Request.Form("txtMessage") & _
	 "|" & Request.Form("prodID") & _ 
	 "|" & Request.Form("txtSubject") & _
	 "|" & Request.Form("txtFriend2")
	
	
	'Response.Write sInformation
	Call createMail ("EmailWishList",sInformation)
	sClosing = "onload=""javascript:window.close()"""
	
End If
%>

<!--Header Begin Modified (Body tag modified)-->
<link rel="stylesheet" href="sfCSS.css" type="text/css">
</head>
<body <% If Request.Form("btnEmailSubmit.x") <> "" Then response.write sClosing %>  link="<%= C_LINK %>" vlink="<%= C_VLINK %>" alink="<%= C_ALINK %>">

<form method="post" name="frmEmailfriend" onsubmit="javascript:this.txtFriend2.optional=true;return sfCheck(this)">  
<table border="0" cellpadding="1" cellspacing="0"  class="tdbackgrnd" width="100%" align="center">
  <tr>
    <td>
      <table width="100%" border="0" cellspacing="1" cellpadding="3">
        <tr>
          <td align="middle"  class="tdTopBanner"><%= C_STORENAME %></td>
        </tr>
<!--Header End -->
        <tr>
          <td class="tdContent2">
<table border=0 width="100%">
    <tr>
      <td colspan=2 nowrap><b>Use this form to e-mail your Wish List to others.</b>
        &nbsp;<br>
        <br>
      </td>
    </tr>
    <tr>
      <td nowrap width="20%">To
        E-Mail Address:</td>
      <td nowrap><input type=text size=40 name="txtFriend" title="Your Firend's Email Address"></td>
    </tr>
    <tr>
      <td nowrap width="20%">CC
        E-Mail Address:</td>
      <td nowrap><input type=text size=40 name="txtFriend2"></td>
    </tr>
    <tr>
      <td nowrap width="20%">From
        E-Mail Address:</td>
      <td nowrap><input type=text size=40 name="txtEmail" title="Your Email Address"></td>
    </tr>
    <tr>
      <td nowrap width="20%">Subject:</td>
      <td nowrap><input type=text size=50 name="txtSubject" title="Email Subject"></td>
    </tr>
    <tr>
      <td colspan=2 nowrap>Message:</td>
    </tr>
    <tr>
      <td colspan=2 align=center nowrap><textarea cols=70 rows=7 name="txtMessage" title="Message"> <%=sMessage%></textarea></td>
    </tr>
    <tr>
      <td colspan=2 align=center nowrap><input type=image border="0" src="<%= C_BTN18 %>" name="btnEmailSubmit"></td>
    </tr>
  <input type=hidden name=prodID value=<%= Request.QueryString("ProdID") %>>
</table>
<!--Footer Begin -->
            <tr>
              <td class="tdFooter"><p align="center"><b><a href="javascript:window.close()">Close</a></b></td>
                </tr>
                </table>
              </td>
            </tr>
            </table>
</form>
</BODY>
</html>
<!--Footer End -->


<%
closeObj(cnn)
%>



