<%@ Language=VBScript %>
<%	option explicit 
	Response.Buffer = True
	Response.clear

	%>
<!--#include file="SFLib/db.conn.open.asp"-->
<!--#include file="SFLib/adovbs.inc"-->
<!--#include file="SFLib/incGeneral.asp"-->

<!--#include file="sfLib/incDesign.asp"-->
<%sub writehead()%>
	<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<script language="javascript" src="SFLib/incAE.js"></script>
<link rel="stylesheet" href="sfCSS.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>TRIBUTE TEA - Order Confirmation</title>


</head>
<body link="<%= C_LINK %>" vlink="<%= C_VLINK %>" alink="<%= C_ALINK %>" onLoad="javascript:linkCorrect()">
<%end sub%>
<%
'@BEGINVERSIONINFO

'@APPVERSION: 50.4011.0.2

'@FILENAME: 
 


'@DESCRIPTION: 

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000, 2001 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO
Dim vDebug
vDebug = 0
Dim sHas, sProdName, iQuantity, sProdUnit, sProdMessage, sResponseMessage
Dim msgtype
Dim avlQTY,gwQTY,ordQTY
Dim gwQTY_bo,ordQTY_bo
Dim iTmpOrderDetailID,sProdId,sAttDetailID
Dim inv ,bo
Dim sTotalMessage	


	sProdName = Request.QueryString("sProdName")
	iQuantity = Request.QueryString("iQuantity")
	sProdUnit = Request.QueryString("sProdUnit")
	sProdMessage = Request.QueryString("sProdMessage")
	sResponseMessage = Request.QueryString("sResponseMessage")
	iTmpOrderDetailID = Request.QueryString("iTmpOrderDetailID")
	sProdId = Request.QueryString("sProdID")
	
	 'sProdName =GetProductName(sProdid)
	'Initilization
	inv =  CheckInventoryTracked(sProdID)  'inventory tracked for  this product?
	bo = CheckBackOrder(sProdId) ' backorder allowed ?
	ordQTY = iQuantity
	
	'JS "window.opener.location='" & (C_HomePath & "order.asp")	& "'"
	%>
	

	<%
	If inv <> 1  Then 'OR Request.QueryString("gAddCaption") = "Back Order" then
	
		'inventory not tracked so show normal thanks page
		sTotalMessage = iQuantity
		ShowThanks()
		
	Else
		'who let the dogs out!		
call writeHead
		DoInventory()
		
	End If

	


Sub DoInventory()
	
	sAttDetailID = GetAttDetailID(iTmpOrderDetailID,"tmp")
	avlqty = GetAvailableQTY(sProdId,sAttDetailID) 
	gwqty =GetTMPGiftWrapQTY(iTmpOrderDetailID)
	ordQTY =GetTMPQTY(iTmpOrderDetailID) 'We need to get the total qty out of tmporder
	
	if gwqty = "X" then gwqty = 0
	
	'Save these values for backorder processing
	ordQTY_bo = ordQTY
	gwqty_bo = gwqty
		
		
	If avlqty >= ordQty then
		sTotalMessage = iQuantity 
		ShowThanks()
		exit sub
		
	ElseIf avlqty <= 0 and bo <> 1 then '1a
		msgtype = "1a"	
		'delete tmporder
		AdjustQty
		'DeleteTmpOrderDetailsAE(iTmpOrderDetailID) 
		
	Elseif avlqty < ordQTY and bo <> 1 then '1b
		msgtype = "1b"
		'replace ordqty with avlqty
		AdjustQty
		
	Elseif avlqty <=0 and bo = 1 then '2a
		msgtype ="2a"
		'if backorder is clicked don't delete and do nothing, 
		' OTHERWISE delete tmporder 	
		AdjustQty
		'DeleteTmpOrderDetailsAE(iTmpOrderDetailID)
		
				
	Elseif avlqty < ordQTY and bo = 1 then '2b
		msgtype ="2b"
		'replace ordqty with avlqty
		AdjustQty	
	End if
	
	'ElseIf IQuantity <> ordQTY then '3
	'	msgtype="3"
		
	
	
	ShowInventoryPage()
	If vDebug = 1 then ShowDebugValues()
	
End Sub	




Sub AdjustQTY
Dim sql,rst
	'rst.CursorLocation = adUseClient
	
	'set giftwrap quantity 
	'If gwqty <> "X" then 
	If gwqty > 0 and gwqty > avlqty then gwqty = avlqty	
	'	gwqty = 0
	'End If
	
	'set total quantity for this product order
	ordQTY = avlqty
	
	sql = "Select * FROM sfTmporderDetails as A INNER JOIN  sftmpOrderDetailsAE as B on a.odrdttmpID = b.odrdttmpaeID  where a.odrdttmpID=" & iTmpOrderDetailID
	'sql = "Select * FROM sfTmporderDetails as a, sftmpOrderDetailsAE as b where a.odrdttmpIDAE = b.odrdttmpID  AND odrdttmpID=" & iTmpOrderDetailID
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn,  adopenKeySet, adLockOptimistic, adCmdText
	If rst.RecordCount <= 0 then exit sub 'error this should never happen
	rst("odrdttmpGiftWrapQty") = gwqty
	rst.update
	rst("odrdttmpQuantity")= ordQTY
	rst.update
	CloseObj (rst)
	
End Sub


Sub ShowInventoryMessage
	Select Case msgtype
		
	case "1a" 'no backorder , out of stock
			%>	<BR>
The product is out of stock. Please select another product or contact the store 
for more information.
<%
		
		case "1b" 'no backorder , partially out of stock	
			%>
<BR>
Sorry, we don't have your requested quantity in stock. A total of <%=ordQTY%> 
items have been added to your order.
<%	
		
		case "2a" 'with backorder, out of stock 
			%>
<BR>
The product is out of stock. If you like us to send the product as soon as it 
arrives please click the Back Order button below.
<%				
		
		case "2b" 'with backorder, partially out of stock
			%>
<BR>
 Sorry, we don't have your requested quantity in stock. A total of <%=ordQTY%> items have been added to your order. If you would like us to send you the remaining quantity as soon as available, click the Back Order button below.<%
		
		case "3"
			%>	<BR>
Our available stock has been depleted while you were shopping. Sorry, we don't 
have your requested quatity in stock anymore. <%=ordQTY%> items have been added 
to your order. If you would like us to send you the remaining quantity as soon 
as available, click the Back Order button below.
<%	
	End select
End Sub



Sub ShowInventoryOptions
	Select Case msgtype
		Case "1a" 'no backorder , out of stock%>
<input type="submit" value="     Close     " name="btnAction" >
			<%	
			
		Case "1b" 'no backorder , partially out of stock	%>	
			<input type="submit" value="     Close     " name="btnAction" >
		<%	
			
		Case "2a" 'backorder, out of stock%>	
			<input type="submit" value="No Thanks" name="btnAction" >
			<input type="submit" value="Back Order" name="btnAction" ><%
			JS "window.resizeTo (window.screen.availWidth * .50 ,window.screen.availHeight * .52)"
				
		case "2b" 'backorder, partially out of stock%>	
			<input type="submit" value="No Thanks" name="btnAction" >
			<input type="submit" value="Back Order" name="btnAction"><%	
			JS "window.resizeTo (window.screen.availWidth * .50 ,window.screen.availHeight * .52)"
	End select

End Sub



Sub ShowInventoryPage

%>

<!--Header Begin Modified (Head and body tags above as to not duplicate)-->

<table border="0" cellpadding="1" cellspacing="0"  class="tdbackgrnd" width="100%" align="center">
  <tr>
    <td>
      <table width="100%" border="0" cellspacing="1" cellpadding="3">
        <tr>
          <td align="middle" class="tdTopBanner">
            <div align="center"><img src="buttons/tt_blue.gif" width="275" height="36"></div>
          </td>
        </tr>
<!--Header End -->

	<form method="post">
	<tr>
	<td class="tdContent2">
    <hr noshade color="#000000" size="1" width="90%">
    <p align="center"><B><%=sProdName%> Stock Availability</b><br>
	<%ShowInventoryMessage()%>
	<BR>
	<BR>
	</td> </tr> 	  
	<tr>
	<td align="middle" class="tdTopBanner">
	
	<%ShowInventoryOptions()%>

	</td><tr>  
	<input type = hidden name="msgtype" value = "<%=msgtype%>" >
	<input type = hidden name="ordqty_bo" value =<%=ordqty_bo%> >
	<input type = hidden name="gwqty_bo" value =<%=gwqty_bo%> >
	<input type = hidden name="ordqty" value =<%=ordqty%> >
	<input type = hidden name="gwqty" value =<%=gwqty%> >
	</form>
<!--Footer Begin Modified (Close Body and close HTML tags below, as to not repeat-->
            <tr> 
              <td class="tdFooter"></td>
                </tr>
                </table>
              </td>
            </tr>
            </table>
          
<!--Footer End -->
<%
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'After Form Post code

	Select Case Request("btnAction") 

		Case "Back Order" '
			If Request.Form ("msgtype")= "2a" or Request.Form ("MsgType")= "2b" then
				SetValues()
				SetBackOrder()
				sTotalMessage = "A Total of " & ordqty_bo 
				ShowThanks()
				'JS "window.opener.location=" & "'order.asp'" 
				'JS "window.opener.location='" & (C_HomePath & "order.asp")	& "'"
				'JS "self.close()"
			Else
				Response.Write "Action Not Recognized"
				Response.End 
			End if 
		
		Case "No Thanks"
			'JS "window.opener.location='" & (C_HomePath & "order.asp")	& "'"
			'JS "self.close()"
		    js "window.close()"
	   	Case   "     Close     "
	   		'JS "window.opener.location='" & (C_HomePath & "order.asp")	& "'"
	   		'JS "window.opener.location=" & "'order.asp'" 
			'JS "self.close()"
		   js "window.close()"	
				   
	End select


Response.End  'this is to stop asp from processing below code without being called


Sub SetValues 'after form post
	msgtype = Request.Form("msgtype")
	ordqty_bo = Request.Form("ordqty_bo")
	gwqty_bo = Request.Form("gwqty_bo")
	If not isnumeric(gwqty_bo) then gwqty_bo = 0
	
End Sub


Sub SetBackOrder 'after form post

Dim sql,rst
	
	sql = "Select * FROM sftmporderdetails as A  INNER JOIN  sftmpOrderDetailsAE as B on a.odrdttmpID = b.odrdttmpaeID  where a.odrdttmpID=" & iTmpOrderDetailID
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn,  adopenKeySet, adLockOptimistic, adCmdText
	If rst.RecordCount <= 0 then 
		Response.Write "Critical Error!"
		Response.End 
		exit sub 'error this should never happen
	End if
	'rst("odrdttmpBackOrderQty") = 1'ordQTY_bo - rst("odrdttmpQuantity")
	rst("odrdttmpBackOrderQty") = ordQTY_bo - rst("odrdttmpQuantity") 'beta 2
	rst("odrdttmpGiftWrapQty") = gwqty_bo
	rst.update
	
	rst("odrdttmpQuantity")= ordQTY_bo
	rst.update
	CloseObj (rst)
	ordQTY = ordQTY_bo
		
End Sub
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Sub ShowDebugValues
		Response.write "<BR> ordQTY =" & ordQTY
		Response.write "<BR> avlqty =" & avlqty
		Response.write "<BR> sAttDetailID=" & sattdetailid
		Response.write "<BR> sProdID=" & sprodid
		Response.Write "<BR> inv=" & inv
		Response.write "<BR> bo=" & bo
		Response.Write "<BR> msgtype=" & msgtype
		Response.write "<BR> gwQTY=" & gwQTY
		Response.write "<BR> ordQTY_bo =" & ordQTY_bo
		Response.write "<BR> gwQTY_bo =" & gwQTY_Bo
End Sub

'Response.End  'this is to stop asp from processing below code without being called

Sub ShowThanks 
response.clear
call writeHead
	'JS "window.resizeTo (window.screen.availWidth * .50 ,window.screen.availHeight * .50)"
		
	If Request.Cookies("sfThanks")("PreviousAction") <> "" Then
		Response.Cookies("sfThanks").Expires = Now()
	End If
	iQuantity = ordQTY

%>


<!--Header Begin Modified (head and body tag above as to not duplicate)-->

<table border="0" cellpadding="1" cellspacing="0"  class="tdbackgrnd" width="100%" align="center">
  <tr>
    <td>
      <table width="100%" border="0" cellspacing="1" cellpadding="3">
        <tr>
          <td align="middle" class="tdTopBanner">
            <div align="center"><img src="buttons/tt_blue.gif" width="275" height="36"></div>
          </td>
        </tr>
<!--Header End -->
        <tr>
          <td class="tdContent2">
            <hr noshade color="#000000" size="1" width="90%">
            <p align="center"><font class="Content_Large"><b>Thank you!</b></font><br>
            <b><%=sTotalMessage%>&nbsp; <%= sProdUnit %> 
            <%= sProdName %> <%= sResponseMessage %></b>
            <p align="center">
            <div align="center"><b> <%= sProdMessage %><br>
              <br>
              <a href="<%= Session("DomainPath") %>order.asp"><img src="buttons/thanks_checkout.gif" width="129" height="22" border="0" alt="Proceed to Checkout"></a><a href="javascript:window.close()"><img src="buttons/return.gif" width="135" height="22" border="0" alt="Continue Shopping"></a> 
              <br>
              <br>
              <a href="<%= Session("DomainPath") %>order.asp"><img src="buttons/view.gif" width="36" height="33" border="0" alt="View Cart"></a> 
              </b></div>
            <b>
            <hr noshade color="#000000" size="1" width="90%" align="center">
            </b> </td>        
            </tr>
<!--Footer Begin Modified (Close Body and close HTML tags below, as to not repeat-->
            <tr> 
              <td class="tdFooter"></td>
                </tr>
                </table>
              </td>
            </tr>
            </table>
          
<!--Footer End -->

<%
'	closeObj(cnn)	
	If vDebug = 1 then ShowDebugValues()
End Sub
Function CheckInventoryTracked(strProductID)
Dim sql, rst
	
	sql = "Select * FROM sfInventoryInfo WHERE invenProdID= '" & strProductID & "'"
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText	
	
	If rst.recordcount <= 0 then 
		CheckInventoryTracked = 0
		exit function	
	End If
	
		
	If rst("invenbTracked") = 1 then
		CheckInventoryTracked = 1
	Else
		CheckInventoryTracked = 0 
	End If
			
	CloseObj (rst)
		
End Function


Function GetAttDetailID (iDetailID,sType)
dim sql
dim rst
dim i
dim sAttID
	
	Set rst = Server.CreateObject("ADODB.RecordSet")		

	If sType = "svd" then 	sql = "Select * FROM sfSavedOrderAttributes WHERE odrattrsvdOrderDetailID=" & iDetailID & " ORDER BY odrattrsvdAttrID" 
	If sType = "tmp" then 	sql = "Select * FROM sfTmpOrderAttributes WHERE odrattrtmpOrderDetailId=" & iDetailID & " ORDER BY odrattrtmpAttrID" 
	If sType = "odr" then 	sql = "Select * FROM sfOrderDetailsAE WHERE odrdtAEID=" & iDetailID 'b2
	
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText
	If rst.recordcount <= 0 then
		GetAttDetailID = 0
		closeobj(rst)
		exit function
	End if
	
	sAttId = ""
	
	If sType ="odr" then   'b2
		sAttId = rst("odrdtAttDetailID")
		GetAttDetailID = sAttID						
		CloseObj (rst)
		Exit Function
	End If

	rst.movefirst
	For i = 1 to rst.recordcount
		If sType ="svd" then 
			If sAttID <> "" then
				sAttId = sAttId &  "," & rst("odrattrsvdAttrID") 
			else
				sAttId = rst("odrattrsvdAttrID")
			End if
				
		Elseif  sType ="tmp" then 
			if sAttID <> "" then
				sAttId = sAttId & "," & rst("odrattrtmpAttrID") 
			else
				sAttId = rst("odrattrtmpAttrID")
			end if
				
		End if
	
		If not rst.eof then rst.movenext
	Next
		
	GetAttDetailID = sAttID						
	CloseObj (rst)

End Function

Function GetGiftWrapQTY(iTmpOrderDetailID,iType) 'incAE
	
Dim sql,rst,avlqty,gwqty,boqty

	sql = gtmpSQL & " where odrdttmpaeID=" & iTmpOrderDetailID
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn,  adOpenStatic, adLockreadOnly, adCmdText
	
	If rst.RecordCount <= 0 then 
		GetGiftWrapQTY = 0
		exit function
	end if
	
	avlqty= rst("odrdttmpQuantity")  - rst("odrdttmpBackOrderQTY")
	boqty = rst("odrdttmpBackOrderQTY")
	gwqty = rst("odrdttmpGiftWrapQTY")
	
	If iType = 1 then  'shipped gift wraps
		IF rst("odrdttmpGiftWrapQTY") > avlqty then
			GetGiftWrapQTY = avlqty 
		else
			GetGiftWrapQTY = gwqty
		End if
	End If 
	
	If iType = 1 then  'backordered gift wraps
		IF rst("odrdttmpGiftWrapQTY") => avlqty then
			GetGiftWrapQTY = 0
		else
			GetGiftWrapQTY = gwqty
		End if
	end if
	
	CloseObj (rst)
	
	
End Function
Function GetTMPQTY(iTmpOrderDetailID) 'incAE

Dim sql,rst
	
	
	sql = "Select * FROM sftmporderdetails  where odrdttmpID=" & iTmpOrderDetailID
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn,  adOpenKeySet, adLockOptimistic, adCmdText
	If rst.RecordCount <= 0 then 
		GetTMPQTY = "X"
		exit function
	end if
	
	GetTMPQTY = rst("odrdttmpQuantity")
	CloseObj (rst)
	
	
End Function

Function GetAvailableQty(strProductID,AttIDs)
Dim sql, rst
dim ret
	
	
	ret = CheckInventoryTracked (strProductID)
	If ret =  0 then 
		GetAvailableQty = "X" 'No inventoryinfo record
		Exit function
	end if
	
	sql = "Select * FROM sfInventory WHERE invenProdID= '" & strProductID & "' AND  invenAttDetailID='" & AttIDs & "'"
	
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText	
	
	If rst.recordcount <= 0 then 
		GetAvailableQty = "X" 'Inventory record missing
		CloseObj (rst)
		exit function
	end if 
	
	GetAvailableQty = rst("invenInstock")
	
	CloseObj (rst)
	
End Function
Sub JS(sfunction)
	sfunction = replace(sfunction,";","")
	Response.Write "<SCRIPT LANGUAGE=" & chr(34) & "javascript" & chr(34) & ">" & vbcrlf
	Response.write sfunction & ";" & vbcrlf
	Response.Write "</SCRIPT>"
End sub

Function CheckBackOrder(strProductID)
Dim sql, rst
	sql = "Select * FROM sfInventoryInfo WHERE invenProdID= '" & strProductID & "'"
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText	
	
	If rst.recordcount <= 0 then 
		CheckBackOrder = 0
		exit function
	End If
	
		
	IF rst("invenbBackOrder") <> 1 then 
		CheckBackOrder = 0
	Else
		CheckBackOrder = 1
	End If
	
	If rst("invenbTracked") <> 1 then 'if inventory not tracked then no backorder either
		CheckBackOrder = 0
	End If
				
	CloseObj(rst)
		
End Function
Function GetTMPGiftWrapQTY(iTmpOrderDetailID) 'incAE
Dim sql,rst
	
	
	sql = "Select * FROM sftmporderdetailsAE  where odrdttmpaeID=" & iTmpOrderDetailID
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn,  adOpenKeySet, adLockOptimistic, adCmdText
	If rst.RecordCount <= 0 then 
		GetTMPGiftWrapQTY = "X"
		exit function
	end if
	
	GetTMPGiftWrapQTY = rst("odrdttmpGiftWrapQTY")
	CloseObj (rst)
	
	
End Function


%>




	</BODY>
</HTML>

