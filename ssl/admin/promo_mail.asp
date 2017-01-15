<%@ Language=VBScript %>
<%
option explicit
Response.Buffer = True
%>
<!--#include file="../SfLib/sfsecurity.asp"-->
<!--#include file="../SFLib/incDesign.asp"-->
<!--#include file="../SFLib/db.conn.open.asp"-->
<!--#include file="../SFLib/incGeneral.asp"-->
<!--#include file="../SFLib/adovbs.inc"-->
<!--#include file="../SFLib/mail.asp" -->
<%
'--------------------------------------------------------------------
'@BEGINVERSIONINFO

'@APPVERSION: 50.4011.0.2

'@FILENAME: promo_mail.asp
	 
'Access Version

'@DESCRIPTION:   sends promotional mail

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws as an unpublished work, and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO		
%>

<html>
<head>
<script language=javascript>
function checkPromoMail(form) {
	if ((form.startDate.value.length == 0) && (form.endDate.value.length == 0) && (form.product.value.length == 0)) {
		alert("Please enter search criteria.")
		form.startDate.focus();
		return false;
	}
	if ((form.startDate.value.length == 0) && (form.endDate.value.length != 0)) {
		alert("Please enter a starting date.");
		form.startDate.focus();
		return false;
	}
	if ((form.startDate.value.length != 0) && (form.endDate.value.length == 0)) {
		alert("Please enter an ending date.");
		form.endDate.focus();
		return false;
	}
	if (form.subject.value.length == 0) {
		alert("Please enter subject.");
		form.subject.focus();
		return false;		
	}
	if (form.message.value.length == 0) {
		alert("Please enter message.");
		form.message.focus();
		return false;		
	}
	return true;	
}
function checkRemove(form) {
	if (form.custEmail.value.length == "") {
		alert("Please enter an Email Address.")
		form.custEmail.focus()
		return false;
	}
	return true;
}
</script>


<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
	<title><%= C_STORENAME %>-SF Store Promotional Mail Utility</title>


<!--Header Begin -->
<link rel="stylesheet" href="../sfCSS.css" type="text/css">
</head>

<body bgproperties="static" link="<%= C_LINK %>" vlink="<%= C_VLINK %>" alink="<%= C_ALINK %>">

                
<table border="0" cellpadding="1" cellspacing="0" class="tdbackgrnd" width="<%= C_WIDTH %>" align="center">
  <tr>
    <td>
      <table width="100%" border="0" cellspacing="1" cellpadding="3">
        <tr>
          <td align="middle"  class="tdTopBanner"><%If C_BNRBKGRND = "" Then%><%= C_STORENAME %><%Else%><img src="../<%= C_BNRBKGRND %>" border="0"><%End If%></td>
        </tr>
<!--Header End -->
        <tr>
          <td colspan=4 align="center" class="tdMiddleTopBanner">StoreFront Web Store Promotional Mail Utility</td>
        </tr>
        <tr>
          <td align="left" class="tdBottomTopBanner">Use Promo Mail to send promotional mailings to customers who have elected to subscribe to the web store mailing list.  Promotional emails will be sent to all customers who have 
          subscribed to your mailing list within the date range specified in the Date fields.</td>
        </tr>
        <% If Request.Form("SendMail.x") = "" And Request.Form("Remove.x") = "" Then %> 
	    <tr>
	      <td width="100%" class="tdContentBar">Promotional Mailing</td>
	    </tr>
	    <tr>
	      <td class="tdContent">
	        <table border="0" cellpadding="0" cellspacing="5" width="100%">
	          <form action=promo_mail.asp method=post id=form1 name=form1 onsubmit="return checkPromoMail(this);">
	            <tr><td>Start Date:</td><td>
                  <input type=text name=startDate SIZE="20"></td><td>End Date:</td><td>
                  <input type=text name=endDate SIZE="20"></td></tr>
	            <tr><td>Product ID:</td><td colspan=3><input type=text name=product size=20></td></tr>
	            <tr><td colspan=4><hr noshade width="90%"></td></tr>
	            <tr><td>Subject:</td><td colspan=3><input type=text name=subject size=58></td></tr>
	              <tr><td colspan=4><u>Mail Message</u></td></tr>
	              <tr><td colspan=4><textarea rows=10 cols=66 name=message></textarea></td></tr>
	              <tr><td colspan=4 align="center"><input type=image name="SendMail" border="0" src="../images/buttons/submit.gif" alt"Send Mail"></td></tr>
	            </form>
	          </table>
	        </td>
	      </tr>
	      <tr>
	        <td width="100%" class="tdContentBar">Customer Removal</td>
	      </tr>
	      <tr>
	        <form action=promo_mail.asp method=post id=form1 name=form2 onsubmit="return checkRemove(this);">
	          <td align="center" class="tdContent">
	            <b>Note:</b> To remove multiple addresses, enter the addresses 
                separated by commas<br>
	           E-Mail Address(es):&nbsp;<input type=text name=custEmail size=50><input type=image border="0" name="Remove" src="../images/buttons/submit.gif" alt"Remove Email(s)"></td>
	        </form>
	      </tr> 
          <% 
ElseIf Request.Form("SendMail.x") <> "" Then
	'*****************
	'** Run Mailing **
	'*****************
	Dim rsMailer,sInformation,sSubject,sMessage,arrMail,iNum,sName,tTemp,arrMailSent,j,sTemp
	Set rsMailer = Server.CreateObject("ADODB.Recordset")
	
	If Request.Form("product") <> "" and Request.Form("startDate") <> "" And Request.Form("endDate") <> "" Then
		SQL = "SELECT custEmail, custFirstName, custMiddleInitial, custLastName, odrdtProductID FROM sfCustomers INNER JOIN (sfOrders INNER JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId) ON sfCustomers.custID = sfOrders.orderCustId " & _
		" WHERE (custIsSubscribed = 1) AND (custLastAccess >=  # " & MakeUSDate(Request("startDate")) & " #) AND (custLastAccess <= # " & MakeUSDate(Request("endDate")) & " #) AND odrdtProductID = '" & Request.Form("product") & "'"
	ElseIf Request.Form("product") <> "" Then
		SQL = "SELECT custEmail, custFirstName, custMiddleInitial, custLastName, odrdtProductID FROM sfCustomers INNER JOIN (sfOrders INNER JOIN sfOrderDetails ON sfOrders.orderID = sfOrderDetails.odrdtOrderId) ON sfCustomers.custID = sfOrders.orderCustId " & _
		" WHERE (custIsSubscribed = 1) AND odrdtProductID = '" & Request.Form("product") & "'"
	Else
		SQL = "SELECT custEmail, custFirstName, custMiddleInitial, custLastName FROM sfCustomers WHERE (custIsSubscribed = 1) AND (custLastAccess >= # " & MakeUSDate(Request("startDate")) & " #) AND (custLastAccess <= # " & MakeUSDate(Request("endDate")) & " #)"
	End If
	rsMailer.CursorLocation = adUseClient
	rsMailer.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
	
	sSubject = Request.Form("subject")
	sMessage = Request.Form("message")
	If Not (rsMailer.BOF And rsMailer.EOF) Then arrMail = rsMailer.GetRows 
	
	If isArray(arrMail) Then
		sTemp=""
		j = 0
		Redim arrMailSent(uBound(arrMail,2))
		For i=0 to uBound(arrMail,2)
			If arrMail(0,i) <> sTemp Then	
				sName = trim(arrMail(1,i)) & " " & trim(arrMail(2,i)) & " " & trim(arrMail(3,i))
				sInformation = arrMail(0,i) & "|" & sSubject & "|" & "Dear " & sName & "," & vbcrlf & vbtab & sMessage
				createMail "PromoMail", sInformation
				arrMailSent(j) = arrMail(0,i)
				j = j + 1
				sTemp = arrMail(0,i)
			End If
		Next
		Redim Preserve arrMailSent(j-1)
	End If 
	%>
	      <tr>
		    <td width="100%" class="tdContentBar">Promotional Mailing Results</td>
	      </tr>
	      <tr>
	        <td class="tdContent">
	          <table border="0" cellpadding="0" cellspacing="5" width="100%">
	            <% If isArray(arrMailSent) Then %>
		        <tr><td align=center colspan=4>The promotional mailing was successfully completed.</td></tr>
	            <% Else %>
		        <tr><td align=center colspan=4>There are no Subscribed Customers who 
                  Ordered between <%= Request.Form("startDate") %> and <%= Request.Form("endDate") %><% If Request.Form("product") <> "" Then %> for Product ID <%= Request.Form("product") %><%End If%></td></tr>
	            <% End If %>
	            <%
	If isArray(arrMailSent) Then
		iNum = uBound(arrMailSent)
		For i=0 to iNum Step 4
			Response.Write "<tr><td><font class='Content_Small'>" & arrMailSent(i) & "</font></td>" 
			If iNum >=i+1 Then 
				Response.Write "<td><font class='Content_Small'>" & arrMailSent(i+1) & "</font></td>"
			Else
				Response.Write "<td><font class='Content_Small'>&nbsp;</font></td>"
			End If
			If iNum >=i+2 Then 
				Response.Write "<td><font class='Content_Small'>" & arrMailSent(i+2) & "</font></td>"
			Else
				Response.Write "<td><font class='Content_Small'>&nbsp;</font></td>"
			End If
			If iNum >=i+3 Then 
				Response.Write "<td><font class='Content_Small'>" & arrMailSent(i+3) & "</font></td></tr>" 
			Else
				Response.Write "<td><font class='Content_Small'>&nbsp;</font></td></tr>"
			End If
		Next 
	End If 
	%>
	            <tr><td align=center colspan=4><a href="promo_mail.asp">Back</a></td></tr>
	          </table>
	        </td>
	      </tr> 
	      <%
	closeObj(rsMailer)
ElseIf Request.Form("Remove.x") <> "" Then
	'**********************
	'** Remove recipient **
	'**********************
	Dim aAddresses, i, SQL,sNot_Found,sMsg,sAddresses,aNot_Found
    sAddresses = Request("custEmail")
	sNot_Found = get_Invalid_eMail(sAddresses)
'''''''''''''''''''''''''''''''''''''''''''''''''''''#287	
	aAddresses = Split(sAddresses,",")	
	aNot_Found = split(sNot_Found,",")
	  
	  for i = 0 to Ubound(aNot_Found)
	   sAddresses = replace(sAddresses,aNot_Found(i)," ")
	  next
	   
	   If left(trim(sAddresses),1) = "," then
	    sAddresses = mid(trim(sAddresses),2,len(trim(sAddresses)))
	   End if
	   If right(trim(sAddresses),1) = "," Then
	    sAddresses = Left(Trim(sAddresses),len(Trim(sAddresses)-1))
	   End IF 
	   If Trim(sAddresses) <> "" Then
	   sAddresses = replace(sAddresses,", ," ,",")
	    sMsg = sAddresses & " was successfully removed from the list.<BR>"
	   End if
	   If Trim(sNot_Found) <> "" Then
	  	   sNot_Found = replace(sNot_Found,", ," ,",")
    	   sMsg = sMsg &  sNot_Found & " was not found in your database."
       
       End if
       
''''''''''''''''''''''''''''''''''''''''''''''' 
	SQL = "UPDATE sfCustomers SET custIsSubscribed = 0 WHERE "
	For i = 0 To UBound(aAddresses)-1
		SQL = SQL & "custEmail = '" & trim(aAddresses(i)) & "' OR "
	Next 
	SQL = SQL & "custEmail = '" & aAddresses(i) & "'"
	'response.write SQL 
	cnn.Execute SQL
	%>
	      <tr>
		    <td width="100%" class="tdContentBar">Customer Removal</td>
	      </tr>
	      <tr>
	        <td class="tdContent">
	          <table border="0" cellpadding="0" cellspacing="5" width="100%">
	            <tr><td align=center colspan=4><%=smsg %></td></tr>
	            <tr><td align=center colspan=4><a href="promo_mail.asp">Back</a></td></tr>
	          </table>
	        </td>
	      </tr>
          <% 
End If 
closeObj(cnn)
%>
	   <!--Footer begin-->
                <tr>
		<td class="tdFooter"><p align="center"><font class="Footer"><b><a href="sfReports.asp">Reports</a> | <a href="<%= C_HOMEPATH %>"><%= C_STORENAME %></a></b></font></p></td>
	    </tr>
              </table>
            </td>
          </tr>
        </table>
       </body>

    </html>
<!--Footer End-->



