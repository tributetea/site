<%@ Language=VBScript %>
<%	option explicit 
	Response.Buffer = True
	%>
<!--#include file="SFLib/db.conn.open.asp"-->
<!--#include file="SFLib/adovbs.inc"-->
<!--#include file="SFLib/incGeneral.asp"-->
<!--#include file="sfLib/incDesign.asp"-->
<%
'@BEGINVERSIONINFO

'@APPVERSION: 50.4011.0.2

'@FILENAME: 

'@FILEVERSION: 



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
Dim sMessage

Js "self.resizeTo(500,350)"
%>


<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<!--Header Begin -->
<link rel="stylesheet" href="sfCSS.css" type="text/css">
<title>TRIBUTE TEA | Discount Coupon</title></head>
<body>
<form method="post" name=CoupForm id=CoupForm>
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
            <td  class="tdContent2"> 
              <hr noshade color="#000000" size="1" width="90%" align="center">
              <div align="center"><font class="Content_Large"><b> <font face="Arial, Helvetica, sans-serif">Coupon 
                Code<font size="-7"><br>
                <br>
                </font></font> </b></font></div>
              <div align="center"> 
                <table>
                  <tr> 
                    <td width="10%" align="left"  class="tdContent2" valign="top" nowrap> 
                      <div align="center"><font face="Arial, Helvetica, sans-serif"><i>If 
                        you have a coupon, please enter it below:</i><br>
                        <input type="text" style="<%= C_FORMDESIGN %>" size="22" name="FormCouponCode">
                        <font size="-2"><br>
                        (BE SURE TO CLAIM YOUR EARL GREY GIFT IN THE<br>
                        SCENTED TEA SECTION BEFORE CHECKING OUT.)</font></font></div>
                    </td>
                  </tr>
                </table>
              </div>
              <p align="center"> <font face="Arial, Helvetica, sans-serif">
                <input type=submit onClick="window.document.CoupForm.submit();" value="Submit Coupon" name="btnAction" >
                <b><br>
                </b></font></p>
              <hr noshade color="#000000" size="1" width="90%" align="center">
              <div align="center"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif"><a href="javascript:window.close()"><font color="#FF0000"><b><font color="#CC0000">CLOSE 
                WINDOW</font></b></font></a></font><font size="-7" face="Arial, Helvetica, sans-serif"><br>
                </font> <font face="Arial, Helvetica, sans-serif"><b></b></font></div>
            </td>        
            </tr>
<!--Footer Begin -->
            <tr> 
              <td class="tdFooter"></td>
                </tr>
                </table>
              </td>
            </tr>
            </table>
            </form>
            </body>
         </html>
<!--Footer End -->
<%	
dim sql
dim rst
dim isCouponGood
dim CouponMin
dim CoupMessage

	if Request.Form.Count > 0 then
			sql = "Select * FROM sfCoupons WHERE cpCouponCode= '" & Request.Form("FormCouponCode") & "' AND cpActivate = 1"
			Set rst = Server.CreateObject("ADODB.RecordSet")		
			rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText	
	
			If rst.recordcount <= 0 then 
				 isCouponGood="no"
			else
				If rst("cpNeverExpire") = 0 AND rst("cpExpirationDate") < date then
					isCouponGood="expired"
				else
					isCouponGood="yes"
					CouponMin=rst("cpMin")
				end if
			end if
			
			if isCouponGood="yes" then
				Order_WriteCouponCode Request.Form("FormCouponCode")
				if CouponMin <>"0" then
				CoupMessage="This Coupon has a minimum purchase amount of $" & CouponMin & " .  If your order is not above this amount, the coupon will not be applied."
				JS "alert(" & chr(34) & CoupMessage & chr(34) & ");"
				end if
				JS "window.opener.location='" & (C_HomePath & "order.asp")	& "';"
				JS "self.close()"

			elseif isCouponGood="expired" then
				CoupMessage="This Coupon has expired.  Please enter a new Coupon Code."
				JS "alert(" & chr(34) & CoupMessage & chr(34) & ");"
			else
				CoupMessage="Coupon not found.  Please enter a new Coupon Code."
				JS "alert(" & chr(34) & CoupMessage & chr(34) & ");"
			end if
		
	end if


Sub Order_WriteCouponCode (sCouponCode)
dim sql,rst
	
	'Write to sfTmpOrdersAE
	sql = "Select * FROM sfTmpOrdersAE where odrtmpsessionid=" & Session("SessionID")
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenKeyset, adLockOptimistic, adCmdText
	if not rst.recordcount > 0 then
		rst.AddNew
		rst("odrtmpSessionId") = Session("SessionID")
	end if
    if trim(sCouponCode) <> "" then
		rst("odrtmpCouponCode") = sCouponCode
	end if
    rst.update
    
    CloseObj (rst)
End sub

Sub JS(sfunction)
	sfunction = replace(sfunction,";","")
	Response.Write "<SCRIPT LANGUAGE=" & chr(34) & "javascript" & chr(34) & ">" & vbcrlf
	Response.write sfunction & ";" & vbcrlf
	Response.Write "</SCRIPT>"
End sub



cnn.Close
Set cnn = Nothing


%>



