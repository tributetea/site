<%@ Language=VBScript %>
<%	option explicit 
	Response.Buffer = True
	%>
<!--#include file="SFLib/db.conn.open.asp"-->
<!--#include file="SFLib/incGeneral.asp"-->
<!--#include file="SFLib/adovbs.inc"-->
<!--#include file="sfLib/incDesign.asp"-->
<!--#include file="sfLib/incText.asp"-->
<!--#include file="sflib/incAE.asp"--> 


<%
	server.ScriptTimeout = 900

	'@BEGINVERSIONINFO


	'@APPVERSION: 50.4011.0.2

	'@FILENAME: detail.asp
 
	'

	'@DESCRIPTION: Product detail Page

	'@STARTCOPYRIGHT
	'The contents of this file is protected under the United States
	'copyright laws and is confidential and proprietary to
	'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
	'expressed written permission of LaGarde, Incorporated is expressly prohibited.
	'
	'(c) Copyright 2000, 2001 by LaGarde, Incorporated.  All rights reserved.
	'@ENDCOPYRIGHT

	'@ENDVERSIONINFO

	' Constant Declarations
	const vDebug  = 0		'DeBug Setting
	const iDesign	= 3		'Layout Selection
	const iPageSize	= 10	'unchangable Page size 
	const iMaxRecords = 0   'Maximum amount of records returned, 0 is no maximum
	
	Dim txtProdId, rsProdDetail, rsProdAttributes, aProdDetail, SQL, txtOutput

	txtProdId = Request.QueryString("product_id")

If iConverion = 1 Then Response.Write "<script language=""JavaScript"" src=""http://www.oanda.com/cgi-bin/fxcommerce/fxcommerce.js?user=" & sUserName & """></script>"
Dim CurrencyISO 
CurrencyISO = getCurrencyISO(Session.LCID )	
	Set rsProdDetail = Server.CreateObject("ADODB.Recordset")
	Set rsProdAttributes = Server.CreateObject("ADODB.Recordset")
   if Application("AppName")="StoreFrontAE" then	
     SQL = " SELECT sfProducts.prodID, sfProducts.prodName, sfProducts.prodShortDescription, sfProducts.prodImageSmallPath," _
     & " sfProducts.prodImageLargePath, sfProducts.prodLink, sfProducts.prodPrice, sfProducts.prodAttrNum," _
     & " sfProducts.prodSaleIsActive, sfProducts.prodSalePrice, sfProducts.prodMessage, sfProducts.prodDescription," _
     & " sfSub_Categories.CatHierarchy FROM (sfProducts INNER JOIN sfsubCatdetail ON sfProducts.prodID = sfsubCatdetail.ProdID)" _
     & " INNER JOIN sfSub_Categories ON sfsubCatdetail.subcatCategoryId = sfSub_Categories.subcatID WHERE sfProducts.prodID = '" & txtProdId & "'"
   else
	SQL = "SELECT prodID, prodName, prodShortDescription, prodImageSmallPath, " _
		& "prodImageLargePath, prodLink, prodPrice, prodAttrNum, catName, " _
		& "prodSaleIsActive, prodSalePrice, prodMessage, prodDescription " _
		& "FROM sfProducts " _
		& "INNER JOIN sfCategories ON sfProducts.prodCategoryID = sfCategories.catID " _
		& "WHERE prodID = '" & txtProdId & "'"
   end if
	
	rsProdDetail.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly

	If Not (rsProdDetail.EOF And rsProdDetail.BOF) Then
		If rsProdDetail("prodAttrNum") > 0 Then
			SQL = "SELECT attrName, attrdtId, attrdtName, attrdtPrice, attrdtType, attrdtOrder " _
				& "FROM sfAttributes " _
				& "INNER JOIN sfAttributeDetail ON sfAttributes.attrId = sfAttributeDetail.attrdtAttributeId " _
				& "WHERE attrProdId = '" & rsProdDetail("prodId") & "' ORDER BY AttrName, attrdtOrder"
			rsProdAttributes.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly
		End If
	End If
%>
<html>

<head>
<script language="javascript" src="SFLib/incAE.js"></script>
<SCRIPT language="javascript" src="SFLib/sfCheckErrors.js"></SCRIPT>
<SCRIPT language="javascript" src="SFLib/sfEmailFriend.js"></SCRIPT>
<meta http-equiv="Pragma" content="no-cache">

<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title><%= C_STORENAME %>-StoreFront Product Detail Page</title>



<!--Header Begin -->
<link rel="stylesheet" href="sfCSS.css" type="text/css">
</head>

<body  bgproperties="fixed"  link="<%= C_LINK %>" vlink="<%= C_VLINK %>" alink="<%= C_ALINK %>">

<table border="0" cellpadding="1" cellspacing="0"  class="tdbackgrnd" width="<%= C_WIDTH %>" align="center">
  <tr>
    <td>
      <table width="100%" border="0" cellspacing="1" cellpadding="2" class="tdbackgrnd">
        <tr>
          <td align="middle"  class="tdTopBanner"> 
            <%If C_BNRBKGRND = "" Then%>
            <%Else%>
            <img src="buttons/tt_blue.gif" border="0" width="275" height="36"> 
            <%End If%>
          </td>
        </tr>
<!--Header End -->
        <tr>
          <td width=100%>
           <table border="0" width="100%"> <!--  class="tdContent" -->
						<tr>
							<td align="left" class="tdMiddleTopBanner" width = "20%">&nbsp;</td>
							<td align="center" class="tdMiddleTopBanner" width="60%"><center>Product Detail</center></td>
							
                <td align="right"class="tdMiddleTopBanner" width="20%"><a href="<%= C_HomePath %>order.asp"><img src="buttons/checkout.gif" border="0" align="right" valign="top" alt="Check Out" width="107" height="22"></a></td>
						</tr>
	        </table>
	      </td>
        </tr>
        <tr>
          <td class="tdContent2"> <%
	' -------------------------------------------
    ' Empty Search Results ----------------------
    ' -------------------------------------------
    If rsProdDetail.EOF Then
		closeObj(rsProdDetail)
		closeObj(cnn)
        txtOutput = "<table border=0 width=100&#37;>" _
        & "<tr>" _
        & "<td><center>Product <b>" & txtProdId &"</b> was not found in the current product inventory</center></td></tr>" _
        & "<tr>" _
        & "<td width=""100%""; colspan=""2""><hr noshade color=""#000000"" size=""1"" width=""90%"">" _
        & "</td>" _
        & "</tr>" _
        & "</table>"
			response.write txtOutput
		 Else %>
	        <form method="post" name="<%= rsProdDetail("prodId") %>" action="<%= C_HomePath %>addproduct.asp" onSubmit="this.QUANTITY.quantityBox=true;return sfCheck(this);">            
              <input TYPE="hidden" NAME="PRODUCT_ID" VALUE="<%= rsProdDetail("prodId") %>">         
              <table border="0" width="100%" class="tdContent2" >
                <tr>
                  <td width="30%" align="center">
                  <% If Trim(rsProdDetail("prodLink")) <> "" Then %><% End If %><% If Trim(rsProdDetail("prodImageLargePath")) <> "" Then %><img border="1" src="<%= rsProdDetail("prodImageLargePath") %>"><% ElseIf Trim(rsProdDetail("prodImageSmallPath")) <> "" Then %><img border="1" src="<%= rsProdDetail("prodImageSmallPath") %>"><% End If %><% If Trim(rsProdDetail("prodLink")) <> "" Then response.write "" %></td>
                  <td width="70%" valign="top">
		            <b><%= C_ProductID %>:</b>&nbsp;<%= rsProdDetail("prodId") %>&nbsp;&nbsp;&nbsp;
                   <%if Application("AppName")="StoreFrontAE" then %>	
                     <b><%= C_CategoryNameS %>:</b>&nbsp;<%=getfullpath(rsProdDetail("CatHierarchy"),0) %><br>
                  <%else%> 
                    <b><%= C_CategoryNameS %>:</b>&nbsp;<%=rsProdDetail("catName") %><br>
                   <%end if%> 
                  <b><font class='Content_Large'><%= rsProdDetail("prodName") %></font></b><br>
                  <b><%= C_Description %>:</b>&nbsp;<%If rsProdDetail("prodDescription") <> "" Then %><% rsProdDetail.MoveFirst %><%= rsProdDetail("prodDescription") %><%Else%><%= rsProdDetail("prodShortDescription") %><%End If%><br>
                  <b><%= C_Price %>:</b>&nbsp; 
		          <%
		If iConverion = 1 Then
			If rsProdDetail("prodSaleIsActive") = "1" Then 
				Response.Write "<i><strike><script>document.write(""" & FormatCurrency(rsProdDetail("prodPrice")) & " = ("" + OANDAconvert(" & trim(rsProdDetail("prodPrice")) & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script></strike></i><br>"
				Response.Write "<font color=#FF0000><b>" & C_SPrice & ": <script>document.write(""" & FormatCurrency(rsProdDetail("prodSalePrice")) & " = ("" + OANDAconvert(" & trim(rsProdDetail("prodSalePrice")) & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script></b></font><br>"
				Response.Write "<font color=#FF0000><i>" & C_YSave & ": <script>document.write(""" & FormatCurrency(CDbl(rsProdDetail("prodPrice"))-CDbl(rsProdDetail("prodSalePrice"))) & " = ("" + OANDAconvert(" & trim(CDbl(rsProdDetail("prodPrice"))-CDbl(rsProdDetail("prodSalePrice"))) & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script></i></font><br>"
			Else
				Response.Write "<script>document.write(""" & FormatCurrency(rsProdDetail("prodPrice")) & " = ("" + OANDAconvert(" & trim(rsProdDetail("prodPrice")) & ", " & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script>"
			End If 
		 Else
		   If rsProdDetail("prodSaleIsActive") = "1" Then 
				Response.Write "<i><strike>" & FormatCurrency(rsProdDetail("prodPrice")) & "</strike></i><br>"
				Response.Write "<font color=#FF0000><b>" & C_SPrice & ": " & FormatCurrency(rsProdDetail("prodSalePrice")) & "</b></font><br>"
				Response.Write "<font color=#FF0000><i>" & C_YSave & ": " & FormatCurrency(CDbl(rsProdDetail("prodPrice"))-CDbl(rsProdDetail("prodSalePrice"))) & "</i></font><br>"
		   Else
				Response.Write FormatCurrency(rsProdDetail("prodPrice"))
		   End If
		 End If
		
		  SearchResults_GetProductInventory rsProdDetail("prodid") 'SFAE 
		  SearchResults_ShowMTPricesLink rsProdDetail("prodid") 'SFAE
 %>
		        <br>
		        <table border="0" align="center"> <%
			' -----------------------------
			' ATTRIBUTE OUTPUT ::: BEGIN --
			' -----------------------------
			If rsProdDetail("prodAttrNum") > 0 Then
				Dim attrName, attrNamePrev, icounter, strOut, iAttrNum, strAttrPrice
				strOut = ""
				icounter = 0
				iAttrNum = 0
				
				Function iAttrCounter(iAttrNum)
					iAttrNum = iAttrNum + 1
					iAttrCounter = iAttrNum				
				End Function
				
				Do While Not rsProdAttributes.EOF
					attrName = rsProdAttributes("attrName")
					strAttrPrice = ""
					If Trim(attrName) <> Trim(attrNamePrev) Then
						If iCounter > 0 Then
							strOut = strOut & "</select></td></tr>"
						End If
						strOut = strOut & "<tr><td align=right>" & attrName & "</td><td><select style=""" & C_FORMDESIGN  & """ name=""attr" & (iAttrCounter(iAttrNum)) & """>"
					End If

					Select Case rsProdAttributes("attrdtType")
						Case 1
							strAttrPrice = " (Add " & FormatCurrency(rsProdAttributes("attrdtPrice")) & ")"
						Case 2
							strAttrPrice = " (Subtract " & FormatCurrency(rsProdAttributes("attrdtPrice")) & ")"
					End Select

					strOut = strOut & "<option value=""" & rsProdAttributes("attrdtId") & """>" & rsProdAttributes("attrdtName") & strAttrPrice & "</option>"

					attrNamePrev = attrName
					icounter = icounter + 1
					rsProdAttributes.MoveNext
				Loop
				strOut = strOut & "</select></td></tr>"
				
				Response.Write strOut
			End If 
			' ------------------------
			' ATTRIBUTE OUTPUT ::: END
			' ------------------------ %>
		        </table>
		        <p align="center"><center>Quantity: <input style="<%= C_FORMDESIGN %>"  type="text" name="QUANTITY" title="Quantity" size="3"><br>
		        <%SearchResults_GetGiftWrap rsProdDetail("prodid") 'SFAE%>
		                <input type="image" name="AddProduct" border="0" src="buttons/addtocart3.gif" alt="Add To Cart" width="107" height="28">
                        <br>
                <% If iSaveCartActive = 1 Then%>
                        <input type="image" name="SaveCart" border="0" src="buttons/savetocart3.gif" alt="Save To Cart" width="138" height="22">
                        <%
        End if
        If iEmailActive = 1 Then
        %>
                        <a href="javascript:emailFriend('<%= txtProdId %>')"><img border="0" src="buttons/email.gif" alt="Email a Friend" width="151" height="22"></a> 
                        <% End If %>
                      </center></p>
		      </td>
		    </tr>
		  </table>
		</form>

        <table border="0" width="100%">
	        <tr>
				<td width="100%"></td>
			</tr>
        </table>
      </td>
    </tr>
<% End If %>
<!--Footer begin-->
    <!--#include file="footer.txt"-->
    </table>
</td>
</tr>
</table>
</body>
</html>
<!--Footer End-->
<%
	closeObj(rsProdDetail)
	closeObj(rsProdAttributes)
	closeObj(cnn)
Private Function GetFullPath(Vdata,justMain) 
Dim sSql ,X
Dim iCatId
Dim sFirst
Dim rsCat,rsSubCat
Dim arrTemp ,bMain
'Response.Write vdata & "<BR>"
bMain = false
if left(vData,4)= "none" then
 bMain = True
 arrTemp = split(vdata,"-")
  vdata = arrtemp(1)
 
  'Response.Write "1"
 ' exit function
elseif vData = "" then
'  Response.Write "2"
  GetFullPath = "" 
  exit function
elseif instr(Vdata,"-") = 0  then
'  Response.Write "3"
  'GetFullPath  = vdata
 ' exit function 
 vData = vData 
end if 
 arrTemp = split(vData,"-")

 Set rsCat = Server.CreateObject("ADODB.RecordSet")
 Set rsSubCat = Server.CreateObject("ADODB.RecordSet")

 rsSubCat.Open "sfSub_Categories",cnn,adOpenStatic,adLockReadOnly ,adcmdtable 
 
  For X = 0 To UBound(arrTemp)
    rsSubCat.Requery
    if arrTemp(X)<> "" then
     rsSubCat.Find "SubCatId = " & CInt(arrTemp(X))
     GetFullPath = GetFullPath & rsSubCat("SubCatName") & "-"
    end if
  Next
 ' Response.Write GetFullPath
 '  Response.End 
  sSql  = "Select catName From sfCategories Where catId =" & rsSubCat("subcatCategoryId")   
  
 ' Response.Write sSql
  
  rsCat.Open sSql,cnn,adOpenStatic,adLockReadOnly ,adcmdText
 if justmain = 1 then
 '   Response.Write "$$$"
    GetFullPath = rsCat("catName")
 else 
  ' Response.Write "FFFFF"
   if bMain = True Then
      GetFullPath = rsCat("catName")
   else
     GetFullPath = rsCat("catName") & "-" &  Left(GetFullPath, Len(GetFullPath) - 1)
   end if 
 end if
 
  
Set rsCat = Nothing
Set rsSubCat = Nothing

End Function
	
	
%>




