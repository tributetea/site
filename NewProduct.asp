<%@ Language=VBScript %>
<%
option explicit
Response.Buffer = True
	%>
	<!--#include file="SFLib/db.conn.open.asp"-->
	<!--#include file="SFLib/adovbs.inc"-->
	<!--#include file="SFLib/incSearchResult.asp"-->
	<!--#include file="SFLib/incGeneral.asp"-->
	<!--#include file="SFLib/incDesign.asp"-->
	<!--#include file="SFLib/incText.asp"-->
	
	<!--#include file="sflib/incAE.asp"-->
	<%	

'@BEGINVERSIONINFO

'@APPVERSION: 50.4011.0.4

'@FILENAME: NewProduct.asp
	 
'Access Version

'@DESCRIPTION:   Retrieves New Products for Display

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws as an unpublished work, and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO	

	' Constant Declarations
	const varDebug  = 0		'DeBug Setting
	const iPageSize	= 10	'Number of Records per Page 
	const iMaxRecords = 0   'Maximum amount of records returned, 0 is no maximum
	const iDaysBack = 14 'Number of days back to look for Added products
	
	Dim txtsearchParamTxt, txtsearchParamType, txtsearchParamCat, txtFromSearch,  txtsearchParamMan, srchDate
	Dim txtCatName, txtsearchParamVen, txtImagePath, txtOutput, txtDateAddedStart
	Dim txtDateAddedEnd, txtPriceStart, txtPriceEnd, txtSale, SQL, sAmount
	Dim iAttCounter, irsSearchAttRecordCount, iAttDetailCounter, irsSearchAttDetailRecordCount
	Dim iPage, iRec, iNumOfPages, iDesignCounter, iVarPageSize, iSearchRecordCount, icounter
	Dim rsCat, iDesign,rsSearch, rsSearchAtt, rsSearchAttDetail, rsCatImage, arrAttDetail, arrProduct, arrAtt, CurrencyISO


	CurrencyISO = getCurrencyISO(Session.LCID)
	If iConverion = 1 Then Response.Write "<script language=""JavaScript"" src=""http://www.oanda.com/cgi-bin/fxcommerce/fxcommerce.js?user=" & sUserName & """></script>"
	iDesign = C_DesignType 'Page Layout
	iDesignCounter = 2
	iVarPageSize = iPageSize ' Records Per Page
	txtsearchParamCat = "ALL"
	srchDate = MakeUSDate(Date()-iDaysBack)
	Set rsSearch = Server.CreateObject("ADODB.RecordSet")
	' -------------------------------------------
	' RecordSet Paging Setup --------------------
	' -------------------------------------------
	
   if Application("AppName")="StoreFrontAE" then	
    'SQL = "SELECT sfProducts.prodID, sfProducts.prodName, sfProducts.prodImageSmallPath, sfProducts.prodLink, sfProducts.prodPrice, sfProducts.prodSaleIsActive," _
    ' & " sfProducts.prodSalePrice, sfProducts.prodDescription, sfProducts.prodAttrNum, sfProducts.prodCategoryId, sfProducts.prodShortDescription," _
    ' & " sfSub_Categories.CatHierarchy "_ 
    ' & " FROM (sfProducts INNER JOIN sfSubCatDetail ON sfProducts.prodID = sfSubCatDetail.ProdID) INNER JOIN sfSub_Categories ON sfSubCatDetail.subcatCategoryId = sfSub_Categories.subcatID" _
    ' & " WHERE (sfProducts.prodDateAdded >= #" & srchDate & "#  AND sfProducts.prodEnabledIsActive = 1)"
     SQL = "SELECT ProdID, prodName, prodImageSmallPath, prodLink, prodPrice, prodSaleIsActive, prodSalePrice, ProdID, prodDescription, prodAttrNum, prodCategoryId, prodShortDescription " _
			& "FROM sfProducts  WHERE (prodDateAdded >= #" & srchDate & "#  AND prodEnabledIsActive = 1)"
	
   else
	SQL = "SELECT ProdID, prodName, prodImageSmallPath, prodLink, prodPrice, prodSaleIsActive, prodSalePrice, catName, prodDescription, prodAttrNum, prodCategoryId, prodShortDescription " _
			& "FROM sfProducts INNER JOIN sfCategories ON sfProducts.prodCategoryId = sfCategories.catID WHERE (prodDateAdded >= #" & srchDate & "#  AND prodEnabledIsActive = 1)"
	
   end if
  ' Response.Write sql
	If varDebug = 1 Then Response.Write SQL & "<br><br>"
	With rsSearch
		.CursorLocation = adUseClient
		.CacheSize = iVarPageSize
		.MaxRecords = iMaxRecords
		.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText 
		.PageSize = iVarPageSize
	End With
	
	' Determine the page user is requesting
	If Request.QueryString("PAGE") = "" Then
		iPage = 1
	Else
		iPage = CInt(Request.QueryString("PAGE"))
		' Protect against out of range pages, in case
		' of a user specified page number
		If iPage < 1 Then
			iPage = 1
		Else
			If iPage > rsSearch.PageCount Then
				iPage = rsSearch.PageCount
			Else
				iPage = CInt(Request.QueryString("PAGE"))
			End If
		End If
	End If
	
	iSearchRecordCount = rsSearch.RecordCount
	iNumOfPages = Int(iSearchRecordCount / iPageSize)
	
	If CInt(iNumOfPages+1) = CInt(iPage) Then iVarPageSize = rsSearch.RecordCount - (iNumOfPages * iPageSize) 

	If NOT rsSearch.EOF Then rsSearch.AbsolutePage = CInt(iPage)
	
	' Create Attribute Record Sets for product on page
	SQL = getAttributeSQL(rsSearch, iVarPageSize, iPage)
	If varDebug = 1 Then Response.Write SQL & "<br><br>"	
	Set rsSearchAtt = Server.CreateObject("ADODB.RecordSet")
	If SQL <> "" Then 
		rsSearchAtt.Open SQL, cnn, adOpenKeyset, adLockReadOnly, adCmdText
		SQL = getAttributeDetailSQL(rsSearchAtt)
		If varDebug = 1 Then Response.Write SQL & "<br><br>"
		If SQL <> "" Then
			Set rsSearchAttDetail = Server.CreateObject("ADODB.RecordSet")
			rsSearchAttDetail.Open SQL, cnn, adOpenKeyset, adLockReadOnly, adCmdText
		End If 
	End If 
	
	If txtsearchParamCat = "ALL" Then
		txtCatName = "All " & C_CategoryNameP
	Else
		If Not rsSearch.EOF Then
		  if Application("AppName")="StoreFrontAE" then	
			txtCatName =GetFullPath(rsSearch.Fields("ProdID"),1,"ALL")
		  else
		   txtCatName =rsSearch.Fields("CatNAME")
		  end if 
		Else
			Set rsCat = Server.CreateObject("ADODB.RecordSet")
			SQL = getCategorySQL(txtsearchParamCat)
			rsCat.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
			txtCatName = rsCat.Fields("catName")
			closeObj(rsCat)
		End If 
	End If
	If txtsearchParamTxt = "" Then txtsearchParamTxt = "*"
	'Corrects Number of Pages if there is overflow less then records per page 
	If iSearchRecordCount mod iPageSize <> 0 Then iNumOfPages = iNumOfPages + 1
                     	
%>
<html>
<head>
<SCRIPT language="javascript" src="SFLib/sfCheckErrors.js"></SCRIPT>
<SCRIPT language="javascript" src="SFLib/sfEmailFriend.js"></SCRIPT>
<SCRIPT language="javascript" src="SFLib/incAE.js"></SCRIPT>

<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title><%= C_STORENAME %>-SF New Products Page</title>


<!--Header Begin -->
<link rel="stylesheet" href="sfCSS.css" type="text/css">
</head>

<body bgproperties="static" link="<%= C_LINK %>" vlink="<%= C_VLINK %>" alink="<%= C_ALINK %>">

                
<table border="0" cellpadding="1" cellspacing="0" class="tdbackgrnd" width="<%= C_WIDTH %>" align="center">
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
        <tr> 
          <td> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td align="left" class="tdMiddleTopBanner" width = "20%">&nbsp;</td>
                <td align="center" class="tdMiddleTopBanner">
                  <center>
                    New Products
                  </center>
                </td>
                <td align="right" class="tdMiddleTopBanner" width="20%"><a href="<%= C_HomePath %>order.asp"><img src="buttons/checkout.gif" border="0" align="right" valign="top" alt="Check Out" width="107" height="22"></a></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <!-- ### SEARCH RESULT QUERY OUTPUT ::: BEGIN ### -->
			<td align="middle" class="tdBottomTopBanner">
          <%IF Application("AppName")="StoreFrontAE" Then
          if rsSearch.EOF and rsSearch.BOF then%>
            <%else%>
          Search for: <b><%= txtsearchParamTxt %></b> 
          in <%= C_CategoryNameS %>: <b><%= GetFullPath(rsSearch.Fields("ProdId"),0,"ALL") %></b> 
          <%end if
         Else
          %>
          Search 
            for: <b><%= txtsearchParamTxt %></b> in <%= C_CategoryNameS %>: <b><%= txtCatName %></b> 
            <%
         End if
         %>
            <br>
            <b><%= iSearchRecordCount %></b> Records Returned <b><%= iVarPageSize %></b> 
            Records Displayed.<br>
            <a href="search.asp"><b>New Search</b></a>|<a href="advancedsearch.asp"><b>Advanced 
            Search</b></a></td>
          <!-- ### SEARCH RESULT QUERY OUTPUT ::: END ### -->
          </tr>
        <tr> 
          <td class="tdContent2"> 
            <table border="0" width="100%">
              <tr> 
                <td width="100%">
                  <p align="center"><font class="Content_Small"> 
                    <%
        If iNumOfPages <> 1 And iNumOfPages <> 0 Then Response.Write bottomPaging(iPage, iPageSize, iSearchRecordCount, iNumOfPages, "NewProducts")
		%>
                    </font> 
                  <hr noshade color="#000000" size="1" width="90%">
                </td>
              </tr>
            </table>
            <%
	' -------------------------------------------
    ' Empty Search Results ----------------------
    ' -------------------------------------------
    If rsSearch.EOF Then
%>
<table border=0 width=100&#37;>
<tr>
<td><center><font class='Content_Large'>Sorry, No Items are Newly Added!</font></center></td></tr>
<tr>
<td width="100%"; colspan="2"><hr noshade color="#000000" size="1" width="90%">
</td>
</tr>
</table>
<%
        'Response.End
    Else
        ' -------------------------------------------
        ' SEARCH RESULT PRODUCT OUTPUT ::: BEGIN ----
        ' -------------------------------------------
        'create arrays for display
        arrProduct = rsSearch.GetRows(iVarPageSize)
        If Not rsSearchAtt.EOF Then 
			arrAtt = rsSearchAtt.GetRows
			irsSearchAttRecordCount = rsSearchAtt.RecordCount-1
			If Not rsSearchAttDetail.EOF Then 
				arrAttDetail = rsSearchAttDetail.GetRows
				irsSearchAttDetailRecordCount = rsSearchAttDetail.RecordCount-1 
			End If
		End If
		'sfAE
		 
     For iRec = 0 to iVarPageSize-1 
			
				
				
				SearchResults_CheckInventoryTracked 'SFAE  b2   
                ' Set Default Image if none specified for product
                If arrProduct(2, iRec) = "" Then
                   ' Set rsCatImage = Server.CreateObject("ADODB.RecordSet")
                   ' SQL = "SELECT catImage FROM sfCategories WHERE catID = " & arrProduct(10, iRec)
                   ' rsCatImage.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
                    'txtImagePath = rsCatImage.Fields("catImage")
                  '  closeObj(rsCatImage) 
                Else
                    txtImagePath = arrProduct(2, iRec)
                End If
      			icounter = 1 
				'Response.Write   arrProduct(11, iRec) & "<bR>"
				'Response.Write   arrProduct(7, iRec)
				       

%>
            <form method="post" name="<%= arrProduct(0, iRec)%>" action="<%= C_HomePath %>addproduct.asp" onSubmit="this.QUANTITY.quantityBox=true;return sfCheck(this);">
              <input type="hidden" name="PRODUCT_ID" value="<%= arrProduct(0, iRec)%>">
              <table border="0" width="100%" class="tdContent" >
                <tr> 
                  <% If iDesign = "1" Then %>
                  <td width="30%" align="center"><a href="<%= replace(arrProduct(3, iRec)," ","+") %>">
                    <%If txtImagePath <> "" Then%>
                    <img border="1" src="<%= txtImagePath %>">
                    <%ElseIf Trim(arrProduct(3, iRec)) <> "" Then %>
                    Link
                    <%End If%>
                    </a><br>
                  </td>
                  <% ElseIf iDesign = "3" And (iDesignCounter / 2) = Int(iDesignCounter / 2) Then%>
                  <td width="30%" align="center"><a href="<%= replace(arrProduct(3, iRec)," ","+") %>">
                    <%If txtImagePath <> "" Then%>
                    <img border="1" src="<%= txtImagePath %>">
                    <%ElseIf Trim(arrProduct(3, iRec)) <> "" Then %>
                    Link
                    <%End If%>
                    </a><br>
                  </td>
                  <% End If %>
                  <td width="70%" valign="top"> <b><%= C_ProductID %>:</b>&nbsp;<%= arrProduct(0, iRec) %>&nbsp;&nbsp;&nbsp; 
                    <b><%= C_CategoryNameS %>:</b>&nbsp;
                    <% if Application("AppName")="StoreFrontAE" then	
                        Response.Write  GetFullPath(arrProduct(7, iRec),0,"ALL")
                     else
                        Response.Write arrProduct(7, iRec)
                     end if   
                      %>
                    <br>
                    <b><font class='Content_Large'><%= arrProduct(1, iRec) %></font></b><br>
                    <b><%= C_Description %>:</b>&nbsp; 
                    <%
                          If arrProduct(11, iRec) <> "" Then
                            Response.Write arrProduct(11, iRec) 
                          Else 
                            Response.Write arrProduct(8, iRec) 
                          End If
                      
                           %>
                    <br>
                    <b><%= C_Price %>:</b>&nbsp; 
                    <% If iConverion = 1 Then
					If arrProduct(5, iRec) = 1 Then 
							Response.Write "<i><strike><script>document.write(""" & FormatCurrency(arrProduct(4, iRec)) & " = ("" + OANDAconvert(" & trim(arrProduct(4, iRec)) & "," & chr(34) & CurrencyISO & chr(34) & ") + "")"");</script></strike></i><br>"
							Response.Write "<font color=#FF0000><b>" & C_SPrice & ": <script>document.write(""" & FormatCurrency(arrProduct(6, iRec)) & " = ("" + OANDAconvert(" & trim(arrProduct(6, iRec)) & ", " & chr(34) & CurrencyISO & chr(34) & ")+ "")"");</script></b></font><br>"
							Response.Write "<font color=#FF0000><i>" & C_YSave & ": <script>document.write(""" & FormatCurrency(CDbl(arrProduct(4, iRec))-CDbl(arrProduct(6, iRec))) & " = ("" + OANDAconvert(" & trim(CDbl(arrProduct(4, iRec))-CDbl(arrProduct(6, iRec))) & ", " & chr(34) & CurrencyISO & chr(34) & ")+ "")"");</script></i></font><br>"
					Else
							Response.Write "<script>document.write(""" & FormatCurrency(arrProduct(4, iRec)) & " = ("" + OANDAconvert(" & trim(arrProduct(4, iRec)) & ", " & chr(34) & CurrencyISO & chr(34) & ")+ "")"");</script>"
					End If 
			   Else
					If arrProduct(5, iRec) = 1 Then 
							Response.Write "<i><strike>" & FormatCurrency(arrProduct(4, iRec)) & "</strike></i><br>"
							Response.Write "<font color=#FF0000><b>" & C_SPrice & ": " & FormatCurrency(arrProduct(6, iRec)) & "</b></font><br>"
							Response.Write "<font color=#FF0000><i>" & C_YSave & ": " & FormatCurrency(CDbl(arrProduct(4, iRec))-CDbl(arrProduct(6, iRec))) & "</i></font><br>"
					Else
							Response.Write FormatCurrency(arrProduct(4, iRec))
					End If 
			   End If
			%>
                    <% SearchResults_GetProductInventory arrProduct(0, iRec) 'SFAE b2 %>
                    <% SearchResults_ShowMTPricesLink arrProduct(0, iRec) 'SFAE%>
                    <br>
                    <table border="0" align="center">
                      <%
                            ' -------------------------------------------
                            ' SEARCH RESULT ATTRIBUTE OUTPUT ::: BEGIN --
                            ' -------------------------------------------
                            If irsSearchAttRecordCount <> "" Then
								For iAttCounter = 0 to irsSearchAttRecordCount
									If arrProduct(0, iRec) = arrAtt(2, iAttCounter) Then
%>
                      <tr> 
                        <td align="right"><%= arrAtt(1, iAttCounter) %></td>
                        <td>
                          <select size="1" name="attr<%= icounter %>" style="<%= C_FORMDESIGN %>">
                            <%
										For iAttDetailCounter = 0 to irsSearchAttDetailRecordCount
											If arrAtt(0, iAttCounter) = arrAttDetail(1, iAttDetailCounter) Then
														sAmount = ""
												Select Case arrAttDetail(4, iAttDetailCounter)
													Case 1 
														sAmount = " (add " & FormatCurrency(arrAttDetail(3, iAttDetailCounter)) & ")"
													Case 2 
														sAmount = " (subtract " & FormatCurrency(arrAttDetail(3, iAttDetailCounter)) & ")"
												End Select
												Response.Write "<option value=" & arrAttDetail(0, iAttDetailCounter) & ">" & arrAttDetail(2, iAttDetailCounter) & sAmount & "</option>"
											End If
										Next
%>
                          </select>
                          </td>
                      </tr>
                      <%  
									icounter = icounter + 1
									End If 
								Next
							End If 
                            ' -------------------------------------------
                            ' SEARCH RESULT ATTRIBUTE OUTPUT ::: END
                            ' -------------------------------------------
                     
%>
                    </table>
                    <p align="center">
                      <center>
                        <%= C_Quantity %>:
                        <input style="<%= C_FORMDESIGN %>"  type="text" name="QUANTITY" title="Quantity" size="3">
                        <br>
                        <%SearchResults_GetGiftWrap arrProduct(0, iRec) 'SFAE%>
                        <input type="image" name="AddProduct" border="0" src="buttons/addtocart3.gif" alt="Add To Cart" width="107" height="28">
                        <br>
                        <% If iSaveCartActive = 1 Then%>
                        <input type="image" name="SaveCart" border="0" src="buttons/savetocart3.gif" alt="Save To Cart" width="138" height="22">
                        <%
        End if
        If iEmailActive = 1 Then
        %>
                        <a href="javascript:emailFriend('<%= server.urlEncode(arrProduct(0, iRec)) %>')"><img border="0" src="buttons/emailwishlist.gif" alt="Email a Friend" width="151" height="22"></a> 
                        <% End If %>
                      </center>
                    </p>
                  </td>
                  <%  If iDesign = "2" Then %>
                  <td width="30%" align="center"><a href="<%= replace(arrProduct(3, iRec)," ","+") %>">
                    <%If txtImagePath <> "" Then%>
                    <img border="1" src="<%= txtImagePath %>">
                    <%ElseIf Trim(arrProduct(3, iRec)) <> "" Then %>
                    Link
                    <%End If%>
                    </a><br>
                  </td>
                  <%  ElseIf iDesign = "3" And (iDesignCounter / 2) <> Int(iDesignCounter / 2) Then %>
                  <td width="30%" align="center"><a href="<%= replace(arrProduct(3, iRec)," ","+") %>">
                    <%If txtImagePath <> "" Then%>
                    <img border="1" src="<%= txtImagePath %>">
                    <%ElseIf Trim(arrProduct(3, iRec)) <> "" Then %>
                    Link
                    <%End If%>
                    </a><br>
                  </td>
                  <%  End If %>
                </tr>
                <tr> 
                  <td width="100%" colspan="2">
                    <hr noshade color="#000000" size="1" width="90%">
                  </td>
                </tr>
              </table>
            </form>
            <%              
				Response.Flush
            If iDesign = "3" Then iDesignCounter = iDesignCounter + 1
          
        Next
        ' -------------------------------------------
        ' SEARCH RESULT PRODUCT OUTPUT ::: END ------
        ' -------------------------------------------
    End If 
%>
            <table border="0" width="100%">
              <tr> 
                <td width="100%">
                  <p align="center"><font class="Content_Small"> 
                    <%
                ' -------------------------------------------   
                ' SEARCH RESULT PAGING OUTPUT ::: BEGIN -----
                ' -------------------------------------------
If iNumOfPages <> 1 And iNumOfPages <> 0 Then Response.Write bottomPaging(iPage, iPageSize, iSearchRecordCount, iNumOfPages, "NewProducts")
                ' -------------------------------------------   
                ' SEARCH RESULT PAGING OUTPUT ::: END -------
                ' -------------------------------------------
%>
                    </font></p>
                </td>
              </tr>
            </table>
          
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
	' Object Cleanup
	closeObj(rsSearch)
	closeObj(rsSearchAtt)
	closeObj(rsSearchAttDetail)
	closeObj(cnn)
%>


























