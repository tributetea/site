


<%
'@BEGINVERSIONINFO

'@APPVERSION: 50.4011.0.3

'@FILENAME: product.asp
	 

'

'@DESCRIPTION: Include File for Product Page

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws as an unpublished work, and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO

Set rsAdmin = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT adminOandaID,adminActivateOanda FROM sfAdmin"
rsAdmin.Open SQL, cnn,3,3 , 1
sUserName = rsAdmin("adminOandaID")
iConverion = rsAdmin("adminActivateOanda")
If iConverion = 1 Then Response.Write "<script language=""JavaScript"" src=""http://www.oanda.com/cgi-bin/fxcommerce/fxcommerce.js?user=" & sUserName & """></script>"
closeObj(rsAdmin)


Function getProductInfo(sProdID, sCase)
	Dim SQL, rsProd, iAEcatID
	SQL = "SELECT * FROM sfProducts WHERE prodID = '" & sProdID & "'"
	Set rsProd = Server.CreateObject("ADODB.Recordset")
	rsProd.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
	If Not (rsProd.BOF and rsProd.EOF) Then
		Select Case sCase
			Case 1
				getProductInfo = sProdID
			Case 2
				getProductInfo = rsProd("prodName")
			Case 3
				getProductInfo = rsProd("prodShortDescription")
			Case 4
				getProductInfo = rsProd("prodDescription")
			Case 5
				getProductInfo = rsProd("prodImageSmallPath")
			Case 6
				getProductInfo = rsProd("prodImageLargePath")
			Case 7	
				getProductInfo = rsProd("prodLink")
			Case 8
				getProductInfo = rsProd("prodPrice")
			Case 9
				getProductInfo = rsProd("prodSalePrice")
			Case 10
				Set rs = Server.CreateObject("ADODB.Recordset")
				If Application("AppName")="StoreFrontAE" then
                  SQL = "SELECT subcatCategoryId FROM sfSubCatDetail WHERE prodID = '" & rsProd("prodId") & "'"
               	  rs.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
     			  iAEcatID = rs("subcatCategoryId")
               	  rs.close
               	   SQL = "SELECT CatHierarchy FROM sfSub_Categories WHERE subcatID = " & iAEcatID
               	   rs.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
     			   getProductInfo = GetFullPath(rs("CatHierarchy"),2)
               	Else
					SQL = "SELECT catName FROM sfCategories WHERE catID = " & rsProd("prodCategoryId")
					rs.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
     				getProductInfo = rs(0)
				end if
				rs.Close 
				Set rs = Nothing
			Case 11
				Set rs = Server.CreateObject("ADODB.Recordset")
				SQL = "SELECT mfgName FROM sfManufacturers WHERE mfgID = " & rsProd("prodManufacturerId")
				rs.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
				getProductInfo = rs(0)
				rs.Close 
				Set rs = Nothing
			Case 12
				Set rs = Server.CreateObject("ADODB.Recordset")
				SQL = "SELECT vendName FROM sfVendors WHERE vendID = " & rsProd("prodVendorId")
				rs.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
				getProductInfo = rs(0)
				rs.Close 
				Set rs = Nothing
			Case 15
				getProductInfo = rsProd("prodSaleIsActive")
			Case Else
				Set rs = Server.CreateObject("ADODB.Recordset")
				SQL = "SELECT * FROM sfAttributes WHERE attrProdId = '" & sProdID & "'"
				rs.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
				
				sDetails = ""
				If Not (rs.EOF And rs.BOF) Then
				    iCounter = 1
				    While Not rs.EOF
				        sDetails = sDetails & "<br>" & rs("attrName") & "<br>" & getAttributeDetails(rs("attrID"), iCounter, sCase)
				        rs.MoveNext
				        iCounter = iCounter + 1
				    Wend
				End If
				getProductInfo = sDetails
		End Select
	Else
		Response.Write "Product Could Not Be Found"
	End If
	closeObj(rsProd)
End Function



Function getAttributeDetails(attrID, iCounter, sCase)
	Set rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM sfAttributeDetail WHERE attrdtAttributeId = " & attrID & " ORDER BY attrdtOrder"
	rs.Open SQL, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
	
    If sCase = 14 Then
        sTemp = "<select name=""attr" & iCounter & """>"
    Else
        sTemp = ""
    End If
    
    sChecked = "CHECKED"
    
    While Not rs.EOF
        sAmount = ""
        Select Case rs("attrdtType")
            Case 1
                sAmount = " (add " & FormatCurrency(rs("attrdtPrice")) & ")"
            Case 2
                sAmount = " (subtract " & FormatCurrency(rs("attrdtPrice")) & ")"
        End Select
        
        If sCase = 14 Then
            sTemp = sTemp & "<option value=" & rs("attrdtID") & ">" & rs("attrdtName") & sAmount & "</option>"
        Else
            sTemp = sTemp & "<input type=""radio"" " & sChecked & " name=""attr" & iCounter & """ value=""" & rs("attrdtID") & """>" & rs("attrdtName") & sAmount & "<br>"
        End If
        rs.MoveNext
        sChecked = ""
    Wend
    If sCase = 14 Then
        sTemp = sTemp & "</select>"
    End If
    closeObj(rs)
    getAttributeDetails = sTemp
End Function


'--------------------------------AE CODE BELOW -------------------------------'

Sub ShowGiftWrap(sProdID)
dim gwprice
	gwprice = GetGiftWrapPrice(sProdID)
	if gwprice <> "X" then	
		Response.Write "<br><INPUT name=chkGiftWrap type=checkbox value = 1 >"
		Response.Write "Gift wrap (add " & formatcurrency(gwprice) & " per item)"	
	End if

End Sub 
Sub ShowMTPricesLink(sProdId) 

Dim sql
Dim rst
Dim i

	
	sql = "Select * FROM sfMTPrices WHERE mtprodid= '" & sProdID & "' ORDER By mtIndex ASC"
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText
	If rst.recordcount > 0 then
		Dim spath,stype,jsvar
		sType = 0
		sPath = "MTPrices.asp?sProdId=" & sProdID
		jsvar = "javascript:show_page(" & "'" & sPath & "')"
		
		%><BR>
		<a href="<%=jsvar%>">Check Volume Discounts</a>
		<%
	End if

	CloseObj (rst)

End Sub

Function CheckShowStatus(strProductID)
Dim sql, rst
	
	sql = "Select * FROM sfInventoryInfo WHERE invenProdID= '" & strProductID & "'"
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText	
	
	If rst.recordcount <= 0 then 
		CheckShowStatus = 0
		exit function	
	End If
	
		
	If rst("invenbStatus") = 1 then
		CheckShowStatus = 1
	Else
		CheckShowStatus = 0
	End If
			
	CloseObj (rst)
		
End Function

Function CheckInStock(strProductID)
Dim sql, rst

	sql = "Select Sum(invenInstock) as instock FROM sfInventory WHERE invenProdID= '" & strProductID & "'"
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText	
	If rst.recordcount <= 0 then 
		CheckInStock = "X" 'error
	else
		CheckInStock= rst("Instock")
		
	End if
	CloseObj (rst)
	
End Function

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

Sub ShowProductInventory (strProduct,sType)
Dim ret	,instock, sPath,jsvar

	ret = CheckInventoryTracked (strProduct)
	If ret = 1 then 
		
		If CheckShowStatus(strProduct) <> 1 then exit sub
			
		Instock = CheckInStock(strProduct)	
		Select case instock
			case "X" 'inventory not tracked for this product
				
				
			case 0 'inventory tracked
				If sType = "Dynamic" Then
					Response.Write "<BR> Out of Stock!"
					If checkbackorder(strProduct) = 1 then	
						%><BR> Click "Add to Cart" to BackOrder!<%
					End If
				End If
					
			case else
				sPath = "StockInfo.asp?sProdId=" & strProduct
				jsvar = "javascript:show_page(" & "'" & sPath & "')"
					
				If sType ="Dynamic" then
					%><BR><a href="<%=jsvar%>">In Stock!</a><%	
				eLse
					%><BR><a href="<%=jsvar%>">Stock Information</a><%	
				end If	
								
		End  Select
	end if
	
End sub

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
Function GetGiftWrapPrice (strProductID) 
Dim rst,sql
		
	sql = "Select * FROM sfgiftwraps WHERE gwProdID= '" & strProductID & "'"
	Set rst = Server.CreateObject("ADODB.RecordSet")		
	rst.Open sql, cnn, adOpenStatic, adLockReadOnly, adCmdText	
	
	if rst.recordcount <= 0 then  
		GetGiftWrapPrice = "X"
		CloseObj (rst)
		exit function
	end if
	if rst("gwActivate") = 0 then 
		GetGiftWrapPrice = "X"
		CloseObj (rst)
		exit function
	end if
	
	GetGiftWrapPrice =  rst("gwPrice") 
	CloseObj (rst)
	   
End Function
Private Function GetFullPath(Vdata,justMain) 
Dim sSql ,X
Dim iCatId
Dim sFirst
Dim rsCat,rsSubCat
Dim arrTemp ,bMain
bMain = false
 if left(vData,4)= "none" then
  bMain = True
  arrTemp = split(vdata,"-")
  vdata = arrtemp(1)
 elseif vData = "" then
   GetFullPath = "" 
  exit function
 elseif instr(Vdata,"-") = 0  then
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
  sSql  = "Select catName From sfCategories Where catId =" & rsSubCat("subcatCategoryId")   
 rsCat.Open sSql,cnn,adOpenStatic,adLockReadOnly ,adcmdText
 if justmain = 1 then
    GetFullPath = rsCat("catName")
 else 
   if bMain = True Then
      GetFullPath = rsCat("catName")
   else
     GetFullPath = rsCat("catName") & "-" &  Left(GetFullPath, Len(GetFullPath) - 1)
   end if 
 end if
 Set rsCat = Nothing
 Set rsSubCat = Nothing
 Exit Function
End Function

%>


