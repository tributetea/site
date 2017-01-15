<%
	
'@BEGINVERSIONINFO

'@APPVERSION: 50.4014.0.2

'@FILENAME: orderform.asp
	 


'@DESCRIPTION: Include File for process_order.asp

'@STARTCOPYRIGHT
'The contents of this file is protected under the United States
'copyright laws and is confidential and proprietary to
'LaGarde, Incorporated.  Its use or disclosure in whole or in part without the
'expressed written permission of LaGarde, Incorporated is expressly prohibited.
'
'(c) Copyright 2000,2001 by LaGarde, Incorporated.  All rights reserved.
'@ENDCOPYRIGHT

'@ENDVERSIONINFO
	
Dim sShipMethods, iStepCounter, iCustID, sCtryList, sStateList, sFormDesign, sCCList, rsProcAdmin, sMailIsActive
   	 
   	 sShipMethods = getShippingList(blnFree)
	iCustID		= Request.Cookies("sfCustomer")("custID")
	sFormDesign	= C_FORMDESIGN
	sCtryList		= getCountryList() 'C_CTRYLIST	
	sStateList		= getStateList() 'C_STATELIST
	
	Set rsProcAdmin = Server.CreateObject("ADODB.Recordset")	
	rsProcAdmin.Open "SELECT adminSubscribeMailIsActive FROM sfAdmin", cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
   sMailIsActive = trim(rsProcAdmin("adminSubscribeMailIsActive"))
	closeObj(rsProcAdmin)

	If bLoggedIn AND iCustID <> "" Then
		Dim rsGetCustDetails, rsGetCustShipDetails, sCustFirstName, sCustMiddleInitial, sCustLastName, sCustCompany, sCustAddress1, sCustAddress2
		Dim sCustCity, sCustState, sCustStateName, sCustZip, sCustCountry, sCustCountryName, sCustPhone, sCustFax, sCustEmail, sCustCardType, sCustCardTypeName, sCustSubscribed, sShipCustFirstName, sShipCustMiddleInitial,sShipCustLastName
		Dim sShipCustAddress1, sShipCustAddress2, sShipCustCity, sShipCustState, sShipCustStateName, sShipCustZip, sShipCustCountry,sShipCustCountryName, sShipCustPhone
		Dim sShipCustFax, sShipCustCompany, sShipCustEmail,sSubmitAction
	   
		' Get RecordSet of customer details and shipping details and credit card details
		Set rsGetCustDetails		= getCustomerRow(iCustID)
		Set rsGetCustShipDetails	= getCustomerShippingRow(iCustID)
		'Set rsGetCustCardDetails	= getCustomerCardRow(iCustID)
				
		' Collect billing address
		sCustFirstName		= rsGetCustDetails.Fields("custFirstName")
		sCustMiddleInitial	= rsGetCustDetails.Fields("custMiddleInitial")
		sCustLastName		= rsGetCustDetails.Fields("custLastName")
		sCustCompany		= rsGetCustDetails.Fields("custCompany")
		sCustAddress1		= rsGetCustDetails.Fields("custAddr1")
		sCustAddress2		= rsGetCustDetails.Fields("custAddr2")	   
		sCustCity			= rsGetCustDetails.Fields("custCity")
		sCustState			= rsGetCustDetails.Fields("custState")		
		sCustStateName		= getNameWithID("sfLocalesState",sCustState,"loclstAbbreviation","loclstName",1)		
		sCustZip			= rsGetCustDetails.Fields("custZip")
		sCustCountry		= rsGetCustDetails.Fields("custCountry")
		sCustCountryName	= getNameWithID("sfLocalesCountry",sCustCountry,"loclctryAbbreviation","loclctryName",1)	
		sCustPhone			= rsGetCustDetails.Fields("custPhone")
		sCustFax			= rsGetCustDetails.Fields("custFax")
		sCustEmail			= rsGetCustDetails.Fields("custEmail")
		sCustSubscribed		= rsGetCustDetails.Fields("custIsSubscribed")
	   
	    ' Change display for saved cart customers
	    If instr(1,sCustFirstName,"Saved Cart Customer",1) Then
			sCustFirstName = ""
		End If	 
	   
		' Get Ship Address
		If Not rsGetCustShipDetails.EOF Then
			sShipCustFirstName		= rsGetCustShipDetails.Fields("cshpaddrShipFirstName")
			sShipCustMiddleInitial	= rsGetCustShipDetails.Fields("cshpaddrShipMiddleInitial")
			sShipCustLastName		= rsGetCustShipDetails.Fields("cshpaddrShipLastName")
			sShipCustCompany		= rsGetCustShipDetails.Fields("cshpaddrShipCompany")
			sShipCustAddress1		= rsGetCustShipDetails.Fields("cshpaddrShipAddr1")
			sShipCustAddress2		= rsGetCustShipDetails.Fields("cshpaddrShipAddr2")	   
			sShipCustCity			= rsGetCustShipDetails.Fields("cshpaddrShipCity")
			sShipCustState			= rsGetCustShipDetails.Fields("cshpaddrShipState")
			sShipCustStateName		= getNameWithID("sfLocalesState",sShipCustState,"loclstAbbreviation","loclstName",1)	
			sShipCustZip			= rsGetCustShipDetails.Fields("cshpaddrShipZip")
			sShipCustCountry		= rsGetCustShipDetails.Fields("cshpaddrShipCountry")
			sShipCustCountryName	= getNameWithID("sfLocalesCountry",sShipCustCountry,"loclctryAbbreviation","loclctryName",1)
			sShipCustPhone			= rsGetCustShipDetails.Fields("cshpaddrShipPhone")
			sShipCustFax			= rsGetCustShipDetails.Fields("cshpaddrShipFax")
			sShipCustEmail			= rsGetCustShipDetails.Fields("cshpaddrShipEmail")
		Else
			sShipCustFirstName		= sCustFirstName
			sShipCustMiddleInitial	= sCustMiddleInitial
			sShipCustLastName		= sCustLastName
			sShipCustCompany		= sCustCompany
			sShipCustAddress1		= sCustAddress1
			sShipCustAddress2		= sCustAddress2
			sShipCustCity			= sCustCity
			sShipCustState			= sCustState
			sShipCustStateName		= sCustStateName
			sShipCustZip			= sCustZip
			sShipCustCountry		= sCustCountry
			sShipCustCountryName	= sCustCountryName
			sShipCustPhone			= sCustPhone
			sShipCustFax			= sCustFax
			sShipCustEmail			= sCustEmail
		End If 
		
	' End logged in if	
	End If
	
	' Cleanup
	closeobj(rsGetCustDetails)
	closeobj(rsGetCustShipDetails)


	' Used for iterating steps -- useful if some step is skipped
	iStepCounter = 0
	Function getStepCounter(iStepCounter)
		iStepCounter = iStepCounter + 1
		getStepCounter = iStepCounter
	End Function
	
	sSubmitAction = ""
	If NOT (bLoggedIn) OR iCustID = "" Then sSubmitAction = "this.Password.password=true;this.Password.optional = true;this.Password2.optional = true;"
	sSubmitAction = sSubmitAction & "this.Company.optional = true;this.Address2.optional = true;this.Fax.optional = true;this.Address2.optional = true;this.Instructions.optional = true;this.Email.eMail = true;this.Phone.phoneNumber = true;this.ShipState.optional = true;this.ShipCountry.optional = true;this.MiddleInitial.optional = true;"		
	If sPaymentMethod = "Credit Card" Then
		sSubmitAction = sSubmitAction & "this.CardNumber.creditCardNumber = true;this.CardExpiryMonth.creditCardExpMonth = true;this.CardExpiryYear.creditCardExpYear = true;return validate_Me(this);"
	Elseif sPaymentMethod = "PhoneFax" Then 
		sSubmitAction = sSubmitAction & "this.CardType.optional = true;this.CardName.special = true;this.CardNumber.special = true;this.CardExpiryMonth.special = true;this.CardExpiryYear.special = true;this.CheckNumber.special = true;this.BankName.special = true;this.RoutingNumber.special = true;this.CheckingAccountNumber.special = true;this.POName.special = true;this.PONumber.special = true;return validate_Me(this);"
	Else
		sSubmitAction = sSubmitAction & "return validate_Me(this);"
	End If
%>
<script language="JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>

	  <form method="post" action="verify.asp" name="form1" onSubmit="<%= sSubmitAction %>">
    
  <table class="tdContent2" border="0" width="600" cellpadding="2" cellspacing="0" height="1042">
    <tr> 
      <td width="3" height="2"></td>
      <td width="583"></td>
      <td width="2"></td>
    </tr>
    <tr> 
      <td colspan="3" height="405"> 
        <table class="tdContent2" border="0" width="100%" cellpadding="2" cellspacing="0" height="371">
          <tr> 
            <td height="18" valign="top" colspan="3"> 
              <div align="center"><font size="1"><img src="bill.gif" width="210" height="14"></font></div>
            </td>
          </tr>
          <tr> 
            <td nowrap align="right" height="32" width="183"><b><font color="#990000">*First 
              Name:</font></b></td>
            <td valign="top" nowrap width="396"> 
              <input type="text" maxlength="25" name="FirstName" title="First Name" size="20" Style="<%= sFormDesign%>" value="<%= sCustFirstName %>">
              &nbsp;&nbsp;<b>MI:&nbsp;</b> 
              <input type="text" name="MiddleInitial" size="1" Style="<%= sFormDesign%>" value="<%= sCustMiddleInitial %>" maxlength="1">
            </td>
            <td width="5"></td>
          </tr>
          <tr> 
            <td nowrap align="right" height="32" width="183"><b><font color="#990000">*Last 
              Name:</font></b></td>
            <td valign="top" nowrap width="396"> 
              <input type="text" maxlength="25" name="LastName" title="Last Name" size="20" Style="<%= sFormDesign%>" value="<%= sCustLastName %>">
            </td>
            <td width="5"></td>
          </tr>
          <tr> 
            <td nowrap align="right" height="32" width="183"><b><font face="Arial, Helvetica, sans-serif">How 
              did you find us?</font></b></td>
            <td valign="center" nowrap width="396"> 
              <select name="Company" title="Company" style="<%= sFormDesign%>">
                <option value="<%= sCustCompany %>"><%= sCustCompany %></option>
                <option selected>SELECT</option>
                <option>Search Engine</option>
                <option>Referal</option>
                <option>Tea Time</option>
				<option>J.A.M.A.</option>
                <option>Bust</option>
				<option>Yoga</option>
                <option>Body Sense</option>
                <option>Dwell</option>
                <option>Lucky</option>
                <option>Other</option>
              </select>
            </td>
            <td width="5"></td>
          </tr>
          <tr> 
            <td nowrap align="right" height="32" width="183"><b><font color="#990000">*Address 
              1:</font></b></td>
            <td valign="top" nowrap width="396"> 
              <input type="text" Style="<%= sFormDesign%>" maxlength="25" name="Address1" title="Street Address" size="20" value="<%= sCustAddress1 %>">
            </td>
            <td width="5"></td>
          </tr>
          <tr> 
            <td nowrap align="right" height="32" width="183"><b>Address 2:</b></td>
            <td valign="top" nowrap width="396"> 
              <input type="text" Style="<%= sFormDesign%>" maxlength="25" name="Address2" title="Address2" size="20" value="<%= sCustAddress2 %>">
            </td>
            <td width="5"></td>
          </tr>
          <tr> 
            <td nowrap align="right" height="32" width="183"><b><font color="#990000">*City:</font></b></td>
            <td valign="top" nowrap width="396"> 
              <input type="text" maxlength="25" name="City" title="City" size="20" Style="<%= sFormDesign%>" value="<%= sCustCity %>">
            </td>
            <td width="5"></td>
          </tr>
          <tr> 
            <td nowrap align="right" height="32" width="183"><b><font color="#990000">*State/Prov:</font></b></td>
            <td valign="center" nowrap width="396"> 
              <select name="State" title="State" style="<%= sFormDesign%>">
                <option value="<%= sCustState %>"><%= sCustStateName %></option>
                <%= sStateList %> 
              </select>
              <br>
              <b><font size="1" face="Arial, Helvetica, sans-serif">OUTSIDE USA 
              &amp; CAN SELECT &quot;INTERNATIONAL&quot;</font></b> </td>
            <td width="5"></td>
          </tr>
          <tr> 
            <td nowrap align="right" height="32" width="183"><b><font color="#990000">*Zip 
              Code:</font></b></td>
            <td valign="top" nowrap width="396"> 
              <input type="text" maxlength="25" name="Zip" title="Zip Code" size="5" Style="<%= sFormDesign%>" value="<%= sCustZip %>">
            </td>
            <td width="5"></td>
          </tr>
          <tr> 
            <td nowrap align="right" height="32" width="183"><b><font color="#990000">*Country:</font></b></td>
            <td valign="center" nowrap width="396"> 
              <select name="Country" title="Country" style="<%= sFormDesign%>">
                <option value="<%= sCustCountry %>"><%= sCustCountryName %></option>
                <%= sCtryList %> 
              </select>
            </td>
            <td width="5"></td>
          </tr>
          <tr> 
            <td nowrap align="right" height="32" width="183"><b><font color="#990000">*Phone:</font></b></td>
            <td valign="top" nowrap width="396"> 
              <input type="text" name="Phone" maxlength="25" title="Phone Number" size="20" Style="<%= sFormDesign%>" value="<%= sCustPhone %>">
              <b><font size="1" face="Arial, Helvetica, sans-serif"> &nbsp;&nbsp;</font></b> 
              <font face="Arial, Helvetica, sans-serif"><a href="#"><font size="-1" onClick="MM_openBrWindow('info.htm','','scrollbars=yes,width=425,height=400')"><i><font size="1" color="#336699">WHY 
              WE REQUEST YOUR NUMBER</font></i></font></a></font></td>
            <td width="5"></td>
          </tr>
          <tr> 
            <td nowrap align="right" height="32" width="183"><b>Fax:</b></td>
            <td valign="top" nowrap width="396"> 
              <input type="text" name="Fax" maxlength="25" title="Fax Number" size="20" Style="<%= sFormDesign%>" value="<%= sCustFax %>">
            </td>
            <td width="5"></td>
          </tr>
          <tr> 
            <td nowrap align="right" height="26" width="183"><b><font color="#990000">*e-mail:</font></b></td>
            <td valign="top" nowrap height="26" width="396"> 
              <input type="text" Style="<%= sFormDesign%>" maxlength="50" name="Email" title="Email Address" size="20" value="<%= sCustEmail %>">
            </td>
            <td height="26" width="5"></td>
          </tr>
          <tr> 
            <td height="2" width="183"></td>
            <td width="396" height="2"></td>
            <td width="5" height="2"></td>
          </tr>
        </table>
      </td>
    </tr>
    <!-- Shipping Info -->
    <tr> 
      <td height="2"></td>
      <td height="2"></td>
      <td height="2"></td>
    </tr>
    <tr>
      <td height="429"></td>
      <td colspan="2" valign="top" class="tdContent2"> 
        <table class="tdContent2" cellpadding="2" cellspacing="0" width="100%" height="429">
          <tr> 
            <td height="2" width="175"></td>
            <td height="2" width="110"></td>
            <td height="2" width="283"></td>
            <td height="2" width="4"></td>
          </tr>
          <tr>
            <td height="25" colspan="4" valign="top"> 
              <div align="center"><img src="ship.gif" width="216" height="14"><br>
                (if different from above)</div>
            </td>
            </tr>
          <tr> 
            <td nowrap align="right" height="32" width="175"><b>First Name:</b></td>
            <td valign="top" nowrap colspan="2"> 
              <input type="text" maxlength="25" name="ShipFirstName" size="20" Style="<%= sFormDesign%>" value="<%= sShipCustFirstName %>">
              &nbsp;&nbsp;<b>MI:&nbsp;</b> 
              <input type="text" name="ShipMiddleInitial" size="1" Style="<%= sFormDesign%>" value="<%= sShipCustMiddleInitial %>" maxlength="1">
            </td>
            <td width="4"></td>
          </tr>
          <tr> 
            <td nowrap align="right" height="32" width="175"> 
              <div align="right"><b>Last Name:</b></div>
            </td>
            <td valign="top" colspan="2"> 
              <input type="text" maxlength="25" name="ShipLastName" size="20" style="<%= sFormDesign%>" value="<%= sShipCustLastName %>">
            </td>
            <td width="4"></td>
          </tr>
          <tr> 
            <td nowrap align="right" height="32" width="175"><b>Company:</b></td>
            <td valign="top" nowrap colspan="2"> 
              <input type="text" Style="<%= sFormDesign%>" maxlength="25" name="ShipCompany" size="20" value="<%= sShipCustCompany %>">
            </td>
            <td width="4"></td>
          </tr>
          <tr> 
            <td nowrap align="right" height="32" width="175"><b>Address:</b></td>
            <td valign="top" nowrap colspan="2"> 
              <input type="text" Style="<%= sFormDesign%>" maxlength="25" name="ShipAddress1" size="20" value="<%= sShipCustAddress1%>">
            </td>
            <td width="4"></td>
          </tr>
          <tr> 
            <td nowrap align="right" height="32" width="175"> 
              <div align="right"><b>Address 2:</b></div>
            </td>
            <td valign="top" colspan="2"> 
              <input type="text" style="<%= sFormDesign%>" maxlength="25" name="ShipAddress2" size="20" value="<%= sShipCustAddress2 %>">
            </td>
            <td width="4"></td>
          </tr>
          <tr> 
            <td nowrap align="right" height="32" width="175"><b>City:</b></td>
            <td valign="top" nowrap colspan="2"> 
              <input type="text" maxlength="25" name="ShipCity" size="20" Style="<%= sFormDesign%>" value="<%= sShipCustCity %>">
            </td>
            <td width="4"></td>
          </tr>
          <tr> 
            <td nowrap align="right" height="32" width="175"> 
              <div align="right"><b>State/Prov:</b></div>
            </td>
            <td valign="center" colspan="2"> 
              <select size="1" name="ShipState" style="<%= sFormDesign %>">
                <option value="<%= sShipCustState %>"><%= sShipCustStateName %></option>
                <%= sStateList %> 
              </select>
            </td>
            <td width="4"></td>
          </tr>
          <tr> 
            <td nowrap align="right" height="32" width="175"><b>Zip Code:</b></td>
            <td valign="center" nowrap colspan="2"> 
              <input type="text" maxlength="25" name="ShipZip" size="5" Style="<%= sFormDesign%>" value="<%= sShipCustZip %>">
            </td>
            <td width="4"></td>
          </tr>
          <tr> 
            <td nowrap align="right" height="32" width="175"> 
              <div align="right"><b>Country:</b></div>
            </td>
            <td valign="center" colspan="2"> 
              <select size="1" name="ShipCountry" style="<%= sFormDesign%>">
                <option value="<%= sShipCustCountry %>"><%= sShipCustCountryName %></option>
                <%= sCtryList %> 
              </select>
            </td>
            <td width="4"></td>
          </tr>
          <tr> 
            <td nowrap align="right" height="32" width="175"><b>Phone:</b></td>
            <td valign="top" nowrap colspan="2"> 
              <input type="text" maxlength="20" name="ShipPhone" size="20" Style="<%= sFormDesign%>" value="<%= sShipCustPhone %>">
            </td>
            <td width="4"></td>
          </tr>
          <tr> 
            <td nowrap align="right" height="32" width="175"><b>Fax:</b></td>
            <td valign="top" nowrap colspan="2"> 
              <input type="text" maxlength="20" name="ShipFax" size="20" Style="<%= sFormDesign%>" value="<%= sShipCustFax %>">
            </td>
            <td width="4"></td>
          </tr>
          <tr> 
            <td nowrap align="right" height="39" width="175"> 
              <div align="right"> <b>e-mail:</b></div>
            </td>
            <td width="110" valign="top"> 
              <input type="text" style="<%= sFormDesign%>" maxlength="50" name="ShipEmail" size="20" value="<%= sShipCustEmail %>">
            </td>
            <td width="283"><img src="buttons/clear.gif" onClick="javascript:clearShipping(form1);" width="92" height="22" align="absmiddle"> 
              <b><font size="1" face="Arial, Helvetica, sans-serif">CLEAR SHIPPING 
              FORM</font></b> </td>
            <td width="4"></td>
          </tr>
        </table>
      </td>
      </tr>
    <tr>
      <td height="4"></td>
      <td></td>
      <td></td>
    </tr>
    <% if iShip <> 0 Then 
       If sShipMethods <> "" Then %>
    <tr> 
      <td class="tdContent2" height="69" colspan="2"> 
        <table class="tdContent2" cellpadding="0" cellspacing="0" width="100%" height="45">
          <tr> 
            <td align="center" height="42"><img src="buttons/ship.gif" width="91" height="14"><br>
              <img src="usps.gif" width="156" height="33" align="absmiddle"> 
              <select name="Shipping" style="<%= sFormDesign %>">
                <option selected>SELECT SERVICE</option>
                <%= sShipMethods%> 
              </select>
              <br>
              <font face="Arial, Helvetica, sans-serif">Select &quot;International 
              Mail&quot; for orders shipping outside the U.S.</font></td>
          </tr>
        </table>
        <%  End If 
     End If %>
        <!-- Payment Method -->
      <td height="69"></td>
    
    <tr> 
      <td colspan="3" height="53"> 
        <table border="0" width="100%" cellspacing="0" cellpadding="2" class="tdContent2">
          <tr> 
            <td width="592"  height="48"> 
              <div align="center"><font size="1"><br>
                <img src="pay.gif" width="100" height="14" align="absmiddle"> 
                </font> 
                <select style="<%= C_FORMDESIGN %>" name="PaymentMethod">
                  <%= sPaymentList %> 
                </select>
                <br>
              </div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr> 
      <td height="2"></td>
      <td></td>
      <td></td>
    </tr>
    <tr> 
      <td class="tdContent2" height="99" colspan="2"> 
        <table class="tdContent2" cellpadding="0" cellspacing="0" width="100%" height="80">
          <tr> 
            <td width="100%" height="2"></td>
          </tr>
          <tr> 
            <td width="100%" class="tdContent2" align="center" height="88"> <img src="spec.gif" width="310" height="14"><br>
              <textarea rows="2" name="Instructions" cols="50" style="<%= sFormDesign%>"></textarea>
              <br>
              If you have any comments or suggestions, please enter them here.
          
        </table>
      </td>
      <td height="99"></td>
    </tr>
    <% If NOT (bLoggedIn) Then %>
    <tr> 
      <td class="tdContent2" height="143" colspan="2"> 
        <table class="tdContent2" cellpadding="0" cellspacing="0" width="100%" height="139">
          <tr> 
            <td width="587" height="16"> 
              <table border="0" width="100%" cellspacing="0" cellpadding="2">
                <tr> 
                  <td colspan="2" > 
                    <div align="center"><img src="pass.gif" width="216" height="14"></div>
                  </td>
                </tr>
                <tr> 
                  <td height="2" width="585"></td>
                  <td height="2"></td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td class="tdContent2" height="113"> 
              <div align="center">In order to serve you better, an account will 
                be created for you as part of the checkout process.<br>
                This will facilitate a speedier checkout for future orders.<br>
                To specify a password, please enter it below. Otherwise, one will 
                autpmatically be generated.</div>
              <table border="0" width="100%" class="tdContent2" align="center" height="56">
                <tr> 
                  <td width="50%" align="right" height="26"><b>Password:</b></td>
                  <td width="50%" height="26"> 
                    <input type="password" name="Password" maxlength="10" title="Password" style="<%= sFormDesign%>">
                  </td>
                </tr>
                <tr> 
                  <td width="50%" align="right" height="26"><b>Password Confirmation:</b></td>
                  <td width="50%" height="26"> 
                    <input type="password" name="Password2" maxlength="10" title="Password Confirmation" style="<%= sFormDesign%>">
                  </td>
                </tr>
              </table>
        </table>
      </td>
      <td></td>
    </tr>
    <% 
	End If 
	closeObj(cnn)
	%>
    <tr> 
      <td height="2"></td>
      <td></td>
      <td></td>
    </tr>
    <tr align="center"> 
      <td height="36" colspan="2"> 
        <input type="image" src="buttons/continue.gif" name="Verify" border="0" width="92" height="22">
        <font size="1"><br>
        </font> </td>
      <td height="36"></td>
    </tr>
    <tr> 
      <td height="2"></td>
      <td height="2"></td>
      <td height="2"><img height="1" width="2" src="/transparent.gif"></td>
    </tr>
  </table>

				 
  <div align="center">
    <input type="hidden" name="FreeShip" value="<%= blnFree %>">
    <input type="hidden" name="TotalPrice" value="<%= sTotalPrice %>">
    <input type="hidden" name="bShip" value="<%= iShip %>">
  </div>
</form>