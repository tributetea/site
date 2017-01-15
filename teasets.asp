<html>
<head>
<title>TRIBUTE TEA - Teasets</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript">
<!--
function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function MM_nbGroup(event, grpName) { //v3.0
  var i,img,nbArr,args=MM_nbGroup.arguments;
  if (event == "init" && args.length > 2) {
    if ((img = MM_findObj(args[2])) != null && !img.MM_init) {
      img.MM_init = true; img.MM_up = args[3]; img.MM_dn = img.src;
      if ((nbArr = document[grpName]) == null) nbArr = document[grpName] = new Array();
      nbArr[nbArr.length] = img;
      for (i=4; i < args.length-1; i+=2) if ((img = MM_findObj(args[i])) != null) {
        if (!img.MM_up) img.MM_up = img.src;
        img.src = img.MM_dn = args[i+1];
        nbArr[nbArr.length] = img;
    } }
  } else if (event == "over") {
    document.MM_nbOver = nbArr = new Array();
    for (i=1; i < args.length-1; i+=3) if ((img = MM_findObj(args[i])) != null) {
      if (!img.MM_up) img.MM_up = img.src;
      img.src = (img.MM_dn && args[i+2]) ? args[i+2] : args[i+1];
      nbArr[nbArr.length] = img;
    }
  } else if (event == "out" ) {
    for (i=0; i < document.MM_nbOver.length; i++) {
      img = document.MM_nbOver[i]; img.src = (img.MM_dn) ? img.MM_dn : img.MM_up; }
  } else if (event == "down") {
    if ((nbArr = document[grpName]) != null)
      for (i=0; i < nbArr.length; i++) { img=nbArr[i]; img.src = img.MM_up; img.MM_dn = 0; }
    document[grpName] = nbArr = new Array();
    for (i=2; i < args.length-1; i+=2) if ((img = MM_findObj(args[i])) != null) {
      if (!img.MM_up) img.MM_up = img.src;
      img.src = img.MM_dn = args[i+1];
      nbArr[nbArr.length] = img;
  } }
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" onLoad="MM_preloadImages('teaware/canisters/sm_canister_g.jpg','teaware/canisters/med_canister_g.jpg','teaware/canisters/lg_canister_g.jpg','teaware/canisters/xlg_canister_g.jpg','navigation/1teas_g.gif','navigation/1herbs_g.gif','navigation/1teaware_g.gif','navigation/incense_g.gif','navigation/books_music_g.gif','/teaware/teasets/sm_pot_sm_2.jpg','/teaware/teasets/herb_pot_sm2.jpg','/teaware/teasets/gaiwan_set_sm_2.jpg','/teaware/teasets/lg_pot_sm_2.jpg')">
<div align="center">
  <table border="0" cellpadding="0" cellspacing="0" width="614">
    <tr> 
      <td valign="top" width="277" rowspan="2"><a href="index.html" onClick="MM_nbGroup('down','group1','home','',1)" onMouseOver="MM_nbGroup('over','home','navigation/tribute_tea_g.gif','',1)" onMouseOut="MM_nbGroup('out')"><img name="home" src="navigation/tribute_tea.gif" border="0" onLoad="" width="277" height="51"></a></td>
      <td rowspan="3" width="56"><a href="about.asp" onClick="MM_nbGroup('down','group1','logo','',1)" onMouseOver="MM_nbGroup('over','logo','navigation/logo_g.gif','',1)" onMouseOut="MM_nbGroup('out')"><img name="logo" src="navigation/logo.gif" border="0" onLoad="" width="56" height="57"></a></td>
      <td height="50" width="282" valign="top"><a href="teas.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image81','','navigation/1teas_g.gif',1)"><img name="Image81" border="0" src="navigation/1teas.gif" width="73" height="27"></a><a href="herbs.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image9','','navigation/1herbs_g.gif',1)"><img name="Image9" border="0" src="navigation/1herbs.gif" width="85" height="27"></a><a href="teaware.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image10','','navigation/1teaware_g.gif',1)"><img name="Image10" border="0" src="navigation/1teaware.gif" width="123" height="27"></a><a href="incense.asp" onClick="MM_nbGroup('down','group1','incense','',1)" onMouseOver="MM_nbGroup('over','incense','navigation/incense_g.gif','',1)" onMouseOut="MM_nbGroup('out')"><img name="incense" src="navigation/incense.gif" border="0" onLoad="" width="108" height="23"></a><a href="books.asp" onClick="MM_nbGroup('down','group1','booksmusic','',1)" onMouseOver="MM_nbGroup('over','booksmusic','navigation/books_music_g.gif','',1)" onMouseOut="MM_nbGroup('out')"><img name="booksmusic" src="navigation/books_music.gif" border="0" onLoad="" width="173" height="23"></a></td>
      </tr>
    <tr>
      <td height="1"></td>
    </tr>
    <tr> 
      <td height="6"></td>
      <td></td>
    </tr>
  </table>
</div>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
  
<table border="0" cellpadding="0" cellspacing="0" width="614" align="center" height="438">
  <tr> 
    <td height="19" valign="top" colspan="9"> 
      <div align="center"></div>
    </td>
  </tr>
  <tr> 
    <td height="16" width="67"></td>
    <td width="103"></td>
    <td valign="top" colspan="4"> 
      <div align="center"><font size="1"><img src="teaware/teasets/teasets.gif" width="181" height="16"></font></div>
    </td>
    <td width="1"></td>
    <td width="120"></td>
    <td width="53"></td>
  </tr>
  <tr> 
    <td height="134"></td>
    <td></td>
    <td colspan="4" valign="top"> 
      <div align="center"><a href="tea_drops.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('teadrops','','books_music/tea_drops_g.jpg',1)"><br>
        </a><a href="/teasets_gaiwan.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image14','','/teaware/teasets/gaiwan_set_sm_2.jpg',1)"><font size="1"><img name="Image14" border="0" src="/teaware/teasets/gaiwan_set_sm_1.jpg" width="270" height="84"></font></a><br>
        <font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>GAIWAN 
        (COVERED CUP) SET<br>
        </b></font></div>
    </td>
    <td></td>
    <td></td>
    <td></td>
  </tr>
  <tr> 
    <td height="117"></td>
    <td colspan="3" valign="top"> 
      <div align="center"> <img name="Image12" border="0" src="/teaware/teasets/sm_pot_sm_1.jpg" width="225" height="105"><br>
        <font face="Verdana, Arial, Helvetica, sans-serif" size="-2"><b>SMALL 
        POT SET</b></font></div>
    </td>
    <td width="14"></td>
    <td valign="top" colspan="3"> 
      <div align="center"><a href="/teasets_herb.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image13','','/teaware/teasets/herb_pot_sm2.jpg',1)"><img name="Image13" border="0" src="/teaware/teasets/herb_pot_SM1.jpg" width="249" height="105"></a><br>
        <font face="Verdana, Arial, Helvetica, sans-serif" size="-2"><b>HERB POT 
        SET</b></font></div>
    </td>
    <td></td>
  </tr>
  <tr> 
    <td height="122"></td>
    <td></td>
    <td width="1"></td>
    <td valign="top" colspan="4"> 
      <div align="center"> 
        <p><a href="/teasets_large.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image15','','/teaware/teasets/lg_pot_sm_2.jpg',1)"><img name="Image15" border="0" src="/teaware/teasets/lg_pot_sm_1.jpg" width="270" height="110"><br>
          </a><font face="Verdana, Arial, Helvetica, sans-serif" size="-2"><b>LARGE 
          POT SET<br>
          <br>
          </b></font></p>
      </div>
    </td>
    <td></td>
    <td></td>
  </tr>
  <tr> 
    <td height="33"></td>
    <td colspan="7" valign="top"> 
      <div align="center"><font size="1"><a href="<%= Session("DomainPath") %>order.asp"><font size="1"><img src="buttons/view.gif" border="0" width="36" height="33" alt="View Cart/Checkout"></font></a></font></div>
    </td>
    <td></td>
  </tr>
  <tr> 
    <td height="112" colspan="9" valign="top"> 
      <div align="center"><font color="#999999" size="2" face="Arial, Helvetica, sans-serif"><font color="#999999" size="2" face="Arial, Helvetica, sans-serif"><font color="#999999" size="2" face="Arial, Helvetica, sans-serif"><font color="#999999" size="2" face="Arial, Helvetica, sans-serif"><a href="index.html"><font color="#666666">HOME</font></a><font color="#666666"> 
        </font></font></font><font color="#666666">| </font><font color="#999999" size="2" face="Arial, Helvetica, sans-serif"><a href="teas.asp"><font color="#666666">TEAS</font></a><font color="#666666"> 
        </font></font><font color="#666666" size="2">| </font><font color="#999999" size="2" face="Arial, Helvetica, sans-serif"><a href="herbs.asp"><font color="#666666">HERBS</font></a><font color="#666666"> 
        </font></font><font color="#666666" size="2">| </font><font color="#999999" size="2" face="Arial, Helvetica, sans-serif"><a href="teaware.asp"><font color="#666666">TEAWARE</font></a></font></font> 
        <font color="#999999" size="2" face="Arial, Helvetica, sans-serif"><font color="#999999" size="2" face="Arial, Helvetica, sans-serif"><font color="#666666"> 
        </font></font><font color="#666666" size="2">| </font><font color="#999999" size="2" face="Arial, Helvetica, sans-serif"><a href="teaware.asp"><font color="#666666"></font></a></font></font> 
        <font color="#999999" size="2" face="Arial, Helvetica, sans-serif"><a href="incense.asp"><font color="#666666">INCENSE</font></a></font><font color="#666666" size="2"> 
        </font><font color="#999999" size="2" face="Arial, Helvetica, sans-serif"><font color="#666666" size="2">|</font></font><font size="2" color="#666666"> 
        </font><font color="#999999" size="2" face="Arial, Helvetica, sans-serif"><a href="books.asp"><font color="#666666">BOOKS 
        &amp; MUSIC</font></a></font><font color="#666666" size="2"> </font><font color="#999999" size="2" face="Arial, Helvetica, sans-serif"><font color="#666666" size="2">|</font></font><font size="2" color="#666666"> 
        </font><font color="#999999" size="2" face="Arial, Helvetica, sans-serif"><a href="teagifts.asp"><font color="#666666">GIFTS</font></a></font><font color="#666666" size="2"></font><font color="#666666" size="2"><br>
        </font><font color="#999999" size="2" face="Arial, Helvetica, sans-serif"><a href="teaschool.asp"><font color="#666666">TEA 
        SCHOOL</font></a><font color="#666666"> | </font><a href="news.asp"><font color="#666666">NEWS</font></a><font color="#666666"> 
        | </font><a href="about.asp"><font color="#666666">ABOUT TT</font></a><font color="#666666"> 
        | </font><a href="faq.asp"><font color="#666666">Q &amp; A</font></a><font color="#666666"> 
        | </font><a href="wholesale.asp"><font color="#666666">WHOLESALE</font></a><font color="#666666"> 
        | </font><a href="contact.asp"><font color="#666666">CONTACT</font></a><font color="#666666"> 
        | </font><a href="search.asp" target="_blank"><font color="#666666">SEARCH</font></a></font></font></div>
      <p align="center"><font color="#999999" size="2" face="Arial, Helvetica, sans-serif"><font color="#999999" size="2" face="Arial, Helvetica, sans-serif"><font color="#999999" size="1" face="Arial, Helvetica, sans-serif"><a href="privacy_security.asp"><font color="#CC0000">PRIVACY</font></a><font color="#CC0000"> 
        <font color="#000000"><b>|</b></font> </font><font color="#999999" size="2" face="Arial, Helvetica, sans-serif"><font color="#999999" size="1" face="Arial, Helvetica, sans-serif"><a href="privacy_security.asp"><font color="#CC0000">SECURITY</font></a></font></font></font></font></font></p>
      <p align="center"><font color="#666600"><font color="#990000" face="Arial, Helvetica, sans-serif"><b><font size="1" color="#CC0000">&copy; 
        2002 TT. All rights reserved.</font></b></font></font></p>
    </td>
  </tr>
  <tr> 
    <td height="1"></td>
    <td></td>
    <td></td>
    <td width="127"></td>
    <td></td>
    <td width="128"></td>
    <td></td>
    <td></td>
    <td></td>
  </tr>
</table>
</body>
</html>