<%@ LANGUAGE="VBScript" %>
<% '<!--#include file="IASUtil.asp"--> %>
<%

	Set PNPObj = CreateObject("pnpcom.main")
	'Response.Write pnpObj.GetTime()
	' the url is used here for testing purposes
	Results = PNPObj.doTransaction("http://pay1.plugnpay.com/payment/pnpremote.cgi","item1=abc&quantity1=2","pnpdemo","pnpdemo")
	response.write "<br>" & Results
	pnpObj.split(REsults)
	Results = PNPObj.getvalue(10)
	'Results = PNPObj.doTransactionWinInet("http://pay1.plugnpay.com/payment/pnpremotetestlink.cgi","item1=abc&quantity1=2","","")
	
	'Results = PNPObj.doTransaction("http://pay1.plugnpay.com/payment/pnpremotetestlink.cgi","item1=abc&quantity1=2","pnpdemo","pnpdemo")
	response.write "<br>" & Results

%>

<html>
<head>
<meta HTTP-EQUIV="Refresh" CONTENT="1000; URL=05_receipt_2.asp">
<title>-</title>
</head>
<p>
<%
  response.write "bbbb"
  'response.write "Results are " & Results
%>
<p>&nbsp;

<p>&nbsp;

<p>&nbsp;

<div align="center">
<center>
<table border="1" width="425" cellpadding="5">
  <tr>
    <td width="100%" bgcolor="#010776" align="center" valign="middle"><p align="center"><strong><font face="Verdana" size="2" color="#FFFFFF"><b>We are
    processing your order . . .</b></font></strong></p>
    </td>
  </tr>
</table>
  </center>
</div>
</html>

<%	'Connection.Close %>
