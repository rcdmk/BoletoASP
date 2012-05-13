<%
Option Explicit

' Importando as classes necessárias
%>
<!--#include file="boletos.class.asp" -->
<!--#include file="real.class.asp" -->
<%

Dim boleto, banco

set boleto = new BoletoASP

Set boleto.Banco = new BancoReal

boleto.DataDocumento = "13/05/2012"
boleto.DataProcessamento = ""
boleto.DataVencimento = "03/09/2004"

boleto.ValorDocumento = "934,23"

boleto.Banco.Carteira = "57"
boleto.Banco.Agencia = "1018"
boleto.Banco.Conta = "0016324"

boleto.NumeroDocumento = "1234"
boleto.NossoNumero = "00005020"
%><html>
<head>
<title></title>
</head>
<body>
<%= boleto.Write() %>
</body>
</html>

