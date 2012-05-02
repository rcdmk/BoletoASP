<%
Option Explicit

' Importando as classes necessárias
%><!--#include file="boletos.class.asp" --><%

Dim boleto, banco

set boleto = new BoletoASP

boleto.Banco = new BancoItau

boleto.DataDocumento = "28/03/2000"
boleto.DataProcessamento = ""
boleto.DataVencimento = "01/05/2002"

boleto.ValorDocumento = "123,45"

boleto.Banco.Carteira = "110"
boleto.Banco.Agencia = "0057"
boleto.Banco.Conta = "12345"
boleto.Banco.ContaDV = "7"

boleto.NumeroDocumento = "1234567890"
boleto.NossoNumero = "12345678"
%><html>
<head>
<title></title>
</head>
<body>
<%= boleto.Write() %>
</body>
</html>

