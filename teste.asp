<%
Option Explicit

' Importando as classes necessárias
%>
<!--#include file="boletos.class.asp" -->
<!--#include file="santander.class.asp" -->
<%

Dim boleto, banco

set boleto = new BoletoASP

Set boleto.Banco = new BancoSantander

boleto.DataDocumento = "08/05/2012"
boleto.DataProcessamento = "08/05/2012"
boleto.DataVencimento = "10/05/2012"

boleto.ValorDocumento = "66,00"
boleto.PercJuros = 0.16

boleto.Banco.Carteira = "101"
boleto.Banco.Agencia = "3719-2" '2
boleto.Banco.Conta = "3782913"

boleto.NumeroDocumento = "64-12528"
boleto.NossoNumero = "900000000754" '3


boleto.Instrucoes = "Cobrar Mora diária de R$ " & formatNumber(boleto.ValorJuros, 2)
%><html>
<head>
<title></title>
</head>
<body>
<%= boleto.Write() %>
</body>
</html>

