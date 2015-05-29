<%
Option Explicit

' Importando as classes necessárias
%>
<!--#include file="src/boleto.class.asp" -->
<!--#include file="src/santander.class.asp" -->
<%

Dim boleto, banco, sacado

set boleto = new BoletoASP

Set boleto.Banco = new BancoSantander

' informações sobre a cobrança
boleto.DataDocumento = "08/05/2012"
boleto.DataProcessamento = "08/05/2012"
boleto.DataVencimento = "10/05/2012"

boleto.ValorDocumento = 66
boleto.PercJuros = 0.16

boleto.Banco.Carteira = "101"
boleto.Banco.Agencia = "3719-2" ' com DV
boleto.Banco.Conta = "3782913"

boleto.NumeroDocumento = "64-12528"
boleto.NossoNumero = "900000000754" ' sem DV

boleto.Instrucoes = "Cobrar Mora diária de R$ " & formatNumber(boleto.ValorJuros, 2)


' quem recebe
boleto.CedenteNome = "Nome do Cedente Ltda. - CNPJ: 12.345.789/0001-23"

' quem paga
set sacado = new SacadoASP

sacado.Nome = "Nome do sacado"
sacado.Endereco = "Endereço do sacado, 123 - bloco 1 apto 123"
sacado.Bairro = "Bairro"
sacado.CEP = "01234-567"
sacado.Cidade = "São Paulo"
sacado.Estado = "SP"
sacado.CPF = "123.456.789-10"

set boleto.Sacado = sacado

%><html>
<head>
<title></title>
</head>
<body>
<%= boleto.Write() %>
</body>
</html>

