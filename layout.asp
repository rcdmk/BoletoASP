<html>
<head>
<title></title>
<style type="text/css">
td.BoletoCodigoBanco {
	font-size: 6mm;
	font-family: arial, verdana;
	font-weight : bold;
	font-style: italic;
	text-align: center;
	vertical-align: bottom;
	border-bottom: 0.15mm solid #000000;
	border-right: 0.15mm solid #000000;
	padding-bottom : 1mm
}
td.BoletoLogo {
	border-bottom: 0.15mm solid #000000;
	border-right: 0.15mm solid #000000;
	text-align: center;
	height: 10mm
}
td.BoletoLinhaDigitavel {
	font-size: 4 mm;
	font-family: arial, verdana;
	font-weight : bold;
	text-align: center;
	vertical-align: bottom;
	border-bottom: 0.15mm solid #000000;
	padding-bottom : 1mm;
}
td.BoletoTituloEsquerdo {
	font-size: 0.2cm;
	font-family: arial, verdana;
	padding-left : 0.15mm;
	border-right: 0.15mm solid #000000;
	text-align: left
}
td.BoletoTituloDireito {
	font-size: 2mm;
	font-family: arial, verdana;
	padding-left : 0.15mm;
	text-align: left
}
td.BoletoValorEsquerdo {
	font-size: 3mm;
	font-family: arial, verdana;
	text-align: center;
	border-right: 0.15mm solid #000000;
	font-weight: bold;
	border-bottom: 0.15mm solid #000000;
	padding-top: 0.5mm
}
td.BoletoValorDireito {
	font-size: 3mm;
	font-family: arial, verdana;
	text-align:right;
	padding-right: 3mm;
	padding-top: 0.8mm;
	border-bottom: 0.15mm solid #000000;
	font-weight: bold;
}
td.BoletoTituloSacado {
	font-size: 2mm;
	font-family: arial, verdana;
	padding-left : 0.15mm;
	vertical-align: top;
	padding-top : 0.15mm;
	text-align: left
}
td.BoletoValorSacado {
	font-size: 3mm;
	font-family: arial, verdana;
	font-weight: bold;
	text-align : left
}
td.BoletoTituloSacador {
	font-size: 2mm;
	font-family: arial, verdana;
	padding-left : 0.15mm;
	vertical-align: bottom;
	padding-bottom : 0.8mm;
	border-bottom: 0.15mm solid #000000
}
td.BoletoValorSacador {
	font-size: 3mm;
	font-family: arial, verdana;
	vertical-align: bottom;
	padding-bottom : 0.15mm;
	border-bottom: 0.15mm solid #000000;
	font-weight: bold;
	text-align: left
}
td.BoletoPontilhado {
	border-top: 0.3mm dashed #000000;
	font-size: 1mm
}
ul.BoletoInstrucoes {
	font-size : 3mm;
	font-family : verdana, arial
}
</style>
</head>
<body>
<table cellspacing="0" cellpadding="0" border="0" class="Boleto">
	<tr>
		<td style="width: 0.9cm"></td>
		<td style="width: 1cm"></td>
		<td style="width: 1.9cm"></td>
		<td style="width: 0.5cm"></td>
		<td style="width: 1.3cm"></td>
		<td style="width: 0.8cm"></td>
		<td style="width: 1cm"></td>
		<td style="width: 1.9cm"></td>
		<td style="width: 1.9cm"></td>
		<td style="width: 3.8cm"></td>
		<td style="width: 3.8cm"></td>
	</tr>
	<tr>
		<td colspan="11"><ul class="BoletoInstrucoes">
				<li>Imprima em papel A4 ou Carta</li>
				<li>Utilize margens mínimas a direita e a esquerda</li>
				<li>Recorte na linha pontilhada</li>
				<li>Não rasure o código de barras</li>
			</ul></td>
	</tr>
	<tr>
		<td colspan="11" class="BoletoPontilhado">&nbsp;</td>
	</tr>
	<tr>
		<td colspan="4" class="BoletoLogo"><img src="imagens/<%= bancoNumero %>.jpg"></td>
		<td colspan="2" class="BoletoCodigoBanco"><%= bancoNumero %>-<%= bancoDigito %></td>
		<td colspan="6" class="BoletoLinhaDigitavel"><%= linhaDigitavel %></td>
	</tr>
	<tr>
		<td colspan="10" class="BoletoTituloEsquerdo">Local de Pagamento</td>
		<td class="BoletoTituloDireito">Vencimento</td>
	</tr>
	<tr>
		<td colspan="10" class="BoletoValorEsquerdo" style="text-align: left; padding-left : 0.1cm"><%= LocalDePagamento %></td>
		<td class="BoletoValorDireito"><%= dataVencimento %></td>
	</tr>
	<tr>
		<td colspan="10" class="BoletoTituloEsquerdo">Cedente</td>
		<td class="BoletoTituloDireito">Agência/Código do Cedente</td>
	</tr>
	<tr>
		<td colspan="10" class="BoletoValorEsquerdo" style="text-align: left; padding-left : 0.1cm"><%= cedenteNome %></td>
		<td class="BoletoValorDireito"><%= bancoAgencia %>/<%= cedenteCodigo %></td>
	</tr>
	<tr>
		<td colspan="3" class="BoletoTituloEsquerdo">Data do Documento</td>
		<td colspan="4" class="BoletoTituloEsquerdo">Número do Documento</td>
		<td class="BoletoTituloEsquerdo">Espécie</td>
		<td class="BoletoTituloEsquerdo">Aceite</td>
		<td class="BoletoTituloEsquerdo">Data do Processamento</td>
		<td class="BoletoTituloDireito">Nosso Número</td>
	</tr>
	<tr>
		<td colspan="3" class="BoletoValorEsquerdo"><%= DataDocumento %></td>
		<td colspan="4" class="BoletoValorEsquerdo"><%= NumeroDocumento %></td>
		<td class="BoletoValorEsquerdo"><%= especie %></td>
		<td class="BoletoValorEsquerdo"><%= aceite %></td>
		<td class="BoletoValorEsquerdo"><%= DataProcessamento %></td>
		<td class="BoletoValorDireito"><%= NossoNumero %></td>
	</tr>
	<tr>
		<td colspan="3" class="BoletoTituloEsquerdo">Uso do Banco</td>
		<td colspan="2" class="BoletoTituloEsquerdo">Carteira</td>
		<td colspan="2" class="BoletoTituloEsquerdo">Moeda</td>
		<td colspan="2" class="BoletoTituloEsquerdo">Quantidade</td>
		<td class="BoletoTituloEsquerdo">(x) Valor</td>
		<td class="BoletoTituloDireito">(=) Valor do Documento</td>
	</tr>
	<tr>
		<td colspan="3" class="BoletoValorEsquerdo">&nbsp;</td>
		<td colspan="2" class="BoletoValorEsquerdo"><%= bancoCarteira %></td>
		<td colspan="2" class="BoletoValorEsquerdo">R$</td>
		<td colspan="2" class="BoletoValorEsquerdo">&nbsp;</td>
		<td class="BoletoValorEsquerdo">&nbsp;</td>
		<td class="BoletoValorDireito"><%= ValorDocumento %></td>
	</tr>
	<tr>
		<td colspan="10" class="BoletoTituloEsquerdo">Instruco</td>
		<td class="BoletoTituloDireito">(-) Desconto</td>
	</tr>
	<tr>
		<td colspan="10" rowspan="9" class="BoletoValorEsquerdo" style="text-align: left; vertical-align:top; padding-left : 0.1cm"><%= Instrucoes %></td>
		<td class="BoletoValorDireito">&nbsp;</td>
	</tr>
	<tr>
		<td class="BoletoTituloDireito">(-) Outras Deduções/Abatimento</td>
	</tr>
	<tr>
		<td class="BoletoValorDireito">&nbsp;</td>
	</tr>
	<tr>
		<td class="BoletoTituloDireito">(+) Mora/Multa/Juros</td>
	</tr>
	<tr>
		<td class="BoletoValorDireito">&nbsp;</td>
	</tr>
	<tr>
		<td class="BoletoTituloDireito">(+) Outros Acréscimos</td>
	</tr>
	<tr>
		<td class="BoletoValorDireito">&nbsp;</td>
	</tr>
	<tr>
		<td class="BoletoTituloDireito">(=) Valor Cobrado</td>
	</tr>
	<tr>
		<td class="BoletoValorDireito">&nbsp;</td>
	</tr>
	<tr>
		<td rowspan="3" class="BoletoTituloSacado">Sacado:</td>
		<td colspan="8" class="BoletoValorSacado"><%= sacadoNome %></td>
		<td colspan="2" class="BoletoValorSacado"><%= sacadoCPF %></td>
	</tr>
	<tr>
		<td colspan="10" class="BoletoValorSacado"><%= sacadoEndereco %> - <%= sacadoBairro %></td>
	</tr>
	<tr>
		<td colspan="10" class="BoletoValorSacado"><%= sacadoCidade %> - <%= sacadoEstado %>&nbsp;&nbsp;&nbsp;<%= sacadoCEP %></td>
	</tr>
	<tr>
		<td colspan="2" class="BoletoTituloSacador">Sacador / Avalista:</td>
		<td colspan="9" class="BoletoValorSacador"><%= sacadorNome %></td>
	</tr>
	<tr>
		<td colspan="11" class="BoletoTituloDireito" style="text-align: right; padding-right: 0.1cm">Recibo do Sacado - Autenticação Mecânica</td>
	</tr>
	<tr>
		<td colspan="11" height="60" valign="top">&nbsp;</td>
	</tr>
	<tr>
		<td colspan="11" class="BoletoPontilhado">&nbsp;</td>
	</tr>
	<tr>
		<td colspan="4" class="BoletoLogo"><img src="imagens/<%= bancoNumero %>.jpg"></td>
		<td colspan="2" class="BoletoCodigoBanco"><%= bancoNumero %>-<%= bancoDV %></td>
		<td colspan="6" class="BoletoLinhaDigitavel"><%= linhaDigitavel %></td>
	</tr>
	<tr>
		<td colspan="10" class="BoletoTituloEsquerdo">Local de Pagamento</td>
		<td class="BoletoTituloDireito">Vencimento</td>
	</tr>
	<tr>
		<td colspan="10" class="BoletoValorEsquerdo" style="text-align: left; padding-left : 0.1cm"><%= localDePagamento %></td>
		<td class="BoletoValorDireito"><%= dataVencimento %></td>
	</tr>
	<tr>
		<td colspan="10" class="BoletoTituloEsquerdo">Cedente</td>
		<td class="BoletoTituloDireito">Agência/Código do Cedente</td>
	</tr>
	<tr>
		<td colspan="10" class="BoletoValorEsquerdo" style="text-align: left; padding-left : 0.1cm"><%= cedenteNome %></td>
		<td class="BoletoValorDireito"><%= bancoAgencia %>/<%= cedenteCodigo %></td>
	</tr>
	<tr>
		<td colspan="3" class="BoletoTituloEsquerdo">Data do Documento</td>
		<td colspan="4" class="BoletoTituloEsquerdo">Número do Documento</td>
		<td class="BoletoTituloEsquerdo">Espécie</td>
		<td class="BoletoTituloEsquerdo">Aceite</td>
		<td class="BoletoTituloEsquerdo">Data do Processamento</td>
		<td class="BoletoTituloDireito">Nosso Número</td>
	</tr>
	<tr>
		<td colspan="3" class="BoletoValorEsquerdo"><%= dataDocumento %></td>
		<td colspan="4" class="BoletoValorEsquerdo"><%= numeroDocumento %></td>
		<td class="BoletoValorEsquerdo"><%= especie %></td>
		<td class="BoletoValorEsquerdo"><%= aceite %></td>
		<td class="BoletoValorEsquerdo"><%= dataProcessamento %></td>
		<td class="BoletoValorDireito"><%= NossoNumero %></td>
	</tr>
	<tr>
		<td colspan="3" class="BoletoTituloEsquerdo">Uso do Banco</td>
		<td colspan="2" class="BoletoTituloEsquerdo">Carteira</td>
		<td colspan="2" class="BoletoTituloEsquerdo">Moeda</td>
		<td colspan="2" class="BoletoTituloEsquerdo">Quantidade</td>
		<td class="BoletoTituloEsquerdo">(x) Valor</td>
		<td class="BoletoTituloDireito">(=) Valor do Documento</td>
	</tr>
	<tr>
		<td colspan="3" class="BoletoValorEsquerdo">&nbsp;</td>
		<td colspan="2" class="BoletoValorEsquerdo"><%= bancoCarteira %></td>
		<td colspan="2" class="BoletoValorEsquerdo">R$</td>
		<td colspan="2" class="BoletoValorEsquerdo">&nbsp;</td>
		<td class="BoletoValorEsquerdo">&nbsp;</td>
		<td class="BoletoValorDireito"><%= valorDocumento %></td>
	</tr>
	<tr>
		<td colspan="10" class="BoletoTituloEsquerdo">Instruções (TODAS AS INFORMAÇÕES DESTE BLOQUETO SÃO DE EXCLUSIVA RESPONSABILIDADE DO CEDENTE)</td>
		<td class="BoletoTituloDireito">(-) Desconto</td>
	</tr>
	<tr>
		<td colspan="10" rowspan="9" class="BoletoValorEsquerdo" style="text-align: left; vertical-align:top; padding-left : 0.1cm">
			<p>
				<% if percJuros > 0 then %>Após o vencimento, cobrar R$ <%= valorJuros %> por dia de atraso.<br><% end if %>
				<% if percMulta > 0 then %>Após <%= dateAdd("d", 3, dataVencimento) %> cobrar multa de R$ <%= valorMulta %>.<br><% end if %>
				<% if percDesconto > 0 then %>Até <%= dataDesconto %> conceder desconto de R$ <%= valorDesconto %>.<br><% end if %>
				<%= instrucoes %>
			</p>
		</td>
		<td class="BoletoValorDireito">&nbsp;</td>
	</tr>
	<tr>
		<td class="BoletoTituloDireito">(-) Outras Deduções/Abatimento</td>
	</tr>
	<tr>
		<td class="BoletoValorDireito">&nbsp;</td>
	</tr>
	<tr>
		<td class="BoletoTituloDireito">(+) Mora/Multa/Juros</td>
	</tr>
	<tr>
		<td class="BoletoValorDireito">&nbsp;</td>
	</tr>
	<tr>
		<td class="BoletoTituloDireito">(+) Outros Acréscimos</td>
	</tr>
	<tr>
		<td class="BoletoValorDireito">&nbsp;</td>
	</tr>
	<tr>
		<td class="BoletoTituloDireito">(=) Valor Cobrado</td>
	</tr>
	<tr>
		<td class="BoletoValorDireito">&nbsp;</td>
	</tr>
	<tr>
		<td rowspan="3" class="BoletoTituloSacado">Sacado:</td>
		<td colspan="8" class="BoletoValorSacado"><%= sacadoNome %></td>
		<td colspan="2" class="BoletoValorSacado"><%= sacadoCPF %></td>
	</tr>
	<tr>
		<td colspan="10" class="BoletoValorSacado"><%= sacadoEndereco %> - <%= sacadoBairro %></td>
	</tr>
	<tr>
		<td colspan="10" class="BoletoValorSacado"><%= sacadoCidade %> - <%= sacadoEstado %>&nbsp;&nbsp;&nbsp;<%= sacadoCEP %></td>
	</tr>
	<tr>
		<td colspan="2" class="BoletoTituloSacador">Sacador / Avalista:</td>
		<td colspan="9" class="BoletoValorSacador"><%= sacadorNome %></td>
	</tr>
	<tr>
		<td colspan="11" class="BoletoTituloDireito" style="text-align: right; padding-right: 0.1cm">Ficha de Compensação - Autenticação Mecânica</td>
	</tr>
	<tr>
		<td colspan="11" height="60" valign="top"><%= codigoBarras %></td>
	</tr>
	<tr>
		<td colspan="11" class="BoletoPontilhado">&nbsp;</td>
	</tr>
</table>
</body>
</html>
