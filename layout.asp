<%
Dim boleto
Set boleto = Session("BoletoASP")
%>
<style type="text/css">
.Boleto td.BoletoCodigoBanco {
	font-size: 24px;
	font-family: arial, verdana;
	font-weight : bold;
	font-style: italic;
	text-align: center;
	vertical-align: bottom;
	border-bottom: 1px solid #000000;
	border-right: 1px solid #000000;
	padding-bottom : 4px
}
.Boleto td.BoletoLogo {
	border-bottom: 1px solid #000000;
	border-right: 1px solid #000000;
	text-align: center;
	height: 10mm
}
.Boleto td.BoletoLinhaDigitavel {
	font-size: 15px;
	font-family: arial, verdana;
	font-weight : bold;
	text-align: center;
	vertical-align: bottom;
	border-bottom: 1px solid #000000;
	padding-bottom : 4px;
}
.Boleto td.BoletoTituloEsquerdo {
	font-size: 9px;
	font-family: arial, verdana;
	padding-left : 1px;
	border-right: 1px solid #000000;
	text-align: left
}
.Boleto td.BoletoTituloDireito {
	font-size: 9px;
	font-family: arial, verdana;
	padding-left : 1px;
	text-align: left
}
.Boleto td.BoletoValorEsquerdo {
	font-size: 10px;
	font-family: arial, verdana;
	text-align: center;
	border-right: 1px solid #000000;
	font-weight: bold;
	border-bottom: 1px solid #000000;
	padding-top: 2px
}
.Boleto td.BoletoValorDireito {
	font-size: 10px;
	font-family: arial, verdana;
	text-align:right;
	padding-right: 9px;
	padding-top: 2px;
	border-bottom: 1px solid #000000;
	font-weight: bold;
}
.Boleto td.BoletoTituloSacado {
	font-size: 8px;
	font-family: arial, verdana;
	padding-left : 1px;
	vertical-align: top;
	padding-top : 1px;
	text-align: left
}
.Boleto td.BoletoValorSacado {
	font-size: 9px;
	font-family: arial, verdana;
	font-weight: bold;
	text-align : left
}
.Boleto td.BoletoTituloSacador {
	font-size: 8px;
	font-family: arial, verdana;
	padding-left : 1px;
	vertical-align: bottom;
	padding-bottom : 2px;
	border-bottom: 1px solid #000000
}
.Boleto td.BoletoValorSacador {
	font-size: 9px;
	font-family: arial, verdana;
	vertical-align: bottom;
	padding-bottom : 1px;
	border-bottom: 1px solid #000000;
	font-weight: bold;
	text-align: left
}
.Boleto td.BoletoPontilhado {
	border-top: 1px dashed #000000;
	font-size: 4px
}
.Boleto ul.BoletoInstrucoes {
	font-size : 9px;
	font-family : verdana, arial
}
</style>
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
		<td colspan="4" class="BoletoLogo"><img src="imagens/<%= boleto.Banco.Numero %>.jpg"></td>
		<td colspan="2" class="BoletoCodigoBanco"><%= boleto.Banco.Numero %>-<%= boleto.Mod11(boleto.Banco.Numero, "") %></td>
		<td colspan="6" class="BoletoLinhaDigitavel"><%= boleto.LinhaDigitavel() %></td>
	</tr>
	<tr>
		<td colspan="10" class="BoletoTituloEsquerdo">Local de Pagamento</td>
		<td class="BoletoTituloDireito">Vencimento</td>
	</tr>
	<tr>
		<td colspan="10" class="BoletoValorEsquerdo" style="text-align: left; padding-left : 0.1cm"><%= boleto.Banco.LocalPagamento %></td>
		<td class="BoletoValorDireito"><%= boleto.DataVencimento %></td>
	</tr>
	<tr>
		<td colspan="10" class="BoletoTituloEsquerdo">Cedente</td>
		<td class="BoletoTituloDireito">Agência/Código do Cedente</td>
	</tr>
	<tr>
		<td colspan="10" class="BoletoValorEsquerdo" style="text-align: left; padding-left : 0.1cm"><%= boleto.CedenteNome %></td>
		<td class="BoletoValorDireito"><%= boleto.Banco.Agencia %>/<%= boleto.Banco.Conta %>-<%= boleto.Banco.ContaDV %></td>
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
		<td colspan="3" class="BoletoValorEsquerdo"><%= boleto.DataDocumento %></td>
		<td colspan="4" class="BoletoValorEsquerdo"><%= boleto.NumeroDocumento %></td>
		<td class="BoletoValorEsquerdo"><%= boleto.Especie %></td>
		<td class="BoletoValorEsquerdo"><%= boleto.Aceite %></td>
		<td class="BoletoValorEsquerdo"><%= boleto.DataProcessamento %></td>
		<td class="BoletoValorDireito"><%= boleto.NossoNumero %>-<%= boleto.NossoNumeroDV %></td>
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
		<td colspan="2" class="BoletoValorEsquerdo"><%= boleto.Banco.Carteira %></td>
		<td colspan="2" class="BoletoValorEsquerdo">R$</td>
		<td colspan="2" class="BoletoValorEsquerdo">&nbsp;</td>
		<td class="BoletoValorEsquerdo">&nbsp;</td>
		<td class="BoletoValorDireito"><%= boleto.ValorDocumento %></td>
	</tr>
	<tr>
		<td colspan="10" class="BoletoTituloEsquerdo">Instruções</td>
		<td class="BoletoTituloDireito">(-) Desconto</td>
	</tr>
	<tr>
		<td colspan="10" rowspan="9" class="BoletoValorEsquerdo" style="text-align: left; vertical-align:top; padding-left : 0.1cm"><%= boleto.Instrucoes %>&nbsp;</td>
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
		<td colspan="8" class="BoletoValorSacado"><%= boleto.Sacado.Nome %></td>
		<td colspan="2" class="BoletoValorSacado"><%= boleto.Sacado.CPF %></td>
	</tr>
	<tr>
		<td colspan="10" class="BoletoValorSacado"><%= boleto.Sacado.Endereco %> - <%= boleto.Sacado.Bairro %></td>
	</tr>
	<tr>
		<td colspan="10" class="BoletoValorSacado"><%= boleto.Sacado.Cidade %> - <%= boleto.Sacado.Estado %>&nbsp;&nbsp;&nbsp;<%= boleto.Sacado.CEP %></td>
	</tr>
	<tr>
		<td colspan="2" class="BoletoTituloSacador">Sacador / Avalista:</td>
		<td colspan="9" class="BoletoValorSacador"><%= boleto.SacadorNome %></td>
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
		<td colspan="4" class="BoletoLogo"><img src="imagens/<%= boleto.Banco.Numero %>.jpg"></td>
		<td colspan="2" class="BoletoCodigoBanco"><%= boleto.Banco.Numero %>-<%= boleto.Mod11(boleto.Banco.Numero, "") %></td>
		<td colspan="6" class="BoletoLinhaDigitavel"><%= boleto.LinhaDigitavel() %></td>
	</tr>
	<tr>
		<td colspan="10" class="BoletoTituloEsquerdo">Local de Pagamento</td>
		<td class="BoletoTituloDireito">Vencimento</td>
	</tr>
	<tr>
		<td colspan="10" class="BoletoValorEsquerdo" style="text-align: left; padding-left : 0.1cm"><%= boleto.Banco.LocalPagamento %></td>
		<td class="BoletoValorDireito"><%= boleto.dataVencimento %></td>
	</tr>
	<tr>
		<td colspan="10" class="BoletoTituloEsquerdo">Cedente</td>
		<td class="BoletoTituloDireito">Agência/Código do Cedente</td>
	</tr>
	<tr>
		<td colspan="10" class="BoletoValorEsquerdo" style="text-align: left; padding-left : 0.1cm"><%= boleto.CedenteNome %></td>
		<td class="BoletoValorDireito"><%= boleto.Banco.Agencia %>/<%= boleto.Banco.Conta %>-<%= boleto.Banco.ContaDV %></td>
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
		<td colspan="3" class="BoletoValorEsquerdo"><%= boleto.DataDocumento %></td>
		<td colspan="4" class="BoletoValorEsquerdo"><%= boleto.NumeroDocumento %></td>
		<td class="BoletoValorEsquerdo"><%= boleto.Especie %></td>
		<td class="BoletoValorEsquerdo"><%= boleto.Aceite %></td>
		<td class="BoletoValorEsquerdo"><%= boleto.DataProcessamento %></td>
		<td class="BoletoValorDireito"><%= boleto.NossoNumero %>-<%= boleto.NossoNumeroDV %></td>
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
		<td colspan="2" class="BoletoValorEsquerdo"><%= boleto.Banco.Carteira %></td>
		<td colspan="2" class="BoletoValorEsquerdo">R$</td>
		<td colspan="2" class="BoletoValorEsquerdo">&nbsp;</td>
		<td class="BoletoValorEsquerdo">&nbsp;</td>
		<td class="BoletoValorDireito"><%= formatNumber(boleto.ValorDocumento, 2) %></td>
	</tr>
	<tr>
		<td colspan="10" class="BoletoTituloEsquerdo">Instruções</td>
		<td class="BoletoTituloDireito">(-) Desconto</td>
	</tr>
	<tr>
		<td colspan="10" rowspan="9" class="BoletoValorEsquerdo" style="text-align: left; vertical-align:top; padding-left : 0.1cm">
			<%= boleto.Instrucoes %>&nbsp;
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
		<td colspan="8" class="BoletoValorSacado"><%= boleto.Sacado.Nome %>&nbsp;</td>
		<td colspan="2" class="BoletoValorSacado"><%= boleto.Sacado.CPF %>&nbsp;</td>
	</tr>
	<tr>
		<td colspan="10" class="BoletoValorSacado"><%= boleto.Sacado.Endereco %> - <%= boleto.Sacado.Bairro %></td>
	</tr>
	<tr>
		<td colspan="10" class="BoletoValorSacado"><%= boleto.Sacado.Cidade %> - <%= boleto.Sacado.Estado %>&nbsp;&nbsp;&nbsp;<%= boleto.Sacado.CEP %></td>
	</tr>
	<tr>
		<td colspan="2" class="BoletoTituloSacador">Sacador / Avalista:</td>
		<td colspan="9" class="BoletoValorSacador"><%= boleto.SacadorNome %>&nbsp;</td>
	</tr>
	<tr>
		<td colspan="11" class="BoletoTituloDireito" style="text-align: right; padding-right: 0.1cm">Ficha de Compensação - Autenticação Mecânica</td>
	</tr>
	<tr>
		<td colspan="11" height="60" valign="top"><%= boleto.CodigoBarras() %>&nbsp;</td>
	</tr>
	<tr>
		<td colspan="11" class="BoletoPontilhado">&nbsp;</td>
	</tr>
</table>
<% Set boleto = Nothing %>