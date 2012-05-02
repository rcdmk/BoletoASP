<%
'
' Gerador de Boletos em ASP por RCDMK <rcdmk@rcdmk.com>
' Início em 29/04/2012
'
' Esta classe tem o objetivo de gerar boletos para diversos bancos através de uma
' interface comum e simples
'
' ## Constantes #########################################################################
Const MOD11_BARRAS 		= "CB"
Const MOD11_BRADESCO 	= "BRADESCO"
Const MOD11_BRADESCO_P 	= "BRADESCO_P"
Const MOD11_BB			= "BB"


' ## Classe principal para gerar o boleto ###############################################
Class BoletoASP
	' ## Campos ##
	Dim i_pastaImagens, i_numeroDocumento, i_cedenteNome, i_especie, i_aceite, i_sacadorNome
	Dim i_banco, i_nossoNumero, i_nossoNumeroDV, i_sacado, i_moeda, i_dataBase, i_fator
	Dim i_valorDocumento, i_dataDocumento, i_dataVencimento, i_percMulta, i_percJuros
	Dim i_dataProcessamento, i_instrucoes
	

	' ## Propriedades ##
	Public Property Get PastaImagens
		PastaImagens = i_pastaImagens
	End Property
	
	Public Property Let PastaImagens(val)
		i_pastaImagens = Replace(val, "\", "/")
		
		' Remover barra no final
		If Right(i_pastaImagens, 1) = "/" Then i_pastaImagens = Left(i_pastaImagens, Len(i_pastaImagens) - 1)
	End Property
	
	
	Public Property Get NumeroDocumento
		NumeroDocumento = i_numeroDocumento
	End Property
	
	Public Property Let NumeroDocumento(val)
		i_numeroDocumento = val
	End Property
	
	
	Public Property Get CedenteNome
		CedenteNome = i_cedenteNome
	End Property
	
	Public Property Let CedenteNome(val)
		i_cedenteNome = val
	End Property
	
	
	Public Property Get Especie
		Especie = i_especie
	End Property
	
	Public Property Let Especie(val)
		i_especie = val
	End Property
	
	
	Public Property Get Aceite
		Aceite = i_aceite
	End Property
	
	Public Property Let Aceite(val)
		i_aceite = val
	End Property
	
	
	Public Property Get SacadorNome
		SacadorNome = i_sacadorNome
	End Property
	
	Public Property Let SacadorNome(val)
		i_sacadorNome = val
	End Property
	
	
	Public Property Get Banco
		Set Banco = i_banco
	End Property
	
	Public Property Let Banco(val)
		Set i_banco = val
	End Property
	
	
	Public Property Get NossoNumero
		NossoNumero = i_nossoNumero
	End Property
	
	Public Property Let NossoNumero(val)
		i_nossoNumero = Completa(val, 8)
		CalculaNossoNumeroDV
	End Property
	
	
	Public Property Get NossoNumeroDV
		CalculaNossoNumeroDV
		NossoNumeroDV = i_nossoNumeroDV
	End Property

	
	Public Property Get Sacado
		Set Sacado = i_sacado
	End Property
	
	Public Property Let Sacado(val)
		Set i_sacado = val
	End Property
	
	
	' Valor do documento
	Public Property Get ValorDocumento
		ValorDocumento = i_valorDocumento
	End Property
	
	Public Property Let ValorDocumento(val)
		i_valorDocumento = val
	End Property
	
	
	Public Property Get DataDocumento
		DataDocumento = i_dataDocumento
	End Property
	
	Public Property Let DataDocumento(val)
		i_dataDocumento = val
	End Property
	
	
	Public Property Get DataProcessamento
		DataProcessamento = i_dataProcessamento
	End Property
	
	Public Property Let DataProcessamento(val)
		i_dataProcessamento = val
	End Property
	
	
	Public Property Get DataVencimento
		DataVencimento = i_dataVencimento
	End Property
	
	Public Property Let DataVencimento(val)
		i_dataVencimento = val
		CalculaFator
	End Property
	
	
	Public Property Get PercMulta
		PercMulta = i_percMulta
	End Property
	
	Public Property Let PercMulta(val)
		i_percMulta = val
	End Property
	
	
	Public Property Get PercJuros
		PercJuros = i_percJuros
	End Property
	
	Public Property Let PercJuros(val)
		i_percJuros = val
	End Property
	
	
	Public Property Get Instrucoes
		Instrucoes = i_instrucoes
	End Property
	
	Public Property Let Instrucoes(val)
		i_instrucoes = val
	End Property
	
	
	' ## Construtor ##
	Private Sub Class_Initialize()
		i_pastaImagens = "imagens"
		i_nossoNumero = "00000000"
		i_numeroDocumento = "000"
		i_cedenteNome = "Cedente"
		i_sacadorNome = ""
		
		i_especie = "DM"
		i_moeda = 9 ' Real
		i_dataBase = CDate("7/10/1997")
		
		i_aceite = "N"
		i_valorDocumento = 0
		i_dataDocumento = date
		i_dataProcessamento = date
		i_dataVencimento = dateAdd("d", 3, date)
		
		i_percMulta = 2
		i_percJuros = 0.33
		
		i_instrucoes = ""
		
		Set i_banco = New BancoASP
		Set i_sacado = New SacadoASP
		
		CalculaFator
		CalculaNossoNumeroDV
	End Sub
	
	Private Sub Class_Terminate()
		Set i_banco = Nothing
		Set i_sacado = Nothing
	End Sub
	
	
	' ## Metodos ##
	' Cálculo de DV Mode 10
	Public Function Mod10(ByVal strNumero)
		Dim DV, tamanho, i, j, k
		Dim num, soma, somaTotal
		
		strNumero = CStr(strNumero)
		tamanho = Len(strNumero)
		k = 0
		
		For i = tamanho - 1 to 0 Step -1
			k = k + 1
			num = Mid(strNumero, i + 1, 1)
			If k And 1 Then num = num * 2
			
			If num > 9 Then
				soma = 0
				For j = 1 to Len(num)
					soma = soma + CInt(Mid(num, j, 1))
				Next
				
				num = soma
			End If
			
			somaTotal = somaTotal + CInt(num)
		Next
		
		DV = somaTotal Mod 10
		
		If DV > 0 Then DV = 10 - DV
		
		Mod10 = DV
	End Function
	
	
	' Cálculo de DV Mode 11
	Public Function Mod11(Byval strNumero, ByVal tipo)
		Dim DV, tamanho, i, fator
		Dim num, soma
		
		strNumero = strNumero
		tamanho = Len(strNumero)
		soma = 0
		fator = 2
		
		For i = tamanho - 1 to 0 Step -1
			num = Mid(strNumero, i + 1, 1)
			num = num * fator
			
			soma = soma + CInt(num)
			
			fator = fator + 1
			If tipo <> "" And fator > 9 Then fator = 2
		Next
	
		'DV = (soma * 10) Mod 11
		DV = soma Mod 11
		DV = 11 - DV
		
		If tipo = MOD11_BARRAS Then
			If DV = 0 Or DV = 10 Then DV = 1
			
		ElseIf tipo = MOD11_BRADESCO Then
			If DV = 10 Then DV = 0
			
		ElseIf tipo = MOD11_BRADESCO_P Then
			If DV = 10 Then DV = "P"
			
		ElseIf tipo = MOD11_BB Then
			If DV = 10 Then DV = "X"
			
		Else
			If DV = 10 Then DV = 0
		End If
		
		Mod11 = DV
	End Function
	
	
	' Número da linha digitável ou representação numérica
	Public Function LinhaDigitavel()
		Dim retorno, i, posicoes(16), numero
		
		retorno = ""
		numero = NumCodigoBarras()
		
		posicoes(1)		= Left(numero, 3) 		' Número do banco
		posicoes(2)		= i_moeda				' Moeda
		posicoes(3)		= Mid(numero, 20, 3)	' Carteira
		posicoes(4)		= Mid(numero, 23, 2)	' 2 primeiros dígitos do nosso número
		posicoes(5)		= "" 					' DV do primeiro grupo
		
		posicoes(6)		= Mid(numero, 25, 6)	' Restante do nosso número
		posicoes(7)		= i_nossoNumeroDV		' Dígito do nosso número
		posicoes(8)		= Mid(numero, 32, 3)	' 
		posicoes(9)		= "" ' DV do segundo grupo
		
		posicoes(10)	= Mid(numero, 35, 1) 	' Restante da agéncia
		posicoes(11)	= Mid(numero, 36, 6)	' Conta + DV
		posicoes(12)	= "000"
		posicoes(13)	= "" 					' DV do terceiro grupo
		
		posicoes(14)	= Mid(numero, 5, 1) 	' DV do código de barras
		
		posicoes(15)	= Mid(numero, 6, 4) 	' Fator de vencimento
		posicoes(16)	= Mid(numero, 10, 10) 	' Valor do documento
		
		' Calculando DVs
		posicoes(5) 	= Mod10(posicoes(1) & posicoes(2) & posicoes(3) & posicoes(4))
		posicoes(9) 	= Mod10(posicoes(6) & posicoes(7) & posicoes(8))
		posicoes(13) 	= Mod10(posicoes(10) & posicoes(11) & posicoes(12))
		
		For i = 1 To 16
			retorno = retorno & posicoes(i)
		Next

		LinhaDigitavel = Left(retorno, 5) & "." & Mid(retorno, 6, 5) & " " & Mid(retorno, 11, 5) & "." & Mid(retorno, 16, 6) & " " & Mid(retorno, 22, 5) & "." & Mid(retorno, 27, 6) & " " & Mid(retorno, 33, 1) & " " & Mid(retorno, 34)
	End Function
	
	
	' Número do código de barras
	Public Function NumCodigoBarras()
		Dim retorno, i, posicoes(43)
		retorno = ""
		
		CalculaNossoNumeroDV
		
		posicoes(1) 	= Completa(i_banco.Numero, 3)
		posicoes(4) 	= i_moeda
		posicoes(5) 	= "" ' DV do código Mod11
		
		' Se o valor for maior que 100 milhões, ignora-se o fator de vencimento
		If i_valorDocumento >= 100000000 Then
			posicoes(6)		= Completa(CInt(i_valorDocumento * 100), 14)
		Else
			posicoes(6) 	= Completa(i_fator, 4)
			posicoes(10)	= Completa(CInt(i_valorDocumento * 100), 10)
		End If
		
		posicoes(20) 	= Completa(i_banco.Carteira, 3)
		posicoes(23) 	= Left(i_nossoNumero, 8)
		posicoes(31) 	= i_nossoNumeroDV
		posicoes(32) 	= Completa(i_banco.Agencia, 4)
		posicoes(36) 	= Completa(i_banco.Conta, 5)
		posicoes(41) 	= i_banco.ContaDV
		posicoes(42) 	= "000"
		
		For i = 1 To 43
			retorno = retorno & posicoes(i)
		Next
		
		posicoes(5) = Mod11(retorno, MOD11_BARRAS)
		
		NumCodigoBarras = Left(retorno, 4) & posicoes(5) & Right(retorno, 39)
	End Function
	
	
	' Monta o código de barras em HTML
	Public Function CodigoBarras()
		Dim inicio, fim, codigo, numeroCodigo, retorno
		Dim representacao(9), i, j, digito, barra, digito1, digito2
		
		' Códigos de início e fim
		inicio = "0000"
		fim = "100"
		
		' Base da codificação
		' pesos 		 = "12470"
		representacao(0) = "00110" ' 4 + 7 = 11 - substituido por 0
		representacao(1) = "10001" ' 1 + 0 = 1
		representacao(2) = "01001" ' 2 + 0 = 2
		representacao(3) = "11000" ' 1 + 2 = 3
		representacao(4) = "00101" ' ...
		representacao(5) = "10100"
		representacao(6) = "01100"
		representacao(7) = "00011"
		representacao(8) = "10010"
		representacao(9) = "01010"
		
		' Numeração do código para codificar
		numeroCodigo = NumCodigoBarras()
		retorno = ""
		
		' Pegar os dígitos em pares
		For i = 1 To 43 Step 2
			digito1 = Mid(numeroCodigo, i, 1)
			digito2 = Mid(numeroCodigo, i + 1, 1)
			
			' Converter para representação binária
			digito1 = representacao(CInt(digito1))
			digito2 = representacao(CInt(digito2))
			
			' Intercalar representações
			For j = 1 To 5
				codigo = codigo & Mid(digito1, j, 1) & Mid(digito2, j, 1)
			Next
		Next		
		
		' Montar código final
		codigo = inicio & codigo & fim
		
		' Montar HTML
		barra = "b" ' "b" = barra, "" = espaço
		For i = 1 To Len(codigo)
			digito = Mid(codigo, i, 1)
			
			retorno = retorno & "<img src=""" & i_pastaImagens & "/barras/" & digito & barra & ".gif"" alt="""" " & vbCrLf & " />" ' Quebra de linha para evitar problemas com e-mails
			
			If barra = "b" Then
				barra = ""
			Else
				barra = "b"
			End If
		Next
		
		CodigoBarras = retorno
	End Function
	
	
	Public Function HTML()
		HTML = "<style type=""text/css"">.Boleto td.BoletoCodigoBanco{font-size: 24px;font-family: arial, verdana;font-weight: bold;font-style: italic;text-align: center;" & vbCrLf _
			 & "vertical-align: bottom;border-bottom: 1px solid #000000;border-right: 1px solid #000000;padding-bottom : 4px}.Boleto td.BoletoLogo{border-bottom: 1px solid #000000;" & vbCrLf _
			 & "border-right: 1px solid #000000;text-align: center;height: 10mm}.Boleto td.BoletoLinhaDigitavel{font-size: 15px;font-family: arial, verdana;font-weight : bold;" & vbCrLf _
			 & "text-align: center;vertical-align: bottom;border-bottom: 1px solid #000000;padding-bottom : 4px;}.Boleto td.BoletoTituloEsquerdo{font-size: 9px;font-family: arial, verdana;" & vbCrLf _
			 & "padding-left : 1px;border-right: 1px solid #000000;text-align: left}.Boleto td.BoletoTituloDireito{font-size: 9px;font-family: arial, verdana;padding-left : 1px;" & vbCrLf _
			 & "text-align: left}.Boleto td.BoletoValorEsquerdo{font-size: 10px;font-family: arial, verdana;text-align: center;border-right: 1px solid #000000;font-weight: bold;" & vbCrLf _
			 & "border-bottom: 1px solid #000000;padding-top: 2px}.Boleto td.BoletoValorDireito{font-size: 10px;font-family: arial, verdana;text-align:right;padding-right: 9px;" & vbCrLf _
			 & "padding-top: 2px;border-bottom: 1px solid #000000;font-weight: bold;}.Boleto td.BoletoTituloSacado{font-size: 8px;font-family: arial, verdana;padding-left : 1px;" & vbCrLf _
			 & "vertical-align: top;padding-top : 1px;text-align: left}.Boleto td.BoletoValorSacado{font-size: 9px;font-family: arial, verdana;font-weight: bold;text-align : left}" & vbCrLf _
			 & ".Boleto td.BoletoTituloSacador{font-size: 8px;font-family: arial, verdana;padding-left : 1px;vertical-align: bottom;padding-bottom : 2px;border-bottom: 1px solid #000000}" & vbCrLf _
			 & ".Boleto td.BoletoValorSacador{font-size: 9px;font-family: arial, verdana;vertical-align: bottom;padding-bottom : 1px;border-bottom: 1px solid #000000;font-weight: bold;" & vbCrLf _
			 & "text-align: left}.Boleto td.BoletoPontilhado{border-top: 1px dashed #000000;font-size: 4px}.Boleto ul.BoletoInstrucoes{font-size : 9px;font-family : verdana, arial}</style>" & vbCrLf _
			 & "<table cellspacing=""0"" cellpadding=""0"" border=""0"" class=""Boleto""><tr><td style=""width: 0.9cm"">&nbsp;</td><td style=""width: 1cm"">&nbsp;</td>" & vbCrLf _
			 & "<td style=""width: 1.9cm"">&nbsp;</td><td style=""width: 0.5cm"">&nbsp;</td><td style=""width: 1.3cm"">&nbsp;</td><td style=""width: 0.8cm"">&nbsp;</td>" & vbCrLf _
			 & "<td style=""width: 1cm"">&nbsp;</td><td style=""width: 1.9cm"">&nbsp;</td><td style=""width: 1.9cm"">&nbsp;</td><td style=""width: 3.8cm"">&nbsp;</td>" & vbCrLf _
			 & "<td style=""width: 3.8cm"">&nbsp;</td></tr><tr><td colspan=""11""><ul class=""BoletoInstrucoes""><li>Imprima em papel A4 ou Carta</li>" & vbCrLf _
			 & "<li>Utilize margens mínimas a direita e a esquerda</li><li>Recorte na linha pontilhada</li><li>Não rasure o código de barras</li></ul>&nbsp;</td></tr><tr>" & vbCrLf _
			 & "<td colspan=""11"" class=""BoletoPontilhado"">&nbsp;</td></tr><tr><td colspan=""4"" class=""BoletoLogo"">" & vbCrLf _
			 & "<img src=""imagens/" & boleto.Banco.Numero & ".jpg"">&nbsp;</td>" & vbCrLf _
			 & "<td colspan=""2"" class=""BoletoCodigoBanco"">" & boleto.Banco.Numero & "-" & boleto.Mod11(boleto.Banco.Numero, "") & "</td>" & vbCrLf _
			 & "<td colspan=""6"" class=""BoletoLinhaDigitavel"">" & boleto.LinhaDigitavel() & "</td></tr><tr>" & vbCrLf _
			 & "<td colspan=""10"" class=""BoletoTituloEsquerdo"">Local de Pagamento</td>" & vbCrLf _
			 & "<td class=""BoletoTituloDireito"">Vencimento</td></tr><tr><td colspan=""10"" class=""BoletoValorEsquerdo"" style=""text-align: left;padding-left : 0.1cm"">" & vbCrLf _
			 & boleto.Banco.LocalPagamento & "</td><td class=""BoletoValorDireito"">" & boleto.DataVencimento & "</td>" & vbCrLf _
			 & "</tr><tr><td colspan=""10"" class=""BoletoTituloEsquerdo"">Cedente</td><td class=""BoletoTituloDireito"">Agência/Código do Cedente</td></tr>" & vbCrLf _
			 & "<tr><td colspan=""10"" class=""BoletoValorEsquerdo"" style=""text-align: left;padding-left : 0.1cm"">" & vbCrLf _
			 & boleto.CedenteNome & "</td><td class=""BoletoValorDireito"">" & boleto.Banco.Agencia & "/" & boleto.Banco.Conta & "-" & boleto.Banco.ContaDV & "</td></tr><tr>" & vbCrLf _
			 & "<td colspan=""3"" class=""BoletoTituloEsquerdo"">Data do Documento</td><td colspan=""4"" class=""BoletoTituloEsquerdo"">Número do Documento</td>" & vbCrLf _
			 & "<td class=""BoletoTituloEsquerdo"">Espécie</td><td class=""BoletoTituloEsquerdo"">Aceite</td><td class=""BoletoTituloEsquerdo"">Data do Processamento</td>" & vbCrLf _
			 & "<td class=""BoletoTituloDireito"">Nosso Número</td></tr><tr><td colspan=""3"" class=""BoletoValorEsquerdo"">" & boleto.DataDocumento & "</td>" & vbCrLf _
			 & "<td colspan=""4"" class=""BoletoValorEsquerdo"">" & boleto.NumeroDocumento & "</td><td class=""BoletoValorEsquerdo"">" & boleto.Especie & "</td>" & vbCrLf _
			 & "<td class=""BoletoValorEsquerdo"">" & boleto.Aceite & "</td><td class=""BoletoValorEsquerdo"">" & boleto.DataProcessamento & "</td>" & vbCrLf _
			 & "<td class=""BoletoValorDireito"">" & boleto.NossoNumero & "-" & boleto.NossoNumeroDV & "</td></tr><tr><td colspan=""3"" class=""BoletoTituloEsquerdo"">" & vbCrLf _
			 & "Uso do Banco</td><td colspan=""2"" class=""BoletoTituloEsquerdo"">Carteira</td><td colspan=""2"" class=""BoletoTituloEsquerdo"">Moeda</td>" & vbCrLf _
			 & "<td colspan=""2"" class=""BoletoTituloEsquerdo"">Quantidade</td><td class=""BoletoTituloEsquerdo"">(x) Valor</td>" & vbCrLf _
			 & "<td class=""BoletoTituloDireito"">(=) Valor do Documento</td></tr><tr><td colspan=""3"" class=""BoletoValorEsquerdo"">&nbsp;</td>" & vbCrLf _
			 & "<td colspan=""2"" class=""BoletoValorEsquerdo"">" & boleto.Banco.Carteira & "</td><td colspan=""2"" class=""BoletoValorEsquerdo"">R$</td>" & vbCrLf _
			 & "<td colspan=""2"" class=""BoletoValorEsquerdo"">&nbsp;</td><td class=""BoletoValorEsquerdo"">&nbsp;</td>" & vbCrLf _
			 & "<td class=""BoletoValorDireito"">" & boleto.ValorDocumento & "</td></tr><tr><td colspan=""10"" class=""BoletoTituloEsquerdo"">Instruções</td>" & vbCrLf _
			 & "<td class=""BoletoTituloDireito"">(-) Desconto</td></tr><tr>" & vbCrLf _
			 & "<td colspan=""10"" rowspan=""9"" class=""BoletoValorEsquerdo"" style=""text-align: left;vertical-align:top;padding-left : 0.1cm"">" & vbCrLf _
			 & boleto.Instrucoes & "</td>" & vbCrLf _
			 & "<td class=""BoletoValorDireito"">&nbsp;</td></tr><tr><td class=""BoletoTituloDireito"">(-) Outras Deduções/Abatimento</td>" & vbCrLf _
			 & "</tr><tr><td class=""BoletoValorDireito"">&nbsp;</td></tr><tr><td class=""BoletoTituloDireito"">(+) Mora/Multa/Juros</td></tr><tr>" & vbCrLf _
			 & "<td class=""BoletoValorDireito"">&nbsp;</td></tr><tr><td class=""BoletoTituloDireito"">(+) Outros Acréscimos</td></tr>" & vbCrLf _
			 & "<tr><td class=""BoletoValorDireito"">&nbsp;</td></tr><tr><td class=""BoletoTituloDireito"">(=) Valor Cobrado</td></tr>" & vbCrLf _
			 & "<tr><td class=""BoletoValorDireito"">&nbsp;</td></tr><tr><td rowspan=""3"" class=""BoletoTituloSacado"">Sacado:</td>" & vbCrLf _
			 & "<td colspan=""8"" class=""BoletoValorSacado"">" & boleto.Sacado.Nome & "</td>" & vbCrLf _
			 & "<td colspan=""2"" class=""BoletoValorSacado"">" & boleto.Sacado.CPF & "</td></tr><tr>" & vbCrLf _
			 & "<td colspan=""10"" class=""BoletoValorSacado"">" & boleto.Sacado.Endereco & " - " & boleto.Sacado.Bairro & "</td>" & vbCrLf _
			 & "</tr><tr><td colspan=""10"" class=""BoletoValorSacado"">" & boleto.Sacado.Cidade & "- " & boleto.Sacado.Estado & "" & boleto.Sacado.CEP & "</td>" & vbCrLf _
			 & "</tr><tr><td colspan=""2"" class=""BoletoTituloSacador"">Sacador / Avalista:</td>" & vbCrLf _
			 & "<td colspan=""9"" class=""BoletoValorSacador"">" & boleto.SacadorNome & "</td></tr>" & vbCrLf _
			 & "<tr><td colspan=""11"" class=""BoletoTituloDireito"" style=""text-align: right;padding-right: 0.1cm"">Recibo do Sacado - Autenticação Mecânica</td></tr><tr>" & vbCrLf _
			 & "<td colspan=""11"" height=""60"" valign=""top"">&nbsp;</td></tr><tr><td colspan=""11"" class=""BoletoPontilhado"">&nbsp;</td>" & vbCrLf _
			 & "</tr><tr><td colspan=""4"" class=""BoletoLogo""><img src=""imagens/" & boleto.Banco.Numero & ".jpg"">&nbsp;</td>" & vbCrLf _
			 & "<td colspan=""2"" class=""BoletoCodigoBanco"">" & boleto.Banco.Numero & "-" & boleto.Mod11(boleto.Banco.Numero, "") & "</td>" & vbCrLf _
			 & "<td colspan=""6"" class=""BoletoLinhaDigitavel"">" & boleto.LinhaDigitavel() & "</td></tr><tr>" & vbCrLf _
			 & "<td colspan=""10"" class=""BoletoTituloEsquerdo"">Local de Pagamento</td><td class=""BoletoTituloDireito"">Vencimento</td></tr>" & vbCrLf _
			 & "<tr><td colspan=""10"" class=""BoletoValorEsquerdo"" style=""text-align: left;padding-left : 0.1cm"">" & vbCrLf _
			 & boleto.Banco.LocalPagamento & "</td><td class=""BoletoValorDireito"">" & boleto.dataVencimento & "</td></tr><tr>" & vbCrLf _
			 & "<td colspan=""10"" class=""BoletoTituloEsquerdo"">Cedente</td><td class=""BoletoTituloDireito"">Agência/Código do Cedente</td></tr><tr>" & vbCrLf _
			 & "<td colspan=""10"" class=""BoletoValorEsquerdo"" style=""text-align: left;padding-left : 0.1cm"">" & boleto.CedenteNome & "</td>" & vbCrLf _
			 & "<td class=""BoletoValorDireito"">" & boleto.Banco.Agencia & "/" & boleto.Banco.Conta & "-" & boleto.Banco.ContaDV & "</td></tr><tr>" & vbCrLf _
			 & "<td colspan=""3"" class=""BoletoTituloEsquerdo"">Data do Documento</td><td colspan=""4"" class=""BoletoTituloEsquerdo"">Número do Documento</td>" & vbCrLf _
			 & "<td class=""BoletoTituloEsquerdo"">Espécie</td><td class=""BoletoTituloEsquerdo"">Aceite</td><td class=""BoletoTituloEsquerdo"">Data do Processamento</td>" & vbCrLf _
			 & "<td class=""BoletoTituloDireito"">Nosso Número</td></tr><tr><td colspan=""3"" class=""BoletoValorEsquerdo"">" & boleto.DataDocumento & "</td>" & vbCrLf _
			 & "<td colspan=""4"" class=""BoletoValorEsquerdo"">" & boleto.NumeroDocumento & "</td><td class=""BoletoValorEsquerdo"">" & boleto.Especie & "</td>" & vbCrLf _
			 & "<td class=""BoletoValorEsquerdo"">" & boleto.Aceite & "</td><td class=""BoletoValorEsquerdo"">" & boleto.DataProcessamento & "</td>" & vbCrLf _
			 & "<td class=""BoletoValorDireito"">" & boleto.NossoNumero & "-" & boleto.NossoNumeroDV & "</td></tr><tr>" & vbCrLf _
			 & "<td colspan=""3"" class=""BoletoTituloEsquerdo"">Uso do Banco</td><td colspan=""2"" class=""BoletoTituloEsquerdo"">Carteira</td>" & vbCrLf _
			 & "<td colspan=""2"" class=""BoletoTituloEsquerdo"">Moeda</td><td colspan=""2"" class=""BoletoTituloEsquerdo"">Quantidade</td>" & vbCrLf _
			 & "<td class=""BoletoTituloEsquerdo"">(x) Valor</td><td class=""BoletoTituloDireito"">(=) Valor do Documento</td>" & vbCrLf _
			 & "</tr><tr><td colspan=""3"" class=""BoletoValorEsquerdo"">&nbsp;</td><td colspan=""2"" class=""BoletoValorEsquerdo"">" & boleto.Banco.Carteira & "</td>" & vbCrLf _
			 & "<td colspan=""2"" class=""BoletoValorEsquerdo"">R$</td><td colspan=""2"" class=""BoletoValorEsquerdo"">&nbsp;</td><td class=""BoletoValorEsquerdo"">&nbsp;</td>" & vbCrLf _
			 & "<td class=""BoletoValorDireito"">" & formatNumber(boleto.ValorDocumento, 2) & "</td></tr><tr><td colspan=""10"" class=""BoletoTituloEsquerdo"">Instruções</td>" & vbCrLf _
			 & "<td class=""BoletoTituloDireito"">(-) Desconto</td></tr><tr><td colspan=""10"" rowspan=""9"" class=""BoletoValorEsquerdo"" style=""text-align: left;" & vbCrLf _
			 & "vertical-align:top;padding-left : 0.1cm"">" & boleto.Instrucoes & "</td><td class=""BoletoValorDireito"">&nbsp;</td></tr><tr>" & vbCrLf _
			 & "<td class=""BoletoTituloDireito"">(-) Outras Deduções/Abatimento</td></tr><tr><td class=""BoletoValorDireito"">&nbsp;</td></tr><tr>" & vbCrLf _
			 & "<td class=""BoletoTituloDireito"">(+) Mora/Multa/Juros</td></tr><tr><td class=""BoletoValorDireito"">&nbsp;</td></tr><tr>" & vbCrLf _
			 & "<td class=""BoletoTituloDireito"">(+) Outros Acréscimos</td></tr><tr><td class=""BoletoValorDireito"">&nbsp;</td></tr><tr>" & vbCrLf _
			 & "<td class=""BoletoTituloDireito"">(=) Valor Cobrado</td></tr><tr><td class=""BoletoValorDireito"">&nbsp;</td></tr><tr>" & vbCrLf _
			 & "<td rowspan=""3"" class=""BoletoTituloSacado"">Sacado:</td><td colspan=""8"" class=""BoletoValorSacado"">" & boleto.Sacado.Nome & "</td>" & vbCrLf _
			 & "<td colspan=""2"" class=""BoletoValorSacado"">" & boleto.Sacado.CPF & "</td></tr><tr>" & vbCrLf _
			 & "<td colspan=""10"" class=""BoletoValorSacado"">" & boleto.Sacado.Endereco & "- " & boleto.Sacado.Bairro & "</td>" & vbCrLf _
			 & "</tr><tr><td colspan=""10"" class=""BoletoValorSacado"">" & boleto.Sacado.Cidade & "- " & boleto.Sacado.Estado & "" & boleto.Sacado.CEP & "</td>" & vbCrLf _
			 & "</tr><tr><td colspan=""2"" class=""BoletoTituloSacador"">Sacador / Avalista:</td><td colspan=""9"" class=""BoletoValorSacador"">" & boleto.SacadorNome & "</td>" & vbCrLf _
			 & "</tr><tr><td colspan=""11"" class=""BoletoTituloDireito"" style=""text-align: right;padding-right: 0.1cm"">Ficha de Compensação - Autenticação Mecânica</td></tr>" & vbCrLf _
			 & "<tr><td colspan=""11"" height=""60"" valign=""top"">" & boleto.CodigoBarras() & "</td></tr>" & vbCrLf _
			 & "<tr><td colspan=""11"" class=""BoletoPontilhado"">&nbsp;</td></tr></table>"
	End Function
	
	
	Public Function Write()
		Response.Write HTML()
	End Function
	
	
	' ## Métodos privados ##
	' Fator do vencimento
	Private Function CalculaFator()
		i_fator = DateDiff("d", i_dataBase, i_dataVencimento)
	End Function
	
	
	' DV do nosso número
	Private Function CalculaNossoNumeroDV()
		Dim retorno, posicoes(4), i
		retorno = ""
		
		posicoes(1) = Completa(i_banco.Agencia, 4)
		posicoes(2) = Completa(i_banco.Conta, 5)
		posicoes(3) = Completa(i_banco.Carteira, 3)
		posicoes(4) = Left(i_nossoNumero, 8)
		
		For i = 1 To 4
			retorno = retorno & posicoes(i)
		Next
		
		i_nossoNumeroDV = Mod10(retorno)
	End Function
	
	
	' Auxiliar para completar numeros com zeros a esquerda
	Private Function Completa(ByVal numero, ByVal casas)
		Dim retorno
		retorno = CStr(numero)
		
		While Len(retorno) < casas
			retorno = "0" & retorno
		Wend
		
		Completa = retorno
	End Function
	
	
	' Auxiliar para formatar datas para 10 caracteres DD/MM/AAAA
	Private Function FormataData(ByVal data)
		Dim retorno
		retorno = data
		
		If IsDate(data) Then retorno = Completa(Day(data), 2) & "/" & Completa(Month(data), 2) & "/" & Year(data)
		
		FormataData = retorno
	End Function
	
	' Auxiliar para limpar datas para os códigos DDMMAA
	Private Function LimpaData(ByVal data)
		Dim retorno
		retorno = ""
		
		If IsDate(data) Then retorno = Completa(Day(data), 2) & Completa(Month(data), 2) & Completa(Year(data), 2)
		
		LimpaData = retorno
	End Function
End Class


' ## Classe base para os bancos #########################################################
Class BancoASP
	' ## Campos ##
	Dim i_numero, i_nome, i_carteira, i_agencia, i_conta, i_contaDV
	Dim i_localPagamento
	
	' ## Propriedades ##
	Public Property Get Numero()
		Numero = i_numero
	End Property
	
	Public Property Let Numero(val)
		i_numero = val
	End Property

	
	Public Property Get Nome()
		Nome = i_nome
	End Property
	
	Public Property Let Nome(val)
		i_nome = val
	End Property
	
	
	Public Property Get Carteira()
		Carteira = i_carteira
	End Property
	
	Public Property Let Carteira(val)
		i_carteira = val
	End Property
	
	
	Public Property Get Agencia()
		Agencia = i_agencia
	End Property
	
	Public Property Let Agencia(val)
		i_agencia = val
	End Property
	
	
	Public Property Get Conta()
		Conta = i_conta
	End Property
	
	Public Property Let Conta(val)
		i_conta = val
	End Property
	
	
	Public Property Get ContaDV()
		ContaDV = i_contaDV
	End Property
	
	Public Property Let ContaDV(valor)
		i_contaDV = valor
	End Property
	
	
	Public Property Get LocalPagamento()
		LocalPagamento = i_localPagamento
	End Property
	
	Public Property Let LocalPagamento(val)
		i_localPagamento = val
	End Property
	
	
	' ## Construtor ##
	Private Sub Class_Initialize()
		i_numero = ""
		i_nome = ""
		i_carteira = ""
		i_conta = ""
		i_contaDV = 0
		i_localPagamento = ""
	End Sub
End Class


' ## Classe base para os sacados ########################################################
Class SacadoASP
	' ## Campos ##
	Dim i_nome, i_endereco, i_bairro, i_cep, i_cidade, i_estado, i_cpf
	
	
	' ## Propriedades ##
	Public Property Get Nome()
		Nome = i_nome
	End Property

	Public Property Let Nome(val)
		i_nome = val
	End Property
	
	
	Public Property Get Endereco()
		Endereco = i_endereco
	End Property

	Public Property Let Endereco(val)
		i_endereco = val
	End Property
	
	
	Public Property Get Bairro()
		Bairro = i_bairro
	End Property

	Public Property Let Bairro(val)
		i_bairro = val
	End Property
	
	
	Public Property Get CEP()
		CEP = i_cep
	End Property

	Public Property Let CEP(val)
		i_cep = val
	End Property
	
	
	Public Property Get Cidade()
		Cidade = i_cidade
	End Property

	Public Property Let Cidade(val)
		i_cidade = val
	End Property
	
	
	Public Property Get Estado()
		Estado = i_estado
	End Property

	Public Property Let Estado(val)
		i_estado = val
	End Property
	
	
	Public Property Get CPF()
		CPF = i_cpf
	End Property

	Public Property Let CPF(val)
		i_cpf = val
	End Property
	
	
	' ## Construtor ##
	Private Sub Class_Initialize()
		i_nome = ""
		i_endereco = ""
		i_bairro = ""
		i_cep = ""
		i_cidade = ""
		i_estado = ""
		i_cpf = ""
	End Sub
End Class



' ## Modelo para Banco Itaú ###########################################################
Class BancoItau
	' ## Campos ##
	Dim i_base
	
	
	' ## Propriedades ##
	Public Property Get Numero()
		Numero = i_base.Numero
	End Property

	
	Public Property Get Nome()
		Nome = i_base.Nome
	End Property
	
	
	Public Property Get Carteira()
		Carteira = i_base.Carteira
	End Property
	
	Public Property Let Carteira(val)
		i_base.Carteira = val
	End Property
	
	
	Public Property Get Agencia()
		Agencia = i_base.Agencia
	End Property
	
	Public Property Let Agencia(val)
		i_base.Agencia = val
	End Property
	
	
	Public Property Get Conta()
		Conta = i_base.Conta
	End Property
	
	Public Property Let Conta(val)
		i_base.Conta = val
	End Property
	
	
	Public Property Get ContaDV()
		ContaDV = i_base.ContaDV
	End Property
	
	Public Property Let ContaDV(val)
		i_base.ContaDV = val
	End Property
	
	
	Public Property Get LocalPagamento()
		LocalPagamento = i_base.LocalPagamento
	End Property
	
	Public Property Let LocalPagamento(val)
		i_base.LocalPagamento = val
	End Property
	
	
	Private Sub Class_Initialize()
		Set i_base = New BancoASP
		i_base.Nome = "Banco Itaú SA"
		i_base.Numero = 341
		i_base.Carteira = 175
		i_base.Conta = 18128
		i_base.ContaDV = 0
		i_base.LocalPagamento = "ATE O VENCIMENTO PAGUE PREFERENCIALMENTE NO ITAU OU BANERJ<br />" & vbCrLf _
							  & "APOS O VENCIMENTO PAGUE SOMENTE NO ITAU OU BANERJ"
	End Sub
	
	Private Sub Class_Terminate()
		Set i_base = Nothing
	End Sub
End Class
%>