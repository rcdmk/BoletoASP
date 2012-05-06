<%
'
' Esta classe representa os padrões de boleto para o banco Itaú com as
' carteiras padrão
'
' ## Modelo para Banco Itaú ##
Class BancoItau 'implements IBancoASP
	' ## Campos ##
	Dim i_pai, i_boleto
	Dim i_numero, i_nome, i_carteira, i_agencia, i_conta, i_contaDV
	Dim i_localPagamento
	
	
	' ## Propriedades ##
	Public Interface
	
	Public Property Get Boleto()
		Set Boleto = i_boleto
	End Property
	
	Public Property Set Boleto(val)
		Set i_boleto = val
	End Property
	
	
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
		Set i_boleto = New BoletoASP
		
		Set Interface = New IBancoASP
		Set Interface.Implementacao = Me
		Interface.Verifica()
		
		i_nome = "Banco Itaú SA"
		i_Numero = 341
		i_carteira = 175
		i_conta = "00000"
		i_contaDV = 0
		i_localPagamento = "ATE O VENCIMENTO PAGUE PREFERENCIALMENTE NO ITAU OU BANERJ<br />" & vbCrLf _
						 & "APOS O VENCIMENTO PAGUE SOMENTE NO ITAU OU BANERJ"
	End Sub
	
	Private Sub Class_Terminate()
		Set Implementacao = Nothing
		Set i_base = Nothing
	End Sub
	
	
	' ## Métodos ##
	Public Function CalculaNossoNumeroDV()
		Dim retorno, posicoes(4), i
		retorno = ""
		
		posicoes(1) = Boleto.Completa(Agencia, 4)
		posicoes(2) = Boleto.Completa(Conta, 5)
		posicoes(3) = Boleto.Completa(Carteira, 3)
		posicoes(4) = Left(Boleto.NossoNumero, 8)
		
		For i = 1 To 4
			retorno = retorno & posicoes(i)
		Next
		
		CalculaNossoNumeroDV = Boleto.Mod10(retorno)
	End Function
	
	
	Public Function NumCodigoBarras()
		Dim retorno, i, posicoes(43)
		retorno = ""
		
		CalculaNossoNumeroDV
		
		posicoes(1) 	= Boleto.Completa(Numero, 3)
		posicoes(4) 	= Boleto.Moeda
		posicoes(5) 	= "" ' DV do código Mod11
		
		' Se o valor for maior que 100 milhões, ignora-se o fator de vencimento
		If Boleto.ValorDocumento >= 100000000 Then
			posicoes(6)		= Boleto.Completa(CInt(Boleto.ValorDocumento * 100), 14)
		Else
			posicoes(6) 	= Boleto.Completa(Boleto.Fator, 4)
			posicoes(10)	= Boleto.Completa(CInt(Boleto.ValorDocumento * 100), 10)
		End If
		
		posicoes(20) 	= Boleto.Completa(Carteira, 3)
		posicoes(23) 	= Left(Boleto.NossoNumero, 8)
		posicoes(31) 	= Boleto.NossoNumeroDV
		posicoes(32) 	= Boleto.Completa(Agencia, 4)
		posicoes(36) 	= Boleto.Completa(Conta, 5)
		posicoes(41) 	= ContaDV
		posicoes(42) 	= "000"
		
		For i = 1 To 43
			retorno = retorno & posicoes(i)
		Next
		
		posicoes(5) = Boleto.Mod11(retorno, MOD11_BARRAS)
		
		NumCodigoBarras = Left(retorno, 4) & posicoes(5) & Right(retorno, 39)
	End Function
	
	
	Public Function LinhaDigitavel()
		Dim retorno, i, posicoes(16), numero
		
		retorno = ""
		numero = NumCodigoBarras()
		
		posicoes(1)		= Left(numero, 3) 				' Número do banco
		posicoes(2)		= Boleto.Moeda			' Moeda
		posicoes(3)		= Mid(numero, 20, 3)			' Carteira
		posicoes(4)		= Mid(numero, 23, 2)			' 2 primeiros dígitos do nosso número
		posicoes(5)		= "" 							' DV do primeiro grupo
		
		posicoes(6)		= Mid(numero, 25, 6)			' Restante do nosso número
		posicoes(7)		= Boleto.NossoNumeroDV	' Dígito do nosso número
		posicoes(8)		= Mid(numero, 32, 3)			' 3 Primeiros numeros da agência
		posicoes(9)		= "" 							' DV do segundo grupo
		
		posicoes(10)	= Mid(numero, 35, 1) 			' Restante da agéncia
		posicoes(11)	= Mid(numero, 36, 6)			' Conta + DV
		posicoes(12)	= "000"
		posicoes(13)	= "" 							' DV do terceiro grupo
		
		posicoes(14)	= Mid(numero, 5, 1) 			' DV do código de barras
		
		posicoes(15)	= Mid(numero, 6, 4) 			' Fator de vencimento
		posicoes(16)	= Mid(numero, 10, 10) 			' Valor do documento
		
		' Calculando DVs
		posicoes(5) 	= Boleto.Mod10(posicoes(1) & posicoes(2) & posicoes(3) & posicoes(4))
		posicoes(9) 	= Boleto.Mod10(posicoes(6) & posicoes(7) & posicoes(8))
		posicoes(13) 	= Boleto.Mod10(posicoes(10) & posicoes(11) & posicoes(12))
		
		For i = 1 To 16
			retorno = retorno & posicoes(i)
		Next

		LinhaDigitavel = Left(retorno, 5) & "." & Mid(retorno, 6, 5) & " " & Mid(retorno, 11, 5) & "." & Mid(retorno, 16, 6) & " " & Mid(retorno, 22, 5) & "." & Mid(retorno, 27, 6) & " " & Mid(retorno, 33, 1) & " " & Mid(retorno, 34)
	End Function
End Class
%>