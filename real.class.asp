<%
'
' Esta classe representa os padrões de boleto para o banco Real com as
' carteiras padrão
'
' ## Modelo para Banco Real ##
Class BancoReal 'implements IBancoASP
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
		i_contaDV = CalculaContaDV
	End Property
	
	
	Public Property Get ContaDV()
		ContaDV = i_contaDV
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
		
		i_nome = "Banco Real SA"
		i_Numero = 356
		i_carteira = ""
		i_conta = "00000"
		i_contaDV = 0
		i_localPagamento = "PAGAVEL EM QUALQUER AGENCIA BANCARIA"
	End Sub
	
	Private Sub Class_Terminate()
		Set Implementacao = Nothing
		Set i_base = Nothing
	End Sub
	
	
	' ## Métodos ##
	Private Function CalculaContaDV()
		Dim retorno, posicoes(2), i
		retorno = ""
		
		posicoes(1) = Boleto.Completa(Agencia, 4)
		posicoes(2) = Boleto.Completa(Conta, 7)
		
		For i = 1 To 2
			retorno = retorno & posicoes(i)
		Next
		
		CalculaContaDV = Boleto.Mod11(retorno, "")
	End Function
	
	
	Public Function CalculaNossoNumeroDV()
		Dim retorno, posicoes(3), i
		retorno = ""
		
		posicoes(1) = Boleto.Completa(Boleto.NossoNumero, 15)
		posicoes(2) = Boleto.Completa(Agencia, 4)
		posicoes(3) = Boleto.Completa(Conta, 7)
		
		For i = 1 To 3
			retorno = retorno & posicoes(i)
		Next
		
		CalculaNossoNumeroDV = Boleto.Mod10(retorno)
	End Function
	
	
	Public Function NumCodigoBarras()
		Dim retorno, i, posicoes(44)
		retorno = ""
		
		CalculaNossoNumeroDV
		
		posicoes(1) 	= Boleto.Completa(Numero, 3)
		posicoes(4) 	= Boleto.Moeda
		posicoes(5) 	= "" ' DV do código Mod11
		
		' Se o valor for maior que 100 milhões, ignora-se o fator de vencimento
		If Boleto.ValorDocumento >= 100000000 Then
			posicoes(6)		= Boleto.Completa(CLng(Boleto.ValorDocumento * 100), 14)
		Else
			posicoes(6) 	= Boleto.Completa(Boleto.Fator, 4)
			posicoes(10)	= Boleto.Completa(CLng(Boleto.ValorDocumento * 100), 10)
		End If
		
		posicoes(20) 	= Boleto.Completa(Agencia, 4)
		posicoes(24) 	= Boleto.Completa(Conta, 7)
		posicoes(31) 	= Boleto.NossoNumeroDV
		posicoes(32) 	= Boleto.Completa(Boleto.NossoNumero, 13)
		
		For i = 1 To 44
			retorno = retorno & posicoes(i)
		Next
		
		posicoes(5) = Boleto.Mod11(retorno, MOD11_BARRAS)
		
		NumCodigoBarras = Left(retorno, 4) & posicoes(5) & Right(retorno, 39)
	End Function
	
	
	Public Function LinhaDigitavel()
		Dim retorno, i, posicoes(14), numero
		
		retorno = ""
		numero = NumCodigoBarras()
		
		posicoes(1)		= Left(numero, 3) 				' Número do banco
		posicoes(2)		= Boleto.Moeda					' Moeda
		posicoes(3)		= Mid(numero, 20, 4)			' Agência
		posicoes(4)		= Mid(numero, 24, 1)			' 1 primeiro dígito da conta
		posicoes(5)		= "" 							' DV do primeiro grupo
		
		posicoes(6)		= Mid(numero, 25, 6)			' Restante da conta corrente
		posicoes(7)		= Boleto.NossoNumeroDV			' Dígito do nosso número
		posicoes(8)		= Mid(numero, 32, 3)			' 3 Primeiros digitos do nosso numero
		posicoes(9)		= "" 							' DV do segundo grupo
		
		posicoes(10)	= Mid(numero, 35, 11) 			' Restante do nosso numero
		posicoes(11)	= "" 							' DV do terceiro grupo
		
		posicoes(12)	= Mid(numero, 5, 1) 			' DV do código de barras
		
		posicoes(13)	= Mid(numero, 6, 4) 			' Fator de vencimento
		posicoes(14)	= Mid(numero, 10, 10) 			' Valor do documento
		
		' Calculando DVs
		posicoes(5) 	= Boleto.Mod10(posicoes(1) & posicoes(2) & posicoes(3) & posicoes(4))
		posicoes(9) 	= Boleto.Mod10(posicoes(6) & posicoes(7) & posicoes(8))
		posicoes(11) 	= Boleto.Mod10(posicoes(10))
		
		For i = 1 To 14
			retorno = retorno & posicoes(i)
		Next

		LinhaDigitavel = Left(retorno, 5) & "." & Mid(retorno, 6, 5) & " " & Mid(retorno, 11, 5) & "." & Mid(retorno, 16, 6) & " " & Mid(retorno, 22, 5) & "." & Mid(retorno, 27, 6) & " " & Mid(retorno, 33, 1) & " " & Mid(retorno, 34)
	End Function
End Class
%>