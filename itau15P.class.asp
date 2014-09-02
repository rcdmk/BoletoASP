<%
' ####################################################################################
'
' Gerador de Boletos em ASP por RCDMK <rcdmk@hotmail.com>
' Início em 29/04/2012
'
' Esta classe representa os padrões de boleto para o banco Itaú com as
' carteiras de nosso número com 15 posições
'
' ## Lisença #########################################################################
'
' The MIT License (MIT)  - http://opensource.org/licenses/MIT
' 
' Copyright (c) 2012 Ricardo Souza (RCDMK) - rcdmk@hotmail.com
' 
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
' 
' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.
' 
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
' THE SOFTWARE.
'
' ####################################################################################
'
' ## Modelo para Banco Itaú ##
' Depende de: IBancoASP (já incluso no banco base, que é utilizado no boleto)
Class BancoItau15P 'implements IBancoASP
	' ## Campos ##
	Dim i_pai, i_boleto, i_codigoCliente
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
	
	
	Public Property Get CodigoCliente()
		CodigoCliente = i_codigoCliente
	End Property
	
	Public Property Let CodigoCliente(val)
		i_codigoCliente = val
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
		i_carteira = 198
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
		
		posicoes(1) = Boleto.Completa(Carteira, 3)
		posicoes(2) = Left(Boleto.NossoNumero, 8)
		posicoes(3) = Boleto.Completa(Boleto.NumeroDocumento, 7)
		posicoes(4) = Boleto.Completa(i_codigoCliente, 5)
		
		For i = 1 To 4
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
		
		posicoes(20) 	= Boleto.Completa(Carteira, 3)
		posicoes(23) 	= Left(Boleto.NossoNumero, 8)
		posicoes(31) 	= Boleto.Completa(Boleto.NumeroDocumento, 7)
		posicoes(38) 	= Boleto.Completa(i_codigoCliente, 5)
		posicoes(43) 	= Boleto.NossoNumeroDV
		posicoes(44)	= "0"
		
		For i = 1 To 44
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
		posicoes(2)		= Boleto.Moeda					' Moeda
		posicoes(3)		= Mid(numero, 20, 3)			' Carteira
		posicoes(4)		= Mid(numero, 23, 2)			' 2 primeiros dígitos do nosso número
		posicoes(5)		= "" 							' DV do primeiro grupo
		
		posicoes(6)		= Mid(numero, 25, 6)			' Restante do nosso número
		posicoes(7)		= Mid(numero, 31, 4)			' 4 primeiros dígitos do número do documento
		posicoes(8)		= "" 							' DV do segundo grupo
		
		posicoes(9)		= Mid(numero, 35, 3) 			' Restante do número do documento
		posicoes(10)	= Mid(numero, 38, 5) 			' Código do cliente
		posicoes(11)	= Boleto.NossoNumeroDV			' DV do nosso número (Carteira/Nosso Número (sem o DAC) / Seu Número (sem o DAC) / Código do Cliente)
		posicoes(12)	= "0" 							
		posicoes(13)	= "" 							' DV do terceiro grupo
		
		posicoes(14)	= Mid(numero, 5, 1) 			' DV do código de barras
		
		posicoes(15)	= Mid(numero, 6, 4) 			' Fator de vencimento
		posicoes(16)	= Mid(numero, 10, 10) 			' Valor do documento
		
		' Calculando DVs
		posicoes(5) 	= Boleto.Mod10(posicoes(1) & posicoes(2) & posicoes(3) & posicoes(4))
		posicoes(8) 	= Boleto.Mod10(posicoes(6) & posicoes(7))
		posicoes(13) 	= Boleto.Mod10(posicoes(9) & posicoes(10) & posicoes(11) & posicoes(12))
		
		For i = 1 To 16
			retorno = retorno & posicoes(i)
		Next

		LinhaDigitavel = Left(retorno, 5) & "." & Mid(retorno, 6, 5) & " " & Mid(retorno, 11, 5) & "." & Mid(retorno, 16, 6) & " " & Mid(retorno, 22, 5) & "." & Mid(retorno, 27, 6) & " " & Mid(retorno, 33, 1) & " " & Mid(retorno, 34)
	End Function
End Class
%>