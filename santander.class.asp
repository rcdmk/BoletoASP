<%
' ####################################################################################
'
' Gerador de Boletos em ASP por RCDMK <rcdmk@hotmail.com>
' Início em 29/04/2012
'
' Esta classe representa os padrões de boleto para o banco Santander com as
' carteiras padrão
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
' ## Modelo para Banco Santander #####################################################
' Depende de: IBancoASP (já incluso no banco base, que é utilizado no boleto)
Class BancoSantander 'implements IBancoASP
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
		i_contaDV = "" ' não possui DV
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
		
		i_nome = "Banco Santander (Brasil) S.A."
		i_Numero = "033"
		i_carteira = "101"
		i_agencia = "00000-0"
		i_conta = "00000"
		i_contaDV = "" ' não possui DV na conta
		i_localPagamento = "Pagar preferencialmente no Grupo Santander - GC"
	End Sub
	
	Private Sub Class_Terminate()
		Set Implementacao = Nothing
		Set i_base = Nothing
	End Sub
	
	
	' ## Métodos ##
	Public Function CalculaNossoNumeroDV()
		Dim retorno

		retorno = Boleto.Completa(Boleto.NossoNumero, 12)		
	
		CalculaNossoNumeroDV = Boleto.Mod11(retorno, MOD11_BRADESCO)
	End Function
	
	
	Public Function NumCodigoBarras()
		Dim retorno, i, posicoes(44), regex
		retorno = ""
		
		CalculaNossoNumeroDV
		
		Set regex = new RegExp
		regex.Pattern = "\D"
		regex.Global = True
		
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
		
		posicoes(20) 	= Boleto.Moeda
		posicoes(21) 	= Boleto.Completa(Left(regex.Replace(Conta, ""), 7), 7)
		posicoes(28) 	= Boleto.Completa(Left(Boleto.NossoNumero, 12), 12)
		posicoes(40) 	= Boleto.NossoNumeroDV
		posicoes(41) 	= 0 ' IOS – Seguradoras (Se 7% informar 7. Limitado a 9%) Demais clientes usar 0 (zero)
		posicoes(42)	= Boleto.Completa(Carteira, 3)
		
		Set regex = Nothing
		
		For i = 1 To 44
			retorno = retorno & posicoes(i)
		Next
		
		posicoes(5) = Boleto.Mod11(retorno, MOD11_BARRAS)
		
		NumCodigoBarras = Left(retorno, 4) & posicoes(5) & Right(retorno, 39)
	End Function
	
	
	Public Function LinhaDigitavel()
		Dim retorno, i, posicoes(15), numero
		
		retorno = ""
		numero = NumCodigoBarras()
		
		posicoes(1)		= Left(numero, 3) 				' Número do banco
		posicoes(2)		= Boleto.Moeda					' Moeda
		posicoes(3)		= Boleto.Moeda					' Moeda
		posicoes(4)		= Mid(numero, 21, 4)			' 4 primeiros da conta (cód. cedente)
		posicoes(5)		= "" 							' DV do primeiro grupo
		
		posicoes(6)		= Mid(numero, 25, 3)			' Restante da conta (cód. cedente)
		posicoes(7)		= Mid(numero, 28, 7)			' 7 Primeiros digitos do nosso numero
		posicoes(8)		= "" 							' DV do segundo grupo
		
		posicoes(9)		= Mid(numero, 35, 6) 			' Restante do nosso numero com DV
		posicoes(10)	= 0 							' IOS – Seguradoras (Se 7% informar 7. Limitado a 9%) Demais clientes usar 0 (zero)
		posicoes(11)	= Mid(numero, 42, 3)  			' Carteira
		posicoes(12)	= "" 							' DV do terceiro grupo
		
		posicoes(13)	= Mid(numero, 5, 1) 			' DV do código de barras
		
		posicoes(14)	= Mid(numero, 6, 4) 			' Fator de vencimento
		posicoes(15)	= Mid(numero, 10, 10) 			' Valor do documento
		
		' Calculando DVs
		posicoes(5) 	= Boleto.Mod10(posicoes(1) & posicoes(2) & posicoes(3) & posicoes(4))
		posicoes(8) 	= Boleto.Mod10(posicoes(6) & posicoes(7))
		posicoes(12) 	= Boleto.Mod10(posicoes(9) & posicoes(10) & posicoes(11))
		
		For i = 1 To 15
			retorno = retorno & posicoes(i)
		Next

		LinhaDigitavel = Left(retorno, 5) & "." & Mid(retorno, 6, 5) & " " & Mid(retorno, 11, 5) & "." & Mid(retorno, 16, 6) & " " & Mid(retorno, 22, 5) & "." & Mid(retorno, 27, 6) & " " & Mid(retorno, 33, 1) & " " & Mid(retorno, 34)
	End Function
End Class
%>