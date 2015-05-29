<%
' ####################################################################################
'
' Gerador de Boletos em ASP por RCDMK <rcdmk[at]hotmail[dot]com>
' Inнcio em 29/04/2012
'
' Implementaзгo bбsica de um banco. Deve ser tulizado como modelo para criar bancos,
' que definem cбlculos especнficos para dнgitos, linha digitбvel e cуdigo de barras
'
' ## Lisenзa #########################################################################
'
' The MIT License (MIT)  - http://opensource.org/licenses/MIT
' 
' Copyright (c) 2015 RCDMK - rcdmk[at]hotmail[dot]com
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
' ## Modelo para Banco Real ##
' Depende de: IBancoASP (jб incluso no banco base, que й utilizado no boleto)
' ## Implementaзгo bбsica ##
Class BancoASP 'implements IBancoASP
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
		
		i_nome = "Banco"
		i_Numero = "000"
		i_carteira = "00"
		i_conta = "00000"
		i_contaDV = 0
		i_localPagamento = ""
	End Sub
	
	Private Sub Class_Terminate()
		Set Implementacao = Nothing
		Set i_base = Nothing
	End Sub
	
	
	' ## Mйtodos ##
	Public Function CalculaNossoNumeroDV()
		CalculaNossoNumeroDV = 0
	End Function
	
	
	Public Function NumCodigoBarras()
		NumCodigoBarras = Boleto.Completa("0", 44)
	End Function
	
	
	Public Function LinhaDigitavel()
		Dim retorno
		retorno = Boleto.Completa("0", 44)

		LinhaDigitavel = Left(retorno, 5) & "." & Mid(retorno, 6, 5) & " " & Mid(retorno, 11, 5) & "." & Mid(retorno, 16, 6) & " " & Mid(retorno, 22, 5) & "." & Mid(retorno, 27, 6) & " " & Mid(retorno, 33, 1) & " " & Mid(retorno, 34)
	End Function
End Class
%>