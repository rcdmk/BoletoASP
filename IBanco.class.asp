<%
' ####################################################################################
'
' Gerador de Boletos em ASP por RCDMK <rcdmk@hotmail.com>
' Inнcio em 29/04/2012
'
' Implementaзгo bбsica de um banco. Esta classe serve como modelo para implementaзгo
' de bancos, que tem cбlculos especнficos para dнgitos e dados extras para os campos
' livres do boleto
'
' ## Lisenзa #########################################################################
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
' ## Classe base para os bancos #########################################################
Class IBancoASP 'extends Interface
	' ## Campos ##
	Dim i_pai


	' ## Interface ##
	Public Property Get Pai()
		Set Pai = i_pai
	End Property
	
	Public Property Set Pai(val)
		Set i_pai = val
	End Property
	
	
	Public Property Get Implementacao()
		Set Implementacao = i_pai.Implementacao
	End Property
	
	Public Property Set Implementacao(val)
		Set i_pai.Implementacao = val
	End Property
	
	
	Public Property Get Obrigatorios()
		Obrigatorios = i_pai.Obrigatorios
	End Property
	
	
	Public Property Get Verifica()
		Verifica = i_pai.Verifica()
	End Property
	
	
	' ## Propriedades ##
	Public Property Get Boleto()
	End Property
	
	Public Property Set Boleto(val)
	End Property
	
	
	Public Property Get Numero()
	End Property
	
	Public Property Let Numero(val)
	End Property

	
	Public Property Get Nome()
	End Property
	
	Public Property Let Nome(val)
	End Property
	
	
	Public Property Get Carteira()
	End Property
	
	Public Property Let Carteira(val)
	End Property
	
	
	Public Property Get Agencia()
	End Property
	
	Public Property Let Agencia(val)
	End Property
	
	
	Public Property Get Conta()
	End Property
	
	Public Property Let Conta(val)
	End Property
	
	
	Public Property Get ContaDV()
	End Property
	
	Public Property Let ContaDV(valor)
	End Property
	
	
	Public Property Get LocalPagamento()
	End Property
	
	Public Property Let LocalPagamento(val)
	End Property
	
	
	' ## Construtor ##
	Private Sub Class_Initialize()
		Set i_pai = New Interface
		i_pai.Obrigatorios = Array("Boleto", "Numero", "Nome", "Carteira", "Agencia", "Conta", "ContaDV", "LocalPagamento", "CalculaNossoNumeroDV", "NumCodigoBarras", "LinhaDigitavel")
	End Sub
	
	Private Sub Class_Terminate()
		Set i_pai = Nothing
	End Sub
	
	
	' ## Mйtodos ##
	Public Function CalculaNossoNumeroDV()
	End Function
	
	Public Function NumCodigoBarras()
	End Function
	
	Public Function LinhaDigitavel()
	End Function
End Class
%>