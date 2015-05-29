<%
' ####################################################################################
'
' Gerador de Boletos em ASP por RCDMK <rcdmk[at]hotmail[dot]com>
' Início em 29/04/2012
'
' Esta classe representa os dados básicos do sacado
'
' ## Lisença #########################################################################
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
%>