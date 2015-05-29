<%
' ####################################################################################
'
' Gerador de Boletos em ASP por RCDMK <rcdmk[at]hotmail[dot]com>
' Início em 29/04/2012
'
' Esta classe representa a base de uma implementação de "interfaces" em VBScript
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
' ## Classes base de código para interfaces #############################################
Class Interface
	' ## Campos ##
	Dim i_implementacao, i_obrigatorios
	
	
	' ## Propriedades ##
	Public Property Get Implementacao()
		Set Implementacao = i_implementacao
	End Property
	
	Public Property Set Implementacao(val)
		Set i_implementacao = val
	End Property
	
	
	Public Property Get Obrigatorios()
		Obrigatorios = i_obrigatorios
	End Property
	
	Public Property Let Obrigatorios(val)
		i_obrigatorios = val
	End Property
	
	
	' ## Construtor ##
	Private Sub Class_Initialize()
		i_obrigatorios = Array()
	End Sub
	
	Private Sub Class_Terminate()
		Set i_implementacao = Nothing
	End Sub
	
	
	' ## Métodos ##
	Public Function Verifica()
		Dim prop, check, i, resultado
		resultado = true
		
		'On Error Resume Next
		
		For i = 0 To UBound(i_obrigatorios)
			prop = i_obrigatorios(i)
			check = TypeName(Eval("i_implementacao." & prop))
			
			If Err.number <> 0 and Err.number <> 5 and Err.number <> 450 Then
				resultado = false
				Err.Clear()
				Exit For
			End If
		Next
		
		'On Error GoTo 0
		
		Verifica = resultado
		
		If Not resultado Then
			'Err.Raise 17, "Implementação de Interface", "A Interface não foi corretamente implementada. Falta a implementação de " & prop & " em " & TypeName(i_implementacao) & "."
		End If
	End Function
End Class
%>