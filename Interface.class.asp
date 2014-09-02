<%
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