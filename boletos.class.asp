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
	Dim i_pastaImagens, i_layout, i_numeroDocumento, i_cedenteNome, i_especie, i_aceite, i_sacadorNome
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
	
	
	Public Property Get Layout()
		Layout = i_layout
	End Property
	
	Public Property Let Layout(val)
		i_layout = Replace(val, "\", "/")
	End Property
	
	
	Public Property Get NumeroDocumento()
		NumeroDocumento = i_numeroDocumento
	End Property
	
	Public Property Let NumeroDocumento(val)
		i_numeroDocumento = val
	End Property
	
	
	Public Property Get CedenteNome()
		CedenteNome = i_cedenteNome
	End Property
	
	Public Property Let CedenteNome(val)
		i_cedenteNome = val
	End Property
	
	
	Public Property Get Especie()
		Especie = i_especie
	End Property
	
	Public Property Let Especie(val)
		i_especie = val
	End Property
	
	
	Public Property Get Aceite()
		Aceite = i_aceite
	End Property
	
	Public Property Let Aceite(val)
		i_aceite = val
	End Property
	
	
	Public Property Get SacadorNome()
		SacadorNome = i_sacadorNome
	End Property
	
	Public Property Let SacadorNome(val)
		i_sacadorNome = val
	End Property
	
	
	Public Property Get Banco()
		Set Banco = i_banco
		Set i_banco.Boleto = Me
	End Property
	
	Public Property Set Banco(val)
		Set i_banco = val
	End Property
	
	
	Public Property Get NossoNumero()
		NossoNumero = i_nossoNumero
	End Property
	
	Public Property Let NossoNumero(val)
		i_nossoNumero = Completa(val, 8)
		calculaNossoNumeroDV
	End Property
	
	
	Public Property Get NossoNumeroDV()
		calculaNossoNumeroDV
		NossoNumeroDV = i_nossoNumeroDV
	End Property

	
	Public Property Get Sacado()
		Set Sacado = i_sacado
	End Property
	
	Public Property Set Sacado(val)
		Set i_sacado = val
	End Property
	
	
	Public Property Get Moeda()
		Moeda = i_moeda
	End Property
	
	
	Public Property Get Fator()
		Fator = i_fator
	End Property

	
	' Valor do documento
	Public Property Get ValorDocumento()
		ValorDocumento = i_valorDocumento
	End Property
	
	Public Property Let ValorDocumento(val)
		i_valorDocumento = val
	End Property
	
	
	Public Property Get DataDocumento()
		DataDocumento = i_dataDocumento
	End Property
	
	Public Property Let DataDocumento(val)
		i_dataDocumento = val
	End Property
	
	
	Public Property Get DataProcessamento()
		DataProcessamento = i_dataProcessamento
	End Property
	
	Public Property Let DataProcessamento(val)
		i_dataProcessamento = val
	End Property
	
	
	Public Property Get DataVencimento()
		DataVencimento = i_dataVencimento
	End Property
	
	Public Property Let DataVencimento(val)
		i_dataVencimento = val
		calculaFator
	End Property
	
	
	Public Property Get PercMulta()
		PercMulta = i_percMulta
	End Property
	
	Public Property Let PercMulta(val)
		i_percMulta = val
	End Property
	
	
	Public Property Get PercJuros()
		PercJuros = i_percJuros
	End Property
	
	Public Property Let PercJuros(val)
		i_percJuros = val
	End Property
	
	
	Public Property Get Instrucoes()
		Instrucoes = i_instrucoes
	End Property
	
	Public Property Let Instrucoes(val)
		i_instrucoes = val
	End Property
	
	
	' ## Construtor ##
	Private Sub Class_Initialize()
		i_pastaImagens = "imagens"
		i_layout = "layout.asp"
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
		
		calculaFator
		calculaNossoNumeroDV
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
	
		DV = (soma * 10) Mod 11
		
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
		LinhaDigitavel = i_banco.LinhaDigitavel
	End Function
	
	
	' Número do código de barras
	Public Function NumCodigoBarras()
		NumCodigoBarras = i_banco.NumCodigoBarras()
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
	
	
	' Monta e retorna o layout do boleto em HTML
	Public Function HTML()
		Dim erro, fso
		Set fso = Server.CreateObject("Scripting.FileSystemObject")
		erro = Not fso.FileExists(Server.MapPath(i_layout))
		Set fso = Nothing
		
		If erro then
			Err.Raise 17, "Caminho inválido", "O arquivo de layout informado não existe."
		Else		
			Set Session("BoletoASP") = Me
			HTML = Server.Execute(i_layout)
		End If
	End Function
	
	
	' Escreve o HTML do boleto direto na página
	Public Function Write()
		Response.Write HTML()
	End Function
	
	
	' Auxiliar para completar numeros com zeros a esquerda
	Public Function Completa(ByVal numero, ByVal casas)
		Dim retorno
		retorno = CStr(numero)
		
		While Len(retorno) < casas
			retorno = "0" & retorno
		Wend
		
		Completa = retorno
	End Function
	
	
	' Auxiliar para formatar datas para 10 caracteres DD/MM/AAAA
	Public Function FormataData(ByVal data)
		Dim retorno
		retorno = data
		
		If IsDate(data) Then retorno = Completa(Day(data), 2) & "/" & Completa(Month(data), 2) & "/" & Year(data)
		
		FormataData = retorno
	End Function
	
	' Auxiliar para limpar datas para os códigos DDMMAA
	Public Function LimpaData(ByVal data)
		Dim retorno
		retorno = ""
		
		If IsDate(data) Then retorno = Completa(Day(data), 2) & Completa(Month(data), 2) & Completa(Year(data), 2)
		
		LimpaData = retorno
	End Function
	
	
	' ## Métodos privados ##
	' Fator do vencimento
	Private Function calculaFator()
		i_fator = DateDiff("d", i_dataBase, i_dataVencimento)
	End Function
	
	
	' DV do nosso número
	Private Function calculaNossoNumeroDV()
		i_nossoNumeroDV = i_banco.CalculaNossoNumeroDV()
		calculaNossoNumeroDV = i_nossoNumeroDV
	End Function
End Class


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
	
	
	' ## Métodos ##
	Public Function CalculaNossoNumeroDV()
	End Function
	
	Public Function NumCodigoBarras()
	End Function
	
	Public Function LinhaDigitavel()
	End Function
End Class


' ## Implementação básica ##
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
	
	
	' ## Métodos ##
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