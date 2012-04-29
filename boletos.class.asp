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
	Dim i_banco, i_nossoNumero, i_sacado
	Dim i_valor, i_dataDocumento, i_dataVencimento, i_percMulta, i_percJuros


	' ## Propriedades ##
	' Valor do documento
	Public Property Get Banco
		Set Banco = i_banco
	End Property
	
	Public Property Let Banco(val)
		Set i_banco = val
	End Property
	
	
	Public Property Get NossoNumero
		NossoNumero = i_nossoNumero
	End Property
	
	Public Property Let NossoNumero(val)
		i_nossoNumero = val
	End Property
	
	
	Public Property Get Sacado
		Set Sacado = i_sacado
	End Property
	
	Public Property Let Sacado(val)
		Set i_sacado = val
	End Property
	
	
	Public Property Get Valor
		Valor = i_valor
	End Property
	
	Public Property Let Valor(val)
		i_valor = val
	End Property
	
	
	Public Property Get DataDocumento
		DataDocumento = i_dataDocumento
	End Property
	
	Public Property Let DataDocumento(val)
		i_dataDocumento = val
	End Property
	
	
	Public Property Get DataVencimento
		DataVencimento = i_dataVencimento
	End Property
	
	Public Property Let DataVencimento(val)
		i_dataVencimento = val
	End Property
	
	
	Public Property Get PercMulta
		PercMulta = i_percMulta
	End Property
	
	Public Property Let PercMulta(val)
		i_percMulta = val
	End Property
	
	
	Public Property Get PercJuros
		PercJuros = i_percJuros
	End Property
	
	Public Property Let PercJuros(val)
		i_percJuros = val
	End Property
	
	
	' ## Construtor ##
	Private Sub Class_Initialize()
		i_nossoNumero = "000000"
		
		i_valor = 0
		i_dataDocumento = date
		i_dataVencimento = dateAdd("d", 3, date)
		i_percMulta = 2
		i_percJuros = 0.33
		
		Set i_banco = New BancoASP
		Set i_sacado = New SacadoASP
	End Sub
	
	Private Sub Class_Terminate()
		Set i_banco = Nothing
	End Sub
	
	' ## Metodos ##
	Public Function Mod10(ByVal strNumero)
		Dim DV, tamanho, i, j
		Dim num, soma, somaTotal
		
		strNumero = CStr(strNumero)
		tamanho = Len(strNumero)
		
		For i = tamanho - 1 to 0 Step -1
			num = Mid(strNumero, i + 1, 1)
			If i And 1 Then num = num * 2
			
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
	
	Public Function Mod11(Byval strNumero, ByVal tipo)
		Dim DV, tamanho, i
		Dim num, soma
		
		strNumero = CStr(strNumero)
		tamanho = Len(strNumero)
		soma = 0
		
		For i = tamanho - 1 to 0 Step -1
			num = Mid(strNumero, i + 1, 1)
			num = num * ((tamanho + 1) - i)
			
			soma = soma + CInt(num)
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
		End If
		
		Mod11 = DV
	End Function
End Class


' ## Classe base para os bancos #########################################################
Class BancoASP
	' ## Campos ##
	Dim i_numero, i_nome, i_carteira, i_conta, i_contaDV
	Dim i_localPagamento
	
	' ## Propriedades ##
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
		i_numero = ""
		i_nome = ""
		i_carteira = ""
		i_conta = ""
		i_contaDV = 0
		i_localPagamento = ""
	End Sub
End Class


' ## Classe base para os sacados ########################################################
Class SacadoASP
	' ## Campos ##
	Dim i_nome, i_endereco, i_bairro, i_cep, i_cidade, i_estado
	
	
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
	
	
	' ## Construtor ##
	Private Sub Class_Initialize()
		i_nome = ""
		i_endereco = ""
		i_bairro = ""
		i_cep = ""
		i_cidade = ""
		i_estado = ""
	End Sub
End Class



'
Class BancoItau
	' ## Campos ##
	Dim i_base
	Dim i_numero, i_nome, i_carteira, i_conta, i_contaDV
	Dim i_localPagamento
	
	
	' ## Propriedades ##
	Public Property Get Numero()
		Numero = i_base.Numero
	End Property

	
	Public Property Get Nome()
		Nome = i_base.Nome
	End Property
	
	
	Public Property Get Carteira()
		Carteira = i_base.Carteira
	End Property
	
	Public Property Let Carteira(val)
		i_base.Carteira = val
	End Property
	
	
	Public Property Get Conta()
		Conta = i_base.Conta
	End Property
	
	Public Property Let Conta(val)
		i_base.Conta = val
	End Property
	
	
	Public Property Get ContaDV()
		ContaDV = i_base.ContaDV
	End Property
	
	Public Property Let ContaDV(val)
		i_base.ContaDV = val
	End Property
	
	
	Public Property Get LocalPagamento()
		LocalPagamento = i_base.LocalPagamento
	End Property
	
	Public Property Let LocalPagamento(val)
		i_base.LocalPagamento = val
	End Property
	
	
	Private Sub Class_Initialize()
		Set i_base = New BancoASP
		i_base.Nome = "Banco Itaú SA"
		i_base.Numero = 341
		i_base.Carteira = 175
		i_base.Conta = 18128
		i_base.ContaDV = 0
		i_base.LocalPagamento = "ATE O VENCIMENTO PAGUE PREFERENCIALMENTE NO ITAU<br />" & vbCrLf _
							  & "APOS O VENCIMENTO PAGUE SOMENTE NO ITAU"
	End Sub
	
	Private Sub Class_Terminate()
		Set i_base = Nothing
	End Sub
End Class
%>