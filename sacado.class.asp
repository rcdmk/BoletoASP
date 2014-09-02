<%
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