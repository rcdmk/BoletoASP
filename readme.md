# Gerador de Boletos HTML em ASP por RCDMK <rcdmk@hotmail.com>

Esta "biblioteca" pretente auxiliar no desenvolvimento e implantação de cobrança
via boletos.

Inicialmente tem apenas a geração do boleto em si, para visualização e impressão.

Posteriormente deverá incluir geração de arquivos de remessa e processamento de
retorno.


## Como Utilizar

Um exemplo de utilização encontra-se no arquivo "teste.asp", na raiz do
repositório e consiste em incluir os arquivos necessários, instanciar a classe,
preencher os campos e chamar o método `Boleto.Write()` para escrever o boleto 
na página ou `Boleto.HTML()` para obter o HTML do boleto.


	set banco = new BancoItau
	
	banco.Agencia = "0123"
	banco.Conta = "12345-6"
	banco.Carteira = "175"

	set sacado = new SacadoASP
	
	sacado.Nome = "Fulano da Silva Sauro"
	sacado.Endereco = "Rua das Flores, 123"
	' ...
	
	
    set boleto = new BoletoASP	
	
	boleto.CedenteNome = "Ricardo A. de Souza"
	' ...
	boleto.ValorDocumento = 1
	boleto.DataVencimento = "01/01/2015"
	
	set boleto.Banco = banco
	set boleto.Sacado = sacado
	
	boleto.Write()


## Bancos já implementados

* Itaú - carteiras padrão
* Itaú - carteiras com 15 posições
* Real - carteiras padrão
* Sandander - carteiras padrão

## Licença

### The MIT License (MIT)  - http://opensource.org/licenses/MIT

Copyright (c) 2012 **Ricardo Souza (RCDMK) - rcdmk@hotmail.com**

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.


