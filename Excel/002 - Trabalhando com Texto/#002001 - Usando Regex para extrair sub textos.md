# #002001 - Usando Regex para extrair sub textos

## Problema

Possui-se um certo padrão de texto, porém ferramentas mais simples não conseguem extrair a parte do texto desejada. Abaixo temos uma lista de textos provenientes de partidas de pagamento e precisa-se extrair dois números principais o NSU e o CV.

| Texto Original | NSU Desejado | CV Desejado |
| :---: | --- | --- |
| PGTO NSU TEF 509178 CV 2000022 | 509178 | 2000022 |
| NSU 507343 - 6 | 507343 | |
| PGTO - NSU prob. extração de dados-CV Aguardando r | | |
| PGTO - NSU -CV 429097347 | | 429097347 |
| NSU NÃO IDENTIFICADA | | |
| PGTO - NSU Sem Tef-CV 15000314 | | 15000314 |
| PGTO - NSU 503532 / 503533 -CV 8000180 / 8000182 | 503532, 503533 | 8000180, 8000182 |
| PGTO - NSU --CV 10000118 | | 10000118 |

## Regex

Regex - Regular Expressions - é um método de reconhecimento de padrões. Com uma escrita padronizada, faz com que o programa facilmente identifique uma cadeia de caracteres dado certo padrão.

Para encontrar as expressões regex necessárias que preciso para resolver o problema costumo utilizar sites que aplicam o regex online como [Regex 101](https://regex101.com/) ou [Regexr](https://regexr.com/).

### Exemplos de Expressões
| Expressão | Descrição |
| :---: | --- |
| **Caracteres** | |
| . | qualquer caractere com exceção de nova linha |
| \w | letra |
| \d | dígito |
| \s | espaço em branco |
| \W | não seja letra |
| \D | não seja dígito |
| \S | não seja espaço em branco |
| [abc] | qualquer uma das três letras a, b ou c |
| [^abc] | não seja a, b ou c |
| [a-g] | qualquer letra entre a e g |
| **Limites** | |
| ^abc$ | inicia / termina uma frase |
| \b \B | limite de letra e não seja letra|
| **Caracteres de Escape** | |
| \. | literalmente . |
| \* | literalmente * |
| \\ | literalmente \ |
| \t | Tab |
| \n | nova linha |
| \r | quebra de linha |
| **Grupos** | |
| (abc) | captura o grupo abc |
| (?:abc) | não captura o grupo abc |
| (?=abc) | positivo a frente |
| (?!abc) | negativo a frente |
| **Quantidades** | |
| a* | 0 ou mais a`s |
| a+ | 1 ou mais a+? a`s |
| a? | 0 ou 1 a |
| a{5} | exatamente 5 a`s |
| a{2, } | dois ou mais a`s |
| a{1,3} | entre um e três a's |
| a+? | menor quantidade possível |
| ab \| cd | encontra ab ou cd |

## Módulo VBA

```
'>Módulo Regex
'>Referências Necessárias: (Ferramentas>Referências)
'- Microsoft VBScript Regular Expressions 5.5

Function RegexMatch(sStr As String, vPattern As Variant, vReplace As Variant) As String

    ' Função: RegexMatch
    '         Encontrando padrões regulares em textos
    ' Inputs: sStr: Texto principal.
    '         vPattern: Vetor com padrões a serem buscados no texto principal.
    '         vReplace: Vetor com padrões a serem removidos após a busca. Nem sempre o regex busca somente o que se deseja.
    ' Outputs: Valores extraídos do texto principal separados por vírgula - ","
    ' Fonte: https://www.automateexcel.com/vba/regex/
    Dim regexOne As Object
    
    '>Inicializações
    RegexMatch = ""
    Set regexOne = New RegExp
    
    '>Loop nos padrões
    For Each sPat In vPattern
    
        '>Configurando Motor Regex (Padrões Regulares)
        regexOne.Pattern = sPat
        regexOne.Global = True
        regexOne.IgnoreCase = IgnoreCase
        
        '>Executando Motor Regex (Padrões Regulares)
        Set Matches = regexOne.Execute(sStr)
        
        '>Concatenando resultado
        For Each Match In Matches
            
            ' Replace patterns
            aux = Trim(Match)
            For Each vRep In vReplace
                aux = Trim(Replace(aux, vRep, ""))
            Next vRep
            
            'Print
            'Debug.Print "Orignal Match: " & Match.Value & " Replaced: " & aux
            
            ' Verificar se match já consta no resultado
            If InStr(RegexMatch, aux) = 0 Then
                ' Não consta, incluir
                If RegexMatch = "" Then
                    RegexMatch = aux
                Else
                    RegexMatch = RegexMatch & ", " & aux
                End If
            End If
        Next Match
        
    Next sPat
 
End Function

Function Extrair_Regex_NSU(sTexto As String) As String
    
    ' Função: Extrair_Regex_NSU
    '         Extrair um conjunto de 6 dígitos geralmente precedidos do termo "NSU" e que geralmente inicia-se por "50"
    ' Inputs: sTexto: Texto principal.
    ' Outputs: Valores extraídos do texto principal separados por vírgula - "," que se enquadram no padrão acima.
    ' Obs: Essa função pode ser utilizada na própria planilha do excel.
    
    Dim vPattern(1) As Variant
    Dim vReplace(3) As Variant
    
    ' Pattern Setting
    vPattern(0) = "NSU[:\-\s]*[\d]+"
    vPattern(1) = "50[1-9][\d]{3}"
    vReplace(0) = "NSU"
    vReplace(1) = "-"
    vReplace(2) = ":"
    vReplace(3) = " "
    
    ' Extrair Numero do NSU
    Extrair_Regex_NSU = RegexMatch(sTexto, vPattern, vReplace)
    
End Function

Function Extrair_Regex_CV(sTexto As String) As String
    
    ' Função: Extrair_Regex_CV
    '         Extrair um conjunto de dígitos com comprimento variando entre 7 a 9, geralmente precedidos do termo "CV" e
    '         iniciando com um dígito entre 1 e 9 seguido por dois zeros.
    ' Inputs: sTexto: Texto principal.
    ' Outputs: Valores extraídos do texto principal separados por vírgula - "," que se enquadram no padrão acima.
    ' Obs: Essa função pode ser utilizada na própria planilha do excel.
    
    Dim vPattern(6) As Variant
    Dim vReplace(3) As Variant
    
    ' Pattern Setting
    vPattern(0) = "CV[:\-\s]*[14][\d]{8}"
    vPattern(1) = "CV[:\-\s]*[14][\d]{7}"
    vPattern(2) = "CV[:\-\s]*[14][\d]{6}"
    vPattern(3) = "CV[:\-\s]*[1-9][0][0][\d]{4}"
    vPattern(4) = "CV[:\-\s]*[\d]{3}"
    vPattern(5) = "CV[:\-\s]*[\d]{2}"
    vPattern(6) = "[1-9][0][0][\d]{4}"
    vReplace(0) = "CV"
    vReplace(1) = "-"
    vReplace(2) = ":"
    vReplace(3) = " "
    
    ' Extrair Numero do CV
    Extrair_Regex_CV = RegexMatch(sTexto, vPattern, vReplace)
    
End Function

Function Teste_Extrair_CV()
    
    ' Função: Teste_Extrair_CV
    '         Executa uma série de testes para verificar se a função Extrair_Regex_CV(sTexto) está retornando o resultado correto.
    '         O resultado dos testes será printado na janela de inspeção imediata.
    
    Dim sStr As String
    
    'Teste 1
    sStr = "PGTO - NSU 507108-CV 16000186"
    sResultado = Extrair_Regex_CV(sStr)
    sResultadoEsperado = "16000186"
    Debug.Print "Teste 01 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 2
    sStr = "Cancelamento de vendas - NSU 89"
    sResultado = Extrair_Regex_CV(sStr)
    sResultadoEsperado = ""
    Debug.Print "Teste 02 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 3
    sStr = "VOTORAN SÃO JOSÉ"
    sResultado = Extrair_Regex_CV(sStr)
    sResultadoEsperado = ""
    Debug.Print "Teste 03 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 4
    sStr = "Pagamento cliente Std 3689 130028767RC Como ref."
    sResultado = Extrair_Regex_CV(sStr)
    sResultadoEsperado = ""
    Debug.Print "Teste 04 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 5
    sStr = "508947 - 1/6"
    sResultado = Extrair_Regex_CV(sStr)
    sResultadoEsperado = ""
    Debug.Print "Teste 05 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 6
    sStr = "CANCELAMENTO NSU 507078 5/6"
    sResultado = Extrair_Regex_CV(sStr)
    sResultadoEsperado = ""
    Debug.Print "Teste 06 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 7
    sStr = "NSU: 509639 / PARC: 1"
    sResultado = Extrair_Regex_CV(sStr)
    sResultadoEsperado = ""
    Debug.Print "Teste 07 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 8
    sStr = "PGTO NSU TEF 509178 CV 2000022"
    sResultado = Extrair_Regex_CV(sStr)
    sResultadoEsperado = "2000022"
    Debug.Print "Teste 08 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 9
    sStr = "NSU 507343 - 6"
    sResultado = Extrair_Regex_CV(sStr)
    sResultadoEsperado = ""
    Debug.Print "Teste 09 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 10
    sStr = "PGTO - NSU prob. extração de dados-CV Aguardando r"
    sResultado = Extrair_Regex_CV(sStr)
    sResultadoEsperado = ""
    Debug.Print "Teste 10 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 11
    sStr = "PGTO - NSU -CV 429097347"
    sResultado = Extrair_Regex_CV(sStr)
    sResultadoEsperado = "429097347"
    Debug.Print "Teste 11 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 12
    sStr = "PGTO - NSU NÃO IDENTIFICADA"
    sResultado = Extrair_Regex_CV(sStr)
    sResultadoEsperado = ""
    Debug.Print "Teste 12 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 13
    sStr = "PGTO - NSU Sem Tef-CV 15000314"
    sResultado = Extrair_Regex_CV(sStr)
    sResultadoEsperado = "15000314"
    Debug.Print "Teste 13 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 14
    sStr = "PGTO - NSU 503532 / 503533 -CV 8000180 / 8000182"
    sResultado = Extrair_Regex_CV(sStr)
    sResultadoEsperado = "8000180, 8000182"
    Debug.Print "Teste 14 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 15
    sStr = "PGTO - NSU --CV 10000118"
    sResultado = Extrair_Regex_CV(sStr)
    sResultadoEsperado = "10000118"
    Debug.Print "Teste 15 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 16
    sStr = "NSU ADQ 4000688 NSU TEF 508706"
    sResultado = Extrair_Regex_CV(sStr)
    sResultadoEsperado = "4000688"
    Debug.Print "Teste 16 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 17
    sStr = "PGTO-NSU504409-CV2000272"
    sResultado = Extrair_Regex_CV(sStr)
    sResultadoEsperado = "2000272"
    Debug.Print "Teste 17 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 18
    sStr = "PGTO - NSU 507115-CV 20017265063272000202"
    sResultado = Extrair_Regex_CV(sStr)
    sResultadoEsperado = "2001726, 2000202"
    Debug.Print "Teste 18 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 19
    sStr = "PGTO NSU - 507524/507521/507522 CV - 4000244/50000"
    sResultado = Extrair_Regex_CV(sStr)
    sResultadoEsperado = "4000244"
    Debug.Print "Teste 19 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 20
    sStr = "PGTO - NSU 91-CV 91"
    sResultado = Extrair_Regex_CV(sStr)
    sResultadoEsperado = "91"
    Debug.Print "Teste 20 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 21
    sStr = "PGTO - NSU 27/21-CV 27/21"
    sResultado = Extrair_Regex_CV(sStr)
    sResultadoEsperado = "27, 21"
    Debug.Print "Teste 21 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
End Function

Function Teste_Extrair_NSU()
    
    ' Função: Teste_Extrair_NSU
    '         Executa uma série de testes para verificar se a função Extrair_Regex_NSU(sTexto) está retornando o resultado correto.
    '         O resultado dos testes será printado na janela de inspeção imediata.
    
    Dim sStr As String
    
    'Teste 1
    sStr = "PGTO - NSU 507108-CV 16000186"
    sResultado = Extrair_Regex_NSU(sStr)
    sResultadoEsperado = "507108"
    Debug.Print "Teste 01 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 2
    sStr = "Cancelamento de vendas - NSU 89"
    sResultado = Extrair_Regex_NSU(sStr)
    sResultadoEsperado = "89"
    Debug.Print "Teste 02 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 3
    sStr = "VOTORAN SÃO JOSÉ"
    sResultado = Extrair_Regex_NSU(sStr)
    sResultadoEsperado = ""
    Debug.Print "Teste 03 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 4
    sStr = "Pagamento cliente Std 3689 130028767RC Como ref."
    sResultado = Extrair_Regex_NSU(sStr)
    sResultadoEsperado = ""
    Debug.Print "Teste 04 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 5
    sStr = "508947 - 1/6"
    sResultado = Extrair_Regex_NSU(sStr)
    sResultadoEsperado = "508947"
    Debug.Print "Teste 05 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 6
    sStr = "CANCELAMENTO NSU 507078 5/6"
    sResultado = Extrair_Regex_NSU(sStr)
    sResultadoEsperado = "507078"
    Debug.Print "Teste 06 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 7
    sStr = "NSU: 509639 / PARC: 1"
    sResultado = Extrair_Regex_NSU(sStr)
    sResultadoEsperado = "509639"
    Debug.Print "Teste 07 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 8
    sStr = "PGTO NSU TEF 509178 CV 2000022"
    sResultado = Extrair_Regex_NSU(sStr)
    sResultadoEsperado = "509178"
    Debug.Print "Teste 08 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 9
    sStr = "NSU 507343 - 6"
    sResultado = Extrair_Regex_NSU(sStr)
    sResultadoEsperado = "507343"
    Debug.Print "Teste 09 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 10
    sStr = "PGTO - NSU prob. extração de dados-CV Aguardando r"
    sResultado = Extrair_Regex_NSU(sStr)
    sResultadoEsperado = ""
    Debug.Print "Teste 10 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 11
    sStr = "PGTO - NSU -CV 429097347"
    sResultado = Extrair_Regex_NSU(sStr)
    sResultadoEsperado = ""
    Debug.Print "Teste 11 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 12
    sStr = "PGTO - NSU NÃO IDENTIFICADA"
    sResultado = Extrair_Regex_NSU(sStr)
    sResultadoEsperado = ""
    Debug.Print "Teste 12 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 13
    sStr = "PGTO - NSU Sem Tef-CV 15000314"
    sResultado = Extrair_Regex_NSU(sStr)
    sResultadoEsperado = ""
    Debug.Print "Teste 13 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 14
    sStr = "PGTO - NSU 503532 / 503533 -CV 8000180 / 8000182"
    sResultado = Extrair_Regex_NSU(sStr)
    sResultadoEsperado = "503532, 503533"
    Debug.Print "Teste 14 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 15
    sStr = "PGTO - NSU --CV 10000118"
    sResultado = Extrair_Regex_NSU(sStr)
    sResultadoEsperado = ""
    Debug.Print "Teste 15 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 16
    sStr = "NSU ADQ 4000688 NSU TEF 508706"
    sResultado = Extrair_Regex_NSU(sStr)
    sResultadoEsperado = "508706"
    Debug.Print "Teste 16 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 17
    sStr = "PGTO-NSU504409-CV2000272"
    sResultado = Extrair_Regex_NSU(sStr)
    sResultadoEsperado = "504409"
    Debug.Print "Teste 17 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 18
    sStr = "PGTO - NSU 507115-CV 20017265063272000202"
    sResultado = Extrair_Regex_NSU(sStr)
    sResultadoEsperado = "507115, 506327"
    Debug.Print "Teste 18 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 19
    sStr = "PGTO NSU - 507524/507521/507522 CV - 4000244/50000"
    sResultado = Extrair_Regex_NSU(sStr)
    sResultadoEsperado = "507524, 507521, 507522"
    Debug.Print "Teste 19 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 20
    sStr = "PGTO - NSU 91-CV 91"
    sResultado = Extrair_Regex_CV(sStr)
    sResultadoEsperado = "91"
    Debug.Print "Teste 20 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
    'Teste 21
    sStr = "PGTO - NSU 27/21-CV 27/21"
    sResultado = Extrair_Regex_CV(sStr)
    sResultadoEsperado = "27, 21"
    Debug.Print "Teste 21 - " & (sResultado = sResultadoEsperado) & " - " & sStr & " - Resultado: " & sResultado
    
End Function
```
