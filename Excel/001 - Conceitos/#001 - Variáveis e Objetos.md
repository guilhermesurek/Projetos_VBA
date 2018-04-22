# #001 - Variáveis e Objetos

## Variáveis
Existem vários tipos de variáveis dentro do VBA, cada qual armazena um tipo de dado diferente, de forma diferente e de tamanho diferente. Destaco que a variável do tipo objeto será tratada com mais detalhes pois suas funcionalidades são amplas.

### Tipos de Variáveis
| Tipo | Tamanho | Intervalo de Valores |
| :---: | :---: | --- |
| **Byte** | 1 byte | de 0 a 255 |
| **Boolean** | 2 bytes | True (Verdadeiro) ou False (Falso) |
| **Integer** | 2 bytes | de -32.768 a 32.767 |
| **Long** | 4 bytes | de -2.147.483.648 a 2.147.483.647 |
| **Single** | 4 bytes | de -3,402823E38 a - 1,401298E-45 (Para Valores Negativos) |
||| de 1,401298E-45 a 3,402823E38 (Para Valores Positivos) |
| **Double** | 8 bytes | de -1,79769313486232E308 a -4,94065645841247E-324 (Para Valores Negativos) |
|||de 4,94065645841247E-324 a 1,79769313486232E308 (Para Valores Positivos) |
| **Currency** | 8 bytes | de -922.337.203.685.477,5808 a 922.337.203.685.477,5807 |
| **Decimal** | 12 bytes | +/-79.228.162.514.264.337.593.543.950.335 sem casas decimais |
||| +/-7,9228162514264337593543950335 com até 28 casas decimais |
| **Date** | 8 bytes | 01 de Janeiro de 0100 a 31 de Dezembro de 9999 |
| **String (Variável)** | 10 bytes + comprimento da String | 0 a aproximadamente 2 bilhões de caracteres |
| **String (Fixa)** | comprimento da String | 1 a aproximadamente 65.400 caracteres |
| **Variant (Números)** | 16 bytes | Qualquer valor até o valor de um tipo de dados Double. Ele também pode |
||| carregar caracteres especiais como Empty, Error, Nothing e Null |
| **Variant (Caracteres)** | 22 bytes + comprimento da String | 0 a aproximadamente 2 bilhões de caracteres |
| **User-Defined** | Depende | Varia de acordo com os elementos |
| **Object** | 4 bytes | Se referencia a qualquer objeto |

> **Importante**: No VBA utiliza-se tudo em inglês, caso não seja familiarizado procure entender as principais palavras utilizadas.

### Como utilizar os diferentes tipos de variáveis
Se você criar uma macro e dentro dela escrever a seguinte linha de código
```
minha_soma = 3 + 5
```
O VBA irá entender que você está utilizando uma nova variável chamada *minha_soma*, irá efetuar a soma **3 + 5** e alocará na variável *minha_soma* o resultado. Além disso, ele verificará que você não declarou a variável, isto é, não disse para ele qual é o tipo da variável *minha_soma* e, nesse caso, ele irá especificar esta variável como tipo **Variant**.
Ou seja, se você simplesmente escrever as suas variáveis o VBA irá entender, porém não é aconselhável não declarar as variáveis. O correto a se fazer nesse caso seria:
```
Dim minha_soma As Integer
'Ou
Dim minha_soma As Byte 'Caso saiba que o valor dessa variável não passará de 255 e não será negativo.
minha_soma = 3 + 5
```
Agora, por que é aconselhável declarar suas variáveis:
1.Quando não se declara o tipo da variável o VBA define como **Variant**, e o tipo **Variant** é o tipo que mais ocupa espaço na memória do computador. Enquanto o tipo **Integer** (Inteiro) ocupa 2 bytes de memória o tipo **Variant** ocupa 16 bytes, 8 vezes mais. Se a sua aplicação é um projeto de pequeno porte e não irá trabalhar com um volume de dados e variáveis muito grande você não irá notar a diferença, porém em aplicações de grande porte o impacto é considerável a ponto de não rodar em computadores com pouca memória.
