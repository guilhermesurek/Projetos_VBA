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
| **** |  bytes |  |

> **Importante**: No VBA utiliza-se tudo em inglês, caso não seja familiarizado procure entender as principais palavras utilizadas.
