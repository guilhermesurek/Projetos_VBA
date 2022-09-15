# #001001 - Alterando a Língua do Revisor de Textos
## Problema

Quando se começar a trabalhar em um arquivo PowerPoint que estava em outra língua, todos os textos serão revisados com a língua de criação e não a língua desejada.
Exemplo, um power point foi construído em inglês e estou trabalhando em um material em português, todos os textos ficarão com o destaque em vermelho por causa do
revisor de texto.

## VBA

Uma forma de resolver esse problema rapidamente, sem precisar entrar em cada objeto e ajustar a língua manualmente, é utilizar um código VBA para tal.
O código abaixo irá iterar por cada objeto do Power Point em questão, nos casos de grupos de objetos ele entrará dentro dos grupos e subgrupos, e nos casos de tabelas
iterará por cada linha e coluna da tabela ajustando a linguagem do revisor de texto para PT-BR. Caso deseje outro idioma, ajustar o ID do idioma.

## REFERÊNCIAS

Este código foi alterado com base no código criado por Ana Pedretti e Vinicius [Link](https://anapedretti.com.br/2018/11/trocar-o-idioma-de-revisao-de-texto-em-todos-slides-powerpoint/#comment-16). 

## Módulo VBA

```
'>Módulo Uteis PowerPoint
'>Referências Necessárias: (Ferramentas>Referências)
'nenhuma referência necessária para este código

Function PercorrerObjetosETrocarLan(obj As Variant, lang As Variant)
    
    DoEvents ' Para não correr o risco de ficar em um loop infito e travar seu arquivo
    ' Grupos
    If obj.Type = msoGroup Then
        For Each SubObj In obj.GroupItems
            PercorrerObjetosETrocarLan SubObj, lang
        Next SubObj
    ' Tabelas
    ElseIf obj.HasTable Then
        For nRow = 1 To obj.Table.Rows.Count
            DoEvents
            For nCol = 1 To obj.Table.Columns.Count
                DoEvents
                obj.Table.Cell(nRow, nCol).Shape.TextFrame.TextRange.LanguageID = lang
            Next nCol
        Next nRow
    ' Objetos de Texto
    Else
        If obj.Type = msoTextBox Or msoPlaceholder Then
            If obj.HasTextFrame Then
                obj.TextFrame.TextRange.LanguageID = lang
            End If
        End If
        
    End If
    
End Function

Sub Main()
    
Dim pSlide As Slide
Dim pShape As Shape

For Each pSlide In ActivePresentation.Slides
    For Each pShape In pSlide.Shapes
        PercorrerObjetosETrocarLan pShape, msoLanguageIDBrazilianPortuguese ' ID do idioma desejado
    Next pShape
Next pSlide


End Sub
```
