Attribute VB_Name = "Módulo1"

'I am a beginner of VBA.
'I want to write a script which  changes the text in black color to white, but  keep the fonts in green color.
''For example, in a text box, the texts are in two different colors, some bullet sentences in green color and some bullet sentences in black color.
'
'Based  on the program (http://www.pptfaq.com/FAQ00465.htm), I wrote the script below.
'However the code below only works in the text box which contains all text in black color.
'The program does not work if the text box  contains the fonts which  have two colors.

Sub white2black()

    Dim R As Integer
    Dim G As Integer
    Dim B As Integer
 
    Dim oSld As Slide
    Dim oShp As Shape
    Dim oShapes As Shapes
 
    For Each oSld In ActivePresentation.Slides
        Set oShapes = oSld.Shapes
        For Each oShp In oShapes
            If oShp.HasTextFrame Then
                If oShp.TextFrame.HasText Then
                    If oShp.TextFrame.TextRange.Font.Color = RGB(255, 255, 255) Then
                        oShp.TextFrame.TextRange.Font.Color = RGB(0, 0, 0)
                    End If
                End If
            End If
 
        Next oShp
    Next oSld

 End Sub

Sub white2black2()
    Dim oSld As Slide
    Dim oShp As Shape
    Dim x As Long
    For Each oSld In ActivePresentation.Slides
        For Each oShp In oSld.Shapes
            If oShp.HasTextFrame Then
                If oShp.TextFrame.HasText Then
                    With oShp.TextFrame.TextRange
                        For x = .Runs.Count To 1 Step -1
                            If .Runs(x).Font.Color.RGB = RGB(255, 255, 255) Then
                                .Runs(x).Font.Color.RGB = RGB(0, 0, 0)
                            End If
                        Next x
                    End With
                End If 'has text
            End If 'has textframe
        Next oShp
    Next oSld
End Sub
