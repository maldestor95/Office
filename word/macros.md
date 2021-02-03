# Redimensionner des images dans la sélection
```
Function tailleImage(H)

Dim i As Long
With Selection
    For i = 1 To .InlineShapes.Count
        With .InlineShapes(i)
            .ScaleHeight = H
        End With
    Next i
End With

End Function

Sub scale50()
    Dim t
    t = tailleImage(50)
End Sub
Sub scale25()
    Dim t
    t = tailleImage(25)
End Sub
Sub scale100()
    Dim t
    t = tailleImage(100)
End Sub
```
