Attribute VB_Name = "CopiaEcriaMosulacao"
Sub marcar()
    'Range("a1", "a5").Formula = "=Rand()"
    'Range("a1:F10").Formula = "=Rand()"
    
    
    
    
    
    
    
    
    
    
    
    
    
    'Worksheets(1).Range("E5").AddComment "Current Sales"
    'Worksheets("Planilha1").Range("A1:E1").Columns.AutoFit
    'Worksheets("Planilha1").Range("A1:F10").BorderAround _
    'ColorIndex:=1, Weight:=xlThick
    Dim pilha As Integer
    
    For i = 2 To 298 Step 2
        pilha = Range("I1").Value
        pilha = pilha + 26
        If pilha = 26 Then
            Worksheets("Table " & i).Range("A1:G26").Copy _
            Destination:=Worksheets("Modulada").Range("A1:G26")
            Range("I1").Value = pilha
        Else
            Worksheets("Table " & i).Range("A1:G26").Copy _
            Destination:=Worksheets("Modulada").Range("A" & (pilha - 26) & ":G" & pilha)
            Range("I1").Value = pilha
        End If
        
    
    Next i
    
End Sub


