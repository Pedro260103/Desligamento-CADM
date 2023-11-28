Attribute VB_Name = "CopiaCola"
Sub CopiaColaNaplanilha()
    
     
        
        Dim bm() As String
        Dim nome() As String
        Dim salario() As String
        Dim contrib() As String
        
    Dim contador As Integer
    
    contador = 1
    
    Do While Range("A" & contador).Value <> ""
        ReDim Preserve bm(1 To contador) ' Redimensiona o array para acomodar mais um elemento
        ReDim Preserve nome(1 To contador) ' Redimensiona o array para acomodar mais um elemento
        ReDim Preserve salario(1 To contador) ' Redimensiona o array para acomodar mais um elemento
        ReDim Preserve contrib(1 To contador) ' Redimensiona o array para acomodar mais um elemento
        
        bm(contador) = Range("A" & contador).Text
        nome(contador) = Range("B" & contador).Text
        salario(contador) = Range("C" & contador).Text
        contrib(contador) = Range("E" & contador).Text
        
        contador = contador + 1 ' Incrementa o contador para evitar um loop infinito
    Loop
    
        Dim wb As Workbook
        Dim ws As Worksheet
        Dim rng As Range
    
        ' Abra a planilha
        Set wb = Workbooks.Open("C:\Users\pres00314250\Documents\PlanilaDestino.xlsx")
        
        ' Acesse ou crie a planilha Modulada
        On Error Resume Next
        Set ws = wb.Sheets("Modulada")
        On Error GoTo 0
        
        If ws Is Nothing Then
            ' Se a planilha não existir, crie uma nova
            Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
            ws.Name = "Modulada"
        End If
    
        ' Edite a planilha conforme necessário
        
        Dim w As Integer
        w = 1
        
        Do While w <= UBound(bm)
            If Range("L5").Value = "" Then
            
                For a = 1 To UBound(bm) Step 1
                    Set rng = ws.Range("A" & w)
                    rng.Value = bm(w)
                Next a
                For b = 1 To UBound(nome) Step 1
                    Set rng = ws.Range("B" & w)
                    rng.Value = nome(w)
                Next b
                For c = 1 To UBound(salario) Step 1
                    Set rng = ws.Range("C" & w)
                    rng.Value = salario(w)
                Next c
                For e = 1 To UBound(contrib) Step 1
                    Set rng = ws.Range("E" & w)
                    rng.Value = contrib(w)
                Next e
                w = w + 1
            Else
            
                For a = 1 To UBound(bm) Step 1
                    Set rng = ws.Range("A" & w + Range("L5").Value)
                    rng.Value = bm(w)
                Next a
                For b = 1 To UBound(nome) Step 1
                    Set rng = ws.Range("B" & w + Range("L5").Value)
                    rng.Value = nome(w)
                Next b
                For c = 1 To UBound(salario) Step 1
                    Set rng = ws.Range("C" & w + Range("L5").Value)
                    rng.Value = salario(w)
                Next c
                For e = 1 To UBound(contrib) Step 1
                    Set rng = ws.Range("E" & w + Range("L5").Value)
                    rng.Value = contrib(w)
                Next e
                w = w + 1
            End If
        Loop
        Range("L5").Value = Range("L5").Value + contador - 1
        
        
    
        ' Salve as alterações e feche a planilha
        wb.Save
        'wb.Close SaveChanges:=False
        
    
    
End Sub

Function PlanilhaExiste(nomePlanilha As String) As Boolean
    ' Verifica se a planilha existe no workbook ativo
    On Error Resume Next
    PlanilhaExiste = Not Sheets(nomePlanilha) Is Nothing
    On Error GoTo 0
End Function



