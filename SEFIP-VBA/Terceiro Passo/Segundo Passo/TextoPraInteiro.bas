Attribute VB_Name = "TextoPraInteiro"
Function ConverterStringParaInteger(Pis As String) As String
    
    'Dim minhaString As String
    'Dim meuInteiro As Integer
    'Dim linhas() As String
    'Dim Contador As LongLong
    'Contador = 2
    
    Dim todosOsNUmeros As String
        todosOsNUmeros = "0123456789"
    Dim Numeros() As String
    Dim i As Integer
    
    ReDim Numeros(1 To Len(todosOsNUmeros)) ' Redimensiona o array para o tamanho das letras
    
    For i = 1 To Len(todosOsNUmeros)
        Numeros(i) = Mid(todosOsNUmeros, i, 1) ' Divide a string em letras individuais
    Next i

    
    
    
    Dim contador As Integer
    contador = 0
    
    
    Dim NovoPis As String
    Do While contador <= Len(Pis)
        
        For i = 1 To Len(todosOsNUmeros) '108
                    Dim NumeroPis As String
                    NumeroPis = Mid(Pis, contador + 1, 1)
                    If Numeros(i) = NumeroPis Then
                        'Dim teste As Integer
                        
                        
                        NovoPis = NovoPis + Numeros(i)
                        'teste = teste + Numeros(i)
                        'MsgBox (NovoPis)
                        
                        Exit For
                    End If
                    
                
                
        
        Next i
    contador = contador + 1
    Loop
    
    NovoPis = CDec(NovoPis)
    ConverterStringParaInteger = NovoPis ' Declare o recepitor como String memso n�o como inteiro
    
        
        
        'meuInteiro = CDec(Pis)

        ' Agora voc� pode usar o inteiro como desejar
        'MsgBox "A string convertida para inteiro �: " & NovoPis
        'MsgBox "A string convertida para inteiro �: " & NovoPis + 500
    
    'Next i
    
    


    
    ' Atribuir a string desejada � vari�vel
    'Pis = "12345"

    ' Converter a string para um inteiro usando CInt
    
End Function
