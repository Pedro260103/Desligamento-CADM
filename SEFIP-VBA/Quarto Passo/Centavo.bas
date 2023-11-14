Attribute VB_Name = "Centavo"
Sub grava()
    Dim LocalDoArquivo As String
    Dim NumArquivo As Integer
    Dim FileContent As String
    Dim Linha As String
    LocalDoArquivo = "X:\GESFO - ISO\GERINS em ES250-89\SEFIP\2020-RETIFICADORAS\02-2020-(1ª Retificadora-DESLIGAMENTOS)\CADM\SEFIP - Código.RE"
    NumArquivo = FreeFile
    
    Open LocalDoArquivo For Input As #NumArquivo

    FileContent = Input$(LOF(1), #NumArquivo)
    Close #NumArquivo
    
    
    Linha = Trim(FileContent)
    
    Dim SegundaInfo As String
    Dim Nome As String
    Dim NomeFlutuante As String
    Dim BM As String
    Dim eUmBM As String
    Dim NovoConteudo As String
    Dim PrimeiraInfo As String
    Dim linhas() As String
    Dim MudaPessoa As LongLong
    MudaPessoa = 0
    Dim puladas As LongLong
    puladas = 0
    linhas = Split(FileContent, vbCrLf)
    
    Dim MesDelig As String
    
    
    
    Dim contador As Integer
    contador = 1
    For i = 3 To UBound(linhas)
        Nome = Mid(linhas(i), 54, 70)
        BM = Mid(linhas(i), 124, 11)
        
        eUmBM = Mid(linhas(i), 124, 2)
        Range("B2").Value = i
        If InStr(1, linhas(i), eUmBM, vbTextCompare) > 0 And eUmBM = "00" Then ' ficou meio redundante
            puladas = puladas + 1
            If InStr(1, linhas(i + 1), eUmBM, vbTextCompare) <> BM Then
                MudaPessoa = MudaPessoa + 1
                contador = 1
            End If
            
            
        Else
            
            NomeFlutuante = Mid(linhas(i + 1), 54, 70)
            MesDelig = Mid(linhas(i), 128, 2)
            'MesDelig = ConverterStringParaInteger(MesDelig)
            'MsgBox (MesDelig)
            'MsgBox (MesDelig + 30)
            If eUmBM = "I3" Or eUmBM = "I1" Or eUmBM = " J" Then 'And MesDelig <> "02" Then
                
                'MsgBox "Achamos o I3N , I1N , J"
                PrimeiraInfo = Mid(linhas(i - contador), 1, 244)
                SegundaInfo = Mid(linhas(i - contador), 245, 116)
                'MsgBox SegundaInfo
                SegundaInfo = Left(SegundaInfo, 1) & "1" & Right(SegundaInfo, Len(SegundaInfo) - 2)
                NovoConteudo = PrimeiraInfo + SegundaInfo
                'MsgBox NovoConteudo
                ' Abre o arquivo em modo de leitura e escrita
                Open LocalDoArquivo For Input As #NumArquivo
                FileContent = Input$(LOF(1), #NumArquivo)
                Close #NumArquivo
                
                ' Substitui a linha anterior pelo NovoConteudo
                FileContent = Replace(FileContent, linhas(i - contador), NovoConteudo)
                
                ' Abre o arquivo em modo de escrita para gravar as alterações
                Open LocalDoArquivo For Output As #NumArquivo
                Print #NumArquivo, FileContent
                Close #NumArquivo
                contador = 1
            Else
            
            
            
            contador = contador + 1
            
            
            
            End If
                
            
            
            
            
        End If
    
    Next i
    
    
MsgBox "Fim!"
End Sub

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
    ConverterStringParaInteger = NovoPis ' Declare o recepitor como String memso não como inteiro
    
        
        
        'meuInteiro = CDec(Pis)

        ' Agora você pode usar o inteiro como desejar
        'MsgBox "A string convertida para inteiro é: " & NovoPis
        'MsgBox "A string convertida para inteiro é: " & NovoPis + 500
    
    'Next i
    
    


    
    ' Atribuir a string desejada à variável
    'Pis = "12345"

    ' Converter a string para um inteiro usando CInt
    
End Function
