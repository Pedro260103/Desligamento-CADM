Attribute VB_Name = "Grava05"
Sub Grava05()
   
    Dim LocalDoArquivo As String
    Dim NumArquivo As Integer
    Dim FileContent As String
    Dim Linha As String
    LocalDoArquivo = "X:\GESFO - ISO\GERINS em ES250-89\SEFIP\2020-RETIFICADORAS\02-2020-(1� Retificadora-DESLIGAMENTOS)\CADM\SEFIP - C�digo.RE"
    NumArquivo = FreeFile
    
    Open LocalDoArquivo For Input As #NumArquivo

    FileContent = Input$(LOF(1), #NumArquivo)
    Close #NumArquivo
    
    
    Linha = Trim(FileContent)
    
    Dim SegundaInfo As String
    Dim NovoConteudo As String
    Dim PrimeiraInfo As String
    Dim linhas() As String
    Dim Codigo05 As String
    linhas = Split(FileContent, vbCrLf)
   
    For i = 3 To UBound(linhas)
        PrimeiraInfo = Mid(linhas(i), 1, 199)
        SegundaInfo = Mid(linhas(i), 202, 159)
        Codigo05 = Mid(linhas(i), 200, 2)
        Range("A2").Value = UBound(linhas)
        Range("B2").Value = i
            If Codigo05 = "  " And SegundaInfo <> "                                                                                                                                                              *" Then
                NovoConteudo = PrimeiraInfo & "05" & SegundaInfo
                linhas(i) = NovoConteudo
                
                
                
            End If
            
            
    Next i
    FileContent = Join(linhas, vbCrLf)
                
    ' Abre o arquivo em modo de escrita para gravar as altera��es
    Open LocalDoArquivo For Output As #NumArquivo
    Print #NumArquivo, FileContent
    Close #NumArquivo
MsgBox "Fim!"
End Sub


