Attribute VB_Name = "AddLinhas"
Sub AddLinhasEmBranco(CaminhoArquivo As String)
    Dim FilePath As String
    Dim FileContent As String
    Dim FileSize As Long
    Dim i As Long
    
    
    ' Leia o conte�do do arquivo existente
    Open CaminhoArquivo For Input As #1
    FileContent = Input$(LOF(1), #1)
    Close #1
    Dim Contador As Integer
    Contador = 2
    Do While Range("A" & Contador).Value <> ""
        Contador = Contador + 1
    Loop
    
    
    ' Especifique o tamanho desejado do arquivo em bytes
    FileSize = Contador ' Por exemplo, 50 bytes
    
    ' Abra o arquivo em modo de sa�da (para grava��o) para acrescentar asteriscos
    Open CaminhoArquivo For Append As #1
    Dim Metade As Long
    Metade = FileSize '\ 2
    ' Escreva os asteriscos ap�s o conte�do existente
    For i = 0 To (Metade + 1)
        Print #1, "*" '"                                                                                                                                                                                                                                                                                                                                                                        "
    Next i
    
    ' Feche o arquivo
    Close #1
    ' Salve o arquivo
    MsgBox Metade + 1
    ActiveWorkbook.Save
End Sub
