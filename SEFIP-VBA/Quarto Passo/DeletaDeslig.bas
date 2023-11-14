Attribute VB_Name = "DeletaDeslig"
Sub DeletaDesli()
    Dim LocaldoArquivo As String
    Dim NumArquivo As Integer
    Dim FileContent As String
    Dim Linha As String
     
    
    LocaldoArquivo = "X:\GESFO - ISO\GERINS em ES250-89\SEFIP\2021-RETIFICADORAS\10-2021-(2ª Retificadora)\CADM\SEFIP - Copia.RE"
    NumArquivo = FreeFile
    
    Open LocaldoArquivo For Input As #NumArquivo ' Abre o arquivo em modo de leitura
    
    FileContent = Input$(LOF(1), #NumArquivo)
    Close #NumArquivo
 
    ' Extrai informações da linha
    Linha = Trim(FileContent) ' Remove espaços em branco no início e no fim
 
    Dim Codigo As String
    Dim Pis As String
    Dim nome As String
    Dim nomeTXT As String
    Dim DataAdimissao As String
    Dim DataAdimissaoTXT As String
    Dim Codigo20 As String
    Dim CodigoDesligamento As String
    Dim bm As String
    Dim bmTXT As String
    Dim DataNacimento As String
    Dim DataDemissao As String
    Dim NovoConteudo As String
    Dim linhas() As String
    Dim Contador As LongLong
    Dim UltimaLinhaCodigo90 As String
    Dim DataAdmissaoTXTCima As String
    Dim bmTXTCima As String
    Dim PisTXTCima As String
    Dim OriginalConteudo As String
    Contador = 2
    linhas = Split(FileContent, vbCrLf)
    
    
    UltimaLinhaCodigo90 = linhas(UBound(linhas) - 1)
    'AddLinhasEmBranco (UBound(linhas))
    
    Do While Range("A" & Contador).Value <> ""
        bm = Range("A" & Contador).Value
        bm = Replace(bm, "-", "")
        bm = Replace(bm, "X", "0")
        nome = Range("B" & Contador).Value
        DataAdmissao = Range("C" & Contador).Value
        DataAdmissao = Replace(DataAdmissao, "/", "")
        
        If Range("A" & Contador).Interior.Color = RGB(255, 255, 0) Then ' Se a selula for amarela então sabemos que temos que deletar o desligamento dela
            For i = 3 To UBound(linhas)
                bmTXT = Mid(linhas(i), 127, 8) ' pega o Bm do Arquivo TXT
                DataAdimissaoTXT = Mid(linhas(i), 44, 8) ' pega a Data de Adimissao do Arquivo TXT
                
                If bm = bmTXT And DataAdmissao = DataAdimissaoTXT Then
                    For w = i To UBound(linhas) - 2 ' - 2 para ele não estourar o Array
                        linhas(w + 1) = linhas(w + 2)
                        Range("B" & Contador).Interior.Color = vbGreen
                    Next w
                    Contador = Contador + 1
                    Exit For
                End If
                
            Next i
            
            If i = UBound(linhas) Then 'teoricamente nuca vai entrar aqui mas é bom para segurança
                'MsgBox "Não encontrado : " & nome & " " & bm
                'i = UltimaLinha
                Range("B" & Contador).Interior.Color = vbRed
                Contador = Contador + 1
            End If
            
            
        Else 'se cairmos em um que não seja amarelo
            Contador = Contador + 1
        End If
        
        
        
        
    Loop
    ' Reescreve o arquivo com o conteúdo atualizado
    Open LocaldoArquivo For Output As #NumArquivo
    Print #NumArquivo, Join(linhas, vbCrLf) ' Converte o array de volta para uma única string com linhas separadas por vbCrLf
    Close #NumArquivo
        
End Sub



