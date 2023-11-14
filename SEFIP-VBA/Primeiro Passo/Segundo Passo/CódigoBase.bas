Attribute VB_Name = "CódigoBase"
Sub SEFIP(LocaldoArquivo As String)
    Dim LocaldoArquivo As String
    Dim NumArquivo As Integer
    Dim FileContent As String
    Dim Linha As String
     
    
    'LocaldoArquivo = "X:\GESFO - ISO\GERINS em ES250-89\SEFIP\EM ANDAMENTO\07-2021-CADM\SEFIP - Pedro.RE"
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
        
        For i = 3 To UBound(linhas) '6966
        'nome = Mid(linhas(i), 54, 70) ' Ajuste para 30 caracteres
        nomeTXT = Mid(linhas(i), 54, 70)
        bmTXT = Mid(linhas(i), 127, 8)
        DataAdimissaoTXT = Mid(linhas(i), 44, 8)
        
        
        If bmTXT = bm And DataAdimissaoTXT = DataAdmissao Then 'Or InStr(1, linhas(i), nome, vbTextCompare) > 0 Then  ' InStr(1, linhas(i), bm, vbTextCompare) > 0
        Codigo = Mid(linhas(i), 1, 18)
        Pis = Mid(linhas(i), 33, 11) ' Ajuste para 11 caracteres
        'DataAdmissao = Mid(linhas(i), 44, 8) ' Extrai a DataAdmissao corretamente
        Codigo20 = Mid(linhas(i), 48, 2)
        DataDemissao = Range("D" & Contador).Value
        DataDemissao = Replace(DataDemissao, "/", "")
        CodigoDesligamento = MotivoDesligamentoCodigo(DataDemissao, Contador) 'MotivoDesligamentoCodigo(DataDemissao, contador)
        DataNacimento = Mid(linhas(i), 155, 8)
        Codigo = Left(Codigo, 1) & "2" & Right(Codigo, Len(Codigo) - 2)
        
            'MsgBox "Valor encontrado na posição " & i
      
            
            ' Exibe as informações extraídas
            'MsgBox "Código: " & Codigo & vbCrLf & "Pis: " & Pis & vbCrLf & "Nome: " & nome & vbCrLf & "DataAdmissao: " & DataAdmissao & vbCrLf & "DataDemissao: " & DataDemissao & vbCrLf & "Codigo20: " & Codigo20 & vbCrLf & "BM: " & bm & vbCrLf & "DataNacimento: " & DataNacimento
    
            ' Crie o novo conteúdo com as informações modificadas
            NovoConteudo = Codigo & Space(14) & Mid(Pis & Space(11), 1, 11) & Mid(DataAdmissao & Space(8), 1, 8) & Codigo20 & nomeTXT & Mid(CodigoDesligamento, 1, 11) & Space(225) & "*"
 
            
            
            OriginalConteudo = linhas(i)
            ReDim Preserve linhas(UBound(linhas))
            linhas(UBound(linhas)) = "                                                                                                                                                                                                                                                                                                                                                                        " & vbCrLf
            
            ' Adiciona uma linha em branco ao final do array
            For c = UBound(linhas) To i Step -1
            DataAdmissaoTXTCima = Mid(linhas(c - 1), 44, 8)
            bmTXTCima = Mid(linhas(c - 1), 127, 8)
            PisTXTCima = Mid(linhas(c - 1), 33, 11)
                If bm = bmTXTCima And DataAdmissao = DataAdmissaoTXTCima And Pis = PisTXTCima Then
                    linhas(c) = NovoConteudo
                    Range("A" & Contador).Interior.Color = vbGreen
                    Exit For
                Else
                    If DataAdmissao = DataAdmissaoTXTCima And Pis = PisTXTCima Then ' posso tambem colocar o Pis para previnir erros de indentificação
                        linhas(c) = NovoConteudo
                        Range("A" & Contador).Interior.Color = vbYellow
                        Exit For
                    Else
                        linhas(c) = linhas(c - 1)
                    End If
                    
                End If
                
            
            Next c
            
                
                
            
            ' Reescreve o arquivo com o conteúdo atualizado
            Open LocaldoArquivo For Output As #NumArquivo
            Print #NumArquivo, Join(linhas, vbCrLf) ' Converte o array de volta para uma única string com linhas separadas por vbCrLf
            Close #NumArquivo

                
                Contador = Contador + 1 ' para na planilha passar para o BM de baixo
            
            
            Exit For
            Else
                If i = UBound(linhas) Then
                    'MsgBox "Não encontrado : " & nome & " " & bm
                    'i = UltimaLinha
                    Range("A" & Contador).Interior.Color = vbRed
                    Contador = Contador + 1
                    Exit For
                End If
            
        End If
     
     Next i
     i = 3
     
     
     
 Loop

End Sub


