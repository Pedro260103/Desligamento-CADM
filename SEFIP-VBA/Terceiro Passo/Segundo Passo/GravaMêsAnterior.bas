Attribute VB_Name = "GravaM�sAnterior"
Sub GravarMesAnterior()
    Dim LocalDoArquivo As String
    Dim NumArquivo As Integer
    Dim FileContent As String
    Dim Linha As String
     
    
    LocalDoArquivo = "X:\GESFO - ISO\GERINS em ES250-89\SEFIP\2018-RETIFICADORAS\08-2018-(1� Retificadora)\CADM\SEFIP - Codigo.RE"
    NumArquivo = FreeFile
    
    Open LocalDoArquivo For Input As #NumArquivo ' Abre o arquivo em modo de leitura
    
    FileContent = Input$(LOF(1), #NumArquivo)
    Close #NumArquivo
 
    ' Extrai informa��es da linha
    Linha = Trim(FileContent) ' Remove espa�os em branco no in�cio e no fim
 
    Dim Codigo As String
    Dim Pis As String
    Dim Nome As String
    Dim nomeTXT As String
    Dim DataAdimissao As String
    Dim DataAdimissaoTXT As String
    Dim Codigo20 As String
    Dim CodigoDesligamento As String
    Dim BM As String
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
    Dim HValorOriginal As String
    
    Dim IValorDesligamento As String
    
    
    UltimaLinhaCodigo90 = linhas(UBound(linhas) - 1)
    'AddLinhasEmBranco (UBound(linhas))
    
    Do While Range("A" & Contador).Value <> ""
        BM = Range("A" & Contador).Value
        BM = Replace(BM, "-", "")
        BM = Replace(BM, "X", "0")
        Nome = Range("B" & Contador).Value
        DataAdmissao = Range("C" & Contador).Value
        DataAdmissao = Replace(DataAdmissao, "/", "")
        HValorOriginal = Range("H" & Contador).Value
        IValorDesligamento = Range("I" & Contador).Value
        
        If Range("A" & Contador).Interior.Color = vbRed Then
        
            For i = 3 To UBound(linhas) '6966
            'nome = Mid(linhas(i), 54, 70) ' Ajuste para 30 caracteres
            nomeTXT = Mid(HValorOriginal, 54, 70)
            bmTXT = Mid(HValorOriginal, 127, 8)
            DataAdimissaoTXT = Mid(HValorOriginal, 44, 8)
            
            
            'If bmTXT = bm And DataAdimissaoTXT = DataAdmissao Then 'Or InStr(1, linhas(i), nome, vbTextCompare) > 0 Then  ' InStr(1, linhas(i), bm, vbTextCompare) > 0
            Codigo = Mid(HValorOriginal, 1, 18)
            Pis = Mid(HValorOriginal, 33, 11) ' Ajuste para 11 caracteres
            'DataAdmissao = Mid(linhas(i), 44, 8) ' Extrai a DataAdmissao corretamente
            Codigo20 = Mid(HValorOriginal, 48, 2)
            DataDemissao = Range("D" & Contador).Value
            DataDemissao = Replace(DataDemissao, "/", "")
            'CodigoDesligamento = MotivoDesligamentoCodigo(DataDemissao, contador) 'MotivoDesligamentoCodigo(DataDemissao, contador)
            DataNacimento = Mid(HValorOriginal, 155, 8)
            ''Codigo = Left(Codigo, 1) & "2" & Right(Codigo, Len(Codigo) - 2)
            
                'MsgBox "Valor encontrado na posi��o " & i
          
                
                ' Exibe as informa��es extra�das
                'MsgBox "C�digo: " & Codigo & vbCrLf & "Pis: " & Pis & vbCrLf & "Nome: " & nome & vbCrLf & "DataAdmissao: " & DataAdmissao & vbCrLf & "DataDemissao: " & DataDemissao & vbCrLf & "Codigo20: " & Codigo20 & vbCrLf & "BM: " & bm & vbCrLf & "DataNacimento: " & DataNacimento
        
                ' Crie o novo conte�do com as informa��es modificadas
                NovoConteudo = IValorDesligamento 'Codigo & Space(14) & Mid(Pis & Space(11), 1, 11) & Mid(DataAdmissao & Space(8), 1, 8) & Codigo20 & nomeTXT & Mid(CodigoDesligamento, 1, 11) & Space(225) & "*"
     
                
                
                OriginalConteudo = HValorOriginal
                
                
                
                OriginalConteudo = Left(OriginalConteudo, 175) & "0000001" & Mid(OriginalConteudo, 183) ' 177 183
                
                ' Substituir os caracteres na posi��o 57 por 5 zeros
                OriginalConteudo = Left(OriginalConteudo, 210) & "000000" & Mid(OriginalConteudo, 217) ' 212 217
                
                ' Substituir os caracteres na posi��o 71 por 6 zeros
                OriginalConteudo = Left(OriginalConteudo, 224) & "0000000" & Mid(OriginalConteudo, 232) ' 226 232
                ' Adiciona uma linha em branco ao final do array
                
                
                ReDim Preserve linhas(UBound(linhas))
                'linhas(UBound(linhas)) = "                                                                                                                                                                                                                                                                                                                                                                        " & vbCrLf
                
                
                For c = UBound(linhas) To i Step -1
                Dim DataAdmissaoTXTBaixo As String
                    
                    If c <> UBound(linhas) Then
                    DataAdmissaoTXTBaixo = Mid(linhas(c + 1), 44, 8)
                    End If
                    
                DataAdmissaoTXTCima = Mid(linhas(c - 1), 44, 8)
                bmTXTCima = Mid(linhas(c - 1), 127, 8)
                PisTXTCima = Mid(linhas(c - 1), 33, 11)
                
                Dim PisInt As String
                Dim PisTXTCimaInt As String
                
                
                If PisTXTCima <> "*" And PisTXTCima <> "" And Pis <> "" And Pis <> "*" Then
                    PisInt = ConverterStringParaInteger(Pis)
                    PisTXTCimaInt = ConverterStringParaInteger(PisTXTCima)
                    'DataAdmissao = Range("C" & contador).Value
                    
                    Dim DataAdmissao2 As String
                    DataAdmissao2 = Range("C" & Contador).Value
                    DataAdmissao2 = Replace(DataAdmissao, "/", "")
                    DataAdmissaoTXTCima = ConverterDataComBarra(DataAdmissaoTXTCima) ' de 26012003 para 26/01/2003
                    
                    
                    If PisInt > PisTXTCimaInt And Range("G" & Contador).Interior.Color = vbRed Then 'And CompararDatas(DataAdmissaoTXTCima, DataAdmissao2) Then ' erro aqui
                        linhas(c) = OriginalConteudo
                        linhas(c + 1) = NovoConteudo
                        Range("A" & Contador).Interior.Color = vbGreen
                        Range("G" & Contador).Value = "ok"
                        Exit For
                    Else
                        'DataAdmissaoTXTCima = ConverterDataSemBarra(DataAdmissaoTXTCima) ' de 26/01/2003 para 26012003
                        DataAdmissaoTXTCima = Replace(DataAdmissaoTXTCima, "/", "")
                        If DataAdmissao = DataAdmissaoTXTCima And Pis = PisTXTCima And BM = bmTXTCima And Range("G" & Contador).Interior.Color = vbRed Then  ' posso tambem colocar o Pis para previnir erros de indentifica��o
                            linhas(c) = NovoConteudo
                            Range("A" & Contador).Interior.Color = vbYellow
                            Exit For
                        Else
                            For y = 0 To 1 Step 1
                                linhas(c + 1) = linhas(c - 1)
                                linhas(c) = linhas(c - 2)
                            Next y
                            linhas(c) = linhas(c - 1)
                            'linhas(c - 1) = linhas(c - 2)
                        End If
                        
                    End If
                End If
                    
                
                    
                    
                
                Next c
                
                    
                    
                
                ' Reescreve o arquivo com o conte�do atualizado
                Open LocalDoArquivo For Output As #NumArquivo
                Print #NumArquivo, Join(linhas, vbCrLf) ' Converte o array de volta para uma �nica string com linhas separadas por vbCrLf
                Close #NumArquivo
    
                    
                    Contador = Contador + 1 ' para na planilha passar para o BM de baixo
                
                
                Exit For
                
                    If i = UBound(linhas) Then
                        'MsgBox "N�o encontrado : " & nome & " " & bm
                        'i = UltimaLinha
                        Range("A" & Contador).Interior.Color = vbRed
                        Contador = Contador + 1
                        Exit For
                    End If
                
            'End If
         
         Next i
         i = 3
        
        
        
        
        
        Else
            Contador = Contador + 1
        End If
        
            
         
     
     
 Loop






End Sub
Function ConverterDataComBarra(ByVal DataAdmissaoTXTCima As String) As String
    ' Verifique se a string tem o comprimento adequado (8 caracteres)
    If Len(DataAdmissaoTXTCima) <> 8 Then
        ConverterDataComBarra = "Data inv�lida"
        Exit Function
    End If
    
    ' Extrair dia, m�s e ano da string
    Dim dia As String
    Dim mes As String
    Dim ano As String
    
    dia = Left(DataAdmissaoTXTCima, 2)
    mes = Mid(DataAdmissaoTXTCima, 3, 2)
    ano = Right(DataAdmissaoTXTCima, 4)
    
    ' Formatar a data no formato desejado
    ConverterDataComBarra = dia & "/" & mes & "/" & ano
End Function
