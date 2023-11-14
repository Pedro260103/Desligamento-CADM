Attribute VB_Name = "AddOsNaMão"
Sub AddFaltantesArte()
    Dim LocalDoArquivo As String
    Dim NumArquivo As Integer
    Dim FileContent As String
    Dim Linha As String
     
    
    LocalDoArquivo = "X:\GESFO - ISO\GERINS em ES250-89\SEFIP\2020-RETIFICADORAS\02-2020-(1ª Retificadora-DESLIGAMENTOS)\CADM\SEFIP - Código.RE"
    NumArquivo = FreeFile
    
    Open LocalDoArquivo For Input As #NumArquivo ' Abre o arquivo em modo de leitura
    
    FileContent = Input$(LOF(1), #NumArquivo)
    Close #NumArquivo
 
    ' Extrai informações da linha
    Linha = Trim(FileContent) ' Remove espaços em branco no início e no fim
 
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
        
        If Range("A" & Contador).Interior.Color = vbRed Then
        HValorOriginal = "30118715383000140               101081505822805201920THERESA CHRISTINA FURTADO                                             00000587180                    1807195202251000000000000001000000000000000  05000000000000000000000000000000000000000000000000000000000000                                                                                                  *"
        IValorDesligamento = "32118715383000140               101081505822805201920THERESA CHRISTINA FURTADO                                              J31012020                                                                                                                                                                                                                                  *"
        BM = Range("A" & Contador).Value
        BM = Replace(BM, "-", "")
        BM = Replace(BM, "X", "0")
        Nome = Range("B" & Contador).Value
        
        DataAdmissao = Range("C" & Contador).Value
        DataAdmissao = Replace(DataAdmissao, "/", "")
        DataDemissao = Range("D" & Contador).Value
        DataDemissao = Replace(DataDemissao, "/", "")
        'CodigoDesligamento = MotivoDesligamentoCodigo(DataDemissao, contador) 'MotivoDesligamentoCodigo(DataDemissao, contador)
        DataNacimento = InputBox("Escreva a Data de Nacimento do " & Nome & ":", "Data Nacimento " & Contador & " de 68", BM)         'Mid(HValorOriginal, 155, 8) 0121607-2
        DataNacimento = Replace(DataNacimento, "/", "")
        Pis = InputBox("Escreva o Pis do " & Nome & ":", "Pis " & Contador & " de 68", BM)         'Mid(HValorOriginal, 33, 11)
        Pis = Replace(Pis, ".", "")
        
        CodigoDesligamento = MotivoDesligamentoCodigo(DataDemissao, Contador)
        Range("G" & Contador).Interior.Color = vbRed
            
        Codigo20 = Mid(HValorOriginal, 48, 2)
        Codigo = Mid(HValorOriginal, 1, 18)
        'Mid(testString, 14, 4)
        
        HValorOriginal = Replace(HValorOriginal, "10108150582", Pis)
        HValorOriginal = Replace(HValorOriginal, "28052019", DataAdmissao)
        HValorOriginal = Replace(HValorOriginal, "THERESA CHRISTINA FURTADO                                             ", Nome & Space(70 - Len(Nome)))
        HValorOriginal = Replace(HValorOriginal, "00587180", BM)
        HValorOriginal = Replace(HValorOriginal, "18071952", DataNacimento)
        
        OriginalConteudo = HValorOriginal 'Codigo & Space(14) & Mid(Pis & Space(11), 1, 11) & Mid(DataAdmissao & Space(8), 1, 8) & "20" & Nome & "000" & BM & Space(20) & DataNacimento & "02251000000000000001000000000000000  05000000000000000000000000000000000000000000000000000000000000" & "                                                                                                  *" 'Mid(linhas(i), 127, 8)
        HValorOriginal = OriginalConteudo
        Codigo = Left(Codigo, 1) & "2" & Right(Codigo, Len(Codigo) - 2)
        
            'MsgBox "Valor encontrado na posição " & i
      
            
            ' Exibe as informações extraídas
            'MsgBox "Código: " & Codigo & vbCrLf & "Pis: " & Pis & vbCrLf & "Nome: " & nome & vbCrLf & "DataAdmissao: " & DataAdmissao & vbCrLf & "DataDemissao: " & DataDemissao & vbCrLf & "Codigo20: " & Codigo20 & vbCrLf & "BM: " & bm & vbCrLf & "DataNacimento: " & DataNacimento
    
            ' Crie o novo conteúdo com as informações modificadas
            
            NovoConteudo = Codigo & Space(14) & Mid(Pis & Space(11), 1, 11) & Mid(DataAdmissao & Space(8), 1, 8) & "20" & Nome & Space(70 - Len(Nome)) & Mid(CodigoDesligamento, 1, 11) & Space(225) & "*"
            IValorDesligamento = NovoConteudo
            
            
            
            'If bmTXT = bm And DataAdimissaoTXT = DataAdmissao Then 'Or InStr(1, linhas(i), nome, vbTextCompare) > 0 Then  ' InStr(1, linhas(i), bm, vbTextCompare) > 0
            
            'Pis = Mid(HValorOriginal, 33, 11) ' Ajuste para 11 caracteres
            'DataAdmissao = Mid(linhas(i), 44, 8) ' Extrai a DataAdmissao corretamente
           
            
            
            ''Codigo = Left(Codigo, 1) & "2" & Right(Codigo, Len(Codigo) - 2)
            
                'MsgBox "Valor encontrado na posição " & i
          
                
                ' Exibe as informações extraídas
                'MsgBox "Código: " & Codigo & vbCrLf & "Pis: " & Pis & vbCrLf & "Nome: " & nome & vbCrLf & "DataAdmissao: " & DataAdmissao & vbCrLf & "DataDemissao: " & DataDemissao & vbCrLf & "Codigo20: " & Codigo20 & vbCrLf & "BM: " & bm & vbCrLf & "DataNacimento: " & DataNacimento
        
                ' Crie o novo conteúdo com as informações modificadas
                NovoConteudo = IValorDesligamento 'Codigo & Space(14) & Mid(Pis & Space(11), 1, 11) & Mid(DataAdmissao & Space(8), 1, 8) & Codigo20 & nomeTXT & Mid(CodigoDesligamento, 1, 11) & Space(225) & "*"
     
                
                
                OriginalConteudo = HValorOriginal
                
                
                
                OriginalConteudo = Left(OriginalConteudo, 175) & "0000001" & Mid(OriginalConteudo, 183) ' 177 183
                
                ' Substituir os caracteres na posição 57 por 5 zeros
                OriginalConteudo = Left(OriginalConteudo, 210) & "000000" & Mid(OriginalConteudo, 217) ' 212 217
                
                ' Substituir os caracteres na posição 71 por 6 zeros
                OriginalConteudo = Left(OriginalConteudo, 224) & "0000000" & Mid(OriginalConteudo, 232) ' 226 232
                ' Adiciona uma linha em branco ao final do array
                
        
        
        
        
        
        
        
        
            Range("H" & Contador).Value = HValorOriginal
            Range("I" & Contador).Value = IValorDesligamento
            
        
            Contador = Contador + 1
        Else
            Contador = Contador + 1
        End If
 Loop
End Sub
