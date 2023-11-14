Attribute VB_Name = "MotivoDesligamento"
Function MotivoDesligamentoCodigo(d As String, c As LongLong) As String
    Select Case Trim(Range("E" & c).Value) ' Trim tira os espa�os
        Case "RESCISAO A PEDIDO"
            MotivoDesligamentoCodigo = " J" + d + " "
         
        Case "EXONERA��O A PEDIDO"
            MotivoDesligamentoCodigo = " J" + d + " "
            
        Case "DEMISS�O/RESCIS�O POR JUSTA CAUSA"
            MotivoDesligamentoCodigo = " H" + d + " "
            
        Case "EXONERA��O POR INICIATIVA DO EXECUTIVO"
            MotivoDesligamentoCodigo = "I1" + d + "N"
            
        Case "T�RMINO DO CONTRATO"
            MotivoDesligamentoCodigo = "I3" + d + "N"
            
        Case "FALECIMENTO SERVIDOR - CELETISTA"
            MotivoDesligamentoCodigo = "S2" + d + " "
            
        Case "RESCIS�O DE CONTRATO SEM JUSTA CAUSA"
            MotivoDesligamentoCodigo = "I1" + d + "N"
         
        Case Else
            MotivoDesligamentoCodigo = "Valor invalido"
    End Select
 
End Function
