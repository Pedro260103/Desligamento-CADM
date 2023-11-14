VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Home 
   Caption         =   "Home"
   ClientHeight    =   12195
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14745
   OleObjectBlob   =   "Home.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CodigoBase_Click()
    If TextBox2.Text <> "" Then
        SEFIP (TextBox2.Text)
    Else
        MsgBox "Campo Obrigatorio!"
    End If
    If TextBox2.Text <> "" Then
        If VerificarCaminhoArquivo(TextBox2.Text) Then
            TextBox2.Enabled = False
            CodigoBase.Enabled = False
        Else
            MsgBox "O caminho do arquivo não é válido."
        End If
    Else
        MsgBox "Campo obrigatório!"
    End If
End Sub

Public Sub EscolherArquivo_Click()
    Filename = Application.GetOpenFilename("Arquivos de Texto (*.txt), *.txt")
    If Filename <> False Then
        TextBox1.Text = Filename
        TextBox2.Text = Filename
    End If
    
End Sub

Private Sub Executar_Click()
    If TextBox1.Text <> "" Then
        If VerificarCaminhoArquivo(TextBox1.Text) Then
            ' Chamar a função que adiciona linhas em branco
            AddLinhasEmBranco TextBox1.Text
            AddLinhasEmBranco TextBox1.Text
            TextBox2.Enabled = True
            CodigoBase.Enabled = True
        Else
            MsgBox "O caminho do arquivo não é válido."
        End If
    Else
        MsgBox "Campo obrigatório!"
    End If
End Sub



Private Sub MultiPage1_Change()
    MsgBox "Teste"
End Sub




Private Sub UserForm_Initialize()
    
    
    TextBox2.Enabled = False
    CodigoBase.Enabled = False
    
    
End Sub









Function VerificarCaminhoArquivo(caminho As String) As Boolean
    ' Verificar se o caminho do arquivo é válido
    If Dir(caminho) <> "" Then
        VerificarCaminhoArquivo = True
    Else
        VerificarCaminhoArquivo = False
    End If
End Function

