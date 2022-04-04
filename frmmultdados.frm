VERSION 5.00
Begin VB.Form frmmultdados 
   Caption         =   "Formulário de Dados"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5625
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   3720
   ScaleWidth      =   5625
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdeditar 
      Caption         =   "&Editar"
      Height          =   315
      Left            =   525
      TabIndex        =   0
      Top             =   3300
      Width           =   915
   End
   Begin VB.CommandButton cmdsalvar 
      Caption         =   "&Salvar"
      Height          =   315
      Left            =   525
      TabIndex        =   19
      Top             =   3300
      Width           =   915
   End
   Begin VB.CommandButton cmdsair 
      Caption         =   "&Sair"
      Height          =   315
      Left            =   4185
      TabIndex        =   13
      Top             =   3300
      Width           =   915
   End
   Begin VB.CommandButton cmdatualizar 
      Caption         =   "&Atualiza"
      Height          =   315
      Left            =   3270
      TabIndex        =   12
      Top             =   3300
      Width           =   915
   End
   Begin VB.CommandButton cmdexcluir 
      Caption         =   "&Excluir"
      Height          =   315
      Left            =   2355
      TabIndex        =   11
      Top             =   3300
      Width           =   915
   End
   Begin VB.CommandButton cmdincluir 
      Caption         =   "&Incluir"
      Height          =   315
      Left            =   1425
      TabIndex        =   1
      Top             =   3300
      Width           =   915
   End
   Begin VB.CommandButton cmdMoveLast 
      Height          =   300
      Left            =   4920
      Picture         =   "frmmultdados.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2925
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton cmdMoveNext 
      Height          =   300
      Left            =   4575
      Picture         =   "frmmultdados.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2925
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton cmdMoveFirst 
      Height          =   300
      Left            =   300
      Picture         =   "frmmultdados.frx":0684
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2925
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton cmdMovePrevious 
      Height          =   300
      Left            =   645
      Picture         =   "frmmultdados.frx":09C6
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2925
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Em Promoção"
      DataField       =   "ProdutoPromocao"
      Height          =   315
      Left            =   1200
      TabIndex        =   5
      Top             =   2250
      Width           =   1665
   End
   Begin VB.TextBox Text3 
      DataField       =   "ProdutoCategoria"
      Height          =   315
      Left            =   1200
      TabIndex        =   4
      Top             =   1725
      Width           =   765
   End
   Begin VB.TextBox Text2 
      DataField       =   "ProdutoNome"
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Top             =   1125
      Width           =   2265
   End
   Begin VB.TextBox Text1 
      DataField       =   "ProdutoID"
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Top             =   525
      Width           =   915
   End
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   1425
      TabIndex        =   20
      Top             =   3300
      Width           =   915
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   75
      X2              =   5550
      Y1              =   2700
      Y2              =   2700
   End
   Begin VB.Label Label6 
      Caption         =   "Este campo somente aceita categorias cadastradas "
      ForeColor       =   &H8000000D&
      Height          =   390
      Left            =   2100
      TabIndex        =   18
      Top             =   1725
      Width           =   3015
   End
   Begin VB.Label Label5 
      Caption         =   "Categoria"
      Height          =   240
      Left            =   75
      TabIndex        =   17
      Top             =   1800
      Width           =   915
   End
   Begin VB.Label Label4 
      Caption         =   "Nome Produto "
      Height          =   390
      Left            =   75
      TabIndex        =   16
      Top             =   1050
      Width           =   990
   End
   Begin VB.Label Label3 
      Caption         =   "Este Campo não pode ser editado pois é gerado incrementalmente pelo banco de dados"
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   2325
      TabIndex        =   15
      Top             =   300
      Width           =   3165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Código"
      Height          =   195
      Left            =   75
      TabIndex        =   14
      Top             =   525
      Width           =   495
   End
   Begin VB.Label lblstatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   975
      TabIndex        =   6
      Top             =   2925
      Width           =   3615
   End
End
Attribute VB_Name = "frmmultdados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum ModoNavegacaoBotao
       Addmode = 3
       EditMode = 4
End Enum

Enum MostraModoEdicao
       Navegacao = False
       Editando = True
End Enum

Private WithEvents mrsPrimary As Recordset
Attribute mrsPrimary.VB_VarHelpID = -1
Private Sub cmdatualizar_Click()
  ' exibe os dados mais recentes
  On Error GoTo Trata_Erro
  
  'altera a ampulheta
  Screen.MousePointer = vbHourglass
  mrsPrimary.Requery
  Screen.MousePointer = vbNormal
  Atualiza_Botoes_Navegacao_Posicao
    
    Exit Sub
Trata_Erro:
    Exibe_Erros (Err.Description)

End Sub
Private Sub cmdcancelar_Click()
 'Cancela as alteracoes
    
    On Error GoTo Trata_Erro
    mrsPrimary.CancelUpdate
    
    'Reverte os valores anteriores dos campos
    Modo_Edicao (Navegacao)
    Atualiza_Botoes_Navegacao_Posicao
    
    Exit Sub
Trata_Erro:
    Exibe_Erros (Err.Description)

End Sub
Private Sub cmdeditar_Click()
    On Error GoTo Trata_Erro
    
    mrsPrimary.Resync adAffectCurrent
    'pega o ultimo dado para editar
    mrsPrimary.Move 0  'atualiza os controles
    
    Modo_Edicao (Editando)
    Atualiza_Botoes_Navegacao_Posicao (EditMode)
    
    Exit Sub

Trata_Erro:
    Select Case Err.Number
        Case -2147217885 'a linha foi excluida
            MsgBox "Esta linha foi excluida por outro usuario...", vbInformation
        Case Else
            Exibe_Erros (Err.Description)
    End Select
        
End Sub
Private Sub cmdexcluir_Click()
    
    If MsgBox("O registro será excluido definitivamente. Continua ?", vbYesNoCancel + vbExclamation, "Confirma Exclusão") <> vbYes Then Exit Sub
        
    On Error Resume Next
    mrsPrimary.Delete
    
    Select Case Err.Number
        Case 0: 'exclusao foi um sucesso
            If cmdMoveNext.Enabled Then
                cmdMoveNext_Click
            Else
                If cmdMovePrevious.Enabled Then cmdMovePrevious_Click
            End If
        Case -2147217864
            MsgBox "Esta linha já foi excluida por outro usuário!", vbInformation
            mrsPrimary.CancelUpdate
        Case -2147467259
            MsgBox "As alterações feitas não podem ser salvas no momento. O registro encontra-se bloqueado pelo por outro usuario." & vbCr & "Voce pode cancelar as alteracoes ou tentar salvar mais tarde...", vbExclamation, "Erro de gravacao"
            Exit Sub
        Case Else
            MsgBox "O registro nao pode ser excluido." + vbCrLf + Err.Description
            mrsPrimary.CancelUpdate
    End Select
    Exit Sub
Trata_Erro:
    Exibe_Erros (Err.Description)

End Sub
Private Sub cmdsalvar_Click()
 'Salva

    Dim vFieldArray(), x As Integer, intUpdateError As Integer, strErrorMessage As String, oError As Error
    Dim blnAdd As Boolean
    
    On Error Resume Next 'limpa o objeto error
    
    blnAdd = mrsPrimary.EditMode = adEditAdd
    mrsPrimary.ActiveConnection.Errors.Clear
    
    Screen.MousePointer = vbHourglass
    mrsPrimary.Update
    Screen.MousePointer = vbNormal
    
    intUpdateError = Err.Number 'armazena os erros(O objeto error sera resetado pelo proxima linha)
    
    On Error GoTo TrataErros
    
    Select Case intUpdateError
        Case 0:
            If mrsPrimary.ActiveConnection.Errors.Count = 0 Then 'nao ocorreu nenhum erro
                Modo_Edicao (Navegacao)
                Atualiza_Botoes_Navegacao_Posicao
                If blnAdd Then
                    mrsPrimary.Resync adAffectCurrent 'exibe os valores padrao
                    mrsPrimary.Move 0 'forca uma atualiacao dos controles para exibir os dados
                End If
            Else
                For Each oError In mrsPrimary.ActiveConnection.Errors
                        If oError.Number = -2147217864 Then
                            strErrorMessage = "Este registro já foi recentemente alterado por outro usuário ! "
                            MsgBox strErrorMessage
                            Exit Sub
                        Else
                            strErrorMessage = strErrorMessage & oError.Description & vbCr
                        End If
                Next
                If mrsPrimary.ActiveConnection.Errors.Count = 1 Then strErrorMessage = "Os seguintes erros" & _
                    IIf(mrsPrimary.ActiveConnection.Errors.Count > 1, "foi(ram)", " foram ") & " definidos pelo provedor : " & vbCr & strErrorMessage
                MsgBox strErrorMessage ' exibe todos os erros
            End If
            
            
        Case 3640 + vbObjectError 'registro alterado por outro usuario
            If MsgBox("Another user has changed this record since you started editing it. If you save the record, you will overwrite the changes the other user made." & vbCr & vbCr & "Do you want to overwrite the other user's changes?", vbExclamation + vbYesNoCancel, "Write Conflict") = vbYes Then
                'forca uma sobrescrita dos dados devemos armazenar os dados em um buffer
                ReDim vFieldArray(mrsPrimary.Fields.Count - 1)
                For x = 0 To mrsPrimary.Fields.Count - 1
                    vFieldArray(x) = mrsPrimary.Fields(x).Value
                Next x
                
                mrsPrimary.CancelUpdate
                mrsPrimary.Resync adAffectCurrent
                
                For x = 0 To mrsPrimary.Fields.Count - 1 'salva as alteracoes no banco de dados
                    If mrsPrimary.Fields(x).Value <> vFieldArray(x) Then mrsPrimary.Fields(x) = vFieldArray(x)
                Next x
                
                mrsPrimary.Update
                Atualiza_Botoes_Navegacao_Posicao (EditMode)
                
                Modo_Edicao (Navegacao)
                Atualiza_Botoes_Navegacao_Posicao
            Else 'usuario escolheu nao sobrescrever os dados
                mrsPrimary.CancelUpdate
                mrsPrimary.Resync adAffectCurrent 'exibe os ultimos dados
                mrsPrimary.Move 0
            End If
        Case -2147467259
            MsgBox "As alterações feitas não podem ser salvas no momento. O registro encontra-se bloqueado pelo por outro usuario." & vbCr & "Voce pode cancelar as alteracoes ou tentar salvar mais tarde...", vbExclamation, "Erro de gravacao"
            Exit Sub
        Case Else:
                MsgBox Err.Description + vbCr & "(Origem: cmdsalvar_Click)", vbExclamation, "Error"
    End Select
    
    Exit Sub
TrataErros:
    If Err.Number = -2147217885 Then ' A chave para este registros foi alterado ou excluida
        MsgBox "Se for usar o Access 97 ele nao atualiza  o campo autonumeracao apos a atualizacao. Para exibir o codigo do produto atualizado pressione o botao atualizar...", vbInformation
        If cmdatualizar.Visible Then cmdatualizar.SetFocus
    Else
        Exibe_Erros (Err.Description)
    End If

End Sub

Private Sub Form_Load()

    On Error GoTo Trata_Erro
    
    Dim db As Connection
    Set db = New Connection
    db.CursorLocation = adUseClient
    db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\estoque.mdb;"
    
    Set mrsPrimary = New ADODB.Recordset
    mrsPrimary.Open "Select ProdutoID, ProdutoNome, ProdutoCategoria, ProdutoPromocao from Produtos", db, adOpenStatic, adLockOptimistic

    Conecta_Controles
    Modo_Edicao (Navegacao)

    cmdMoveFirst_Click
    
    Exit Sub
Trata_Erro:
    Exibe_Erros (Err.Description)

End Sub
Private Sub cmdSair_Click()
On Error GoTo TrataErros
    Unload Me

    Exit Sub
TrataErros:
    Exibe_Erros (Err.Description)
End Sub
Private Sub cmdMoveFirst_Click()
' vai para primeiro registro

    On Error GoTo TrataErros
    
    mrsPrimary.MoveFirst
    Atualiza_Botoes_Navegacao_Posicao
     
    Exit Sub
TrataErros:
    Exibe_Erros (Err.Description)
   
End Sub
Private Sub cmdMoveLast_Click()
 
    On Error GoTo TrataErros
    
    If mrsPrimary.EOF And mrsPrimary.BOF Then Exit Sub
    mrsPrimary.MoveLast
    Atualiza_Botoes_Navegacao_Posicao
     
    Exit Sub
TrataErros:
    Exibe_Erros (Err.Description)
     
End Sub
Private Sub cmdMoveNext_Click()
 'vai para o proximo registro
 
    On Error GoTo TrataErros
    
    mrsPrimary.MoveNext
        
    ' salta automaticamente as linhas excluidas
    Do While mrsPrimary.Status = adRecDBDeleted
        mrsPrimary.MoveNext
        If mrsPrimary.EOF Then
            mrsPrimary.MovePrevious
            Exit Do
        End If
    Loop
    Atualiza_Botoes_Navegacao_Posicao
     
    Exit Sub
TrataErros:
    Exibe_Erros (Err.Description)
         
End Sub
Private Sub cmdMovePrevious_Click()
 'move para o registro anterior

    On Error GoTo TrataErros

    mrsPrimary.MovePrevious
    
    ' salta registros deletados
    Do While mrsPrimary.Status = adRecDBDeleted
        mrsPrimary.MovePrevious
        If mrsPrimary.EOF Then
            mrsPrimary.MoveNext
            Exit Do
        End If
    Loop
    Atualiza_Botoes_Navegacao_Posicao
     
    Exit Sub
TrataErros:
    Exibe_Erros (Err.Description)
   
End Sub
Private Sub cmdIncluir_Click()
 'Inclui um novo registro

    On Error GoTo TrataErros

    Modo_Edicao (Editando)
    Atualiza_Botoes_Navegacao_Posicao Addmode
    mrsPrimary.AddNew
    Text2.SetFocus
         
    Exit Sub
TrataErros:
    Exibe_Erros (Err.Description)
   
End Sub
Sub Modo_Edicao(rblnEditMode As MostraModoEdicao)
 'navegacao: trava os controles e esconde os botoes salvar e cancela Lock Databound Controls and hide cancel and save buttons
 'editando: Destrava os controles e exibe botoes salvar e cancelar Unlock Databound Controls and show cancel and save buttons
    
    Dim oControl As Control
    Const EDIT_BACKCOLOUR As Long = vbYellow
    Const LOCKED_BACKCOLOUR As Long = &H8000000F
    
    On Error GoTo TrataErros
    
    For Each oControl In Me.Controls  'aplica um efeito visual aos controles
        Select Case TypeName(oControl)
            Case "TextBox", "DataCombo", "CheckBox"
                 'If oControl.DataField <> "" Then oControl.Enabled = rblnEditMode
                 'If mrsPrimary.Fields(oControl.DataField).Properties("ISAUTOINCREMENT") Then oControl.Enabled = False
                 'altera a cor de fundo
                 'oControl.Locked = rblnEditMode - nao funciona para todos os controles
                 oControl.BackColor = IIf(rblnEditMode, EDIT_BACKCOLOUR, LOCKED_BACKCOLOUR)
        End Select
    Next oControl
   
   'esconde ou exibe os botoes
    cmdincluir.Visible = Not rblnEditMode
    cmdexcluir.Visible = Not rblnEditMode
    cmdeditar.Visible = Not rblnEditMode
    cmdatualizar.Visible = Not rblnEditMode
    cmdsair.Visible = Not rblnEditMode
    cmdsalvar.Visible = rblnEditMode
    cmdcancelar.Visible = rblnEditMode
         
    Exit Sub
TrataErros:
    Exibe_Erros (Err.Description)
   
End Sub
Sub Conecta_Controles()
 'forca a atualizacao da exibicao dos dados
 
    On Error GoTo Trata_Erro
        
    Dim oControle As Control
    Dim i As Integer
    
    For Each oControle In Me.Controls
        If TypeName(oControle) = "TextBox" Or TypeName(oControle) = "CheckBox" Then
            Set oControle.DataSource = mrsPrimary
        End If
   Next oControle
        
    Exit Sub
Trata_Erro:
    Exibe_Erros (Err.Description)
   
End Sub
Sub Atualiza_Botoes_Navegacao_Posicao(Optional rModo As ModoNavegacaoBotao)
' desabilita os botoes de navegacao ao fim do recordset
' e atualiza a posicao do registro

    Dim blnCanMoveForward As Boolean, blnCanMoveBack As Boolean
    On Error GoTo TrataErros
    
    blnCanMoveForward = True
    blnCanMoveBack = True
    
    Select Case rModo
        Case Addmode
            lblStatus = "Novo registro"
        Case EditMode
                lblStatus = " Editando o registro -> " & CStr(mrsPrimary.AbsolutePosition)
        Case Else
                lblStatus = "  Registro  " & CStr(mrsPrimary.AbsolutePosition) & " de " & mrsPrimary.RecordCount
    End Select
        
    'desabilida a navegacao enquanto estiver na edicao
    If rModo = Addmode Or rModo = EditMode Then
        cmdMoveLast.Enabled = False
        cmdMoveNext.Enabled = False
        cmdMoveFirst.Enabled = False
        cmdMovePrevious.Enabled = False
    Else
        'habilita botoes apropriados
        mrsPrimary.MoveNext
        If mrsPrimary.EOF Then blnCanMoveForward = False
        mrsPrimary.MovePrevious 'volta
        
        mrsPrimary.MovePrevious
        If mrsPrimary.BOF Then blnCanMoveBack = False
        mrsPrimary.MoveNext 'volta
        
        'define botoes
        cmdMoveLast.Enabled = blnCanMoveForward
        cmdMoveNext.Enabled = blnCanMoveForward
        cmdMoveFirst.Enabled = blnCanMoveBack
        cmdMovePrevious.Enabled = blnCanMoveBack
    End If
    
    Exit Sub
TrataErros:
    MsgBox Err.Description & " (Fonte: Atualiza_Botoes_Navegacao_Posicao)", vbExclamation, "Erro !"
   
End Sub
Private Sub Exibe_Erros(strMensagem_Erro As String)
    On Error Resume Next
    MsgBox strMensagem_Erro, vbInformation, "Ocorreu um erro !"
End Sub
Private Sub Form_Unload(Cancel As Integer)
 'libera os recursos
  On Error GoTo Trata_Erro

    mrsPrimary.Close
    Set mrsPrimary = Nothing
    
    Exit Sub
Trata_Erro:
    Exibe_Erros (Err.Description)

End Sub
Private Sub mrsPrimary_FieldChangeComplete(ByVal cFields As Long, ByVal Fields As Variant, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
 ' verifica dados invalidos inseridos pelo usuario
 ' este evento e disparado quando o foco deixa o controle que esta sendo editado
 ' é chamado novamente depois do metodo update ser invocado
    
    On Error GoTo TrataErros
    
    If adStatus = adStatusErrorsOccurred Then
        Beep
        MsgBox "O valor informado não é valido para este campo." & vbCr & vbCr & _
            "Ex: Voce entrou um valor numerico para um campo texto.", vbInformation
        Me.ActiveControl.Text = mrsPrimary.Fields(Me.ActiveControl.DataField).OriginalValue
        'restaura os valores anteriores no controle
    End If
    
    Exit Sub
TrataErros:
    Exibe_Erros (Err.Description)
    
End Sub


