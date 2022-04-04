VERSION 5.00
Begin VB.Form frmprincipal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Criando uma aplicação Multiusuário com ADO"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5190
      Left            =   -75
      ScaleHeight     =   5160
      ScaleWidth      =   1185
      TabIndex        =   0
      Top             =   -300
      Width           =   1215
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "Para criar mais de uma instância de sua aplicação (simulando um ambiente multiusuário) , clique duas vezes no ícone - Gravar Dados"
         ForeColor       =   &H8000000D&
         Height          =   2295
         Left            =   75
         TabIndex        =   3
         Top             =   375
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Sair"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   450
         TabIndex        =   2
         Top             =   4200
         Width           =   225
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         Caption         =   "Gravar Dados"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   225
         TabIndex        =   1
         Top             =   3225
         Width           =   840
      End
      Begin VB.Image imgsair 
         Height          =   480
         Left            =   375
         Picture         =   "frmprincipal.frx":0000
         Top             =   3675
         Width           =   480
      End
      Begin VB.Image imgincluir 
         Height          =   480
         Left            =   300
         Picture         =   "frmprincipal.frx":0442
         Top             =   2775
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmprincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    On Error GoTo trata_erros
    
    While Forms.Count > 1
        i = 0
        While Forms(i).Caption = Me.Caption
             i = i + 1
        Wend
        Unload Forms(i)
    Wend
    
    Unload Me
    End

    Exit Sub
trata_erros:
    MsgBox Err.Description
End Sub
Private Sub imgincluir_Click()
'Abre uma instância dodo formulario de dados

    Dim frmNovo As frmmultdados
    Static intFormContador As Integer
    
    On Error GoTo trata_erros
    
    intFormContador = intFormContador + 1
    
    Set frmNovo = New frmmultdados
    Load frmNovo
    frmNovo.Caption = "Formulários de Dados Multiusuário, Instancia => #" & intFormContador
    frmNovo.Show
    
    Exit Sub
trata_erros:
     MsgBox Err.Description, vbInformation, " Erros "

End Sub

Private Sub imgsair_Click()
  Unload Me
End Sub
