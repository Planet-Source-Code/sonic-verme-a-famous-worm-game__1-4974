VERSION 5.00
Begin VB.Form frmHiscores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hi-Scores..."
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmHiscores.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   150
      Picture         =   "frmHiscores.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   150
      Width           =   510
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      Caption         =   "10."
      Height          =   195
      Index           =   11
      Left            =   135
      TabIndex        =   32
      Top             =   3825
      Width           =   225
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      Caption         =   "09."
      Height          =   195
      Index           =   10
      Left            =   135
      TabIndex        =   31
      Top             =   3495
      Width           =   225
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      Caption         =   "08."
      Height          =   195
      Index           =   9
      Left            =   135
      TabIndex        =   30
      Top             =   3180
      Width           =   225
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      Caption         =   "07."
      Height          =   195
      Index           =   8
      Left            =   135
      TabIndex        =   29
      Top             =   2850
      Width           =   225
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      Caption         =   "06."
      Height          =   195
      Index           =   7
      Left            =   135
      TabIndex        =   28
      Top             =   2520
      Width           =   225
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      Caption         =   "05."
      Height          =   195
      Index           =   6
      Left            =   135
      TabIndex        =   27
      Top             =   2205
      Width           =   225
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      Caption         =   "04."
      Height          =   195
      Index           =   5
      Left            =   135
      TabIndex        =   26
      Top             =   1875
      Width           =   225
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      Caption         =   "03."
      Height          =   195
      Index           =   4
      Left            =   135
      TabIndex        =   25
      Top             =   1545
      Width           =   225
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      Caption         =   "02."
      Height          =   195
      Index           =   3
      Left            =   135
      TabIndex        =   24
      Top             =   1230
      Width           =   225
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      Caption         =   "01."
      Height          =   195
      Index           =   2
      Left            =   135
      TabIndex        =   23
      Top             =   900
      Width           =   225
   End
   Begin VB.Label lblPontos 
      Alignment       =   1  'Right Justify
      Caption         =   "1000000000"
      Height          =   210
      Index           =   0
      Left            =   3375
      TabIndex        =   22
      Top             =   900
      Width           =   1110
   End
   Begin VB.Label lblNome 
      Caption         =   "Player1"
      Height          =   210
      Index           =   0
      Left            =   435
      TabIndex        =   21
      Top             =   900
      Width           =   2820
   End
   Begin VB.Label lblPontos 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   210
      Index           =   9
      Left            =   3375
      TabIndex        =   20
      Top             =   3825
      Width           =   1110
   End
   Begin VB.Label lblNome 
      Caption         =   "Player10"
      Height          =   210
      Index           =   9
      Left            =   435
      TabIndex        =   19
      Top             =   3825
      Width           =   2820
   End
   Begin VB.Label lblPontos 
      Alignment       =   1  'Right Justify
      Caption         =   "10"
      Height          =   210
      Index           =   8
      Left            =   3375
      TabIndex        =   18
      Top             =   3500
      Width           =   1110
   End
   Begin VB.Label lblNome 
      Caption         =   "Player9"
      Height          =   210
      Index           =   8
      Left            =   435
      TabIndex        =   17
      Top             =   3495
      Width           =   2820
   End
   Begin VB.Label lblPontos 
      Alignment       =   1  'Right Justify
      Caption         =   "100"
      Height          =   210
      Index           =   7
      Left            =   3375
      TabIndex        =   16
      Top             =   3175
      Width           =   1110
   End
   Begin VB.Label lblNome 
      Caption         =   "Player8"
      Height          =   210
      Index           =   7
      Left            =   435
      TabIndex        =   15
      Top             =   3180
      Width           =   2820
   End
   Begin VB.Label lblPontos 
      Alignment       =   1  'Right Justify
      Caption         =   "1000"
      Height          =   210
      Index           =   6
      Left            =   3375
      TabIndex        =   14
      Top             =   2850
      Width           =   1110
   End
   Begin VB.Label lblNome 
      Caption         =   "Player7"
      Height          =   210
      Index           =   6
      Left            =   435
      TabIndex        =   13
      Top             =   2850
      Width           =   2820
   End
   Begin VB.Label lblPontos 
      Alignment       =   1  'Right Justify
      Caption         =   "10000"
      Height          =   210
      Index           =   5
      Left            =   3375
      TabIndex        =   12
      Top             =   2525
      Width           =   1110
   End
   Begin VB.Label lblNome 
      Caption         =   "Player6"
      Height          =   210
      Index           =   5
      Left            =   435
      TabIndex        =   11
      Top             =   2520
      Width           =   2820
   End
   Begin VB.Label lblPontos 
      Alignment       =   1  'Right Justify
      Caption         =   "100000"
      Height          =   210
      Index           =   4
      Left            =   3375
      TabIndex        =   10
      Top             =   2200
      Width           =   1110
   End
   Begin VB.Label lblNome 
      Caption         =   "Player5"
      Height          =   210
      Index           =   4
      Left            =   435
      TabIndex        =   9
      Top             =   2205
      Width           =   2820
   End
   Begin VB.Label lblPontos 
      Alignment       =   1  'Right Justify
      Caption         =   "1000000"
      Height          =   210
      Index           =   3
      Left            =   3375
      TabIndex        =   8
      Top             =   1875
      Width           =   1110
   End
   Begin VB.Label lblNome 
      Caption         =   "Player4"
      Height          =   210
      Index           =   3
      Left            =   435
      TabIndex        =   7
      Top             =   1875
      Width           =   2820
   End
   Begin VB.Label lblPontos 
      Alignment       =   1  'Right Justify
      Caption         =   "10000000"
      Height          =   210
      Index           =   2
      Left            =   3375
      TabIndex        =   6
      Top             =   1550
      Width           =   1110
   End
   Begin VB.Label lblNome 
      Caption         =   "Player3"
      Height          =   210
      Index           =   2
      Left            =   435
      TabIndex        =   5
      Top             =   1545
      Width           =   2820
   End
   Begin VB.Label lblPontos 
      Alignment       =   1  'Right Justify
      Caption         =   "100000000"
      Height          =   210
      Index           =   1
      Left            =   3375
      TabIndex        =   4
      Top             =   1225
      Width           =   1110
   End
   Begin VB.Label lblNome 
      Caption         =   "Player2"
      Height          =   210
      Index           =   1
      Left            =   435
      TabIndex        =   3
      Top             =   1230
      Width           =   2820
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The Hall of Fame"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   480
      Index           =   1
      Left            =   960
      TabIndex        =   2
      Top             =   180
      Width           =   3450
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The Hall of Fame"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Index           =   0
      Left            =   945
      TabIndex        =   1
      Top             =   165
      Width           =   3450
   End
End
Attribute VB_Name = "frmHiscores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' /////////////////////////////////////////////////////////////
' //
' // This form is used to show the hi-score table.
' // Welcome to Hall of Fame!
' //
' // Feel free to mail me if you have any trouble with this
' // code. I´ll help you, if possible
' //
' //    Recife, 17 de Dezembro de 1999
' //    Rafael Winter Teixeira (sonic@torricelli.com.br)
' //
' /////////////////////////////////////////////////////////////

Option Explicit

' Declare hi-score array
Private Recordes(0 To 9) As TRecorde


Private Sub Form_Load()

    Dim i As Integer
    Dim Recorde As TRecorde
    Dim FileHandle As Integer
    Dim bSwap As Boolean
    
    ' If RECORDES.DAT (hi-score file) is not found, so create
    ' it with sample values
    If Dir(App.Path & "\Recordes.dat") = "" Then
        For i = 0 To 9
            Recorde.Nome = "Worm" & Str(i + 1)
            Recorde.Nivel = (i \ 3) + 1
            Recorde.Score = (i \ 2) + i * 10
            Recordes(i) = Recorde
        Next i
            
        FileHandle = FreeFile
        Open App.Path & "\Recordes.dat" For Binary As #FileHandle
        Put #FileHandle, , Recordes
        Close #FileHandle
    End If
    
    ' Open RECORDES.DAT to get the hi-scores
    FileHandle = FreeFile
    Open App.Path & "\Recordes.dat" For Binary As #FileHandle
    Get #FileHandle, , Recordes
    Close #FileHandle

    ' Make decreasing sort
    Do
        bSwap = False
        For i = 0 To 8
            If Recordes(i).Score < Recordes(i + 1).Score Then
                Recorde = Recordes(i + 1)
                Recordes(i + 1) = Recordes(i)
                Recordes(i) = Recorde
                bSwap = True
            End If
        Next i
    Loop While bSwap
    
    ' Put the hi-scores in the labels
    For i = 0 To 9
        lblNome(i).Caption = Recordes(i).Nome
        lblPontos(i).Caption = Recordes(i).Score
    Next
    
End Sub
