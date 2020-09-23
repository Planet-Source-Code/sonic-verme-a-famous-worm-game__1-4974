VERSION 5.00
Begin VB.Form frmNewHiscore 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Hi-Score!!"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3765
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   2235
      TabIndex        =   5
      Top             =   2220
      Width           =   1335
   End
   Begin VB.TextBox txtName 
      Height          =   330
      Left            =   225
      MaxLength       =   30
      TabIndex        =   0
      Text            =   "Worm Name!"
      Top             =   1665
      Width           =   3345
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   270
      Picture         =   "frmNewHiscore.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   225
      Width           =   510
   End
   Begin VB.Label LabelX 
      Caption         =   "You just achieved a hi-score. Type your name to enter in the Hall of Fame."
      Height          =   465
      Index           =   2
      Left            =   270
      TabIndex        =   4
      Top             =   1035
      Width           =   3300
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Congratulations!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   300
      Width           =   2625
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Congratulations!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   945
      TabIndex        =   2
      Top             =   285
      Width           =   2625
   End
End
Attribute VB_Name = "frmNewHiscore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' /////////////////////////////////////////////////////////////
' //
' // This form is used to pick player´s name, when a hi-score
' // is achieved
' //
' // Feel free to mail me if you have any trouble with this
' // code. I´ll help you, if possible
' //
' //    Recife, 17 de Dezembro de 1999
' //    Rafael Winter Teixeira (sonic@torricelli.com.br)
' //
' /////////////////////////////////////////////////////////////

Option Explicit

Private mWormName As String


Public Property Get WormName() As String
    WormName = mWormName
End Property

Private Sub cmdOK_Click()
    mWormName = txtName.Text
    Unload Me
End Sub
