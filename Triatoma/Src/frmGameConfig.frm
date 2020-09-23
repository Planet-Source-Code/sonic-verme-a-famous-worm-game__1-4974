VERSION 5.00
Begin VB.Form frmGameConfig 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Game Settings"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2670
   Icon            =   "frmGameConfig.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   2670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboLevel 
      Height          =   315
      ItemData        =   "frmGameConfig.frx":0442
      Left            =   705
      List            =   "frmGameConfig.frx":045B
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   150
      Width           =   1005
   End
   Begin VB.TextBox txtVertSize 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1395
      TabIndex        =   3
      Top             =   1185
      Width           =   1125
   End
   Begin VB.TextBox txtHorzSize 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1395
      TabIndex        =   1
      Top             =   765
      Width           =   1125
   End
   Begin VB.Line LineX 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   105
      X2              =   2550
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line LineX 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   105
      X2              =   2550
      Y1              =   585
      Y2              =   585
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      Caption         =   "Vertical size"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   165
      TabIndex        =   4
      Top             =   1245
      Width           =   840
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      Caption         =   "Horizontal size"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   165
      TabIndex        =   2
      Top             =   825
      Width           =   1035
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      Caption         =   "Level"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   165
      TabIndex        =   0
      Top             =   195
      Width           =   375
   End
End
Attribute VB_Name = "frmGameConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' /////////////////////////////////////////////////////////////
' //
' // This form is used to show and change game configuration
' //
' // Feel free to mail me if you have any trouble with this
' // code. I´ll help you, if possible
' //
' //    Recife, 17 de Dezembro de 1999
' //    Rafael Winter Teixeira (sonic@torricelli.com.br)
' //
' /////////////////////////////////////////////////////////////

Option Explicit

' Open file VERME.INI to load settings
Private Sub Form_Load()
    
    Dim ConfigFile As CArquivoINI
    Dim strTemp As String
    
    ' Create instance of INI files class
    Set ConfigFile = New CArquivoINI
    ConfigFile.Nome = App.Path & "\Verme.ini"
    ConfigFile.Secao = "Matriz"
    
    ' Try to get Horizontal Size from ini file
    ConfigFile.Chave = "Dimensao_Matriz_X"
    strTemp = ConfigFile.Valor
    If strTemp = "" Then
    
        ' If not found, assign default values
        txtHorzSize.Text = "20"
        ConfigFile.Valor = "20"
    Else
    
        ' Assign text box text with value
        txtHorzSize.Text = Val(strTemp)
    End If
    
    ' Try to get Vertical Size from file
    ConfigFile.Chave = "Dimensao_Matriz_Y"
    strTemp = ConfigFile.Valor
    If strTemp = "" Then
        txtVertSize.Text = "15"
        ConfigFile.Valor = "15"
    Else
        txtVertSize.Text = Val(strTemp)
    End If
    
    ' Try to get Level from file
    ConfigFile.Chave = "Nivel"
    strTemp = ConfigFile.Valor
    If strTemp = "" Then
        cboLevel.ListIndex = 1
        ConfigFile.Valor = "2"
    Else
        If ((Val(strTemp) > 7) Or (Val(strTemp) < 1)) Then
            cboLevel.ListIndex = 1
        Else
            cboLevel.ListIndex = Val(strTemp) - 1
        End If
    End If

    ' Kill instance
    Set ConfigFile = Nothing
End Sub

' Save game settings to file VERME.INI
Private Sub Form_Unload(Cancel As Integer)
    
    Dim ConfigFile As CArquivoINI
    
    ' Create instance of INI files class
    Set ConfigFile = New CArquivoINI
    ConfigFile.Nome = App.Path & "\Verme.ini"
    ConfigFile.Secao = "Matriz"
    
    ' Save horizontal size
    ConfigFile.Chave = "Dimensao_Matriz_X"
    ConfigFile.Valor = txtHorzSize.Text
    
    ' Save vertical size
    ConfigFile.Chave = "Dimensao_Matriz_Y"
    ConfigFile.Valor = txtVertSize.Text
        
    ' Save level
    ConfigFile.Chave = "Nivel"
    ConfigFile.Valor = cboLevel.ListIndex + 1
    
    ' Kill instance
    Set ConfigFile = Nothing
End Sub

Private Sub txthorzsize_LostFocus()
    If Not IsNumeric(txtHorzSize.Text) Then
        MsgBox "Type a valid numeric value.", vbCritical, "Setting Verme"
        txtHorzSize.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtvertsize_LostFocus()
    If Not IsNumeric(txtVertSize.Text) Then
        MsgBox "Type a valid numeric value.", vbCritical, "Setting Verme"
        txtVertSize.SetFocus
        Exit Sub
    End If
End Sub
