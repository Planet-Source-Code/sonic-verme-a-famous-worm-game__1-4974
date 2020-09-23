VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "O Verme..."
   ClientHeight    =   5130
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6135
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   342
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   409
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picAmarelo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   195
      Picture         =   "frmMain.frx":0442
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   5
      Top             =   210
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox picPreto 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   435
      Picture         =   "frmMain.frx":04E4
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   4
      Top             =   210
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox picVermelho 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   675
      Picture         =   "frmMain.frx":0586
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   3
      Top             =   210
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox picFrontBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   4950
      Left            =   75
      ScaleHeight     =   326
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   395
      TabIndex        =   0
      Top             =   90
      Width           =   5985
   End
   Begin VB.PictureBox picBuffer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5070
      Left            =   -15
      ScaleHeight     =   334
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   395
      TabIndex        =   1
      Top             =   15
      Visible         =   0   'False
      Width           =   5985
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Left            =   15
         Top             =   0
      End
      Begin VB.Label Label1 
         Caption         =   "Back buffer - para desenho mais rápido na tela"
         Height          =   240
         Left            =   735
         TabIndex        =   2
         Top             =   45
         Width           =   3375
      End
   End
   Begin VB.Menu mnuJogo 
      Caption         =   "&Game"
      Begin VB.Menu mnuNovo 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuConfigurar 
         Caption         =   "&Customize..."
      End
      Begin VB.Menu mnuPausa 
         Caption         =   "&Pause"
      End
      Begin VB.Menu mnuRecordes 
         Caption         =   "&High scores..."
      End
      Begin VB.Menu mnuSair 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuAjuda 
      Caption         =   "&Help"
      Begin VB.Menu mnuSobre 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' /////////////////////////////////////////////////////////////
' //
' // This form is main game´s window. It has the menus, and
' // game screen
' //
' // Feel free to mail me if you have any trouble with this
' // code. I´ll help you, if possible
' //
' //    Recife, 17 de Dezembro de 1999
' //    Rafael Winter Teixeira (sonic@torricelli.com.br)
' //
' /////////////////////////////////////////////////////////////

Option Explicit

' Well, this I think you may know what it is
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

' Constants for BitBlt
Private Const SRCCOPY = &HCC0020
Private Const SRCINVERT = &H660046

' Enum for game state:
' * Nenhum (none)      - there is no game running
' * Iniciado (started) - the game was started, and it´s running
' * Pausado (paused)   - player paused game
' * Finalizado (ended) - the game is over
Public Enum EStatusJogo
    Nenhum
    Iniciado
    Pausado
    Finalizado
End Enum

' Enum for cell type
' * Nada (nothing)             - Nothing, there´s nothing on this cell
' * CorpoVerme (worm body)     - There´s a piece of worm´s body on this cell
' * ComidaNormal (normal food) - There´s a little of protein
' * ComidaGorda (fat food)     - For later use
' * ComidaDiet (diet food)     - For later use
' * ComidaUltra (ultra food)   - For later use
' * Obstaculo (obstacle)       - For later use
Private Enum ETipoCelula
    Nada
    CorpoVerme
    ComidaNormal
    ComidaGorda
    ComidaDiet
    ComidaUltra
    Obstaculo
End Enum

' Enum for worm´s head direction
' * Direita (right)
' * Baixo (down)
' * Esquerda (left)
' * Cima (up)
Private Enum EDirecao
    Direita
    Baixo
    Esquerda
    Cima
End Enum

' Point record (no explanation needed, ok?)
Private Type TPonto
    x As Integer
    y As Integer
End Type

' Worm record
' * Direcao (direction) - Holds worm´s head direction
' * Cabeca (head)       - Holds worm´s head position in the cell´s array
' * Tamanho (size)      - Number of pieces from tail plus one (head too)
' * Cauda (tail)        - Array of pieces from worm´s tail
Private Type TVerme
    Direcao As EDirecao
    Cabeca As TPonto
    Tamanho As Integer
    Cauda(0 To 98) As TPonto
End Type

' Game record
' * DIMENSAO_MATRIZ_X - Cells´ array horizontal size
' * DIMENSAO_MATRIZ_Y - Cells´ array vertical size
' * DIMENSAO_CELULA_X - Cells´ horizontal size (in pixels, depends on bitmaps used)
' * DIMENSAO_CELULA_Y - Cells´ vertical size (in pixels, depends on bitmaps used)
' * POS_X_MATRIZ      - Horizontal position of cells´ array
' * POS_Y_MATRIZ      - Vertical position of cells´ array
' * Status            - Game current state
' * Nivel             - Game level
' * Score             - Self explicative, isn´t?
' * UltimaTecla       - Last key pressed
' * Matriz            - Cell array (The world!!)
' * Verme             - Our friend, worm
' * PosComida         - Position of protein
Private Type TJogo
    DIMENSAO_MATRIZ_X As Long
    DIMENSAO_MATRIZ_Y As Long
    DIMENSAO_CELULA_X As Long
    DIMENSAO_CELULA_Y As Long
    POS_X_MATRIZ As Long
    POS_Y_MATRIZ As Long
    Status As EStatusJogo
    Nivel As Integer
    Score As Integer
    UltimaTecla As Integer
    Matriz() As ETipoCelula
    Verme As TVerme
    PosComida As TPonto
End Type

' Private module level variable for holding game data
Private Jogo As TJogo

' Private module level variable for class that opens INI files
Private ArqConfig As CArquivoINI


' DesenharMatriz (Draw Array) - Method that draws the cell array in the screen
Private Sub DesenharMatriz()

    Dim i As Integer, j As Integer
    
    ' Draw each bitmap on the back buffer
    For i = 0 To Jogo.DIMENSAO_MATRIZ_X - 1
        For j = 0 To Jogo.DIMENSAO_MATRIZ_Y - 1
            Select Case Jogo.Matriz(i, j)
            Case ETipoCelula.ComidaDiet
                
            Case ETipoCelula.ComidaGorda
            
            Case ETipoCelula.ComidaNormal
                BitBlt picBuffer.hDC, i * Jogo.DIMENSAO_CELULA_X, j * Jogo.DIMENSAO_CELULA_Y, (i * Jogo.DIMENSAO_CELULA_X) + Jogo.DIMENSAO_CELULA_X, (j * Jogo.DIMENSAO_CELULA_Y) + Jogo.DIMENSAO_CELULA_Y, picAmarelo.hDC, 0, 0, SRCCOPY
                
            Case ETipoCelula.ComidaUltra
            
            Case ETipoCelula.CorpoVerme
                BitBlt picBuffer.hDC, i * Jogo.DIMENSAO_CELULA_X, j * Jogo.DIMENSAO_CELULA_Y, (i * Jogo.DIMENSAO_CELULA_X) + Jogo.DIMENSAO_CELULA_X, (j * Jogo.DIMENSAO_CELULA_Y) + Jogo.DIMENSAO_CELULA_Y, picVermelho.hDC, 0, 0, SRCCOPY
                
            Case ETipoCelula.Nada
                BitBlt picBuffer.hDC, i * Jogo.DIMENSAO_CELULA_X, j * Jogo.DIMENSAO_CELULA_Y, (i * Jogo.DIMENSAO_CELULA_X) + Jogo.DIMENSAO_CELULA_X, (j * Jogo.DIMENSAO_CELULA_Y) + Jogo.DIMENSAO_CELULA_Y, picPreto.hDC, 0, 0, SRCCOPY
                
            Case ETipoCelula.Obstaculo
            
            End Select
        Next j
    Next i
    
    ' Draw back buffer bitmap in screen
    BitBlt picFrontBuffer.hDC, Jogo.POS_X_MATRIZ, Jogo.POS_Y_MATRIZ, Jogo.DIMENSAO_MATRIZ_X * Jogo.DIMENSAO_CELULA_X, Jogo.DIMENSAO_MATRIZ_Y * Jogo.DIMENSAO_CELULA_Y, picBuffer.hDC, 0, 0, SRCCOPY
    
    ' If the game is paused, print "PAUSA" in screen
    If Jogo.Status = Pausado Then
        picFrontBuffer.Font.Name = "Tahoma"
        picFrontBuffer.Font.Size = 12
        picFrontBuffer.Font.Italic = True
        picFrontBuffer.Font.Bold = True
        picFrontBuffer.ForeColor = vbBlack
        picFrontBuffer.CurrentX = Jogo.POS_X_MATRIZ + ((Jogo.DIMENSAO_CELULA_X * Jogo.DIMENSAO_MATRIZ_X) / 2) - (Me.TextWidth("Pausa")) ' / 2)
        picFrontBuffer.CurrentY = Jogo.POS_Y_MATRIZ + ((Jogo.DIMENSAO_CELULA_Y * Jogo.DIMENSAO_MATRIZ_Y) / 2) - (Me.TextHeight("Pausa")) ' / 2)
        picFrontBuffer.Print "Pausa"
        picFrontBuffer.ForeColor = vbYellow
        picFrontBuffer.CurrentX = Jogo.POS_X_MATRIZ + ((Jogo.DIMENSAO_CELULA_X * Jogo.DIMENSAO_MATRIZ_X) / 2) - (Me.TextWidth("Pausa")) + 1 ' / 2) + 2
        picFrontBuffer.CurrentY = Jogo.POS_Y_MATRIZ + ((Jogo.DIMENSAO_CELULA_Y * Jogo.DIMENSAO_MATRIZ_Y) / 2) - (Me.TextHeight("Pausa")) + 1 ' / 2) + 2
        picFrontBuffer.Print "Pausa"
    End If
End Sub

' Inicializar (Initialize) - Procedure executed when begining a new game
Private Sub Inicializar()

    Dim i As Integer, j As Integer
    Dim Ponto As TPonto
    Dim strTemp As String
    
    ' Clear front buffer
    picFrontBuffer.Cls
    
    ' Open INI file for loading custom configuration
    Set ArqConfig = New CArquivoINI
    ArqConfig.Nome = App.Path & "\Verme.ini"
    ArqConfig.Secao = "Matriz"
    ArqConfig.Chave = "Dimensao_Matriz_X"
    strTemp = ArqConfig.Valor
    If strTemp = "" Then
        Jogo.DIMENSAO_MATRIZ_X = 20
        ArqConfig.Valor = "20"
    Else
        Jogo.DIMENSAO_MATRIZ_X = Val(strTemp)
    End If
    
    ArqConfig.Chave = "Dimensao_Matriz_Y"
    strTemp = ArqConfig.Valor
    If strTemp = "" Then
        Jogo.DIMENSAO_MATRIZ_Y = 15
        ArqConfig.Valor = "15"
    Else
        Jogo.DIMENSAO_MATRIZ_Y = Val(strTemp)
    End If
    
    ArqConfig.Chave = "Nivel"
    strTemp = ArqConfig.Valor
    If strTemp = "" Then
        Jogo.Nivel = 2
        ArqConfig.Valor = "2"
    Else
        If ((Val(strTemp) > 7) Or (Val(strTemp) < 1)) Then
            Jogo.Nivel = 2
        Else
            Jogo.Nivel = Val(strTemp)
        End If
    End If
    
    ' Set timer´s interval, depending on level choice
    Select Case Jogo.Nivel
    Case 1
        Timer1.Interval = 150
        
    Case 2
        Timer1.Interval = 120
        
    Case 3
        Timer1.Interval = 80
        
    Case 4
        Timer1.Interval = 65
        
    Case 5
        Timer1.Interval = 40
        
    Case 6
        Timer1.Interval = 25
        
    Case 7
        Timer1.Interval = 10
        
    End Select
        
    ' As the bitmaps i´ve made are 8x8 pixels, i put bitmap
    ' width here...
    Jogo.DIMENSAO_CELULA_X = 8
    
    ' ...and height here
    Jogo.DIMENSAO_CELULA_Y = 8
    
    ' This bitmaps settings were made to give you flexibility
    ' to use any bitmap set that you make. It´s important to
    ' change this values if you change bitmaps...
    
    ' Game status is now INICIADO (started)
    Jogo.Status = Iniciado
    
    ' Last key pressed was right key (default)
    Jogo.UltimaTecla = vbKeyRight
    
    ' Get array´s position on screen depending on it´s size, and
    ' cell´s size too
    Jogo.POS_X_MATRIZ = (picFrontBuffer.Width - (Jogo.DIMENSAO_CELULA_X * Jogo.DIMENSAO_MATRIZ_X)) / 2
    Jogo.POS_Y_MATRIZ = (picFrontBuffer.Height - (Jogo.DIMENSAO_CELULA_Y * Jogo.DIMENSAO_MATRIZ_Y)) / 2
    
    With Jogo
        
        ' Resize cell array
        ReDim .Matriz(0 To .DIMENSAO_MATRIZ_X - 1, 0 To .DIMENSAO_MATRIZ_Y - 1) As ETipoCelula
        
        ' Fill array with NADA (nothing)
        For i = 0 To .DIMENSAO_MATRIZ_X - 1
            For j = 0 To .DIMENSAO_MATRIZ_Y - 1
                .Matriz(i, j) = Nada
            Next j
        Next i
        
        ' Set worm´s position
        .Verme.Direcao = Direita
        .Verme.Tamanho = 6   ' Tamanho = Cabeça + Cauda
        .Verme.Cabeca.x = 5
        .Verme.Cabeca.y = .DIMENSAO_MATRIZ_Y - 1
        
        ' Add worm´s tail (obs: Array element of index 0 (zero) is the start of tail)
        For i = 0 To .Verme.Tamanho - 2
            .Verme.Cauda(i).x = .Verme.Cabeca.x - i - 1
            .Verme.Cauda(i).y = .DIMENSAO_MATRIZ_Y - 1
        Next i
        
        ' Add worm to cell array
        .Matriz(.Verme.Cabeca.x, .Verme.Cabeca.y) = CorpoVerme
        For i = 0 To .Verme.Tamanho - 2
            .Matriz(.Verme.Cauda(i).x, .Verme.Cauda(i).y) = CorpoVerme
        Next i
        
        ' Add protein in cell array
        PontoLivre .PosComida
        .Matriz(.PosComida.x, .PosComida.y) = ComidaNormal
    End With
    
    ' Enables timer
    Timer1.Enabled = True
End Sub

' PontoLivre (Free Point) - Method that returns a randomic free point (cell holding nothing) on cell´s array
Private Sub PontoLivre(ByRef Ponto As TPonto)

    Randomize
    
    Do
        Ponto.x = Int(Jogo.DIMENSAO_MATRIZ_X * Rnd)
        Ponto.y = Int(Jogo.DIMENSAO_MATRIZ_Y * Rnd)
    Loop While Jogo.Matriz(Ponto.x, Ponto.y) <> Nada
    
End Sub

' ProximaJogada (Next turn) - Method that controls game´s turn
Private Sub ProximaJogada()
    
    Dim i As Integer
    Dim blnComeu As Boolean
    Dim PosBlocoCauda As TPonto
    Dim PontoTemp As TPonto
        
    With Jogo
    
        ' Defines new direction, depending on last key pressed
        Select Case .UltimaTecla
        Case vbKeyUp
            If Not .Verme.Direcao = Baixo Then
                .Verme.Direcao = Cima
            End If
            
        Case vbKeyDown
            If Not .Verme.Direcao = Cima Then
                .Verme.Direcao = Baixo
            End If
            
        Case vbKeyLeft
            If Not .Verme.Direcao = Direita Then
                .Verme.Direcao = Esquerda
            End If
            
        Case vbKeyRight
            If Not .Verme.Direcao = Esquerda Then
                .Verme.Direcao = Direita
            End If
            
        End Select
        
        ' Check if worm´s head will hit borders (array bounds),
        ' or if it will hit the tail, or if worm will eat
        ' some protein
        Select Case .Verme.Direcao
        Case EDirecao.Baixo
            If .Verme.Cabeca.y = .DIMENSAO_MATRIZ_Y - 1 Then
                .Status = Finalizado
                Exit Sub
            End If
            
            If .Matriz(.Verme.Cabeca.x, .Verme.Cabeca.y + 1) = CorpoVerme Then
                .Status = Finalizado
                Exit Sub
            End If
            
            If .Matriz(.Verme.Cabeca.x, .Verme.Cabeca.y + 1) = ComidaNormal Then
                blnComeu = True
            End If
            
        Case EDirecao.Cima
            If .Verme.Cabeca.y = 0 Then
                .Status = Finalizado
                Exit Sub
            End If
        
            If .Matriz(.Verme.Cabeca.x, .Verme.Cabeca.y - 1) = CorpoVerme Then
                .Status = Finalizado
                Exit Sub
            End If
        
            If .Matriz(.Verme.Cabeca.x, .Verme.Cabeca.y - 1) = ComidaNormal Then
                blnComeu = True
            End If
        
        Case EDirecao.Direita
            If .Verme.Cabeca.x = .DIMENSAO_MATRIZ_X - 1 Then
                .Status = Finalizado
                Exit Sub
            End If
        
            If .Matriz(.Verme.Cabeca.x + 1, .Verme.Cabeca.y) = CorpoVerme Then
                .Status = Finalizado
                Exit Sub
            End If
    
            If .Matriz(.Verme.Cabeca.x + 1, .Verme.Cabeca.y) = ComidaNormal Then
                blnComeu = True
            End If
    
        Case EDirecao.Esquerda
            If .Verme.Cabeca.x = 0 Then
                .Status = Finalizado
                Exit Sub
            End If
        
            If .Matriz(.Verme.Cabeca.x - 1, .Verme.Cabeca.y) = CorpoVerme Then
                .Status = Finalizado
                Exit Sub
            End If
    
            If .Matriz(.Verme.Cabeca.x - 1, .Verme.Cabeca.y) = ComidaNormal Then
                blnComeu = True
            End If
    
        End Select
        
        ' Moves head
        PosBlocoCauda = .Verme.Cabeca
        
        Select Case .Verme.Direcao
        Case EDirecao.Baixo
            .Verme.Cabeca.y = .Verme.Cabeca.y + 1
            
        Case EDirecao.Cima
            .Verme.Cabeca.y = .Verme.Cabeca.y - 1
        
        Case EDirecao.Direita
            .Verme.Cabeca.x = .Verme.Cabeca.x + 1
        
        Case EDirecao.Esquerda
            .Verme.Cabeca.x = .Verme.Cabeca.x - 1
        
        End Select
    
        ' Update array with new head´s position
        .Matriz(.Verme.Cabeca.x, .Verme.Cabeca.y) = CorpoVerme
        
        ' Move tail, swapping tails positions
        For i = 0 To .Verme.Tamanho - 2
        
            .Matriz(.Verme.Cauda(i).x, .Verme.Cauda(i).y) = Nada
            .Matriz(PosBlocoCauda.x, PosBlocoCauda.y) = CorpoVerme
            
            PontoTemp = PosBlocoCauda
            PosBlocoCauda = .Verme.Cauda(i)
            .Verme.Cauda(i) = PontoTemp
        Next i
        
        ' If worm eated, increase it´s size
        If blnComeu Then
            .Verme.Tamanho = .Verme.Tamanho + 1
            .Verme.Cauda(.Verme.Tamanho - 2) = PosBlocoCauda
            .Matriz(PosBlocoCauda.x, PosBlocoCauda.y) = CorpoVerme
            PontoLivre .PosComida
            .Matriz(.PosComida.x, .PosComida.y) = ComidaNormal
            .Score = .Score + .Nivel
        End If
        
        ' Finally, draw cell array
        DesenharMatriz
    End With
End Sub

' GameOver - Method used for terminating the game
Private Sub GameOver()

    Dim FileHandle As Integer
    Dim Recordes(0 To 9) As TRecorde
    Dim Recorde As TRecorde
    Dim blnSwap As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim strNomeJogador As String
    Dim NovoRecorde As frmNewHiscore
    
    ' Disable timer
    Timer1.Enabled = False
    
    ' No current game is running now, so set game status to
    ' NENHUM (none)
    Jogo.Status = Nenhum
    
    ' Show message box with player´s score
    MsgBox "Game Over!" & vbNewLine & vbNewLine & "Pontuação: " & Str(Jogo.Score), vbExclamation, "Fim do jogo!"
    
    ' Check if hi-score file exists
    If Dir(App.Path & "\Recordes.dat") = "" Then
    
        ' If not found, create one
        For i = 0 To 9
            Recorde.Nome = "Verme" & Str(i + 1)
            Recorde.Nivel = (i \ 3) + 1
            Recorde.Score = (i \ 2) + i * 10
            Recordes(i) = Recorde
        Next i
            
        FileHandle = FreeFile
        Open App.Path & "\Recordes.dat" For Binary As #FileHandle
        Put #FileHandle, , Recordes
        Close #FileHandle
    End If
    
    ' Open hi-score file to check if player has achieved
    ' a record
    FileHandle = FreeFile
    Open App.Path & "\Recordes.dat" For Binary As #FileHandle
    Get #FileHandle, , Recordes
    Close #FileHandle
    
    ' Make a decreasing sort
    Do
        blnSwap = False
        For i = 0 To 8
            If Recordes(i).Score < Recordes(i + 1).Score Then
                Recorde = Recordes(i + 1)
                Recordes(i + 1) = Recordes(i)
                Recordes(i) = Recorde
                blnSwap = True
            End If
        Next i
    Loop While blnSwap
    
    ' Check if player´s score is a new hi-score
    For i = 0 To 9
        If Recordes(i).Score < Jogo.Score Then
            
            ' If it is, pick his name
            Set NovoRecorde = New frmNewHiscore
            NovoRecorde.Show vbModal, Me
            strNomeJogador = NovoRecorde.WormName
            Set NovoRecorde = Nothing
            
            ' Update hi-score table
            Recorde = Recordes(i)
            Recordes(i).Nivel = Jogo.Nivel
            Recordes(i).Score = Jogo.Score
            Recordes(i).Nome = strNomeJogador
            If i < 9 Then
                For j = i + 1 To 8
                    Recordes(j) = Recorde
                    Recorde = Recordes(j + 1)
                Next j
            End If
            i = 9
            
            ' Open hi-score file and put new table
            FileHandle = FreeFile
            Open App.Path & "\Recordes.dat" For Binary As #FileHandle
            Put #FileHandle, , Recordes
            Close #FileHandle
            
            ' Show hi-score form
            frmHiscores.Show vbModal, Me
        End If
    Next i
    
    ' Reset game score
    Jogo.Score = 0
End Sub

Private Sub mnuConfigurar_Click()
    frmGameConfig.Show vbModal, Me
End Sub

Private Sub mnuNovo_Click()
    Inicializar
End Sub

Private Sub mnuPausa_Click()
    If Jogo.Status = Iniciado Then
        Jogo.Status = Pausado
    ElseIf Jogo.Status = Pausado Then
        Jogo.Status = Iniciado
    End If
    DesenharMatriz
End Sub

Private Sub mnuRecordes_Click()
    frmHiscores.Show vbModal, Me
End Sub

Private Sub mnuSair_Click()
    Unload Me
End Sub

Private Sub mnuSobre_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeySpace
    
        ' If space key was pressed, pause the game
        If Jogo.Status = Iniciado Then
            Jogo.Status = Pausado
        ElseIf Jogo.Status = Pausado Then
            Jogo.Status = Iniciado
        End If
         
        ' Redraw cell array
        DesenharMatriz
    
    Case Else
        ' If other, hold it for further processing
        If Jogo.Status = Iniciado Then
            Jogo.UltimaTecla = KeyCode
        End If
    End Select
End Sub

Private Sub Form_Paint()

    ' Redraw array
    DesenharMatriz
End Sub

Private Sub Timer1_Timer()
    ' Check game state
    If Jogo.Status = Iniciado Then
    
        ' If it´s running, do next turn and redraw cell array
        ProximaJogada
        picFrontBuffer.Refresh
    End If
    
    If Jogo.Status = Finalizado Then
        ' If it´s ended, run GameOver method
        GameOver
    End If
End Sub
