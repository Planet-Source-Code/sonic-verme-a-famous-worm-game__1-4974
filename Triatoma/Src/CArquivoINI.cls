VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CArquivoINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' /////////////////////////////////////////////////////////////
' //
' // This class is used to access INI Files. It has reading/
' // writing routines
' //
' // Feel free to mail me if you have any trouble with this
' // code. I´ll help you, if possible
' //
' //    Recife, 17 de Dezembro de 1999
' //    Rafael Winter Teixeira (sonic@torricelli.com.br)
' //
' /////////////////////////////////////////////////////////////

Option Explicit

' Declare INI files API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As Any, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

' Private variables for class' attributes
Private m_Nome As String
Private m_Chave As String
Private m_Secao As String
Private m_Default As String
Private m_CodigoRetorno As Long


' If any error happened during last function call returns 0 (zero)
Property Get CodigoRetorno() As Long
    CodigoRetorno = m_CodigoRetorno
End Property

' Default return value, if any error happens
Property Let Default(ByVal vNewValue As String)
    m_Default = vNewValue
End Property

Property Get Default() As String
    Default = m_Default
End Property

' Filename
Property Let Nome(ByVal vNewValue As String)
    m_Nome = vNewValue
End Property

Property Get Nome() As String
    Nome = m_Nome
End Property

' ;INI File section sample:
'
' [SecaoX]
' ;[SectionX]
' ChaveY = VALOR
' ;KeyY = VALUE
Property Let Secao(ByVal vNewValue As String)
    m_Secao = vNewValue
End Property

Property Get Secao() As String
    Secao = m_Secao
End Property

' INI file Key
Property Let Chave(ByVal vNewValue As String)
    m_Chave = vNewValue
End Property

Property Get Chave() As String
    Chave = m_Chave
End Property

' Key´s value within INI section
Property Get Valor() As String

    Dim sBuf As String
    Dim iSize As String
    Dim iRetCode As Integer
    Dim strERRO As String
    Dim beta As Long
    Dim alfa As Long
        
    On Error GoTo ERRO
    sBuf = Space$(255)
    iSize = Len(sBuf)
    iRetCode = GetPrivateProfileString(m_Secao, m_Chave, m_Default, sBuf, iSize, m_Nome)
    
    If (iSize > 0) Then
        Valor = Left$(sBuf, iRetCode)
    Else
        Valor = ""
    End If
    Exit Property
    
ERRO:
    Err.Raise Err.Number, Err.Source, Err.Description
End Property

Property Let Valor(ByVal vNewValue As String)

    Dim iPos As Integer
    
    On Error GoTo ERRO
    
    ' Strip chr$(0):
    iPos = InStr(vNewValue, Chr$(0))

    Do While iPos <> 0
        vNewValue = Left$(vNewValue, (iPos - 1)) & Mid$(vNewValue, (iPos + 1))
        iPos = InStr(vNewValue, Chr$(0))
    Loop

    m_CodigoRetorno = WritePrivateProfileString(m_Secao, m_Chave, vNewValue, m_Nome)
    
    Exit Property
    
ERRO:
    Err.Raise Err.Number, Err.Source, Err.Description
End Property

' Kill Key´s Value within section
Public Sub DeletarValor()
    m_CodigoRetorno = WritePrivateProfileString(m_Secao, m_Chave, 0&, m_Nome)
End Sub

' Kill Section
Public Sub DeletarSecao()
    m_CodigoRetorno = WritePrivateProfileString(m_Secao, 0&, 0&, m_Nome)
End Sub

' Return all Keys and Values from current section, separated by null char
Property Get SecaoINI() As String

    Dim sBuf As String
    Dim iSize As String
    Dim iRetCode As Integer
    sBuf = Space$(255)
    iSize = Len(sBuf)
    iRetCode = GetPrivateProfileString(m_Secao, 0&, m_Default, sBuf, iSize, m_Secao)

    If (iSize > 0) Then
        SecaoINI = Left$(sBuf, iRetCode)
    Else
        SecaoINI = ""
    End If


End Property

' Set all Keys and Values from current section, separated by null char
Property Let SecaoINI(sSection As String)
    m_CodigoRetorno = WritePrivateProfileString(m_Secao, 0&, sSection, m_Secao)
End Property
