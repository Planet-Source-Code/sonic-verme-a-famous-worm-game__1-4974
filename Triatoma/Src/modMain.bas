Attribute VB_Name = "modMain"
' /////////////////////////////////////////////////////////////
' //
' // This module is used only to declare the hi-score type
' //
' // Feel free to mail me if you have any trouble with this
' // code. I´ll help you, if possible
' //
' //    Recife, 17 de Dezembro de 1999
' //    Rafael Winter Teixeira (sonic@torricelli.com.br)
' //
' /////////////////////////////////////////////////////////////

Option Explicit

' Declare type TRecorde, used for hi-score storage
Public Type TRecorde
    Nome As String * 30
    Nivel As Byte
    Score As Long
End Type

