Module ConsoleTools

#Region "[  Déclarations  ]"

    ' /// Appel aux API Win32 ///
    Private Declare Function SetConsoleTextAttribute Lib "kernel32" (ByVal hConsoleOutput As Integer, ByVal wScreenColors As Integer) As Long
    Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Integer

    ' /// Je ne sais pas ///
    Private Const STD_OUTPUT_HANDLE As Integer = -11

    ' /// Enumération des 16 couleurs disponibles ///
    Public Enum ConsoleColor

        black = 0
        darkblue = 1
        darkgreen = 2
        darkaqua = 3
        darkred = 4
        purple = 5
        brown = 6
        grey = 7
        darkgrey = 8
        blue = 9
        green = 10
        aqua = 11
        red = 12
        pink = 13
        yellow = 14
        white = 15

    End Enum

#End Region

#Region "[  Routines  ]"

#Region "[  SetConsoleColors (txt + fond)  ]"

    ' /// On change la couleur du texte et du fond ///
    Public Sub SetConsoleColors(ByVal forecolor As ConsoleColor, ByVal backcolor As ConsoleColor)
        Dim hConsole As Integer = GetStdHandle(STD_OUTPUT_HANDLE)
        backcolor = IIf(backcolor = 0, 256, backcolor * 16)
        SetConsoleTextAttribute(hConsole, forecolor Or backcolor)
    End Sub

#End Region

#Region "[  SetConsoleColors (txt)  ]"

    ' /// On change la couleur du texte ///
    Public Sub SetConsoleColors(ByVal forecolor As ConsoleColor)
        Dim hConsole As Integer = GetStdHandle(STD_OUTPUT_HANDLE)
        SetConsoleTextAttribute(hConsole, forecolor)
    End Sub

#End Region

#End Region

End Module

