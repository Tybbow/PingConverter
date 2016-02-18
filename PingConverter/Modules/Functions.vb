Imports System.Globalization

Module Functions

#Region "[  Fonctions Conversions  ]"

#Region "[  GetUnit  ]"

    ' /// Pour obtenir l'unité bytes ==> bytes, kb, mb, gb ///
    Public Function GetUnit(ByVal bytes As Double) As String

        ' /// Constante de conversion = 1024 ///
        Const byteConversion As Integer = 1024

        ' /// bytes plus grand que 1024^3 ///
        If bytes >= Math.Pow(byteConversion, 3) Then

            ' /// On retourne "Gbit/s" ///
            Return "Gbit/s"

            ' /// bytes plus grand que 1024^2 ///
        ElseIf bytes >= Math.Pow(byteConversion, 2) Then

            ' /// On retourne "Mbit/s" ///
            Return "Mbit/s"

            ' /// bytes plus grand que 1024^1 ///
        ElseIf bytes >= byteConversion Then

            ' /// On retourne "Kbit/s" ///
            Return "Kbit/s"

        Else

            ' /// On retourne "bit/s" ///
            Return "bit/s"

        End If

    End Function

#End Region

#Region "[  Octets en K/M/G octets  ]"

    ' /// Converti des octets en Ko, Mo ou Go ///
    Public Function ConvertByte(ByVal bytes As Double, ByVal unit As String) As Double

        ' /// Contante de conversion = 1024 ///
        Const byteConversion As Integer = 1024

        Select Case unit

            Case "Gbit/s"

                Return Math.Round(bytes / Math.Pow(byteConversion, 3), 2)

            Case "Mbit/s"

                Return Math.Round(bytes / Math.Pow(byteConversion, 2), 2)

            Case "Kbit/s"

                Return Math.Round(bytes / byteConversion, 2)

            Case Else

                Return Math.Round(bytes, 2)

        End Select

    End Function

#End Region

#Region "[  Fonctions Date  ]"

#Region "[  DateToUnix  ]"

    Public Function DateTimeToUnixTimestamp(ByVal _DateTime As DateTime) As Double

        ' /// Retourne le nombre de secondes passées depuis le 01/01/1970 ///
        Return CType((_DateTime - New DateTime(1970, 1, 1, 0, 0, 0)), TimeSpan).TotalSeconds

    End Function

#End Region

#Region "[  UnixToDate  ]"

    Public Function UnixTimestampToDateTime(ByVal _UnixTimeStamp As Long) As DateTime

        ' /// Converti le nombre de secondes passées depuis le 01/01/1970 en DateTime ///
        Return (New DateTime(1970, 1, 1, 0, 0, 0)).AddSeconds(_UnixTimeStamp)

    End Function

#End Region

#End Region

#End Region

#Region "[ Fonctions Récupération DateTo]"

    Public Function GetDateTo(ByVal dateFrom As Date, ByVal diffTime As Integer) As Date

        Return (dateFrom - TimeSpan.FromMilliseconds(diffTime))

    End Function


#End Region
End Module

