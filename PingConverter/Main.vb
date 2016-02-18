Imports System.Windows.Forms.DataVisualization.Charting
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Drawing

Module Main

#Region "[  Déclarations  ]"

    ' /// Fichier en cours de lecture ///
    Dim myFile As String = String.Empty
    Dim fileName As String = String.Empty

    ' /// Créer le tableau de valeurs ///
    Dim nbPing As ArrayList

    Dim title As String = String.Empty
    Dim stats As String = String.Empty
    Dim rtt As String = String.Empty

    ' /// Date To ///
    Dim dateTo As Date

#End Region

#Region "[  Main  ]"

    ' /// Lancement de l'application ///
    Sub Main()

        ' /// On vérifie si des arguments sont passés avec l'exécutable ///
        If My.Application.CommandLineArgs.Count > 0 Then

            ' /// Pour chaque argument passé ///
            For Each arg As String In My.Application.CommandLineArgs

                ' /// Si l'argument n'est pas un fichier .TXT, on affiche une erreur ///
                If Not Path.GetExtension(arg).ToLower = ".txt" Then ShowError("Fichier avec extension TXT uniquement !") : Exit Sub

                Try

                    ' /// Création d'un répertoire ///
                    If Not Directory.Exists(arg.Substring(0, arg.LastIndexOf("\")) + "\PingConverter\") Then Directory.CreateDirectory(arg.Substring(0, arg.LastIndexOf("\")) + "\PingConverter\")

                    ' /// Chargement du fichier TXT
                    Using sr As New StreamReader(arg)

                        ' /// On lit le fichier jusqu'à la fin et on met tout ça dans myFile ///

                        myFile = sr.ReadToEnd


                    End Using

                    myFile = myFile.Replace(vbLf, vbCrLf)

                    ' /// On  compte le nombre de ligne dans myFile
                    Dim line() As String = Split(myFile, vbCrLf)
                    Dim lineCount As Integer = line.Count

                    ' /// On lit les lignes souhaitées.
                    title = line(lineCount - 4)
                    Console.WriteLine(title)
                    stats = line(lineCount - 3)
                    Console.WriteLine(stats)
                    rtt = line(lineCount - 2)
                    Console.WriteLine(rtt)

                    ' /// On récupère les différentes dates
                    fileName = New FileInfo(arg).Name
                    dateTo = New FileInfo(arg).CreationTime

                    ' /// On parse le fichier; si erreur on sort ///
                    If Not ParseFile(myFile) Then Exit Sub

                    ' /// On crée le chart et on le sauvegarde en image; si erreur on sort ///
                    If Not CreateChart(arg) Then Exit Sub

                Catch ex As Exception

                    ' /// Erreur ///
                    Console.WriteLine(ex.Message)
                    Console.ReadKey()

                End Try

            Next

            ' /// Affichage final ///
            Console.WriteLine("Conversion réussie !")
            Console.ReadKey()

        Else

            ' /// On affiche une erreur ///
            ShowError("Aucun argument passé !" & ControlChars.CrLf & "Veuillez faire un drag'n'drop de fichiers TXT...")

        End If

    End Sub

#End Region

#Region "[  Routines  ]"

#Region "[  Show Error  ]"

    ' /// Routine d'affichage d'une erreur (pas d'argument) ///
    Private Sub ShowError(ByVal [text] As String)

        ' /// Couleur rouge du texte de la console ///
        SetConsoleColors(ConsoleColor.red)

        ' /// Affichage texte d'erreur ///
        Console.WriteLine([text] & ControlChars.CrLf)

        ' /// Affichage pub en darkblue ///
        SetConsoleColors(ConsoleColor.darkblue)
        Console.WriteLine(" PingConverter By Tybbow v1.0")

        ' /// Affichage du drapeau tricolore ///
        SetConsoleColors(ConsoleColor.blue)
        Console.Write("##########")
        SetConsoleColors(ConsoleColor.white)
        Console.Write("##########")
        SetConsoleColors(ConsoleColor.red)
        Console.WriteLine("##########")
        SetConsoleColors(ConsoleColor.blue)
        Console.Write("##########")
        SetConsoleColors(ConsoleColor.white)
        Console.Write("##########")
        SetConsoleColors(ConsoleColor.red)
        Console.WriteLine("##########")
        SetConsoleColors(ConsoleColor.blue)
        Console.Write("##########")
        SetConsoleColors(ConsoleColor.white)
        Console.Write("##########")
        SetConsoleColors(ConsoleColor.red)
        Console.WriteLine("##########")
        SetConsoleColors(ConsoleColor.blue)
        Console.Write("##########")
        SetConsoleColors(ConsoleColor.white)
        Console.Write("##########")
        SetConsoleColors(ConsoleColor.red)
        Console.WriteLine("##########")

        ' /// Couleur verte du texte de la console ///
        SetConsoleColors(ConsoleColor.green)

        ' /// Erreur sur console ///
        Console.WriteLine("")
        Console.WriteLine("Appuyez sur une touche pour quitter le programme...")
        Console.ReadKey()

    End Sub

#End Region

#Region "[  ParseFile  ]"

    ' /// Parseur de fichier ///
    Private Function ParseFile(ByVal dataFile As String) As Boolean

        nbPing = New ArrayList

        Try

            ' /// Chaîne de caractères à trouver - Mesures ///
            Dim reg As Regex = New Regex("(?<ping>\d+\.?\d+ ms)")

            ' /// Intégration des pings dans mcMesures
            Dim mcMesures As MatchCollection = reg.Matches(dataFile)

            ' /// On scrute les résultats ///
            For i As Integer = 0 To mcMesures.Count - 2

                ' /// On ajoute les vitesses des mesures dans les tableaux correspondants ///
                nbPing.Add(Replace(mcMesures.Item(i).Groups("ping").ToString, " ms", ""))

            Next

            ' /// Couleur verte du texte de la console ///
            SetConsoleColors(ConsoleColor.green)

            ' /// Affichage ///
            Console.WriteLine(String.Format("{0}, Analyse du fichier terminé.", fileName.Substring(0, fileName.Length - 4)))

            ' /// OK ///
            Return True

        Catch ex As Exception

            ' /// Affichage de l'erreur et on quitte le programme ///
            ShowError("Erreur lors de l'analyse du fichier" & ControlChars.CrLf & ex.Message)

            ' /// NOK ///
            Return False

        End Try

    End Function

#End Region

#Region "[  CreateChart  ]"

    Private Function CreateChart(ByVal filePath As String) As Boolean

        Try

            ' /// Création d'un chart ///
            Dim myChart As New Chart

            ' /// Création d'une zone de chart et ajout dans myChart ///
            Dim myChartArea As New ChartArea("PingConverter")
            myChart.ChartAreas.Add(myChartArea)

            ' /// Série (BP Entrée) (qui contiendra les DataPoint) ///
            Dim series1 As New Series()
            series1.Name = "Ping"

            ' /// On scrute la collection de dates ///
            For i As Integer = 0 To nbPing.Count - 1

                ' /// On ajoute la date et la BP Entrée à la série 1 ///
                series1.Points.AddXY(CDate(dateTo + TimeSpan.FromSeconds(i)).ToString("dd/MM/yy HH:mm"), nbPing.Item(i))


            Next

            ' /// On ajoute la serie au chart ///
            myChart.Series.Add(series1)

            ' /// Références de l'abscisse : 12 valeurs voulues ///
            myChart.ChartAreas("PingConverter").AxisX.Interval = Math.Round(series1.Points.Count / 6)

            ' /// Type de graphe ///
            series1.ChartType = SeriesChartType.StackedArea

            ' /// Couleur de BP Entrée, Color.Argb ( composante alpha, couleur ) ///
            series1.Color = Color.FromArgb(230, Color.RoyalBlue)

            ' /// On définit nous-même la légende du chart ///
            Dim myLegend As New Legend("MyLegend")
            myLegend.LegendStyle = LegendStyle.Row
            myLegend.TextWrapThreshold = 10
            myLegend.Docking = Docking.Bottom
            myLegend.Alignment = StringAlignment.Center
            myLegend.Font = New Font("Tahoma", 8, FontStyle.Regular)
            myLegend.BorderColor = Color.Gray
            myLegend.BorderDashStyle = ChartDashStyle.Solid
            myLegend.ShadowOffset = 3

            ' /// On ajoute la légende au chart ///
            myChart.Legends.Add(myLegend)

            ' /// Dimensionner le Chart ///
            myChart.Width = 1200
            myChart.Height = 600

            ' /// Création d'un titre qu'on ajoute à la collection Titles du Chart ///
            myChart.Titles.Add("MESURE")
            myChart.Titles(0).Name = "MESURE"
            myChart.Titles("MESURE").Text = String.Format(title)
            myChart.Titles("MESURE").Font = New Font("Tahoma", 11, FontStyle.Bold)
            myChart.Titles("MESURE").ForeColor = Color.WhiteSmoke
            myChart.Titles("MESURE").BackColor = Color.FromArgb(50, 50, 50)
            myChart.Titles("MESURE").Alignment = System.Drawing.ContentAlignment.MiddleCenter

            ' /// Ajout d'un autre titre (PUB) ///
            myChart.Titles.Add("PUB")
            myChart.Titles(1).Name = "PUB"
            myChart.Titles("PUB").Text = "PingConverter" + ControlChars.Lf + "CNMO-R - Section Métrologie"
            myChart.Titles("PUB").Font = New Font("Tahoma", 9, FontStyle.Bold)
            myChart.Titles("PUB").ForeColor = Color.LightGray
            myChart.Titles("PUB").Position.Auto = False
            myChart.Titles("PUB").Position.X = 90
            myChart.Titles("PUB").Position.Y = 96


            myChart.Titles.Add("STATISTIQUE")
            myChart.Titles(2).Name = "STATISTIQUE"
            myChart.Titles("STATISTIQUE").Text = String.Format(stats + ControlChars.Lf + rtt)
            myChart.Titles("STATISTIQUE").Font = New Font("Tahoma", 10, FontStyle.Bold)
            myChart.Titles("STATISTIQUE").ForeColor = Color.DarkGray
            myChart.Titles("STATISTIQUE").Alignment = System.Drawing.ContentAlignment.MiddleCenter


            ' /// Titre de l'axe des ordonnées + Font ///
            myChart.ChartAreas("PingConverter").AxisY.Title = String.Format("Latence (ms)")
            myChart.ChartAreas("PingConverter").AxisY.TitleFont = New Font("Tahoma", 10, FontStyle.Regular)

            ' /// Pas de marges sur l'axe des abscisses ///
            myChart.ChartAreas("PingConverter").AxisX.IsMarginVisible = False

            ' /// Si les dates sont trop longues, elles seront affichées avec un angle de 30° ///
            myChart.ChartAreas("PingConverter").AxisX.LabelStyle.Angle = 30

            ' /// On sauvegarde le chart en image ///
            myChart.SaveImage(filePath.Substring(0, filePath.LastIndexOf("\")) + String.Format("\PingConverter\{0}.png", fileName.Substring(0, fileName.Length - 4)), ChartImageFormat.Png)

            ' ///  Affichage ///
            Console.WriteLine("Conversion en image terminée." & ControlChars.CrLf)

            ' /// OK ///
            Return True

        Catch ex As Exception

            ' /// Affichage de l'erreur et on quitte le programme ///
            ShowError("Erreur lors de la création du graphe." & ControlChars.CrLf & ex.Message)

            ' /// NOK ///
            Return False

        End Try

    End Function

#End Region

#End Region

End Module
