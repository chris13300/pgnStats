Imports System.Management 'ajouter référence system.management
Imports VB = Microsoft.VisualBasic

Module modMain
    Private chainePGN As String
    Private tabPGN() As String
    Private joueurs As String
    Private compteurTaches As Integer
    Private progression As Integer
    Private nbTaches As Integer
    Private tabResultat() As String
    Private tabBGW() As System.ComponentModel.BackgroundWorker

    Sub Main()
        Dim i As Integer, chaine As String, j As Integer, fichierPGN As String, nbCoups As Integer
        Dim tabJoueurs() As String, tabVictoires() As String, tabCompteurs() As Integer, prof As Integer
        Dim tabTmp() As String, compteur As Integer, memo As String, total As Integer, k As Integer, tabChaine() As String
        Dim maxProfondeurs() As Integer, moyProfondeurs() As Integer, nbProfondeurs() As Integer, resultat As String, supplement As Integer
        Dim moyTemps() As Single, nbTemps() As Integer, duree As Single, debut As Integer, fin As Integer
        Dim ordre1 As String, ordre2 As String, tabOrdre1() As String

        Console.Title = My.Computer.Name & " : loading workers..."
        nbTaches = cpu() - 1

        ReDim tabBGW(nbTaches)
        For i = 1 To nbTaches
            tabBGW(i) = New System.ComponentModel.BackgroundWorker
            tabBGW(i).WorkerReportsProgress = True
            tabBGW(i).WorkerSupportsCancellation = True
            AddHandler tabBGW(i).DoWork, AddressOf tabBGW_DoWork
            AddHandler tabBGW(i).RunWorkerCompleted, AddressOf tabBGW_RunWorkerCompleted
        Next

        ordre1 = ""
        ordre2 = ""

        'on demande le fichier PGN
        fichierPGN = Replace(Command(), """", "")
        If Not My.Computer.FileSystem.FileExists(fichierPGN) Then
            End
        End If

        'on charge le fichier PGN
        Console.Title = My.Computer.Name & " : creation of " & nomFichier(Replace(fichierPGN, ".pgn", "_bak.pgn"))
        If My.Computer.FileSystem.FileExists(Replace(fichierPGN, ".pgn", "_bak.pgn")) Then
            My.Computer.FileSystem.DeleteFile(Replace(fichierPGN, ".pgn", "_bak.pgn"))
        End If
        My.Computer.FileSystem.CopyFile(fichierPGN, Replace(fichierPGN, ".pgn", "_bak.pgn"))

        Console.Title = My.Computer.Name & " : reading of " & nomFichier(fichierPGN)

        Dim lecture As IO.StreamReader, nbLignes As Integer
        lecture = New IO.StreamReader(Replace(fichierPGN, ".pgn", "_bak.pgn"))
        nbLignes = 0
        nbCoups = 0
        ReDim tabPGN(1000000) '1 000 000 lignes
        Do
            chaine = lecture.ReadLine
            nbCoups = nbCoups + nbCaracteres(chaine, "{")
            If nbLignes > UBound(tabPGN) Then
                ReDim Preserve tabPGN(nbLignes * 2)
            End If
            tabPGN(nbLignes) = chaine
            nbLignes = nbLignes + 1
        Loop Until lecture.EndOfStream
        lecture.Close()
        ReDim Preserve tabPGN(nbLignes - 1)

        'on liste les joueurs blancs et noirs
        joueurs = ""
        chainePGN = ""
        compteurTaches = 0
        progression = 0
        debut = 0
        fin = -1
        ReDim tabResultat(nbTaches)

        'diviser le travail
        If nbTaches < UBound(tabPGN) Then
            For i = 1 To nbTaches - 1
                debut = fin + 1
                'division théorique
                j = i * UBound(tabPGN) / nbTaches
                Do
                    j = j + 1
                    Threading.Thread.Sleep(10)
                    'division précise
                    If j < UBound(tabPGN) Then
                        If tabPGN(j) = "" And InStr(tabPGN(j + 1), "[", CompareMethod.Text) > 0 And InStr(tabPGN(j + 1), "]", CompareMethod.Text) > 0 Then
                            Exit Do
                        End If
                    End If
                Loop While j < UBound(tabPGN)
                fin = j
                tabBGW(i).RunWorkerAsync(i & ":" & debut & ":" & fin)
            Next
            debut = fin + 1
            fin = UBound(tabPGN)
            tabBGW(i).RunWorkerAsync(i & ":" & debut & ":" & fin)
        Else
            nbTaches = 1
            ReDim Preserve tabBGW(nbTaches)

            i = 1
            debut = 0
            fin = UBound(tabPGN)
            tabBGW(i).RunWorkerAsync(i & ":" & debut & ":" & fin)
        End If

        Do
            Console.Clear()
            Console.Title = My.Computer.Name & " : search of players @ " & Format((progression + 1) / tabPGN.Length, "0.00%")
            Console.WriteLine("There is " & Split(joueurs, "|").Length - 1 & " players including :")
            Console.WriteLine(Replace(joueurs, "|", ", "))
            Threading.Thread.Sleep(1000)
        Loop Until compteurTaches = nbTaches

        chainePGN = ""
        For i = 1 To nbTaches
            chainePGN = chainePGN & tabResultat(i)
        Next

        Console.Title = My.Computer.Name & " : loading of games..."
        tabPGN = Split(chainePGN, vbCrLf)
        tabJoueurs = Split(joueurs, "|")
        ReDim Preserve tabJoueurs(UBound(tabJoueurs) - 1)
        ReDim tabVictoires(UBound(tabJoueurs))
        ReDim tabCompteurs(UBound(tabJoueurs))
        ReDim maxProfondeurs(UBound(tabJoueurs))
        ReDim moyProfondeurs(UBound(tabJoueurs))
        ReDim nbProfondeurs(UBound(tabJoueurs))
        ReDim moyTemps(UBound(tabJoueurs))
        ReDim nbTemps(UBound(tabJoueurs))

        Array.Sort(tabJoueurs)

        'pour chaque joueur blanc, on comptabilise les victoires
        Console.WriteLine()
        Console.WriteLine("Victories :")
        For i = 0 To UBound(tabJoueurs)
            maxProfondeurs(i) = 0
            moyProfondeurs(i) = 0
            nbProfondeurs(i) = 0
            moyTemps(i) = 0
            nbTemps(i) = 0

            For j = 0 To UBound(tabPGN)
                If tabPGN(j) <> "" Then
                    supplement = -1
                    If tabPGN(j) = "[White """ & tabJoueurs(i) & """]" Then
                        supplement = 1
                    ElseIf tabPGN(j) = "[Black """ & tabJoueurs(i) & """]" Then
                        supplement = 0
                    End If

                    If supplement <> -1 Then
                        'on comptabilise les parties du joueur avec les blancs
                        tabCompteurs(i) = tabCompteurs(i) + 1

                        resultat = ""
                        Select Case tabPGN(j + 1 + supplement)
                            Case "[Result ""1-0""]"
                                resultat = "1-0"
                            Case "[Result ""1/2-1/2""]"
                                resultat = "1/2-1/2"
                            Case "[Result ""0-1""]"
                                resultat = "0-1"
                        End Select

                        j = j + 3 + supplement

                        'on retient la liste des coups
                        chaine = ""
                        While InStr(tabPGN(j), resultat) = 0 And j < UBound(tabPGN)
                            If tabPGN(j) <> "" Then
                                chaine = chaine & tabPGN(j) & " "
                            End If
                            j = j + 1
                        End While

                        chaine = chaine & tabPGN(j)
                        j = j + 1
                        If (supplement = 1 And resultat = "1-0") _
                        Or (supplement = 0 And resultat = "0-1") Then
                            tabVictoires(i) = tabVictoires(i) & chaine & "|"
                        End If

                        'on cherche la profondeur maximale atteinte par le joueur blanc uniquement
                        tabChaine = Split(Replace(Replace(chaine, " {", "_{"), "{ ", "{"), " ")
                        k = 0
                        Do
                            If droite(tabChaine(k), 1) = "." Then
                                If supplement = 0 Then
                                    'on évite coups du livre
                                    If InStr(tabChaine(k), "...", CompareMethod.Text) = 0 Then
                                        'on evite le coup blanc
                                        If tabChaine(k + 1).IndexOf("/") > 0 Then
                                            k = k + 2
                                        Else
                                            k = k + 1
                                        End If
                                        If k + 1 = tabChaine.Length Then
                                            Exit Do
                                        End If
                                    End If
                                ElseIf supplement = 1 Then
                                    If InStr(tabChaine(k), "...", CompareMethod.Text) > 0 Then
                                        k = k + 2
                                    End If
                                End If

                                If k >= UBound(tabChaine) Then
                                    Exit Do
                                End If

                                If tabChaine(k + 1).IndexOf("/") > 0 And tabChaine(k + 1) <> "1/2-1/2" Then
                                    tabChaine(k + 1) = Replace(tabChaine(k + 1), "}", "")

                                    prof = CInt(tabChaine(k + 1).Substring(tabChaine(k + 1).IndexOf("/") + 1))

                                    If prof > maxProfondeurs(i) Then
                                        maxProfondeurs(i) = prof
                                    End If

                                    moyProfondeurs(i) = moyProfondeurs(i) + prof
                                    nbProfondeurs(i) = nbProfondeurs(i) + 1

                                    k = k + 2
                                End If

                            End If

                            Do
                                k = k + 1
                                If k >= UBound(tabChaine) Then
                                    Exit Do
                                End If
                            Loop Until droite(tabChaine(k), 1) = "."
                        Loop Until k >= UBound(tabChaine)

                        'on cherche le temps maximal mis par le joueur blanc uniquement
                        tabChaine = Split(Replace(chaine, " {", "_{"), " ")
                        k = 0
                        Do
                            If droite(tabChaine(k), 1) = "." Then
                                If supplement = 0 Then
                                    'on évite coups du livre
                                    If InStr(tabChaine(k), "...", CompareMethod.Text) = 0 Then
                                        'on evite le coup blanc
                                        If tabChaine(k + 1).IndexOf("/") > 0 Then
                                            k = k + 2
                                        Else
                                            k = k + 1
                                        End If
                                        If k + 1 = tabChaine.Length Then
                                            Exit Do
                                        End If
                                    End If
                                ElseIf supplement = 1 Then
                                    If InStr(tabChaine(k), "...", CompareMethod.Text) > 0 Then
                                        k = k + 2
                                    End If
                                End If

                                If k + 2 <= UBound(tabChaine) Then
                                    If InStr(tabChaine(k + 2), "{", CompareMethod.Text) = 0 And tabChaine(k + 2).IndexOf("}") > 0 And InStr(tabChaine(k + 2), "/", CompareMethod.Text) = 0 _
                                    And InStr(tabChaine(k + 2), "mates", CompareMethod.Text) = 0 And InStr(tabChaine(k + 2), "Pat", CompareMethod.Text) = 0 And InStr(tabChaine(k + 2), "Arena", CompareMethod.Text) = 0 _
                                    And InStr(tabChaine(k + 2), "insuffisant", CompareMethod.Text) = 0 And InStr(tabChaine(k + 2), "disconnects", CompareMethod.Text) = 0 And InStr(tabChaine(k + 2), "book", CompareMethod.Text) = 0 Then
                                        moyTemps(i) = moyTemps(i) + CDbl(Replace(Replace(Replace(tabChaine(k + 2), "s", ""), ".", ","), "}", ""))
                                        nbTemps(i) = nbTemps(i) + 1
                                        k = k + 2
                                    End If
                                End If
                            End If

                            Do
                                k = k + 1
                                If k >= UBound(tabChaine) Then
                                    Exit Do
                                End If
                            Loop Until droite(tabChaine(k), 1) = "."
                        Loop Until k >= UBound(tabChaine)
                    End If
                End If
            Next
            Console.Title = My.Computer.Name & " : analysis of games @ " & Format((i + 1) / tabJoueurs.Length, "0.00%")
            If InStr(tabVictoires(i), "|") = 0 Then
                Console.WriteLine("0 for " & tabJoueurs(i))
            Else
                Console.WriteLine(tabVictoires(i).Split("|").Length - 1 & " for " & tabJoueurs(i))
            End If
        Next

        Console.WriteLine()

        'mise en forme
        ordre1 = ""
        chaine = ""
        duree = 0
        For i = 0 To UBound(tabJoueurs)
            If nbProfondeurs(i) > 0 Then
                moyProfondeurs(i) = moyProfondeurs(i) / nbProfondeurs(i)
            End If
            If nbTemps(i) > 0 Then
                duree = duree + moyTemps(i)
                moyTemps(i) = moyTemps(i) / nbTemps(i)
            End If
            If Not tabVictoires(i) Is Nothing Then
                tabTmp = tabVictoires(i).Split("|")
                ReDim Preserve tabTmp(UBound(tabTmp) - 1)
                chaine = chaine & tabJoueurs(i) & " with " & tabTmp.Length & "/" & tabCompteurs(i) & " wins (" & Format(tabTmp.Length / tabCompteurs(i), "0.00%") & ", D" & moyProfondeurs(i) & "/" & maxProfondeurs(i) & ", " & Trim(Format(moyTemps(i) * 1000, "# ### ##0")) & " ms) :" & vbCrLf & vbCrLf

                'trier les parties
                tabTmp = trierTableau(tabTmp)

                For j = 0 To UBound(tabTmp)
                    If InStr(tabTmp(j), "11.", CompareMethod.Text) > 0 Then
                        chaine = chaine & gauche(tabTmp(j), InStr(tabTmp(j), "11.", CompareMethod.Text) - 2) & vbCrLf & vbCrLf
                    Else
                        chaine = chaine & tabTmp(j) & vbCrLf & vbCrLf
                    End If
                Next
                chaine = chaine & vbCrLf
            Else
                'aucune victoire
                chaine = chaine & tabJoueurs(i) & " with " & "0/" & tabCompteurs(i) & " wins " & "(D" & moyProfondeurs(i) & "/" & maxProfondeurs(i) & ", " & Trim(Format(moyTemps(i) * 1000, "# ### ##0")) & " ms) :" & vbCrLf & vbCrLf
            End If
            Console.Title = My.Computer.Name & " : analysis of games @ " & Format(i / UBound(tabJoueurs), "0.00%")
        Next
        ordre1 = chaine

        'on cherche les pourcentages
        ordre2 = ""
        tabOrdre1 = Split(ordre1, vbCrLf)
        memo = ""
        For i = 0 To UBound(tabOrdre1)
            If InStr(tabOrdre1(i), " with ", CompareMethod.Text) = 0 And tabOrdre1(i) <> "" Then
                'pour chaque coup, on regarde si la partie suivante commence par le même
                tabTmp = Split(tabOrdre1(i), " ")
                chaine = ""
                compteur = 1
                For j = 0 To UBound(tabTmp)
                    If InStr(tabTmp(j), "{", CompareMethod.Text) = 0 And InStr(tabTmp(j), "/", CompareMethod.Text) = 0 And (droite(chaine, 3) = "1. " Or chaine = "") Then
                        chaine = chaine & tabTmp(j) & " "
                    End If

                    If InStr(tabTmp(j), ".", CompareMethod.Text) = 0 And InStr(memo, chaine) = 0 Then
                        'on regarde les lignes suivantes
                        For k = i + 1 To tabOrdre1.Length - 1
                            If tabOrdre1(k) <> "" Or (compteur > 0 And k = tabOrdre1.Length - 1) Then
                                If InStr(tabOrdre1(k), chaine) = 0 Then
                                    ordre2 = ordre2 & chaine & "@ " & Format(compteur / total, "0.00%") & vbCrLf
                                    memo = memo & chaine & "|"
                                    chaine = ""
                                    i = i + (2 * compteur) - 1
                                    compteur = 1
                                    Exit For
                                ElseIf InStr(tabOrdre1(k), chaine) = 1 Then
                                    compteur = compteur + 1
                                End If
                            End If
                        Next

                        'on essaie de voir si avec le coup suivant on trouve plusieurs parties
                        If compteur = 1 Then
                            'pas la peine
                            Exit For
                        End If
                    End If
                Next
            Else
                If tabOrdre1(i) <> "" Then
                    total = 1
                    If InStr(tabOrdre1(i), " with ", CompareMethod.Text) > 0 And ordre2 = "" Then
                        ordre2 = ordre2 & tabOrdre1(i) & vbCrLf & vbCrLf
                    ElseIf InStr(tabOrdre1(i), " with ", CompareMethod.Text) > 0 And ordre2 <> "" Then
                        ordre2 = ordre2 & vbCrLf & vbCrLf & tabOrdre1(i) & vbCrLf & vbCrLf
                    End If
                    total = gauche(droite(tabOrdre1(i), Len(tabOrdre1(i)) - InStr(tabOrdre1(i), "with", CompareMethod.Text) - 4), InStr(droite(tabOrdre1(i), Len(tabOrdre1(i)) - InStr(tabOrdre1(i), "with", CompareMethod.Text) - 4), "/") - 1)
                    Console.Title = My.Computer.Name & " : statistics @ " & Format(i / (tabOrdre1.Length - 1), "0.00%")
                End If
                memo = ""
            End If
        Next
        Console.Title = My.Computer.Name
        My.Computer.FileSystem.DeleteFile(Replace(fichierPGN, ".pgn", "_bak.pgn"))

        'nombre de parties
        total = 0
        For i = 0 To UBound(tabCompteurs)
            total = total + tabCompteurs(i)
        Next
        total = total / 2

        Console.WriteLine(ordre2)
        Console.WriteLine()

        Console.WriteLine("Averages :")
        Console.WriteLine(Trim(Format(duree / total, "# ##0")) & " sec/game")
        Console.WriteLine(Format(nbCoups / total, "0") & " plies/game")

        Console.WriteLine()
        Console.WriteLine("Press a key to exit.")
        Console.ReadKey()
    End Sub

    Public Function cpu(Optional reel As Boolean = False) As Integer
        Dim collection As New ManagementObjectSearcher("select * from Win32_Processor"), taches As Integer
        taches = 0

        For Each element As ManagementObject In collection.Get
            If reel Then
                taches = taches + element.Properties("NumberOfCores").Value 'cores
            Else
                taches = taches + element.Properties("NumberOfLogicalProcessors").Value 'threads
            End If
        Next

        Return taches
    End Function

    Public Function droite(texte As String, longueur As Integer) As String
        If longueur > 0 Then
            Return VB.Right(texte, longueur)
        Else
            Return ""
        End If
    End Function

    Public Function gauche(texte As String, longueur As Integer) As String
        If longueur > 0 Then
            Return VB.Left(texte, longueur)
        Else
            Return ""
        End If
    End Function

    Public Function nbCaracteres(ByVal chaine As String, ByVal critere As String) As Integer
        Return Len(chaine) - Len(Replace(chaine, critere, ""))
    End Function

    Public Function nomFichier(chemin As String) As String
        Return My.Computer.FileSystem.GetName(chemin)
    End Function

    Private Sub tabBGW_DoWork(ByVal sender As System.ComponentModel.BackgroundWorker, ByVal e As System.ComponentModel.DoWorkEventArgs)
        Dim i As Integer, tabTampon() As String, debut As Integer, fin As Integer, tab() As String
        Dim chaineTampon As String, chaine As String, index As Integer


        tab = Split(e.Argument, ":")
        index = tab(0)
        debut = tab(1)
        fin = tab(2)

        ReDim tabTampon(fin - debut)
        Array.Copy(tabPGN, debut, tabTampon, 0, fin - debut + 1)

        chaineTampon = ""
        chaine = ""
        For i = 0 To UBound(tabTampon)
            If InStr(tabTampon(i), "[White """, CompareMethod.Text) = 1 Then
                chaineTampon = chaineTampon & tabTampon(i) & vbCrLf

                chaine = Replace(Replace(tabTampon(i), "[White """, ""), """]", "")
                If InStr(joueurs, chaine & "|") = 0 Then
                    joueurs = joueurs & chaine & "|"
                End If
            ElseIf InStr(tabTampon(i), "[Black """, CompareMethod.Text) = 1 Then
                chaineTampon = chaineTampon & tabTampon(i) & vbCrLf

                chaine = Replace(Replace(tabTampon(i), "[Black """, ""), """]", "")
                If InStr(joueurs, chaine & "|") = 0 Then
                    joueurs = joueurs & chaine & "|"
                End If
            ElseIf (InStr(tabTampon(i), "[", CompareMethod.Text) = 0 And InStr(tabTampon(i), "]", CompareMethod.Text) = 0) _
                Or InStr(tabTampon(i), "[Result """, CompareMethod.Text) > 0 Then
                chaineTampon = chaineTampon & tabTampon(i) & vbCrLf
            End If
            progression = progression + 1
        Next

        tabResultat(index) = chaineTampon
        sender.ReportProgress(100)

    End Sub

    Private Sub tabBGW_RunWorkerCompleted(ByVal sender As System.ComponentModel.BackgroundWorker, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs)
        compteurTaches = compteurTaches + 1
    End Sub

    Public Function trierTableau(tableau() As String, Optional ordre As Boolean = True) As String()
        Dim tabChaine() As String

        ReDim tabChaine(UBound(tableau))
        Array.Copy(tableau, tabChaine, tableau.Length)

        Array.Sort(tabChaine)
        If Not ordre Then
            Array.Reverse(tabChaine)
        End If

        Return tabChaine
    End Function


End Module
