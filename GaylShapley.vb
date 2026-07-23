Option Explicit

'================================================================================================
' MACRO 1 : GÉNÉRATION DES DONNÉES DE TEST
' Feuilles générées :
'   - Préférences_Élèves  : préférences individuelles de chaque élève
'   - Équipes             : composition de chaque équipe
'   - Préférences_Équipes : préférences agrégées de chaque équipe (vote majoritaire)
'   - Préférences_Projets : classement des équipes par projet + capacités
'================================================================================================
Sub GenererDonneesDeTest()
    ' === DÉCLARATIONS COMPLÈTES AU DÉBUT ===
    Dim nbEleves As Long, nbProjets As Long, tailleMin As Long, tailleMax As Long
    Dim i As Long, j As Long, k As Long, r As Long, tmpL As Long, projLookup As Long
    Dim wsE As Worksheet, wsEq As Worksheet, wsPE As Worksheet, wsP As Worksheet
    Dim prefsInd() As String, projs() As String, tmpS As String
    Dim elevesShuf() As Long
    Dim eqDebut() As Long, eqTaille() As Long
    Dim nbEquipes As Long, cursor As Long, restants As Long, taille As Long, maxT As Long
    Dim eq As Long, nbM As Long, debut As Long, eleveNum As Long
    Dim wsTemp As Worksheet, sheetName As Variant
    Dim prefsMembres() As String, projUsed() As Boolean
    Dim pos As Long, pIdx As Long, maxV As Long, winner As Long, votesP() As Long
    Dim minEq As Long, maxEq As Long
    Dim projetNom As String, scoreTotal As Double
    Dim scores() As Double, eNo As Long
    Dim rang As Long, rangs() As Long

    ' === DÉBUT DU CODE EXÉCUTABLE ===
    On Error Resume Next
    nbEleves = CLng(InputBox("Nombre d'élèves ?", "Génération des données"))
    If Err.Number <> 0 Or nbEleves <= 1 Then MsgBox "Annulé.", vbInformation: Exit Sub
    Err.Clear
    nbProjets = CLng(InputBox("Nombre de projets ?", "Génération des données"))
    If Err.Number <> 0 Or nbProjets <= 0 Then MsgBox "Annulé.", vbInformation: Exit Sub
    Err.Clear
    tailleMin = CLng(InputBox("Taille minimale des équipes ?", "Génération des données", "2"))
    If Err.Number <> 0 Or tailleMin <= 0 Then MsgBox "Annulé.", vbInformation: Exit Sub
    Err.Clear
    tailleMax = CLng(InputBox("Taille maximale des équipes ?", "Génération des données", "4"))
    If Err.Number <> 0 Or tailleMax < tailleMin Then MsgBox "Annulé.", vbInformation: Exit Sub
    On Error GoTo 0

    If tailleMin > nbEleves Then
        MsgBox "Erreur : la taille minimale d'équipe dépasse le nombre d'élèves.", vbCritical
        Exit Sub
    End If

    ' Créer les feuilles nécessaires si elles n'existent pas
    For Each sheetName In Array("Préférences_Élèves", "Équipes", "Préférences_Équipes", "Préférences_Projets")
        On Error Resume Next
        Set wsTemp = Nothing
        Set wsTemp = ThisWorkbook.Sheets(CStr(sheetName))
        On Error GoTo 0
        If wsTemp Is Nothing Then
            Set wsTemp = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            wsTemp.Name = CStr(sheetName)
        End If
    Next sheetName

    Application.ScreenUpdating = False
    Randomize

    ' ======================================================================
    ' ÉTAPE 1 : Préférences individuelles des élèves (Fisher-Yates)
    ' ======================================================================
    Set wsE = ThisWorkbook.Sheets("Préférences_Élèves")
    wsE.Cells.Clear
    wsE.Cells(1, 1).Value = "Élève"
    For i = 1 To nbProjets
        wsE.Cells(1, 1 + i).Value = "Choix " & i
    Next i
    wsE.Rows(1).Font.Bold = True

    ' Mémoire des préférences individuelles : prefsInd(élève, position) = nom du projet
    ReDim prefsInd(1 To nbEleves, 1 To nbProjets)

    ReDim projs(1 To nbProjets)
    For i = 1 To nbProjets: projs(i) = "Projet " & i: Next i

    For i = 1 To nbEleves
        wsE.Cells(i + 1, 1).Value = "Élève " & i
        For j = nbProjets To 2 Step -1
            r = Int(j * Rnd) + 1
            tmpS = projs(j): projs(j) = projs(r): projs(r) = tmpS
        Next j
        For j = 1 To nbProjets
            wsE.Cells(i + 1, 1 + j).Value = projs(j)
            prefsInd(i, j) = projs(j)
        Next j
    Next i
    wsE.Columns.AutoFit

    ' ======================================================================
    ' ÉTAPE 2 : Formation aléatoire des équipes
    ' ======================================================================
    ' Mélange des indices d'élèves
    ReDim elevesShuf(1 To nbEleves)
    For i = 1 To nbEleves: elevesShuf(i) = i: Next i
    For i = nbEleves To 2 Step -1
        r = Int(i * Rnd) + 1
        tmpL = elevesShuf(i): elevesShuf(i) = elevesShuf(r): elevesShuf(r) = tmpL
    Next i

    ' Construction des équipes par découpe séquentielle
    ReDim eqDebut(1 To nbEleves): ReDim eqTaille(1 To nbEleves)
    nbEquipes = 0
    cursor = 1

    Do While cursor <= nbEleves
        nbEquipes = nbEquipes + 1
        eqDebut(nbEquipes) = cursor
        restants = nbEleves - cursor + 1
        If restants <= tailleMax Then
            taille = restants ' Dernière équipe : prend tous les élèves restants
        Else
            ' S'assurer que les élèves restants après cette équipe formeront une équipe valide
            maxT = restants - tailleMin
            If maxT > tailleMax Then maxT = tailleMax
            If maxT < tailleMin Then
                taille = restants
            Else
                taille = tailleMin + Int((maxT - tailleMin + 1) * Rnd)
            End If
        End If
        eqTaille(nbEquipes) = taille
        cursor = cursor + taille
    Loop

    ' Écriture feuille Équipes
    Set wsEq = ThisWorkbook.Sheets("Équipes")
    wsEq.Cells.Clear
    wsEq.Cells(1, 1).Value = "Équipe"
    For i = 1 To tailleMax: wsEq.Cells(1, 1 + i).Value = "Membre " & i: Next i
    wsEq.Rows(1).Font.Bold = True
    For i = 1 To nbEquipes
        wsEq.Cells(i + 1, 1).Value = "Équipe " & i
        For j = 1 To eqTaille(i)
            wsEq.Cells(i + 1, 1 + j).Value = "Élève " & elevesShuf(eqDebut(i) + j - 1)
        Next j
    Next i
    wsEq.Columns.AutoFit

    ' ======================================================================
    ' ÉTAPE 3 : Calcul des préférences d'équipes par vote majoritaire
    '
    ' À chaque position, on identifie le premier choix restant de chaque membre
    ' et le projet avec le plus de votes obtient cette position dans la liste
    ' de l'équipe. En cas d'égalité, le projet d'indice le plus bas gagne.
    ' ======================================================================
    Set wsPE = ThisWorkbook.Sheets("Préférences_Équipes")
    ReDim projUsed(1 To nbProjets)

    wsPE.Cells.Clear
    wsPE.Cells(1, 1).Value = "Équipe"
    For i = 1 To nbProjets: wsPE.Cells(1, 1 + i).Value = "Choix " & i: Next i
    wsPE.Rows(1).Font.Bold = True

    For eq = 1 To nbEquipes
        wsPE.Cells(eq + 1, 1).Value = "Équipe " & eq

        nbM = eqTaille(eq)
        debut = eqDebut(eq)

        ' Préférences de chaque membre : prefsMembres(m, j) = nom du j-ième choix du membre m
        ReDim prefsMembres(1 To nbM, 1 To nbProjets)
        For i = 1 To nbM
            eleveNum = elevesShuf(debut + i - 1)
            For j = 1 To nbProjets
                prefsMembres(i, j) = prefsInd(eleveNum, j)
            Next j
        Next i

        ' Vote majoritaire position par position
        ReDim projUsed(1 To nbProjets)

        For pos = 1 To nbProjets
            ReDim votesP(1 To nbProjets)

            For i = 1 To nbM
                ' Premier choix restant de ce membre
                For j = 1 To nbProjets
                    ' Chercher l'index du projet dans projs() par son nom
                    projLookup = 0
                    For k = 1 To nbProjets
                        If projs(k) = prefsMembres(i, j) Then
                            projLookup = k
                            Exit For
                        End If
                    Next k
                    If projLookup > 0 And Not projUsed(projLookup) Then
                        votesP(projLookup) = votesP(projLookup) + 1
                        Exit For
                    End If
                Next j
            Next i

            ' Projet gagnant : le plus de votes (ex-aequo : indice le plus bas)
            maxV = -1
            winner = 0
            For j = 1 To nbProjets
                If Not projUsed(j) And votesP(j) > maxV Then
                    maxV = votesP(j): winner = j
                End If
            Next j

            wsPE.Cells(eq + 1, 1 + pos).Value = "Projet " & winner
            projUsed(winner) = True
        Next pos
    Next eq
    wsPE.Columns.AutoFit

    ' ======================================================================
    ' ÉTAPE 4 : Feuille Préférences_Projets
    '
    ' Colonnes : Projet | MinEquipes | MaxEquipes | TailleMinEquipe | TailleMaxEquipe | Eq1 | Eq2 ...
    '
    ' Classement de chaque équipe par un projet = moyenne du rang que les membres
    ' accordent à ce projet dans leur liste individuelle (score bas = équipe enthousiaste).
    ' ======================================================================
    Set wsP = ThisWorkbook.Sheets("Préférences_Projets")

    wsP.Cells.Clear
    wsP.Cells(1, 1).Value = "Projet"
    wsP.Cells(1, 2).Value = "MinEquipes"
    wsP.Cells(1, 3).Value = "MaxEquipes"
    wsP.Cells(1, 4).Value = "TailleMinEquipe"
    wsP.Cells(1, 5).Value = "TailleMaxEquipe"
    For eq = 1 To nbEquipes: wsP.Cells(1, 5 + eq).Value = "Équipe " & eq: Next eq
    wsP.Rows(1).Font.Bold = True
    
    For i = 1 To nbProjets
        wsP.Cells(i + 1, 1).Value = "Projet " & i
        minEq = Application.WorksheetFunction.RandBetween(1, 2)
        maxEq = Application.WorksheetFunction.RandBetween(minEq, minEq + 2)
        wsP.Cells(i + 1, 2).Value = minEq
        wsP.Cells(i + 1, 3).Value = maxEq
        wsP.Cells(i + 1, 4).Value = tailleMin    ' Le projet accepte des équipes de taille comprise
        wsP.Cells(i + 1, 5).Value = tailleMax    ' dans la fourchette globale de génération

        ' Calcul des scores : score(eq) = moyenne du rang du projet chez les membres de eq
        projetNom = "Projet " & i
        ReDim scores(1 To nbEquipes)
        For eq = 1 To nbEquipes
            scoreTotal = 0
            For j = 1 To eqTaille(eq)
                eNo = elevesShuf(eqDebut(eq) + j - 1)
                For k = 1 To nbProjets
                    If prefsInd(eNo, k) = projetNom Then
                        scoreTotal = scoreTotal + k: Exit For
                    End If
                Next k
            Next j
            scores(eq) = scoreTotal / eqTaille(eq)
        Next eq

        ' Attribution des rangs (rang 1 = score le plus bas = équipe la plus enthousiaste)
        ' Ex-aequo : le même rang est attribué (le rang suivant est donc sauté)
        ReDim rangs(1 To nbEquipes)
        For eq = 1 To nbEquipes
            rang = 1
            For k = 1 To nbEquipes
                If scores(k) < scores(eq) Then rang = rang + 1
            Next k
            rangs(eq) = rang
        Next eq
        For eq = 1 To nbEquipes: wsP.Cells(i + 1, 5 + eq).Value = rangs(eq): Next eq
    Next i
    wsP.Columns.AutoFit

    Application.ScreenUpdating = True
    MsgBox nbEleves & " élèves répartis en " & nbEquipes & " équipes, " & nbProjets & " projets générés.", vbInformation

End Sub


'================================================================================================
' MACRO 2 : AFFECTATION DES ÉQUIPES AUX PROJETS (Gale-Shapley)
'
' Feuilles lues    : Équipes, Préférences_Équipes, Préférences_Projets
' Feuille écrite   : Résultats
'================================================================================================
Sub AffectationEquipesProjets()
    ' === DÉCLARATIONS COMPLÈTES AU DÉBUT ===
    Dim wsEq As Worksheet, wsPE As Worksheet, wsP As Worksheet, wsR As Worksheet
    Dim i As Long, j As Long
    Dim nomEq As String, taille As Long, nomP As String
    Dim equipesTaille As Object
    Dim projetsMinEq As Object, projetsMaxEq As Object, projetsTailleMin As Object, projetsTailleMax As Object
    Dim projetsRangs As Object, affectationsProjet As Object
    Dim lastCol As Long, rangsDict As Object, eqH As String
    Dim equipePrefs As Object, celibataires As Collection
    Dim propositionsFaites As Object, affectationEquipe As Object
    Dim tEq As Long, pNom As String, prefsFiltrees As Collection
    Dim equipeActuelle As String, indexProp As Long, projetVise As String
    Dim affectesAuProjet As Object, maxEqP As Long, pireEquipe As String
    Dim rangPire As Long, eAff As Variant, rangNouvelle As Long
    Dim ligneR As Long, projet As Variant, listeEq As String, nbAff As Long

    ' === DÉBUT DU CODE EXÉCUTABLE ===
    On Error Resume Next
    Set wsEq = ThisWorkbook.Sheets("Équipes")
    Set wsPE = ThisWorkbook.Sheets("Préférences_Équipes")
    Set wsP  = ThisWorkbook.Sheets("Préférences_Projets")
    Set wsR  = ThisWorkbook.Sheets("Résultats")
    If wsEq Is Nothing Or wsPE Is Nothing Or wsP Is Nothing Or wsR Is Nothing Then
        MsgBox "Erreur : une ou plusieurs feuilles manquantes.", vbCritical: Exit Sub
    End If
    On Error GoTo 0

    Application.ScreenUpdating = False

    ' ---- Lecture des tailles d'équipes ----
    Set equipesTaille = CreateObject("Scripting.Dictionary")
    For i = 2 To wsEq.Cells(wsEq.Rows.Count, "A").End(xlUp).Row
        nomEq = Trim(wsEq.Cells(i, 1).Value)
        If nomEq <> "" Then
            taille = 0
            For j = 2 To wsEq.Cells(i, wsEq.Columns.Count).End(xlToLeft).Column
                If Trim(wsEq.Cells(i, j).Value) <> "" Then taille = taille + 1
            Next j
            equipesTaille(nomEq) = taille
        End If
    Next i

    ' ---- Lecture des données projets ----
    Set projetsMinEq = CreateObject("Scripting.Dictionary")
    Set projetsMaxEq = CreateObject("Scripting.Dictionary")
    Set projetsTailleMin = CreateObject("Scripting.Dictionary")
    Set projetsTailleMax = CreateObject("Scripting.Dictionary")
    Set projetsRangs = CreateObject("Scripting.Dictionary")
    Set affectationsProjet = CreateObject("Scripting.Dictionary")

    lastCol = wsP.Cells(1, wsP.Columns.Count).End(xlToLeft).Column
    For i = 2 To wsP.Cells(wsP.Rows.Count, "A").End(xlUp).Row
        nomP = Trim(wsP.Cells(i, 1).Value)
        If nomP <> "" Then
            projetsMinEq(nomP)     = CLng(wsP.Cells(i, 2).Value)
            projetsMaxEq(nomP)     = CLng(wsP.Cells(i, 3).Value)
            projetsTailleMin(nomP)  = CLng(wsP.Cells(i, 4).Value)
            projetsTailleMax(nomP)  = CLng(wsP.Cells(i, 5).Value)
            Set rangsDict = CreateObject("Scripting.Dictionary")
            For j = 6 To lastCol
                eqH = Trim(wsP.Cells(1, j).Value)
                If eqH <> "" Then rangsDict(eqH) = CLng(wsP.Cells(i, j).Value)
            Next j
            Set projetsRangs(nomP) = rangsDict
            Set affectationsProjet(nomP) = CreateObject("Scripting.Dictionary")
        End If
    Next i

    ' ---- Lecture des préférences + pré-filtrage des incompatibilités de taille ----
    ' Les projets incompatibles avec la taille de l'équipe sont retirés dès la lecture.
    Set equipePrefs = CreateObject("Scripting.Dictionary")
    Set celibataires = New Collection
    Set propositionsFaites = CreateObject("Scripting.Dictionary")
    Set affectationEquipe = CreateObject("Scripting.Dictionary")

    For i = 2 To wsPE.Cells(wsPE.Rows.Count, "A").End(xlUp).Row
        nomEq = Trim(wsPE.Cells(i, 1).Value)
        If nomEq <> "" Then
            tEq = equipesTaille(nomEq)
            Set prefsFiltrees = New Collection
            For j = 2 To wsPE.Cells(i, wsPE.Columns.Count).End(xlToLeft).Column
                pNom = Trim(wsPE.Cells(i, j).Value)
                If projetsTailleMin.Exists(pNom) Then
                    If tEq >= projetsTailleMin(pNom) And tEq <= projetsTailleMax(pNom) Then
                        prefsFiltrees.Add pNom
                    End If
                End If
            Next j
            Set equipePrefs(nomEq) = prefsFiltrees
            celibataires.Add nomEq
            affectationEquipe(nomEq) = ""
            propositionsFaites(nomEq) = 0
        End If
    Next i

    ' ---- Algorithme de Gale-Shapley sur les équipes ----
    While celibataires.Count > 0
        equipeActuelle = celibataires(1)
        indexProp = propositionsFaites(equipeActuelle) + 1

        If indexProp > equipePrefs(equipeActuelle).Count Then
            ' L'équipe a épuisé tous ses vœux compatibles : reste non affectée
            celibataires.Remove 1
        Else
            projetVise = equipePrefs(equipeActuelle)(indexProp)
            propositionsFaites(equipeActuelle) = indexProp

            Set affectesAuProjet = affectationsProjet(projetVise)
            maxEqP = projetsMaxEq(projetVise)

            If affectesAuProjet.Count < maxEqP Then
                ' Place disponible : acceptation provisoire
                affectesAuProjet(equipeActuelle) = projetsRangs(projetVise)(equipeActuelle)
                affectationEquipe(equipeActuelle) = projetVise
                celibataires.Remove 1
            Else
                ' Projet plein : chercher la pire équipe actuellement acceptée
                pireEquipe = ""
                rangPire = -1
                For Each eAff In affectesAuProjet.Keys
                    If affectesAuProjet(eAff) > rangPire Then
                        rangPire = affectesAuProjet(eAff): pireEquipe = eAff
                    End If
                Next eAff

                rangNouvelle = projetsRangs(projetVise)(equipeActuelle)

                If rangNouvelle < rangPire Then
                    ' La nouvelle équipe est mieux classée : éviction
                    affectesAuProjet.Remove pireEquipe
                    affectationEquipe(pireEquipe) = ""
                    celibataires.Add pireEquipe
                    affectesAuProjet(equipeActuelle) = rangNouvelle
                    affectationEquipe(equipeActuelle) = projetVise
                    celibataires.Remove 1
                Else
                    ' Rejet : l'équipe réessaiera avec son prochain choix
                    celibataires.Remove 1
                    celibataires.Add equipeActuelle
                End If
            End If
        End If
    Wend

    ' ---- Écriture des résultats ----
    wsR.Cells.ClearContents
    wsR.Range("A1:C1").Value = Array("Projet", "Équipes Affectées", "Statut Capacité")
    wsR.Range("A1:C1").Font.Bold = True
    ligneR = 2
    For Each projet In affectationsProjet.Keys
        wsR.Cells(ligneR, 1).Value = projet
        listeEq = IIf(affectationsProjet(projet).Count > 0, Join(affectationsProjet(projet).Keys, ", "), "Aucune")
        wsR.Cells(ligneR, 2).Value = listeEq
        wsR.Cells(ligneR, 2).WrapText = True
        nbAff = affectationsProjet(projet).Count
        If nbAff < projetsMinEq(projet) Then
            wsR.Cells(ligneR, 3).Value = "MINIMUM NON ATTEINT (" & nbAff & "/" & projetsMinEq(projet) & " équipes)"
            wsR.Cells(ligneR, 3).Interior.Color = vbYellow
        Else
            wsR.Cells(ligneR, 3).Value = "OK (" & nbAff & "/" & projetsMaxEq(projet) & " équipes)"
            wsR.Cells(ligneR, 3).Interior.ColorIndex = xlNone
        End If
        ligneR = ligneR + 1
    Next projet
    wsR.Columns.AutoFit

    Application.ScreenUpdating = True
    MsgBox "Affectation des équipes aux projets terminée. Consultez la feuille 'Résultats'.", vbInformation

End Sub


'================================================================================================
' MACRO 3 : AFFECTATION PAS À PAS (avec journal dans Log_Affectation)
'================================================================================================
Sub AffectationEquipesPasAPas()
    ' === DÉCLARATIONS COMPLÈTES AU DÉBUT ===
    Dim wsEq As Worksheet, wsPE As Worksheet, wsP As Worksheet, wsR As Worksheet, wsLog As Worksheet
    Dim ligneLog As Long
    Dim i As Long, j As Long
    Dim nomEq As String, taille As Long, nomP As String
    Dim tEq As Long, pNom As String
    Dim equipesTaille As Object
    Dim projetsMinEq As Object, projetsMaxEq As Object, projetsTailleMin As Object, projetsTailleMax As Object
    Dim projetsRangs As Object, affectationsProjet As Object
    Dim lastCol As Long, rangsDict As Object, eqH As String
    Dim equipePrefs As Object, celibataires As Collection
    Dim propositionsFaites As Object, affectationEquipe As Object
    Dim prefsFiltrees As Collection
    Dim equipeActuelle As String, indexProp As Long, projetVise As String
    Dim action As String, decision As String, statut As String
    Dim affectesAuProjet As Object, maxEqP As Long, pireEquipe As String
    Dim rangPire As Long, eAff As Variant, rangNouvelle As Long
    Dim ligneR As Long, projet As Variant, listeEq As String, nbAff As Long

    ' === DÉBUT DU CODE EXÉCUTABLE ===
    On Error Resume Next
    Set wsEq  = ThisWorkbook.Sheets("Équipes")
    Set wsPE  = ThisWorkbook.Sheets("Préférences_Équipes")
    Set wsP   = ThisWorkbook.Sheets("Préférences_Projets")
    Set wsR   = ThisWorkbook.Sheets("Résultats")
    Set wsLog = ThisWorkbook.Sheets("Log_Affectation")
    If wsEq Is Nothing Or wsPE Is Nothing Or wsP Is Nothing Or wsR Is Nothing Or wsLog Is Nothing Then
        MsgBox "Erreur : feuilles manquantes.", vbCritical: Exit Sub
    End If
    On Error GoTo 0

    Application.ScreenUpdating = False

    wsLog.Cells.Clear
    wsLog.Range("A1:E1").Value = Array("Étape", "Action de l'Équipe", "Décision du Projet", "Statut du Projet", "Équipes libres")
    wsLog.Range("A1:E1").Font.Bold = True
    ligneLog = 1

    ' ---- Lecture des tailles d'équipes ----
    Set equipesTaille = CreateObject("Scripting.Dictionary")
    For i = 2 To wsEq.Cells(wsEq.Rows.Count, "A").End(xlUp).Row
        nomEq = Trim(wsEq.Cells(i, 1).Value)
        If nomEq <> "" Then
            taille = 0
            For j = 2 To wsEq.Cells(i, wsEq.Columns.Count).End(xlToLeft).Column
                If Trim(wsEq.Cells(i, j).Value) <> "" Then taille = taille + 1
            Next j
            equipesTaille(nomEq) = taille
        End If
    Next i

    ' ---- Lecture des données projets ----
    Set projetsMinEq = CreateObject("Scripting.Dictionary")
    Set projetsMaxEq = CreateObject("Scripting.Dictionary")
    Set projetsTailleMin = CreateObject("Scripting.Dictionary")
    Set projetsTailleMax = CreateObject("Scripting.Dictionary")
    Set projetsRangs = CreateObject("Scripting.Dictionary")
    Set affectationsProjet = CreateObject("Scripting.Dictionary")

    lastCol = wsP.Cells(1, wsP.Columns.Count).End(xlToLeft).Column
    For i = 2 To wsP.Cells(wsP.Rows.Count, "A").End(xlUp).Row
        nomP = Trim(wsP.Cells(i, 1).Value)
        If nomP <> "" Then
            projetsMinEq(nomP)     = CLng(wsP.Cells(i, 2).Value)
            projetsMaxEq(nomP)     = CLng(wsP.Cells(i, 3).Value)
            projetsTailleMin(nomP)  = CLng(wsP.Cells(i, 4).Value)
            projetsTailleMax(nomP)  = CLng(wsP.Cells(i, 5).Value)
            Set rangsDict = CreateObject("Scripting.Dictionary")
            For j = 6 To lastCol
                eqH = Trim(wsP.Cells(1, j).Value)
                If eqH <> "" Then rangsDict(eqH) = CLng(wsP.Cells(i, j).Value)
            Next j
            Set projetsRangs(nomP) = rangsDict
            Set affectationsProjet(nomP) = CreateObject("Scripting.Dictionary")
        End If
    Next i

    ' ---- Lecture des préférences + pré-filtrage des incompatibilités de taille ----
    Set equipePrefs = CreateObject("Scripting.Dictionary")
    Set celibataires = New Collection
    Set propositionsFaites = CreateObject("Scripting.Dictionary")
    Set affectationEquipe = CreateObject("Scripting.Dictionary")

    For i = 2 To wsPE.Cells(wsPE.Rows.Count, "A").End(xlUp).Row
        nomEq = Trim(wsPE.Cells(i, 1).Value)
        If nomEq <> "" Then
            tEq = equipesTaille(nomEq)
            Set prefsFiltrees = New Collection
            For j = 2 To wsPE.Cells(i, wsPE.Columns.Count).End(xlToLeft).Column
                pNom = Trim(wsPE.Cells(i, j).Value)
                If projetsTailleMin.Exists(pNom) Then
                    If tEq >= projetsTailleMin(pNom) And tEq <= projetsTailleMax(pNom) Then
                        prefsFiltrees.Add pNom
                    End If
                End If
            Next j
            Set equipePrefs(nomEq) = prefsFiltrees
            celibataires.Add nomEq
            affectationEquipe(nomEq) = ""
            propositionsFaites(nomEq) = 0
        End If
    Next i

    ' ---- Algorithme de Gale-Shapley avec journalisation ----
    While celibataires.Count > 0
        equipeActuelle = celibataires(1)
        indexProp = propositionsFaites(equipeActuelle) + 1

        If indexProp > equipePrefs(equipeActuelle).Count Then
            celibataires.Remove 1
            LogStep wsLog, ligneLog, equipeActuelle & " a épuisé sa liste de vœux.", "Reste non affectée.", "", GetCelibatairesString(celibataires)
        Else
            projetVise = equipePrefs(equipeActuelle)(indexProp)
            propositionsFaites(equipeActuelle) = indexProp

            action   = equipeActuelle & " propose au " & projetVise & " (choix n°" & indexProp & ")."
            decision = ""
            statut   = ""

            Set affectesAuProjet = affectationsProjet(projetVise)
            maxEqP = projetsMaxEq(projetVise)

            If affectesAuProjet.Count < maxEqP Then
                decision = "Le projet a de la place. ACCEPTATION PROVISOIRE."
                affectesAuProjet(equipeActuelle) = projetsRangs(projetVise)(equipeActuelle)
                affectationEquipe(equipeActuelle) = projetVise
                celibataires.Remove 1
            Else
                pireEquipe = ""
                rangPire = -1
                For Each eAff In affectesAuProjet.Keys
                    If affectesAuProjet(eAff) > rangPire Then
                        rangPire = affectesAuProjet(eAff): pireEquipe = eAff
                    End If
                Next eAff

                rangNouvelle = projetsRangs(projetVise)(equipeActuelle)

                If rangNouvelle < rangPire Then
                    decision = "Plein. " & equipeActuelle & " est mieux classée que " & pireEquipe & ". ACCEPTATION et ÉVICTION."
                    affectesAuProjet.Remove pireEquipe
                    affectationEquipe(pireEquipe) = ""
                    celibataires.Add pireEquipe
                    affectesAuProjet(equipeActuelle) = rangNouvelle
                    affectationEquipe(equipeActuelle) = projetVise
                    celibataires.Remove 1
                Else
                    decision = "Plein. " & equipeActuelle & " n'est pas mieux classée. REJET."
                    celibataires.Remove 1
                    celibataires.Add equipeActuelle
                End If
            End If

            statut = projetVise & ": [" & Join(affectationsProjet(projetVise).Keys, ", ") & "]"
            LogStep wsLog, ligneLog, action, decision, statut, GetCelibatairesString(celibataires)
        End If
    Wend

    ' ---- Écriture des résultats ----
    wsR.Cells.ClearContents
    wsR.Range("A1:C1").Value = Array("Projet", "Équipes Affectées", "Statut Capacité")
    wsR.Range("A1:C1").Font.Bold = True
    ligneR = 2
    For Each projet In affectationsProjet.Keys
        wsR.Cells(ligneR, 1).Value = projet
        listeEq = IIf(affectationsProjet(projet).Count > 0, Join(affectationsProjet(projet).Keys, ", "), "Aucune")
        wsR.Cells(ligneR, 2).Value = listeEq
        wsR.Cells(ligneR, 2).WrapText = True
        nbAff = affectationsProjet(projet).Count
        If nbAff < projetsMinEq(projet) Then
            wsR.Cells(ligneR, 3).Value = "MINIMUM NON ATTEINT (" & nbAff & "/" & projetsMinEq(projet) & " équipes)"
            wsR.Cells(ligneR, 3).Interior.Color = vbYellow
        Else
            wsR.Cells(ligneR, 3).Value = "OK (" & nbAff & "/" & projetsMaxEq(projet) & " équipes)"
            wsR.Cells(ligneR, 3).Interior.ColorIndex = xlNone
        End If
        ligneR = ligneR + 1
    Next projet
    wsR.Columns.AutoFit
    wsLog.Columns.AutoFit

    Application.ScreenUpdating = True
    MsgBox "Affectation pas à pas terminée.", vbInformation

End Sub


'================================================================================================
' MACRO 4 : RAPPORT DE SATISFACTION DES ÉQUIPES
' Pour chaque équipe : rang du projet obtenu dans sa liste de vœux
' Statistiques globales : moyenne et écart-type des rangs (écart-type d'échantillon)
'================================================================================================
Sub CreerRapportSatisfaction()

    Dim wsR As Worksheet, wsPE As Worksheet, wsRapport As Worksheet
    Dim affectations As Object, prefsEquipes As Object
    Dim i As Long, j As Long
    Dim projetAff As String, listeStr As String, arr As Variant, eq As Variant
    Dim nomEq As String
    Dim rangsObtenus As Collection, ligneRap As Long
    Dim equipe As Variant
    Dim projetObtenu As String, rang As Long, prefs As Collection
    Dim somme As Double, ri As Variant, moyenne As Double, sc As Double, ecartType As Double
    Set rangsObtenus = New Collection

    On Error Resume Next
    Set wsR       = ThisWorkbook.Sheets("Résultats")
    Set wsPE      = ThisWorkbook.Sheets("Préférences_Équipes")
    Set wsRapport = ThisWorkbook.Sheets("Rapport_Satisfaction")
    If wsR Is Nothing Or wsPE Is Nothing Or wsRapport Is Nothing Then
        MsgBox "Erreur : feuilles manquantes.", vbCritical: Exit Sub
    End If
    On Error GoTo 0

    Application.ScreenUpdating = False

    ' Lecture affectations équipe -> projet depuis Résultats
    Set affectations = CreateObject("Scripting.Dictionary")
    For i = 2 To wsR.Cells(wsR.Rows.Count, "A").End(xlUp).Row
        projetAff = Trim(wsR.Cells(i, 1).Value)
        listeStr  = Trim(wsR.Cells(i, 2).Value)
        If listeStr <> "Aucune" And listeStr <> "" Then
            arr = Split(listeStr, ", ")
            For Each eq In arr: affectations(Trim(eq)) = projetAff: Next eq
        End If
    Next i

    ' Lecture des préférences des équipes
    Set prefsEquipes = CreateObject("Scripting.Dictionary")
    For i = 2 To wsPE.Cells(wsPE.Rows.Count, "A").End(xlUp).Row
        nomEq = Trim(wsPE.Cells(i, 1).Value)
        If nomEq <> "" Then
            Set prefsEquipes(nomEq) = New Collection
            For j = 2 To wsPE.Cells(i, wsPE.Columns.Count).End(xlToLeft).Column
                prefsEquipes(nomEq).Add Trim(wsPE.Cells(i, j).Value)
            Next j
        End If
    Next i

    ' Génération du rapport
    wsRapport.Cells.ClearContents
    wsRapport.Range("A1:C1").Value = Array("Équipe", "Projet Affecté", "Rang du Choix")
    wsRapport.Range("A1:C1").Font.Bold = True

    ligneRap = 1
    For Each equipe In prefsEquipes.Keys
        ligneRap = ligneRap + 1
        wsRapport.Cells(ligneRap, 1).Value = equipe
        If affectations.Exists(equipe) Then
            projetObtenu = affectations(equipe)
            wsRapport.Cells(ligneRap, 2).Value = projetObtenu
            rang = 0
            Set prefs = prefsEquipes(equipe)
            For j = 1 To prefs.Count
                If StrComp(prefs(j), projetObtenu, vbTextCompare) = 0 Then rang = j: Exit For
            Next j
            If rang > 0 Then
                wsRapport.Cells(ligneRap, 3).Value = rang
                rangsObtenus.Add rang
            Else
                wsRapport.Cells(ligneRap, 3).Value = "Erreur : projet introuvable dans les vœux"
            End If
        Else
            wsRapport.Cells(ligneRap, 2).Value = "Non affectée"
            wsRapport.Cells(ligneRap, 3).Value = "N/A"
        End If
    Next equipe

    ' Statistiques (écart-type d'échantillon)
    If rangsObtenus.Count > 0 Then
        somme = 0
        For Each ri In rangsObtenus: somme = somme + ri: Next ri
        moyenne = somme / rangsObtenus.Count

        sc = 0
        For Each ri In rangsObtenus: sc = sc + (ri - moyenne) ^ 2: Next ri
        If rangsObtenus.Count > 1 Then ecartType = Sqr(sc / (rangsObtenus.Count - 1)) Else ecartType = 0

        ligneRap = ligneRap + 2
        wsRapport.Cells(ligneRap, "B").Value = "Satisfaction moyenne :"
        wsRapport.Cells(ligneRap, "C").Value = Round(moyenne, 2)
        ligneRap = ligneRap + 1
        wsRapport.Cells(ligneRap, "B").Value = "Écart-type des rangs :"
        wsRapport.Cells(ligneRap, "C").Value = Round(ecartType, 2)
        wsRapport.Range(wsRapport.Cells(ligneRap - 1, "B"), wsRapport.Cells(ligneRap, "C")).Font.Bold = True
    End If

    wsRapport.Columns.AutoFit
    Application.ScreenUpdating = True
    MsgBox "Rapport de satisfaction des équipes généré.", vbInformation

End Sub


'================================================================================================
' MACRO 5 : BILAN DE PERFORMANCE
'================================================================================================
Sub BilanPerformanceAlgorithme()

    Dim wsR As Worksheet, wsPE As Worksheet, wsP As Worksheet
    Dim wsBilan As Worksheet, wsDetails As Worksheet
    Dim i As Long, j As Long, eq As Variant, projet As Variant
    Dim nomP As String, listeStr As String, arr As Variant, nomEq As String
    Dim affectationsEquipe As Object, affectationsProjet As Object
    Dim prefsEquipes As Object, projetsCapacites As Object
    Dim nbEquipesTotal As Long, nbEquipesAffectees As Long
    Dim nbChoix1 As Long, nbChoix2 As Long, nbChoix3 As Long
    Dim rangsObtenus As Collection, equipesSansProjet As Collection
    Dim rang As Long, prefs As Collection, nbAff As Long
    Dim somme As Double, sc As Double, ri As Variant
    Dim projetsSousMinimum As Collection, projetsVides As Collection
    Dim nbProjetsMinAtteint As Long, rangMoyen As Double, ecartType As Double
    Dim ligne As Long, ld As Long

    On Error Resume Next
    Set wsR       = ThisWorkbook.Sheets("Résultats")
    Set wsPE      = ThisWorkbook.Sheets("Préférences_Équipes")
    Set wsP       = ThisWorkbook.Sheets("Préférences_Projets")
    Set wsBilan   = ThisWorkbook.Sheets("Bilan_Performance")
    Set wsDetails = ThisWorkbook.Sheets("Details_Suivi")
    If wsDetails Is Nothing Then
        Set wsDetails = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsDetails.Name = "Details_Suivi"
    End If
    On Error GoTo 0

    If wsR Is Nothing Or wsPE Is Nothing Or wsP Is Nothing Or wsBilan Is Nothing Then
        MsgBox "Erreur : feuilles manquantes.", vbCritical: Exit Sub
    End If

    Application.ScreenUpdating = False
    wsBilan.Cells.Clear
    wsDetails.Cells.Clear

    ' ---- Lecture des données ----
    Set affectationsEquipe  = CreateObject("Scripting.Dictionary")
    Set affectationsProjet  = CreateObject("Scripting.Dictionary")
    Set prefsEquipes        = CreateObject("Scripting.Dictionary")
    Set projetsCapacites    = CreateObject("Scripting.Dictionary")

    For i = 2 To wsR.Cells(wsR.Rows.Count, "A").End(xlUp).Row
        nomP = Trim(wsR.Cells(i, 1).Value)
        listeStr = Trim(wsR.Cells(i, 2).Value)
        Set affectationsProjet(nomP) = New Collection
        If listeStr <> "Aucune" And listeStr <> "" Then
            arr = Split(listeStr, ", ")
            For Each eq In arr
                affectationsEquipe(Trim(eq)) = nomP
                affectationsProjet(nomP).Add Trim(eq)
            Next eq
        End If
    Next i

    For i = 2 To wsPE.Cells(wsPE.Rows.Count, "A").End(xlUp).Row
        nomEq = Trim(wsPE.Cells(i, 1).Value)
        If nomEq <> "" Then
            Set prefsEquipes(nomEq) = New Collection
            For j = 2 To wsPE.Cells(i, wsPE.Columns.Count).End(xlToLeft).Column
                prefsEquipes(nomEq).Add Trim(wsPE.Cells(i, j).Value)
            Next j
        End If
    Next i

    For i = 2 To wsP.Cells(wsP.Rows.Count, "A").End(xlUp).Row
        nomP = Trim(wsP.Cells(i, 1).Value)
        If nomP <> "" Then
            projetsCapacites(nomP) = Array(CLng(wsP.Cells(i, 2).Value), CLng(wsP.Cells(i, 3).Value))
        End If
    Next i

    ' ---- Calcul des métriques ----
    nbEquipesTotal     = prefsEquipes.Count
    nbEquipesAffectees = affectationsEquipe.Count
    Set rangsObtenus = New Collection
    Set equipesSansProjet = New Collection
    Set projetsSousMinimum = New Collection
    Set projetsVides = New Collection
    nbProjetsMinAtteint = 0

    For Each eq In prefsEquipes.Keys
        If affectationsEquipe.Exists(eq) Then
            rang = 0
            Set prefs = prefsEquipes(eq)
            For j = 1 To prefs.Count
                If StrComp(prefs(j), affectationsEquipe(eq), vbTextCompare) = 0 Then rang = j: Exit For
            Next j
            If rang > 0 Then
                rangsObtenus.Add rang
                Select Case rang
                    Case 1: nbChoix1 = nbChoix1 + 1
                    Case 2: nbChoix2 = nbChoix2 + 1
                    Case 3: nbChoix3 = nbChoix3 + 1
                End Select
            End If
        Else
            equipesSansProjet.Add eq
        End If
    Next eq

    For Each projet In projetsCapacites.Keys
        nbAff = IIf(affectationsProjet.Exists(projet), affectationsProjet(projet).Count, 0)
        If nbAff = 0 Then projetsVides.Add projet
        If nbAff < projetsCapacites(projet)(0) Then
            projetsSousMinimum.Add projet
        Else
            nbProjetsMinAtteint = nbProjetsMinAtteint + 1
        End If
    Next projet

    If rangsObtenus.Count > 0 Then
        somme = 0
        For Each ri In rangsObtenus: somme = somme + ri: Next ri
        rangMoyen = somme / rangsObtenus.Count
        If rangsObtenus.Count > 1 Then
            sc = 0
            For Each ri In rangsObtenus: sc = sc + (ri - rangMoyen) ^ 2: Next ri
            ecartType = Sqr(sc / (rangsObtenus.Count - 1))
        End If
    End If

    ' ---- Écriture du bilan ----
    wsBilan.Columns("A").ColumnWidth = 45: wsBilan.Columns("B").ColumnWidth = 15
    ligne = 1

    wsBilan.Cells(ligne, "A").Value = "BILAN DU POINT DE VUE DES ÉQUIPES"
    wsBilan.Range("A" & ligne & ":B" & ligne).Merge
    wsBilan.Range("A" & ligne).Font.Bold = True
    ligne = ligne + 2
    wsBilan.Cells(ligne, "A").Value = "Nombre total d'équipes":          wsBilan.Cells(ligne, "B").Value = nbEquipesTotal:          ligne = ligne + 1
    wsBilan.Cells(ligne, "A").Value = "Nombre d'équipes affectées":      wsBilan.Cells(ligne, "B").Value = nbEquipesAffectees:      ligne = ligne + 1
    wsBilan.Cells(ligne, "A").Value = "Nombre d'équipes non affectées":  wsBilan.Cells(ligne, "B").Value = equipesSansProjet.Count: ligne = ligne + 1
    If nbEquipesTotal > 0 Then
        wsBilan.Cells(ligne, "A").Value = "Taux d'affectation"
        wsBilan.Cells(ligne, "B").Value = Format(nbEquipesAffectees / nbEquipesTotal, "0.0%")
    End If: ligne = ligne + 2
    wsBilan.Cells(ligne, "A").Value = "Satisfaction moyenne (rang du vœu obtenu)": wsBilan.Cells(ligne, "B").Value = Round(rangMoyen, 2): ligne = ligne + 1
    wsBilan.Cells(ligne, "A").Value = "Écart-type des rangs (équité)":             wsBilan.Cells(ligne, "B").Value = Round(ecartType, 2):  ligne = ligne + 2
    wsBilan.Cells(ligne, "A").Value = "Équipes ayant obtenu leur 1er vœu":   wsBilan.Cells(ligne, "B").Value = nbChoix1: ligne = ligne + 1
    wsBilan.Cells(ligne, "A").Value = "Équipes ayant obtenu leur 2ème vœu":  wsBilan.Cells(ligne, "B").Value = nbChoix2: ligne = ligne + 1
    wsBilan.Cells(ligne, "A").Value = "Équipes ayant obtenu leur 3ème vœu":  wsBilan.Cells(ligne, "B").Value = nbChoix3: ligne = ligne + 3

    wsBilan.Cells(ligne, "A").Value = "BILAN DU POINT DE VUE DES PROJETS"
    wsBilan.Range("A" & ligne & ":B" & ligne).Merge
    wsBilan.Range("A" & ligne).Font.Bold = True
    ligne = ligne + 2
    wsBilan.Cells(ligne, "A").Value = "Nombre total de projets":                       wsBilan.Cells(ligne, "B").Value = projetsCapacites.Count: ligne = ligne + 1
    wsBilan.Cells(ligne, "A").Value = "Projets ayant atteint leur minimum d'équipes":  wsBilan.Cells(ligne, "B").Value = nbProjetsMinAtteint:     ligne = ligne + 1
    If projetsCapacites.Count > 0 Then
        wsBilan.Cells(ligne, "A").Value = "Taux de projets satisfaits"
        wsBilan.Cells(ligne, "B").Value = Format(nbProjetsMinAtteint / projetsCapacites.Count, "0.0%")
    End If: ligne = ligne + 2
    wsBilan.Range("A1:B" & ligne).Borders.Weight = xlThin

    ligne = ligne + 1
    wsBilan.Hyperlinks.Add Anchor:=wsBilan.Cells(ligne, "A"), Address:="", _
        SubAddress:="'Details_Suivi'!A1", TextToDisplay:="Voir les listes de suivi détaillées"
    wsBilan.Cells(ligne, "A").Font.Underline = xlUnderlineStyleSingle
    wsBilan.Cells(ligne, "A").Font.Color = vbBlue

    ' ---- Listes de suivi ----
    wsDetails.Columns("A").ColumnWidth = 35: wsDetails.Columns("B").ColumnWidth = 35
    ld = 1
    wsDetails.Cells(ld, "A").Value = "LISTES DE SUIVI": wsDetails.Cells(ld, "A").Font.Bold = True: ld = ld + 2

    wsDetails.Cells(ld, "A").Value = "Équipes sans projet :": wsDetails.Cells(ld, "A").Font.Bold = True
    If equipesSansProjet.Count > 0 Then
        For i = 1 To equipesSansProjet.Count: wsDetails.Cells(ld + i - 1, "B").Value = equipesSansProjet(i): Next i
        ld = ld + equipesSansProjet.Count
    Else
        wsDetails.Cells(ld, "B").Value = "Aucune": ld = ld + 1
    End If: ld = ld + 1

    wsDetails.Cells(ld, "A").Value = "Projets n'atteignant pas leur minimum d'équipes :": wsDetails.Cells(ld, "A").Font.Bold = True
    If projetsSousMinimum.Count > 0 Then
        For i = 1 To projetsSousMinimum.Count: wsDetails.Cells(ld + i - 1, "B").Value = projetsSousMinimum(i): Next i
        ld = ld + projetsSousMinimum.Count
    Else
        wsDetails.Cells(ld, "B").Value = "Aucun": ld = ld + 1
    End If: ld = ld + 1

    wsDetails.Cells(ld, "A").Value = "Projets sans aucune équipe :": wsDetails.Cells(ld, "A").Font.Bold = True
    If projetsVides.Count > 0 Then
        For i = 1 To projetsVides.Count: wsDetails.Cells(ld + i - 1, "B").Value = projetsVides(i): Next i
    Else
        wsDetails.Cells(ld, "B").Value = "Aucun"
    End If

    Application.ScreenUpdating = True
    MsgBox "Bilan de performance et listes de suivi générés.", vbInformation

End Sub


'================================================================================================
' MACRO 6 : AFFECTATIONS PAR PROJET (vue détaillée avec membres des équipes)
' Feuilles lues  : Résultats, Équipes, Préférences_Projets
' Feuille écrite : Affectations_par_Projet
'================================================================================================
Sub RemplirAffectationsParProjet()

    Dim wsR As Worksheet, wsEq As Worksheet, wsP As Worksheet, wsAff As Worksheet
    Dim i As Long, j As Long
    Dim nomProjet As String, listeStr As String, arr As Variant, nomEq As Variant
    Dim minEq As Long, maxEq As Long, nbAff As Long
    Dim ligne As Long
    Dim capacites As Object, rangsProjet As Object, membres As Object
    Dim rd As Object, lastColP As Long, nomEqStr As String
    Dim memList As String, nbMem As Long, tailleEq As Long
    Dim nomEqClean As String, premiereLigne As Long, statut As String
    Dim eqH As String

    On Error Resume Next
    Set wsR   = ThisWorkbook.Sheets("Résultats")
    Set wsEq  = ThisWorkbook.Sheets("Équipes")
    Set wsP   = ThisWorkbook.Sheets("Préférences_Projets")
    Set wsAff = ThisWorkbook.Sheets("Affectations_par_Projet")
    If wsAff Is Nothing Then
        Set wsAff = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsAff.Name = "Affectations_par_Projet"
    End If
    On Error GoTo 0

    If wsR Is Nothing Or wsEq Is Nothing Or wsP Is Nothing Then
        MsgBox "Erreur : feuilles Résultats, Équipes ou Préférences_Projets manquantes.", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False
    wsAff.Cells.Clear

    ' ---- En-tête ----
    wsAff.Range("A1:F1").Value = Array("Projet", "Statut Capacité", "Équipe", "Membres", "Taille", "Rang dans projet")
    wsAff.Range("A1:F1").Font.Bold = True
    wsAff.Range("A1:F1").Interior.Color = RGB(68, 114, 196)
    wsAff.Range("A1:F1").Font.Color = vbWhite
    ligne = 2

    ' ---- Lecture capacités et rangs projets ----
    Set capacites = CreateObject("Scripting.Dictionary")
    Set rangsProjet = CreateObject("Scripting.Dictionary")
    lastColP = wsP.Cells(1, wsP.Columns.Count).End(xlToLeft).Column
    For i = 2 To wsP.Cells(wsP.Rows.Count, "A").End(xlUp).Row
        nomProjet = Trim(wsP.Cells(i, 1).Value)
        If nomProjet <> "" Then
            capacites(nomProjet) = Array(CLng(wsP.Cells(i, 2).Value), CLng(wsP.Cells(i, 3).Value))
            Set rd = CreateObject("Scripting.Dictionary")
            For j = 6 To lastColP
                eqH = Trim(wsP.Cells(1, j).Value)
                If eqH <> "" Then rd(eqH) = CLng(wsP.Cells(i, j).Value)
            Next j
            Set rangsProjet(nomProjet) = rd
        End If
    Next i

    ' ---- Lecture composition équipes ----
    Set membres = CreateObject("Scripting.Dictionary")
    For i = 2 To wsEq.Cells(wsEq.Rows.Count, "A").End(xlUp).Row
        nomEqStr = Trim(wsEq.Cells(i, 1).Value)
        If nomEqStr <> "" Then
            memList = ""
            For j = 2 To wsEq.Cells(i, wsEq.Columns.Count).End(xlToLeft).Column
                If Trim(wsEq.Cells(i, j).Value) <> "" Then
                    If memList <> "" Then memList = memList & ", "
                    memList = memList & Trim(wsEq.Cells(i, j).Value)
                End If
            Next j
            membres(nomEqStr) = IIf(memList <> "", memList, "(aucun)")
        End If
    Next i

    ' ---- Construction tableau par projet ----
    For i = 2 To wsR.Cells(wsR.Rows.Count, "A").End(xlUp).Row
        nomProjet = Trim(wsR.Cells(i, 1).Value)
        listeStr  = Trim(wsR.Cells(i, 2).Value)
        If nomProjet = "" Then GoTo SuiteLigne

        statut = Trim(wsR.Cells(i, 3).Value)

        If listeStr = "Aucune" Or listeStr = "" Then
            wsAff.Cells(ligne, 1).Value = nomProjet
            wsAff.Cells(ligne, 2).Value = statut
            wsAff.Cells(ligne, 3).Value = "(aucune équipe affectée)"
            wsAff.Range("A" & ligne & ":F" & ligne).Interior.Color = RGB(255, 235, 235)
            ligne = ligne + 1
        Else
            arr = Split(listeStr, ", ")
            nbAff = UBound(arr) - LBound(arr) + 1
            premiereLigne = ligne

            For Each nomEq In arr
                nomEqClean = Trim(CStr(nomEq))
                wsAff.Cells(ligne, 1).Value = IIf(ligne = premiereLigne, nomProjet, "")
                wsAff.Cells(ligne, 2).Value = IIf(ligne = premiereLigne, statut, "")
                wsAff.Cells(ligne, 3).Value = nomEqClean
                wsAff.Cells(ligne, 4).Value = IIf(membres.Exists(nomEqClean), membres(nomEqClean), "?")
                tailleEq = 0
                If membres.Exists(nomEqClean) And membres(nomEqClean) <> "(aucun)" Then
                    tailleEq = UBound(Split(membres(nomEqClean), ", ")) + 1
                End If
                wsAff.Cells(ligne, 5).Value = tailleEq
                If rangsProjet.Exists(nomProjet) Then
                    If rangsProjet(nomProjet).Exists(nomEqClean) Then
                        wsAff.Cells(ligne, 6).Value = rangsProjet(nomProjet)(nomEqClean)
                    End If
                End If
                If InStr(statut, "MINIMUM NON ATTEINT") > 0 Then
                    wsAff.Cells(ligne, 1).Interior.Color = RGB(255, 235, 156)
                Else
                    wsAff.Cells(ligne, 1).Interior.Color = RGB(235, 255, 235)
                End If
                ligne = ligne + 1
            Next nomEq

            If nbAff > 1 Then
                wsAff.Range("A" & premiereLigne & ":A" & ligne - 1).Merge
                wsAff.Range("B" & premiereLigne & ":B" & ligne - 1).Merge
                wsAff.Range("A" & premiereLigne & ":B" & ligne - 1).VerticalAlignment = xlVAlignCenter
            End If
        End If

        wsAff.Range("A" & ligne & ":F" & ligne).Borders(xlEdgeTop).Weight = xlThin

SuiteLigne:
    Next i

    wsAff.Columns("A:F").AutoFit
    wsAff.Columns("D").ColumnWidth = 50
    wsAff.Range("A1:F" & ligne - 1).Borders.Weight = xlThin

    Application.ScreenUpdating = True
    MsgBox "Feuille 'Affectations_par_Projet' générée (" & ligne - 2 & " lignes).", vbInformation

End Sub


'================================================================================================
' MACRO DIAGNOSTIC : Vérifier la disponibilité de System.Collections.ArrayList
'================================================================================================
Sub TestArrayList()
    Dim testList As Object
    On Error Resume Next
    Set testList = CreateObject("System.Collections.ArrayList")
    If Err.Number <> 0 Then
        MsgBox "ERREUR : 'System.Collections.ArrayList' n'est PAS disponible.", vbCritical
    Else
        MsgBox "SUCCÈS : 'System.Collections.ArrayList' est disponible.", vbInformation
    End If
End Sub


'================================================================================================
' UTILITAIRES PRIVÉS
'================================================================================================
Private Sub LogStep(ByVal ws As Worksheet, ByRef ligne As Long, ByVal action As String, _
                    ByVal decision As String, ByVal statut As String, ByVal libre As String)
    ligne = ligne + 1
    ws.Cells(ligne, 1).Value = ligne - 1
    ws.Cells(ligne, 2).Value = action
    ws.Cells(ligne, 3).Value = decision
    ws.Cells(ligne, 4).Value = statut
    ws.Cells(ligne, 5).Value = libre
End Sub

Private Function GetCelibatairesString(ByVal coll As Collection) As String
    Dim arr() As String
    Dim i As Long
    
    If coll.Count = 0 Then GetCelibatairesString = "Aucune": Exit Function
    ReDim arr(1 To coll.Count)
    For i = 1 To coll.Count: arr(i) = coll(i): Next i
    GetCelibatairesString = Join(arr, ", ")
End Function

Private Function CollectionToArray(coll As Collection) As Variant
    Dim arr() As Variant
    Dim i As Long
    
    ReDim arr(1 To coll.Count)
    For i = 1 To coll.Count: arr(i) = coll(i): Next i
    CollectionToArray = arr
End Function
