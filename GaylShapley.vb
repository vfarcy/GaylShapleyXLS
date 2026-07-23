Option Explicit ' Force la déclaration de toutes les variables, une bonne pratique.

'****************************************************************************************
' MACRO UNIFIÉE : Génère les deux feuilles de préférences en une seule fois.
'****************************************************************************************
Sub GenererDonneesDeTest()

    Dim nbProjets As Long, nbEleves As Long
    Dim wsProjets As Worksheet, wsEleves As Worksheet
    
    ' --- ÉTAPE 1 : Saisie unique des données ---
    On Error Resume Next
    nbEleves = CLng(InputBox("Combien d'élèves souhaitez-vous générer ?", "Génération des Données"))
    If Err.Number <> 0 Or nbEleves <= 0 Then
        MsgBox "Génération annulée.", vbInformation
        Exit Sub
    End If
    
    nbProjets = CLng(InputBox("Combien de projets souhaitez-vous générer ?", "Génération des Données"))
    If Err.Number <> 0 Or nbProjets <= 0 Then
        MsgBox "Génération annulée.", vbInformation
        Exit Sub
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = False

    ' --- ÉTAPE 2 : Génération de la feuille Préférences_Projets ---
    Set wsProjets = ThisWorkbook.Sheets("Préférences_Projets")
    wsProjets.Cells.Clear
    
    wsProjets.Range("A1:C1").Value = Array("Projet", "Min", "Max")
    Dim i As Long, j As Long
    For i = 1 To nbEleves
        wsProjets.Cells(1, 3 + i).Value = "Élève " & i
    Next i
    wsProjets.Rows(1).Font.Bold = True
    
    Randomize
    For i = 1 To nbProjets
        wsProjets.Cells(i + 1, 1).Value = "Projet " & Chr(64 + i)
        
        Dim capMin As Long: capMin = Application.WorksheetFunction.RandBetween(1, 3)
        Dim capMax As Long: capMax = Application.WorksheetFunction.RandBetween(capMin, capMin + 5)
        wsProjets.Cells(i + 1, 2).Value = capMin
        wsProjets.Cells(i + 1, 3).Value = capMax
        
        Dim rangs As Object: Set rangs = CreateObject("System.Collections.ArrayList")
        For j = 1 To nbEleves
            rangs.Add j
        Next j
        
        Dim temp As Variant, r As Long
        For j = rangs.Count - 1 To 0 Step -1
            r = Int(j * Rnd)
            temp = rangs(j)
            rangs(j) = rangs(r)
            rangs(r) = temp
        Next j
        
        For j = 0 To rangs.Count - 1
            wsProjets.Cells(i + 1, 4 + j).Value = rangs(j)
        Next j
    Next i
    wsProjets.Columns.AutoFit

    ' --- ÉTAPE 3 : Génération de la feuille Préférences_Élèves ---
    Set wsEleves = ThisWorkbook.Sheets("Préférences_Élèves")
    wsEleves.Cells.Clear
    
    wsEleves.Cells(1, 1).Value = "Élève"
    For i = 1 To nbProjets
        wsEleves.Cells(1, 1 + i).Value = "Choix " & i
    Next i
    wsEleves.Rows(1).Font.Bold = True

    Dim listeProjets As Object: Set listeProjets = CreateObject("System.Collections.ArrayList")
    For i = 1 To nbProjets
        listeProjets.Add "Projet " & Chr(64 + i)
    Next i

    For i = 1 To nbEleves
        wsEleves.Cells(i + 1, 1).Value = "Élève " & i
        
        For j = listeProjets.Count - 1 To 0 Step -1
            r = Int(j * Rnd)
            temp = listeProjets(j)
            listeProjets(j) = listeProjets(r)
            listeProjets(r) = temp
        Next j
        
        For j = 0 To listeProjets.Count - 1
            wsEleves.Cells(i + 1, 2 + j).Value = listeProjets(j)
        Next j
    Next i
    wsEleves.Columns.AutoFit
    
    Application.ScreenUpdating = True
    MsgBox nbEleves & " élèves et " & nbProjets & " projets ont été générés avec succès.", vbInformation

End Sub



'****************************************************************************************
' MACRO D'AFFECTATION (Version Corrigée avec logique focalisée)
'****************************************************************************************
Sub AffectationElevesProjets()

    ' Déclaration des feuilles
    Dim wsEleves As Worksheet, wsProjets As Worksheet, wsResultats As Worksheet
    
    ' Structures de données
    Dim elevesPrefs As Object, projetsPrefs As Object, projetsCapacites As Object
    Dim affectationsProjet As Object, affectationEleve As Object
    Dim celibataires As Collection
    
    ' Initialisation
    Set elevesPrefs = CreateObject("Scripting.Dictionary")
    Set projetsPrefs = CreateObject("Scripting.Dictionary")
    Set projetsCapacites = CreateObject("Scripting.Dictionary")
    Set affectationsProjet = CreateObject("Scripting.Dictionary")
    Set affectationEleve = CreateObject("Scripting.Dictionary")
    Set celibataires = New Collection

    ' --- Association des feuilles ---
    On Error Resume Next
    Set wsEleves = ThisWorkbook.Sheets("Préférences_Élèves")
    Set wsProjets = ThisWorkbook.Sheets("Préférences_Projets")
    Set wsResultats = ThisWorkbook.Sheets("Résultats")
    If wsEleves Is Nothing Or wsProjets Is Nothing Or wsResultats Is Nothing Then
        MsgBox "Erreur : Une ou plusieurs feuilles ('Préférences_Élèves', 'Préférences_Projets', 'Résultats') sont introuvables.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = False

    ' --- 1. LECTURE DES DONNÉES ---
    Dim i As Long, j As Long, eleve As Variant
    ' Lecture Élèves
    For i = 2 To wsEleves.Cells(wsEleves.Rows.Count, "A").End(xlUp).Row
        Dim nomEleve As String: nomEleve = wsEleves.Cells(i, 1).Value
        If nomEleve <> "" Then
            celibataires.Add nomEleve
            affectationEleve(nomEleve) = ""
            Dim prefsList As New Collection
            For j = 2 To wsEleves.Cells(i, wsEleves.Columns.Count).End(xlToLeft).Column
                prefsList.Add wsEleves.Cells(i, j).Value
            Next j
            Set elevesPrefs(nomEleve) = prefsList
        End If
    Next i
    ' Lecture Projets
    For i = 2 To wsProjets.Cells(wsProjets.Rows.Count, "A").End(xlUp).Row
        Dim nomProjet As String: nomProjet = wsProjets.Cells(i, 1).Value
        If nomProjet <> "" Then
            projetsCapacites(nomProjet) = Array(wsProjets.Cells(i, 2).Value, wsProjets.Cells(i, 3).Value)
            Set affectationsProjet(nomProjet) = CreateObject("Scripting.Dictionary")
            Dim rangsDict As Object: Set rangsDict = CreateObject("Scripting.Dictionary")
            For j = 4 To wsProjets.Cells(1, wsProjets.Columns.Count).End(xlToLeft).Column
                rangsDict(wsProjets.Cells(1, j).Value) = wsProjets.Cells(i, j).Value
            Next j
            Set projetsPrefs(nomProjet) = rangsDict
        End If
    Next i
    
    ' --- 2. EXÉCUTION DE L'ALGORITHME (Logique Corrigée) ---
    While celibataires.Count > 0
        Dim eleveActuel As String: eleveActuel = celibataires(1) ' On prend le premier de la liste
        celibataires.Remove 1 ' On le retire de la liste des célibataires à traiter

        Dim estPlace As Boolean: estPlace = False
        
        ' On boucle sur les choix de cet élève JUSQU'À ce qu'il soit placé ou qu'il n'ait plus de choix
        For j = 1 To elevesPrefs(eleveActuel).Count
            Dim projetVise As String: projetVise = elevesPrefs(eleveActuel)(j)
            
            Dim affectesAuProjet As Object: Set affectesAuProjet = affectationsProjet(projetVise)
            Dim capaciteMax As Long: capaciteMax = projetsCapacites(projetVise)(1)

            ' Cas 1: Le projet a de la place
            If affectesAuProjet.Count < capaciteMax Then
                affectesAuProjet(eleveActuel) = projetsPrefs(projetVise)(eleveActuel)
                affectationEleve(eleveActuel) = projetVise
                estPlace = True
                Exit For ' L'élève est placé, on sort de sa boucle de choix
            Else
            ' Cas 2: Le projet est plein
                Dim pireEleve As String: pireEleve = ""
                Dim rangPireEleve As Long: rangPireEleve = -1
                Dim eleveAffecte As Variant
                For Each eleveAffecte In affectesAuProjet.Keys
                    If affectesAuProjet(eleveAffecte) > rangPireEleve Then
                        rangPireEleve = affectesAuProjet(eleveAffecte)
                        pireEleve = eleveAffecte
                    End If
                Next eleveAffecte
                
                Dim rangEleveActuel As Long: rangEleveActuel = projetsPrefs(projetVise)(eleveActuel)
                
                If rangEleveActuel < rangPireEleve Then
                    ' Le nouvel élève est meilleur
                    affectesAuProjet.Remove pireEleve
                    affectationEleve(pireEleve) = ""
                    celibataires.Add pireEleve ' Le pire élève retourne dans la file d'attente
                    
                    affectesAuProjet(eleveActuel) = rangEleveActuel
                    affectationEleve(eleveActuel) = projetVise
                    estPlace = True
                    Exit For ' L'élève est placé, on sort de sa boucle de choix
                ' Else: L'élève est rejeté, la boucle 'For j' continue pour qu'il essaie son choix suivant
                End If
            End If
        Next j ' Passe au choix suivant pour eleveActuel s'il n'est pas encore placé
    Wend

    ' --- 3. ÉCRITURE DU RÉSULTAT FINAL ---
    wsResultats.Cells.ClearContents
    wsResultats.Range("A1:C1").Value = Array("Projet", "Élèves Affectés", "Statut Capacité")
    wsResultats.Range("A1:C1").Font.Bold = True
    
    Dim ligneResultat As Long: ligneResultat = 2
    Dim projet As Variant
    For Each projet In affectationsProjet.Keys
        wsResultats.Cells(ligneResultat, 1).Value = projet
        
        Dim listeEleves As String
        If affectationsProjet(projet).Count > 0 Then
            listeEleves = Join(affectationsProjet(projet).Keys, ", ")
        Else
            listeEleves = "Aucun"
        End If
        wsResultats.Cells(ligneResultat, 2).Value = listeEleves
        wsResultats.Cells(ligneResultat, 2).WrapText = True
        
        Dim statut As String
        Dim nbAffectes As Long: nbAffectes = affectationsProjet(projet).Count
        Dim capMin As Long: capMin = projetsCapacites(projet)(0)
        Dim capMax As Long: capMax = projetsCapacites(projet)(1)
        If nbAffectes < capMin Then
            statut = "MINIMUM NON ATTEINT (" & nbAffectes & "/" & capMin & ")"
            wsResultats.Cells(ligneResultat, 3).Interior.Color = vbYellow
        Else
            statut = "OK (" & nbAffectes & "/" & capMax & ")"
            wsResultats.Cells(ligneResultat, 3).Interior.ColorIndex = xlNone
        End If
        wsResultats.Cells(ligneResultat, 3).Value = statut
        
        ligneResultat = ligneResultat + 1
    Next projet
    
    wsResultats.Columns.AutoFit
    
    Application.ScreenUpdating = True
    MsgBox "L'affectation des élèves aux projets est terminée. Consultez la feuille 'Résultats'.", vbInformation
End Sub



'****************************************************************************************
' Pour test , à effacer
' Corrige l'erreur 450 en utilisant le mot-clé 'Set' pour les affectations d'objets.
'****************************************************************************************

Sub TestArrayList()
    Dim testList As Object
    On Error Resume Next ' Continue même si une erreur se produit

    ' Tente de créer l'objet
    Set testList = CreateObject("System.Collections.ArrayList")

    ' Vérifie si une erreur s'est produite
    If Err.Number <> 0 Then
        MsgBox "ERREUR : Le composant 'System.Collections.ArrayList' n'est PAS disponible sur votre système.", vbCritical
    Else
        MsgBox "SUCCÈS : Le composant 'System.Collections.ArrayList' est bien disponible. Le problème est ailleurs.", vbInformation
    End If
End Sub


Option Explicit ' Force la déclaration de toutes les variables, une bonne pratique.

'****************************************************************************************
' MACRO 4 : ANALYSE DE LA SATISFACTION (avec Moyenne et Écart-Type)
'****************************************************************************************
Sub RapportSatisfactionEleves()

    ' Déclaration des feuilles
    Dim wsResultats As Worksheet, wsPrefsEleves As Worksheet, wsRapport As Worksheet
    
    ' Structures de données
    Dim affectations As Object, prefsEleves As Object
    Dim rangsObtenus As Collection ' Pour stocker les rangs pour le calcul de l'écart-type
    
    ' Variables de boucle et de travail
    Dim i As Long, j As Long, eleve As Variant, prefsList As Collection
    
    ' Initialisation
    Set affectations = CreateObject("Scripting.Dictionary")
    Set prefsEleves = CreateObject("Scripting.Dictionary")
    Set rangsObtenus = New Collection

    ' --- Association des feuilles ---
    On Error Resume Next
    Set wsResultats = ThisWorkbook.Sheets("Résultats")
    Set wsPrefsEleves = ThisWorkbook.Sheets("Préférences_Élèves")
    Set wsRapport = ThisWorkbook.Sheets("Rapport_Satisfaction")
    If wsResultats Is Nothing Or wsPrefsEleves Is Nothing Or wsRapport Is Nothing Then
        MsgBox "Erreur : Assurez-vous que les feuilles 'Résultats', 'Préférences_Élèves' et 'Rapport_Satisfaction' existent.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = False

    ' --- 1. LECTURE DES RÉSULTATS ---
    Dim derniereLigneR As Long
    derniereLigneR = wsResultats.Cells(wsResultats.Rows.Count, "A").End(xlUp).Row
    For i = 2 To derniereLigneR
        Dim projetAffecte As String: projetAffecte = wsResultats.Cells(i, 1).Value
        Dim listeElevesStr As String: listeElevesStr = wsResultats.Cells(i, 2).Value
        If listeElevesStr <> "Aucun" And listeElevesStr <> "" Then
            Dim elevesAffectesArray As Variant: elevesAffectesArray = Split(listeElevesStr, ", ")
            For Each eleve In elevesAffectesArray
                affectations(eleve) = projetAffecte
            Next eleve
        End If
    Next i
    
    ' --- 2. LECTURE DES PRÉFÉRENCES ---
    Dim derniereLigneE As Long
    derniereLigneE = wsPrefsEleves.Cells(wsPrefsEleves.Rows.Count, "A").End(xlUp).Row
    For i = 2 To derniereLigneE
        Dim nomEleve As String: nomEleve = wsPrefsEleves.Cells(i, 1).Value
        If nomEleve <> "" Then
            Set prefsList = New Collection
            Dim derniereColE As Long
            derniereColE = wsPrefsEleves.Cells(i, wsPrefsEleves.Columns.Count).End(xlToLeft).Column
            For j = 2 To derniereColE
                prefsList.Add wsPrefsEleves.Cells(i, j).Value
            Next j
            Set prefsEleves(nomEleve) = prefsList
        End If
    Next i
    
    ' --- 3. GÉNÉRATION DU RAPPORT ---
    wsRapport.Cells.ClearContents
    wsRapport.Range("A1:C1").Value = Array("Élève", "Projet Affecté", "Rang du Choix")
    wsRapport.Range("A1:C1").Font.Bold = True
    
    Dim ligneRapport As Long: ligneRapport = 2
    
    For Each eleve In prefsEleves.Keys
        wsRapport.Cells(ligneRapport, 1).Value = eleve
        If affectations.Exists(eleve) Then
            Dim projetObtenu As String: projetObtenu = affectations(eleve)
            wsRapport.Cells(ligneRapport, 2).Value = projetObtenu
            Dim rang As Long: rang = 0
            Set prefsList = prefsEleves(eleve)
            For j = 1 To prefsList.Count
                If prefsList(j) = projetObtenu Then
                    rang = j
                    Exit For
                End If
            Next j
            If rang > 0 Then
                wsRapport.Cells(ligneRapport, 3).Value = rang
                rangsObtenus.Add rang ' Ajoute le rang à notre collection pour les calculs stats
            Else
                wsRapport.Cells(ligneRapport, 3).Value = "Erreur: Projet non trouvé"
            End If
        Else
            wsRapport.Cells(ligneRapport, 2).Value = "Non affecté"
            wsRapport.Cells(ligneRapport, 3).Value = "N/A"
        End If
        ligneRapport = ligneRapport + 1
    Next eleve
    
    ' --- 4. AFFICHAGE DES STATISTIQUES ---
    If rangsObtenus.Count > 0 Then
        ' Calcul de la moyenne
        Dim sommeRangs As Double: sommeRangs = 0
        Dim rangItem As Variant
        For Each rangItem In rangsObtenus
            sommeRangs = sommeRangs + rangItem
        Next rangItem
        Dim rangMoyen As Double: rangMoyen = sommeRangs / rangsObtenus.Count
        
        ' Calcul de l'écart-type
        Dim sommeCarresEcarts As Double: sommeCarresEcarts = 0
        For Each rangItem In rangsObtenus
            sommeCarresEcarts = sommeCarresEcarts + (rangItem - rangMoyen) ^ 2
        Next rangItem
        Dim ecartType As Double
        If rangsObtenus.Count > 1 Then
            ecartType = Sqr(sommeCarresEcarts / (rangsObtenus.Count - 1)) ' Écart-type d'échantillon (le plus courant)
        Else
            ecartType = 0 ' Pas de dispersion avec une seule valeur
        End If

        ' Affichage dans la feuille
        ligneRapport = ligneRapport + 1 ' Laisse une ligne vide
        wsRapport.Cells(ligneRapport, "B").Value = "Satisfaction moyenne :"
        wsRapport.Cells(ligneRapport, "C").Value = Round(rangMoyen, 2)
        
        ligneRapport = ligneRapport + 1
        wsRapport.Cells(ligneRapport, "B").Value = "Écart-type des rangs :"
        wsRapport.Cells(ligneRapport, "C").Value = Round(ecartType, 2)
        
        wsRapport.Range(wsRapport.Cells(ligneRapport - 1, "B"), wsRapport.Cells(ligneRapport, "C")).Font.Bold = True
    End If
    
    wsRapport.Columns.AutoFit
    Application.ScreenUpdating = True
    
    MsgBox "Le rapport de satisfaction des élèves (avec moyenne et écart-type) a été généré.", vbInformation

End Sub

'****************************************************************************************
' MACRO 5 : AFFECTATION PAS À PAS (Version Finale avec Correction de Lecture)
'****************************************************************************************
Sub AffectationPasAPas()

    ' Déclaration des feuilles
    Dim wsEleves As Worksheet, wsProjets As Worksheet, wsResultats As Worksheet, wsLog As Worksheet
    
    ' Structures de données
    Dim elevesPrefs As Object, projetsPrefs As Object, projetsCapacites As Object
    Dim affectationsProjet As Object, affectationEleve As Object
    Dim celibataires As Collection, propositionsFaites As Object
    
    ' Initialisation
    Set elevesPrefs = CreateObject("Scripting.Dictionary")
    Set projetsPrefs = CreateObject("Scripting.Dictionary")
    Set projetsCapacites = CreateObject("Scripting.Dictionary")
    Set affectationsProjet = CreateObject("Scripting.Dictionary")
    Set affectationEleve = CreateObject("Scripting.Dictionary")
    Set celibataires = New Collection
    Set propositionsFaites = CreateObject("Scripting.Dictionary")

    ' --- Association des feuilles ---
    On Error Resume Next
    Set wsEleves = ThisWorkbook.Sheets("Préférences_Élèves")
    Set wsProjets = ThisWorkbook.Sheets("Préférences_Projets")
    Set wsResultats = ThisWorkbook.Sheets("Résultats")
    Set wsLog = ThisWorkbook.Sheets("Log_Affectation")
    If wsEleves Is Nothing Or wsProjets Is Nothing Or wsResultats Is Nothing Or wsLog Is Nothing Then
        MsgBox "Erreur : Assurez-vous que toutes les feuilles nécessaires existent.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = False
    
    ' --- Préparation du journal ---
    Dim ligneLog As Long: ligneLog = 1
    wsLog.Cells.Clear
    wsLog.Range("A1:E1").Value = Array("Étape", "Action de l'Élève", "Décision du Projet", "Statut du Projet", "Élèves libres")
    wsLog.Range("A1:E1").Font.Bold = True

    ' --- 1. LECTURE SÉCURISÉE DES DONNÉES (CORRIGÉE) ---
    Dim i As Long, j As Long
    Dim prefsList As Collection ' Déclarée une seule fois ici
    
    ' Lecture Élèves
    For i = 2 To wsEleves.Cells(wsEleves.Rows.Count, "A").End(xlUp).Row
        Dim nomEleve As String: nomEleve = Trim(wsEleves.Cells(i, 1).Value)
        If nomEleve <> "" Then
            celibataires.Add nomEleve
            affectationEleve(nomEleve) = ""
            propositionsFaites(nomEleve) = 0
            
            ' CORRECTION CRUCIALE : On force la création d'une NOUVELLE liste pour chaque élève
            Set prefsList = New Collection
            
            For j = 2 To wsEleves.Cells(i, wsEleves.Columns.Count).End(xlToLeft).Column
                prefsList.Add Trim(wsEleves.Cells(i, j).Value)
            Next j
            Set elevesPrefs(nomEleve) = prefsList
        End If
    Next i
    
    ' Lecture Projets
    For i = 2 To wsProjets.Cells(wsProjets.Rows.Count, "A").End(xlUp).Row
        Dim nomProjet As String: nomProjet = Trim(wsProjets.Cells(i, 1).Value)
        If nomProjet <> "" Then
            projetsCapacites(nomProjet) = Array(wsProjets.Cells(i, 2).Value, wsProjets.Cells(i, 3).Value)
            Set affectationsProjet(nomProjet) = CreateObject("Scripting.Dictionary")
            Dim rangsDict As Object: Set rangsDict = CreateObject("Scripting.Dictionary")
            For j = 4 To wsProjets.Cells(1, wsProjets.Columns.Count).End(xlToLeft).Column
                rangsDict(Trim(wsProjets.Cells(1, j).Value)) = wsProjets.Cells(i, j).Value
            Next j
            Set projetsPrefs(nomProjet) = rangsDict
        End If
    Next i
    
    ' --- 2. EXÉCUTION DE L'ALGORITHME (LOGIQUE STANDARD) ---
    While celibataires.Count > 0
        Dim eleveActuel As String: eleveActuel = celibataires(1)
        Dim indexProposition As Long: indexProposition = propositionsFaites(eleveActuel) + 1
        
        If indexProposition > elevesPrefs(eleveActuel).Count Then
            celibataires.Remove 1
            LogStep wsLog, ligneLog, eleveActuel & " a épuisé sa liste de vœux.", "Reste non affecté.", "", GetCelibatairesString(celibataires)
        Else
            Dim projetVise As String: projetVise = elevesPrefs(eleveActuel)(indexProposition)
            propositionsFaites(eleveActuel) = indexProposition
            
            Dim action As String: action = eleveActuel & " propose au " & projetVise & " (son choix n°" & indexProposition & ")."
            Dim decision As String, statut As String
            
            Dim affectesAuProjet As Object: Set affectesAuProjet = affectationsProjet(projetVise)
            Dim capaciteMax As Long: capaciteMax = projetsCapacites(projetVise)(1)

            If affectesAuProjet.Count < capaciteMax Then
                decision = "Le projet a de la place. ACCEPTATION PROVISOIRE."
                affectesAuProjet(eleveActuel) = projetsPrefs(projetVise)(eleveActuel)
                affectationEleve(eleveActuel) = projetVise
                celibataires.Remove 1
            Else
                Dim pireEleve As String: pireEleve = ""
                Dim rangPireEleve As Long: rangPireEleve = -1
                Dim eleveAffecte As Variant
                For Each eleveAffecte In affectesAuProjet.Keys
                    If affectesAuProjet(eleveAffecte) > rangPireEleve Then
                        rangPireEleve = affectesAuProjet(eleveAffecte)
                        pireEleve = eleveAffecte
                    End If
                Next eleveAffecte
                
                Dim rangEleveActuel As Long: rangEleveActuel = projetsPrefs(projetVise)(eleveActuel)
                
                If rangEleveActuel < rangPireEleve Then
                    decision = "Le projet est plein. " & eleveActuel & " est meilleur que " & pireEleve & ". ACCEPTATION et ÉVICTION."
                    affectesAuProjet.Remove pireEleve
                    affectationEleve(pireEleve) = ""
                    celibataires.Add pireEleve
                    affectesAuProjet(eleveActuel) = rangEleveActuel
                    affectationEleve(eleveActuel) = projetVise
                    celibataires.Remove 1
                Else
                    decision = "Le projet est plein. " & eleveActuel & " n'est pas meilleur. REJET."
                    celibataires.Remove 1
                    celibataires.Add eleveActuel
                End If
            End If
            
            statut = projetVise & ": " & Join(affectationsProjet(projetVise).Keys, ", ")
            LogStep wsLog, ligneLog, action, decision, statut, GetCelibatairesString(celibataires)
        End If
    Wend

    ' --- 3. ÉCRITURE DU RÉSULTAT FINAL ---
    wsResultats.Cells.ClearContents
    wsResultats.Range("A1:C1").Value = Array("Projet", "Élèves Affectés", "Statut Capacité")
    wsResultats.Range("A1:C1").Font.Bold = True
    Dim ligneResultat As Long: ligneResultat = 2
    Dim projet As Variant
    For Each projet In affectationsProjet.Keys
        wsResultats.Cells(ligneResultat, 1).Value = projet
        Dim listeEleves As String
        If affectationsProjet(projet).Count > 0 Then
            listeEleves = Join(affectationsProjet(projet).Keys, ", ")
        Else
            listeEleves = "Aucun"
        End If
        wsResultats.Cells(ligneResultat, 2).Value = listeEleves
        Dim statutCapacite As String, nbAffectes As Long, capMin As Long, capMax As Long
        nbAffectes = affectationsProjet(projet).Count
        capMin = projetsCapacites(projet)(0)
        capMax = projetsCapacites(projet)(1)
        If nbAffectes < capMin Then
            statutCapacite = "MINIMUM NON ATTEINT (" & nbAffectes & "/" & capMin & ")"
            wsResultats.Cells(ligneResultat, 3).Interior.Color = vbYellow
        Else
            statutCapacite = "OK (" & nbAffectes & "/" & capMax & ")"
            wsResultats.Cells(ligneResultat, 3).Interior.ColorIndex = xlNone
        End If
        wsResultats.Cells(ligneResultat, 3).Value = statutCapacite
        ligneResultat = ligneResultat + 1
    Next projet
    wsResultats.Columns.AutoFit

    wsLog.Columns.AutoFit
    Application.ScreenUpdating = True
    MsgBox "L'affectation pas à pas est terminée.", vbInformation

End Sub

' === Les procédures utilitaires ci-dessous n'ont pas changé et sont nécessaires ===
Private Sub LogStep(ByVal ws As Worksheet, ByRef ligne As Long, ByVal action As String, ByVal decision As String, ByVal statut As String, ByVal celibataires As String)
    ligne = ligne + 1
    ws.Cells(ligne, 1).Value = ligne - 1
    ws.Cells(ligne, 2).Value = action
    ws.Cells(ligne, 3).Value = decision
    ws.Cells(ligne, 4).Value = statut
    ws.Cells(ligne, 5).Value = celibataires
End Sub

Private Function GetCelibatairesString(ByVal celibatairesColl As Collection) As String
    If celibatairesColl.Count = 0 Then
        GetCelibatairesString = "Aucun"
        Exit Function
    End If
    Dim arr() As String, i As Long
    ReDim arr(1 To celibatairesColl.Count)
    For i = 1 To celibatairesColl.Count
        arr(i) = celibatairesColl(i)
    Next i
    GetCelibatairesString = Join(arr, ", ")
End Function


'****************************************************************************************
' MACRO DE DIAGNOSTIC (Version Finale et Définitivement Corrigée)
'****************************************************************************************
Sub DiagnostiqueAffectation()

    ' Déclaration des feuilles et objets
    Dim wsEleves As Worksheet, wsProjets As Worksheet, wsLog As Worksheet
    Dim elevesPrefs As Object, projetsPrefs As Object, celibataires As Collection
    
    ' Initialisation
    Set elevesPrefs = CreateObject("Scripting.Dictionary")
    Set celibataires = New Collection
    
    ' Association des feuilles
    On Error Resume Next
    Set wsEleves = ThisWorkbook.Sheets("Préférences_Élèves")
    Set wsProjets = ThisWorkbook.Sheets("Préférences_Projets")
    Set wsLog = ThisWorkbook.Sheets("Log_Affectation")
    If wsEleves Is Nothing Or wsProjets Is Nothing Or wsLog Is Nothing Then
        MsgBox "Erreur : Feuilles manquantes.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = False
    wsLog.Cells.Clear
    wsLog.Range("A1:C1").Value = Array("Élève", "Vérification Choix 1", "Vérification Choix 2")
    wsLog.Range("A1:C1").Font.Bold = True

    ' --- LECTURE DES DONNÉES (CORRIGÉE) ---
    Dim i As Long, j As Long
    For i = 2 To wsEleves.Cells(wsEleves.Rows.Count, "A").End(xlUp).Row
        Dim nomEleve As String: nomEleve = Trim(wsEleves.Cells(i, 1).Value)
        If nomEleve <> "" Then
            ' CORRECTION DÉFINITIVE : On déclare et crée la liste ICI, à l'intérieur de la boucle.
            ' Cela garantit une nouvelle liste indépendante pour chaque élève.
            Dim prefsList As New Collection
            
            celibataires.Add nomEleve
            
            For j = 2 To wsEleves.Cells(i, wsEleves.Columns.Count).End(xlToLeft).Column
                prefsList.Add Trim(wsEleves.Cells(i, j).Value)
            Next j
            
            Set elevesPrefs(nomEleve) = prefsList
        End If
    Next i
    
    ' --- VÉRIFICATION DES DONNÉES LUES ---
    Dim ligneLog As Long: ligneLog = 1
    Dim eleve As Variant
    For Each eleve In elevesPrefs.Keys
        ligneLog = ligneLog + 1
        wsLog.Cells(ligneLog, 1).Value = eleve
        
        ' On affiche dans le log les 2 premiers choix que la macro a en mémoire
        If elevesPrefs(eleve).Count >= 1 Then
            wsLog.Cells(ligneLog, 2).Value = elevesPrefs(eleve)(1)
        End If
        If elevesPrefs(eleve).Count >= 2 Then
            wsLog.Cells(ligneLog, 3).Value = elevesPrefs(eleve)(2)
        End If
    Next eleve
    
    wsLog.Columns.AutoFit
    Application.ScreenUpdating = True
    
    MsgBox "Le diagnostic de lecture est terminé. Veuillez vérifier la feuille 'Log_Affectation' pour voir les préférences que la macro a réellement lues. Si elles sont correctes, l'algorithme fonctionnera.", vbInformation

End Sub

'****************************************************************************************
' MACRO : Créer un rapport de satisfaction (Version Finale et Définitive)
'****************************************************************************************
Sub CreerRapportSatisfaction()

    ' Déclaration des feuilles et des objets
    Dim wsResultats As Worksheet, wsPrefsEleves As Worksheet, wsRapport As Worksheet
    Dim affectations As Object, prefsEleves As Object
    Dim rangsObtenus As Collection
    
    ' Déclaration des variables de boucle et de travail
    Dim i As Long, j As Long
    Dim eleve As Variant
    
    ' Initialisation
    Set affectations = CreateObject("Scripting.Dictionary")
    Set prefsEleves = CreateObject("Scripting.Dictionary")
    Set rangsObtenus = New Collection

    ' --- Association des feuilles ---
    On Error Resume Next
    Set wsResultats = ThisWorkbook.Sheets("Résultats")
    Set wsPrefsEleves = ThisWorkbook.Sheets("Préférences_Élèves")
    Set wsRapport = ThisWorkbook.Sheets("Rapport_Satisfaction")
    If wsResultats Is Nothing Or wsPrefsEleves Is Nothing Or wsRapport Is Nothing Then
        MsgBox "Erreur : Assurez-vous que les feuilles 'Résultats', 'Préférences_Élèves' et 'Rapport_Satisfaction' existent.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = False

    ' --- ÉTAPE 1 : Lire les résultats de l'affectation ---
    For i = 2 To wsResultats.Cells(wsResultats.Rows.Count, "A").End(xlUp).Row
        Dim projetAffecte As String: projetAffecte = Trim(wsResultats.Cells(i, 1).Value)
        Dim listeElevesStr As String: listeElevesStr = Trim(wsResultats.Cells(i, 2).Value)
        
        If listeElevesStr <> "Aucun" And listeElevesStr <> "" Then
            Dim elevesAffectesArray As Variant: elevesAffectesArray = Split(listeElevesStr, ", ")
            For Each eleve In elevesAffectesArray
                affectations(Trim(eleve)) = projetAffecte
            Next eleve
        End If
    Next i
    
    ' --- ÉTAPE 2 : Lire les préférences de chaque élève (Nouvelle Logique Fiable) ---
    For i = 2 To wsPrefsEleves.Cells(wsPrefsEleves.Rows.Count, "A").End(xlUp).Row
        Dim nomEleve As String: nomEleve = Trim(wsPrefsEleves.Cells(i, 1).Value)
        If nomEleve <> "" Then
            ' On crée la collection DIRECTEMENT dans le dictionnaire pour cet élève
            Set prefsEleves(nomEleve) = New Collection
            
            ' On remplit la collection qui est DÉJÀ dans le dictionnaire
            For j = 2 To wsPrefsEleves.Cells(i, wsPrefsEleves.Columns.Count).End(xlToLeft).Column
                prefsEleves(nomEleve).Add Trim(wsPrefsEleves.Cells(i, j).Value)
            Next j
        End If
    Next i
    
    ' --- ÉTAPE 3 : Générer le rapport ---
    wsRapport.Cells.ClearContents
    wsRapport.Range("A1:D1").Value = Array("Élève", "Projet Affecté", "Rang du Choix", "Diagnostic")
    wsRapport.Range("A1:D1").Font.Bold = True
    
    Dim ligneRapport As Long: ligneRapport = 1
    
    For Each eleve In prefsEleves.Keys
        ligneRapport = ligneRapport + 1
        wsRapport.Cells(ligneRapport, 1).Value = eleve
        
        If affectations.Exists(eleve) Then
            Dim projetObtenu As String: projetObtenu = affectations(eleve)
            wsRapport.Cells(ligneRapport, 2).Value = projetObtenu
            
            Dim listeDeChoix As Collection: Set listeDeChoix = prefsEleves(eleve)
            Dim rang As Long: rang = 0
            
            For j = 1 To listeDeChoix.Count
                If StrComp(listeDeChoix(j), projetObtenu, vbTextCompare) = 0 Then
                    rang = j
                    Exit For
                End If
            Next j
            
            If rang > 0 Then
                wsRapport.Cells(ligneRapport, 3).Value = rang
                rangsObtenus.Add rang
                wsRapport.Cells(ligneRapport, 4).Value = "OK"
            Else
                wsRapport.Cells(ligneRapport, 3).Value = "Erreur"
                wsRapport.Cells(ligneRapport, 4).Value = "Projet '" & projetObtenu & "' non trouvé dans les vœux."
            End If
            
        Else
            wsRapport.Cells(ligneRapport, 2).Value = "Non affecté"
            wsRapport.Cells(ligneRapport, 3).Value = "N/A"
        End If
    Next eleve
    
    ' --- ÉTAPE 4 : Calculer et afficher les statistiques ---
    If rangsObtenus.Count > 0 Then
        Dim sommeRangs As Double: sommeRangs = 0
        Dim rangItem As Variant
        For Each rangItem In rangsObtenus
            sommeRangs = sommeRangs + rangItem
        Next rangItem
        Dim rangMoyen As Double: rangMoyen = sommeRangs / rangsObtenus.Count
        
        Dim sommeCarresEcarts As Double: sommeCarresEcarts = 0
        For Each rangItem In rangsObtenus
            sommeCarresEcarts = sommeCarresEcarts + (rangItem - rangMoyen) ^ 2
        Next rangItem
        Dim ecartType As Double
        If rangsObtenus.Count > 1 Then
            ecartType = Sqr(sommeCarresEcarts / (rangsObtenus.Count - 1))
        Else
            ecartType = 0
        End If

        ligneRapport = ligneRapport + 2
        wsRapport.Cells(ligneRapport, "B").Value = "Satisfaction moyenne :"
        wsRapport.Cells(ligneRapport, "C").Value = Round(rangMoyen, 2)
        
        ligneRapport = ligneRapport + 1
        wsRapport.Cells(ligneRapport, "B").Value = "Écart-type des rangs :"
        wsRapport.Cells(ligneRapport, "C").Value = Round(ecartType, 2)
        
        wsRapport.Range(wsRapport.Cells(ligneRapport - 1, "B"), wsRapport.Cells(ligneRapport, "C")).Font.Bold = True
    End If
    
    wsRapport.Columns.AutoFit
    Application.ScreenUpdating = True
    
    MsgBox "Le rapport de satisfaction des élèves a été généré.", vbInformation

End Sub

'****************************************************************************************
' MACRO : Rapport complet par projet avec double analyse de satisfaction (élèves et projet)
'****************************************************************************************
Sub CreerRapportProjets()

    ' --- Déclarations ---
    Dim wsResultats As Worksheet, wsPrefsEleves As Worksheet, wsPrefsProjets As Worksheet, wsRapportProjet As Worksheet
    Dim affectationsProjet As Object, prefsEleves As Object, projetsPrefs As Object
    
    ' --- Initialisation ---
    Set affectationsProjet = CreateObject("Scripting.Dictionary")
    Set prefsEleves = CreateObject("Scripting.Dictionary")
    Set projetsPrefs = CreateObject("Scripting.Dictionary")

    ' --- Association des feuilles ---
    On Error Resume Next
    Set wsResultats = ThisWorkbook.Sheets("Résultats")
    Set wsPrefsEleves = ThisWorkbook.Sheets("Préférences_Élèves")
    Set wsPrefsProjets = ThisWorkbook.Sheets("Préférences_Projets")
    Set wsRapportProjet = ThisWorkbook.Sheets("Affectations_par_Projet")
    If wsResultats Is Nothing Or wsPrefsEleves Is Nothing Or wsPrefsProjets Is Nothing Or wsRapportProjet Is Nothing Then
        MsgBox "Erreur : Assurez-vous que toutes les feuilles nécessaires existent.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = False

    ' --- ÉTAPE 1 : Lire toutes les données nécessaires ---
    Dim i As Long, j As Long, eleve As Variant, projet As Variant
    
    ' Lecture des affectations
    For i = 2 To wsResultats.Cells(wsResultats.Rows.Count, "A").End(xlUp).Row
        Dim projetAffecte As String: projetAffecte = Trim(wsResultats.Cells(i, 1).Value)
        Dim listeElevesStr As String: listeElevesStr = Trim(wsResultats.Cells(i, 2).Value)
        Set affectationsProjet(projetAffecte) = New Collection
        If listeElevesStr <> "Aucun" And listeElevesStr <> "" Then
            Dim elevesAffectesArray As Variant: elevesAffectesArray = Split(listeElevesStr, ", ")
            For Each eleve In elevesAffectesArray
                affectationsProjet(projetAffecte).Add Trim(eleve)
            Next eleve
        End If
    Next i
    
    ' Lecture des préférences des élèves
    For i = 2 To wsPrefsEleves.Cells(wsPrefsEleves.Rows.Count, "A").End(xlUp).Row
        Dim nomEleve As String: nomEleve = Trim(wsPrefsEleves.Cells(i, 1).Value)
        If nomEleve <> "" Then
            Set prefsEleves(nomEleve) = New Collection
            For j = 2 To wsPrefsEleves.Cells(i, wsPrefsEleves.Columns.Count).End(xlToLeft).Column
                prefsEleves(nomEleve).Add Trim(wsPrefsEleves.Cells(i, j).Value)
            Next j
        End If
    Next i
    
    ' Lecture des préférences des projets
    For i = 2 To wsPrefsProjets.Cells(wsPrefsProjets.Rows.Count, "A").End(xlUp).Row
        Dim nomProjet As String: nomProjet = Trim(wsPrefsProjets.Cells(i, 1).Value)
        If nomProjet <> "" Then
            Set projetsPrefs(nomProjet) = CreateObject("Scripting.Dictionary")
            For j = 4 To wsPrefsProjets.Cells(1, wsPrefsProjets.Columns.Count).End(xlToLeft).Column
                projetsPrefs(nomProjet)(Trim(wsPrefsProjets.Cells(1, j).Value)) = wsPrefsProjets.Cells(i, j).Value
            Next j
        End If
    Next i
    
    ' --- ÉTAPE 2 : Générer le rapport ---
    wsRapportProjet.Cells.Clear
    wsRapportProjet.Range("A1:E1").Value = Array("Projet", "Moy. Satisfaction Élèves", "Écart-type Satis. Élèves", "Moy. Qualité Équipe", "Écart-type Qualité Équipe")
    wsRapportProjet.Range("A1:E1").Font.Bold = True
    
    Dim ligneRapport As Long: ligneRapport = 1
    
    For Each projet In affectationsProjet.Keys
        ligneRapport = ligneRapport + 1
        wsRapportProjet.Cells(ligneRapport, 1).Value = projet
        
        Dim listeEleves As Collection: Set listeEleves = affectationsProjet(projet)
        Dim nbEleves As Long: nbEleves = listeEleves.Count
        
        If nbEleves > 0 Then
            Dim rangsSatisEleves As New Collection
            Dim rangsQualiteProjet As New Collection
            
            ' Pour chaque élève du projet, trouver les deux types de rangs
            For Each eleve In listeEleves
                ' 1. Satisfaction de l'élève
                If prefsEleves.Exists(eleve) Then
                    Dim rangSatis As Long: rangSatis = 0
                    Dim listeDeChoix As Collection: Set listeDeChoix = prefsEleves(eleve)
                    For j = 1 To listeDeChoix.Count
                        If StrComp(listeDeChoix(j), projet, vbTextCompare) = 0 Then
                            rangSatis = j
                            Exit For
                        End If
                    Next j
                    If rangSatis > 0 Then rangsSatisEleves.Add rangSatis
                End If
                
                ' 2. "Qualité" de l'élève pour le projet
                If projetsPrefs.Exists(projet) And projetsPrefs(projet).Exists(eleve) Then
                    rangsQualiteProjet.Add projetsPrefs(projet)(eleve)
                End If
            Next eleve
            
            ' Calculer et afficher les statistiques de satisfaction des élèves
            Dim moyenneSatis As Double, ecartTypeSatis As Double
            If rangsSatisEleves.Count > 0 Then
                moyenneSatis = Application.WorksheetFunction.Average(CollectionToArray(rangsSatisEleves))
                If rangsSatisEleves.Count > 1 Then ecartTypeSatis = Application.WorksheetFunction.StDev_S(CollectionToArray(rangsSatisEleves))
            End If
            wsRapportProjet.Cells(ligneRapport, 2).Value = Round(moyenneSatis, 2)
            wsRapportProjet.Cells(ligneRapport, 3).Value = Round(ecartTypeSatis, 2)
            
            ' Calculer et afficher les statistiques de qualité de l'équipe
            Dim moyenneQualite As Double, ecartTypeQualite As Double
            If rangsQualiteProjet.Count > 0 Then
                moyenneQualite = Application.WorksheetFunction.Average(CollectionToArray(rangsQualiteProjet))
                If rangsQualiteProjet.Count > 1 Then ecartTypeQualite = Application.WorksheetFunction.StDev_S(CollectionToArray(rangsQualiteProjet))
            End If
            wsRapportProjet.Cells(ligneRapport, 4).Value = Round(moyenneQualite, 2)
            wsRapportProjet.Cells(ligneRapport, 5).Value = Round(ecartTypeQualite, 2)
            
            ' Afficher la liste des élèves dans les colonnes suivantes
            For j = 1 To nbEleves
                If j = 1 Then
                     wsRapportProjet.Cells(1, 6).Value = "Membres de l'équipe ->"
                     wsRapportProjet.Cells(1, 6).Font.Bold = True
                End If
                wsRapportProjet.Cells(ligneRapport, 5 + j).Value = listeEleves(j)
            Next j
        Else
            ' Si le projet est vide, mettre N/A partout
            wsRapportProjet.Range("B" & ligneRapport & ":E" & ligneRapport).Value = "N/A"
        End If
    Next projet
    
    wsRapportProjet.Columns.AutoFit
    Application.ScreenUpdating = True
    
    MsgBox "Le rapport complet par projet a été généré.", vbInformation

End Sub

' Fonction utilitaire pour convertir une Collection en Array pour les WorksheetFunctions
Private Function CollectionToArray(coll As Collection) As Variant
    Dim arr() As Variant
    ReDim arr(1 To coll.Count)
    Dim i As Long
    For i = 1 To coll.Count
        arr(i) = coll(i)
    Next i
    CollectionToArray = arr
End Function

'****************************************************************************************
' MACRO : Mesurer la performance (Version Finale Corrigée)
'****************************************************************************************
Sub BilanPerformanceAlgorithme()

    ' --- Déclarations ---
    Dim wsResultats As Worksheet, wsPrefsEleves As Worksheet, wsPrefsProjets As Worksheet, wsBilan As Worksheet, wsDetails As Worksheet
    Dim affectationsEleve As Object, affectationsProjet As Object, prefsEleves As Object, projetsCapacites As Object
    Dim rangsObtenus As Collection
    Dim projetsSousMinimum As Collection, projetsVides As Collection, elevesSansProjet As Collection
    
    ' --- Initialisation ---
    Set affectationsEleve = CreateObject("Scripting.Dictionary")
    Set affectationsProjet = CreateObject("Scripting.Dictionary")
    Set prefsEleves = CreateObject("Scripting.Dictionary")
    Set projetsCapacites = CreateObject("Scripting.Dictionary")
    Set rangsObtenus = New Collection
    Set projetsSousMinimum = New Collection
    Set projetsVides = New Collection
    Set elevesSansProjet = New Collection

    ' --- Association des feuilles ---
    On Error Resume Next
    Set wsResultats = ThisWorkbook.Sheets("Résultats")
    Set wsPrefsEleves = ThisWorkbook.Sheets("Préférences_Élèves")
    Set wsPrefsProjets = ThisWorkbook.Sheets("Préférences_Projets")
    Set wsBilan = ThisWorkbook.Sheets("Bilan_Performance")
    Set wsDetails = ThisWorkbook.Sheets("Details_Suivi")
    If wsDetails Is Nothing Then
        Set wsDetails = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsDetails.Name = "Details_Suivi"
    End If
    On Error GoTo 0
    
    If wsResultats Is Nothing Or wsPrefsEleves Is Nothing Or wsPrefsProjets Is Nothing Or wsBilan Is Nothing Then
        MsgBox "Erreur : Une ou plusieurs feuilles de base sont manquantes.", vbCritical
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    wsBilan.Cells.Clear
    wsDetails.Cells.Clear
    wsBilan.Columns("A").ColumnWidth = 35
    wsBilan.Columns("B").ColumnWidth = 15
    wsDetails.Columns("A").ColumnWidth = 35
    wsDetails.Columns("B").ColumnWidth = 35

    ' --- ÉTAPE 1 : Lire toutes les données ---
    Dim i As Long, j As Long, eleve As Variant, projet As Variant
    For i = 2 To wsResultats.Cells(wsResultats.Rows.Count, "A").End(xlUp).Row
        Dim projetAffecte As String: projetAffecte = Trim(wsResultats.Cells(i, 1).Value)
        Dim listeElevesStr As String: listeElevesStr = Trim(wsResultats.Cells(i, 2).Value)
        Set affectationsProjet(projetAffecte) = New Collection
        If listeElevesStr <> "Aucun" And listeElevesStr <> "" Then
            Dim elevesAffectesArray As Variant: elevesAffectesArray = Split(listeElevesStr, ", ")
            For Each eleve In elevesAffectesArray
                affectationsEleve(Trim(eleve)) = projetAffecte
                affectationsProjet(projetAffecte).Add Trim(eleve)
            Next eleve
        End If
    Next i
    For i = 2 To wsPrefsEleves.Cells(wsPrefsEleves.Rows.Count, "A").End(xlUp).Row
        Dim nomEleve As String: nomEleve = Trim(wsPrefsEleves.Cells(i, 1).Value)
        If nomEleve <> "" Then
            Set prefsEleves(nomEleve) = New Collection
            For j = 2 To wsPrefsEleves.Cells(i, wsPrefsEleves.Columns.Count).End(xlToLeft).Column
                prefsEleves(nomEleve).Add Trim(wsPrefsEleves.Cells(i, j).Value)
            Next j
        End If
    Next i
    For i = 2 To wsPrefsProjets.Cells(wsPrefsProjets.Rows.Count, "A").End(xlUp).Row
        Dim nomProjet As String: nomProjet = Trim(wsPrefsProjets.Cells(i, 1).Value)
        If nomProjet <> "" Then
            projetsCapacites(nomProjet) = Array(wsPrefsProjets.Cells(i, 2).Value, wsPrefsProjets.Cells(i, 3).Value)
        End If
    Next i

    ' --- ÉTAPE 2 : Calculer les métriques ---
    Dim nbElevesTotal As Long: nbElevesTotal = prefsEleves.Count
    Dim nbElevesAffectes As Long: nbElevesAffectes = affectationsEleve.Count
    Dim nbChoix1 As Long, nbChoix2 As Long, nbChoix3 As Long
    For Each eleve In prefsEleves.Keys
        If affectationsEleve.Exists(eleve) Then
            Dim rang As Long: rang = 0
            Dim listeDeChoix As Collection: Set listeDeChoix = prefsEleves(eleve)
            For j = 1 To listeDeChoix.Count
                If StrComp(listeDeChoix(j), affectationsEleve(eleve), vbTextCompare) = 0 Then
                    rang = j
                    Exit For
                End If
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
            elevesSansProjet.Add eleve
        End If
    Next eleve
    Dim nbProjetsMinAtteint As Long: nbProjetsMinAtteint = 0
    Dim totalPlacesRemplies As Long: totalPlacesRemplies = nbElevesAffectes
    Dim totalCapaciteMax As Long: totalCapaciteMax = 0
    For Each projet In projetsCapacites.Keys
        totalCapaciteMax = totalCapaciteMax + projetsCapacites(projet)(1)
        Dim nbAffectesSurProjet As Long
        If affectationsProjet.Exists(projet) Then
            nbAffectesSurProjet = affectationsProjet(projet).Count
        Else
            Set affectationsProjet(projet) = New Collection
        End If
        If nbAffectesSurProjet = 0 Then
            projetsVides.Add projet
        End If
        If nbAffectesSurProjet < projetsCapacites(projet)(0) Then
            projetsSousMinimum.Add projet
        Else
            nbProjetsMinAtteint = nbProjetsMinAtteint + 1
        End If
    Next projet
    Dim rangMoyen As Double, ecartType As Double
    If rangsObtenus.Count > 0 Then
        Dim sommeRangs As Double: sommeRangs = 0
        Dim rangItem As Variant
        For Each rangItem In rangsObtenus
            sommeRangs = sommeRangs + rangItem
        Next rangItem
        rangMoyen = sommeRangs / rangsObtenus.Count
        If rangsObtenus.Count > 1 Then
            Dim sommeCarresEcarts As Double: sommeCarresEcarts = 0
            For Each rangItem In rangsObtenus
                sommeCarresEcarts = sommeCarresEcarts + (rangItem - rangMoyen) ^ 2
            Next rangItem
            ecartType = Sqr(sommeCarresEcarts / (rangsObtenus.Count - 1))
        End If
    End If
    
    ' --- ÉTAPE 3 : Écrire le Bilan Synthétique et le lien ---
    Dim ligne As Long: ligne = 1
    wsBilan.Cells(ligne, "A").Value = "BILAN DU POINT DE VUE DES ÉLÈVES": wsBilan.Range("A" & ligne & ":B" & ligne).Merge: wsBilan.Range("A" & ligne).Font.Bold = True
    ligne = ligne + 2
    wsBilan.Cells(ligne, "A").Value = "Nombre total d'élèves": wsBilan.Cells(ligne, "B").Value = nbElevesTotal
    ligne = ligne + 1
    wsBilan.Cells(ligne, "A").Value = "Nombre d'élèves affectés": wsBilan.Cells(ligne, "B").Value = nbElevesAffectes
    ligne = ligne + 1
    wsBilan.Cells(ligne, "A").Value = "Nombre d'élèves non affectés": wsBilan.Cells(ligne, "B").Value = elevesSansProjet.Count
    ligne = ligne + 1
    If nbElevesTotal > 0 Then wsBilan.Cells(ligne, "A").Value = "Taux d'affectation": wsBilan.Cells(ligne, "B").Value = Format(nbElevesAffectes / nbElevesTotal, "0.0%")
    ligne = ligne + 2
    wsBilan.Cells(ligne, "A").Value = "Satisfaction moyenne (rang du vœu)": wsBilan.Cells(ligne, "B").Value = Round(rangMoyen, 2)
    ligne = ligne + 1
    wsBilan.Cells(ligne, "A").Value = "Écart-type des rangs (équité)": wsBilan.Cells(ligne, "B").Value = Round(ecartType, 2)
    ligne = ligne + 2
    wsBilan.Cells(ligne, "A").Value = "Élèves ayant obtenu leur 1er vœu": wsBilan.Cells(ligne, "B").Value = nbChoix1
    ligne = ligne + 1
    wsBilan.Cells(ligne, "A").Value = "Élèves ayant obtenu leur 2ème vœu": wsBilan.Cells(ligne, "B").Value = nbChoix2
    ligne = ligne + 1
    wsBilan.Cells(ligne, "A").Value = "Élèves ayant obtenu leur 3ème vœu": wsBilan.Cells(ligne, "B").Value = nbChoix3
    ligne = ligne + 3
    wsBilan.Cells(ligne, "A").Value = "BILAN DU POINT DE VUE DES PROJETS": wsBilan.Range("A" & ligne & ":B" & ligne).Merge: wsBilan.Range("A" & ligne).Font.Bold = True
    ligne = ligne + 2
    wsBilan.Cells(ligne, "A").Value = "Nombre total de projets": wsBilan.Cells(ligne, "B").Value = projetsCapacites.Count
    ligne = ligne + 1
    wsBilan.Cells(ligne, "A").Value = "Projets ayant atteint leur capacité Min.": wsBilan.Cells(ligne, "B").Value = nbProjetsMinAtteint
    ligne = ligne + 1
    If projetsCapacites.Count > 0 Then wsBilan.Cells(ligne, "A").Value = "Taux de projets 'satisfaits'": wsBilan.Cells(ligne, "B").Value = Format(nbProjetsMinAtteint / projetsCapacites.Count, "0.0%")
    ligne = ligne + 2
    wsBilan.Cells(ligne, "A").Value = "Nombre total de places remplies": wsBilan.Cells(ligne, "B").Value = totalPlacesRemplies
    ligne = ligne + 1
    wsBilan.Cells(ligne, "A").Value = "Capacité d'accueil maximale totale": wsBilan.Cells(ligne, "B").Value = totalCapaciteMax
    ligne = ligne + 1
    If totalCapaciteMax > 0 Then wsBilan.Cells(ligne, "A").Value = "Taux d'occupation global": wsBilan.Cells(ligne, "B").Value = Format(totalPlacesRemplies / totalCapaciteMax, "0.0%")
    wsBilan.Range("A1:B" & ligne).Borders.Weight = xlThin
    ligne = ligne + 2
    wsBilan.Hyperlinks.Add Anchor:=wsBilan.Cells(ligne, "A"), Address:="", SubAddress:="'Details_Suivi'!A1", TextToDisplay:="Cliquer ici pour voir les listes de suivi détaillées"
    wsBilan.Cells(ligne, "A").Font.Underline = xlUnderlineStyleSingle
    wsBilan.Cells(ligne, "A").Font.Color = vbBlue

    ' --- ÉTAPE 4 : Écrire les listes de suivi ---
    ' CORRECTION : La déclaration en double a été supprimée d'ici.
    ' La variable wsDetails est déjà déclarée et définie en haut de la macro.
    wsDetails.Cells.Clear
    Dim ligneDetails As Long: ligneDetails = 1
    wsDetails.Columns("A").ColumnWidth = 35
    wsDetails.Columns("B").ColumnWidth = 35
    wsDetails.Cells(ligneDetails, "A").Value = "LISTES DE SUIVI DÉTAILLÉES"
    wsDetails.Cells(ligneDetails, "A").Font.Bold = True
    ligneDetails = ligneDetails + 2
    wsDetails.Cells(ligneDetails, "A").Value = "Élèves sans projet :"
    wsDetails.Cells(ligneDetails, "A").Font.Bold = True
    If elevesSansProjet.Count > 0 Then
        For i = 1 To elevesSansProjet.Count
            wsDetails.Cells(ligneDetails + i - 1, "B").Value = elevesSansProjet(i)
        Next i
        ligneDetails = ligneDetails + elevesSansProjet.Count
    Else
        wsDetails.Cells(ligneDetails, "B").Value = "Aucun"
        ligneDetails = ligneDetails + 1
    End If
    ligneDetails = ligneDetails + 1
    wsDetails.Cells(ligneDetails, "A").Value = "Projets n'atteignant pas leur minimum :"
    wsDetails.Cells(ligneDetails, "A").Font.Bold = True
    If projetsSousMinimum.Count > 0 Then
        For i = 1 To projetsSousMinimum.Count
            wsDetails.Cells(ligneDetails + i - 1, "B").Value = projetsSousMinimum(i)
        Next i
        ligneDetails = ligneDetails + projetsSousMinimum.Count
    Else
        wsDetails.Cells(ligneDetails, "B").Value = "Aucun"
        ligneDetails = ligneDetails + 1
    End If
    ligneDetails = ligneDetails + 1
    wsDetails.Cells(ligneDetails, "A").Value = "Projets sans aucun élève :"
    wsDetails.Cells(ligneDetails, "A").Font.Bold = True
    If projetsVides.Count > 0 Then
        For i = 1 To projetsVides.Count
            wsDetails.Cells(ligneDetails + i - 1, "B").Value = projetsVides(i)
        Next i
    Else
        wsDetails.Cells(ligneDetails, "B").Value = "Aucun"
    End If

    ' --- Finalisation ---
    CreerDashboard
    ThisWorkbook.Sheets("Dashboard").Activate
    Application.ScreenUpdating = True
    MsgBox "Le bilan de performance, les listes de suivi et le tableau de bord ont été générés avec succès.", vbInformation

End Sub

'****************************************************************************************
' MACRO DE VÉRIFICATION FINALE (Version 2, Fiable)
' Lit et affiche les données brutes pour "Élève 2" de manière robuste.
'****************************************************************************************
Sub VerificationFinale_V2()

    Dim wsResultats As Worksheet, wsPrefsEleves As Worksheet, wsLog As Worksheet
    
    On Error Resume Next
    Set wsResultats = ThisWorkbook.Sheets("Résultats")
    Set wsPrefsEleves = ThisWorkbook.Sheets("Préférences_Élèves")
    Set wsLog = ThisWorkbook.Sheets("Log_Affectation")
    If wsResultats Is Nothing Or wsPrefsEleves Is Nothing Or wsLog Is Nothing Then
        MsgBox "Erreur : Feuilles manquantes.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = False
    wsLog.Cells.Clear
    wsLog.Range("A1:Z1").Font.Bold = True
    
    ' --- ÉTAPE 1 : Trouver le projet de "Élève 2" (MÉTHODE FIABLE) ---
    Dim projetObtenu As String: projetObtenu = "Non trouvé"
    Dim i As Long
    For i = 2 To wsResultats.Cells(wsResultats.Rows.Count, "A").End(xlUp).Row
        Dim listeElevesStr As String: listeElevesStr = Trim(wsResultats.Cells(i, 2).Value)
        
        If listeElevesStr <> "" And listeElevesStr <> "Aucun" Then
            Dim elevesArray As Variant: elevesArray = Split(listeElevesStr, ", ")
            Dim eleve As Variant
            For Each eleve In elevesArray
                If Trim(eleve) = "Élève 2" Then
                    projetObtenu = Trim(wsResultats.Cells(i, 1).Value)
                    Exit For ' Sort de la boucle des élèves
                End If
            Next eleve
        End If
        
        If projetObtenu <> "Non trouvé" Then Exit For ' Sort de la boucle des projets
    Next i
    
    ' --- ÉTAPE 2 : Lire la liste de vœux de "Élève 2" ---
    Dim ligneEleve As Long: ligneEleve = 0
    For i = 2 To wsPrefsEleves.Cells(wsPrefsEleves.Rows.Count, "A").End(xlUp).Row
        If Trim(wsPrefsEleves.Cells(i, "A").Value) = "Élève 2" Then
            ligneEleve = i
            Exit For
        End If
    Next i
    
    ' --- ÉTAPE 3 : Afficher les résultats bruts dans le journal ---
    wsLog.Cells(1, 1).Value = "Vérification Finale V2 pour 'Élève 2'"
    wsLog.Cells(2, 1).Value = "Projet affecté (lu depuis 'Résultats'):"
    wsLog.Cells(2, 2).Value = projetObtenu
    
    wsLog.Cells(4, 1).Value = "Liste de vœux (lue depuis 'Préférences_Élèves'):"
    If ligneEleve > 0 Then
        Dim j As Long
        For j = 2 To wsPrefsEleves.Cells(ligneEleve, wsPrefsEleves.Columns.Count).End(xlToLeft).Column
            wsLog.Cells(4, j).Value = "Vœu " & (j - 1)
            wsLog.Cells(5, j).Value = Trim(wsPrefsEleves.Cells(ligneEleve, j).Value)
        Next j
    Else
        wsLog.Cells(5, 2).Value = "'Élève 2' non trouvé dans la feuille des préférences."
    End If
    
    wsLog.Columns.AutoFit
    Application.ScreenUpdating = True
    MsgBox "Vérification V2 terminée. Veuillez inspecter la feuille 'Log_Affectation'."

End Sub






'****************************************************************************************
' MACRO DE DIAGNOSTIC FINAL : Calcul du rang pour "Élève 2"
'****************************************************************************************
Sub DebugRang()

    Dim wsResultats As Worksheet, wsPrefsEleves As Worksheet, wsLog As Worksheet
    
    ' --- Initialisation et association des feuilles ---
    On Error Resume Next
    Set wsResultats = ThisWorkbook.Sheets("Résultats")
    Set wsPrefsEleves = ThisWorkbook.Sheets("Préférences_Élèves")
    Set wsLog = ThisWorkbook.Sheets("Log_Affectation")
    If wsResultats Is Nothing Or wsPrefsEleves Is Nothing Or wsLog Is Nothing Then
        MsgBox "Erreur : Feuilles manquantes.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = False
    wsLog.Cells.Clear
    wsLog.Range("A1:C1").Value = Array("Étape", "Information", "Résultat de la Vérification")
    wsLog.Range("A1:C1").Font.Bold = True
    Dim ligneLog As Long: ligneLog = 1

    ' --- ÉTAPE 1 : Trouver le projet affecté à "Élève 2" ---
    ligneLog = ligneLog + 1
    wsLog.Cells(ligneLog, 1).Value = "1. Recherche du projet affecté"
    
    Dim projetObtenu As String: projetObtenu = "Non trouvé"
    Dim i As Long
    For i = 2 To wsResultats.Cells(wsResultats.Rows.Count, "A").End(xlUp).Row
        Dim listeElevesStr As String: listeElevesStr = Trim(wsResultats.Cells(i, 2).Value)
        If listeElevesStr <> "" And listeElevesStr <> "Aucun" Then
            Dim elevesArray As Variant: elevesArray = Split(listeElevesStr, ", ")
            Dim eleve As Variant
            For Each eleve In elevesArray
                If Trim(eleve) = "Élève 2" Then
                    projetObtenu = Trim(wsResultats.Cells(i, 1).Value)
                    Exit For
                End If
            Next eleve
        End If
        If projetObtenu <> "Non trouvé" Then Exit For
    Next i
    wsLog.Cells(ligneLog, 2).Value = "Le projet lu dans 'Résultats' pour 'Élève 2' est :"
    wsLog.Cells(ligneLog, 3).Value = projetObtenu
    wsLog.Cells(ligneLog, 3).Font.Bold = True

    ' --- ÉTAPE 2 : Lire la liste de vœux de "Élève 2" ---
    ligneLog = ligneLog + 2
    wsLog.Cells(ligneLog, 1).Value = "2. Recherche du rang dans les préférences"
    
    Dim ligneEleve As Long: ligneEleve = 0
    For i = 2 To wsPrefsEleves.Cells(wsPrefsEleves.Rows.Count, "A").End(xlUp).Row
        If Trim(wsPrefsEleves.Cells(i, "A").Value) = "Élève 2" Then
            ligneEleve = i
            Exit For
        End If
    Next i

    ' --- ÉTAPE 3 : Comparer le projet affecté avec chaque vœu ---
    If ligneEleve > 0 And projetObtenu <> "Non trouvé" Then
        Dim j As Long, rangFinal As Long: rangFinal = 0
        For j = 2 To wsPrefsEleves.Cells(ligneEleve, wsPrefsEleves.Columns.Count).End(xlToLeft).Column
            ligneLog = ligneLog + 1
            Dim voeuActuel As String: voeuActuel = Trim(wsPrefsEleves.Cells(ligneEleve, j).Value)
            wsLog.Cells(ligneLog, 1).Value = "Vœu n°" & (j - 1)
            wsLog.Cells(ligneLog, 2).Value = "Est-ce que '" & voeuActuel & "' = '" & projetObtenu & "' ?"
            
            If StrComp(voeuActuel, projetObtenu, vbTextCompare) = 0 Then
                wsLog.Cells(ligneLog, 3).Value = "OUI"
                wsLog.Range("A" & ligneLog & ":C" & ligneLog).Interior.Color = vbGreen
                rangFinal = j - 1
                Exit For ' On a trouvé, on arrête
            Else
                wsLog.Cells(ligneLog, 3).Value = "NON"
            End If
        Next j
        
        ligneLog = ligneLog + 2
        wsLog.Cells(ligneLog, 1).Value = "3. Conclusion du calcul"
        wsLog.Cells(ligneLog, 2).Value = "Le rang calculé est :"
        wsLog.Cells(ligneLog, 3).Value = rangFinal
        wsLog.Cells(ligneLog, 3).Font.Bold = True
        
    Else
        ligneLog = ligneLog + 1
        wsLog.Cells(ligneLog, 2).Value = "Impossible de calculer le rang car les données de base sont introuvables."
    End If
    
    wsLog.Columns.AutoFit
    Application.ScreenUpdating = True
    MsgBox "Le diagnostic du calcul de rang est terminé. Veuillez inspecter la feuille 'Log_Affectation'."

End Sub

'****************************************************************************************
' MACRO DE VÉRIFICATION FINALE (Nouvelle Approche de Lecture)
'****************************************************************************************
Sub VerifierToutesLesDonnees()

    ' --- Déclarations ---
    Dim wsResultats As Worksheet, wsPrefsEleves As Worksheet, wsVerif As Worksheet
    Dim affectations As Object, prefsEleves As Object
    
    ' --- Initialisation ---
    Set affectations = CreateObject("Scripting.Dictionary")
    Set prefsEleves = CreateObject("Scripting.Dictionary")

    ' --- Association des feuilles ---
    On Error Resume Next
    Set wsResultats = ThisWorkbook.Sheets("Résultats")
    Set wsPrefsEleves = ThisWorkbook.Sheets("Préférences_Élèves")
    ' Crée ou vide la feuille de vérification
    Set wsVerif = ThisWorkbook.Sheets("Verification_Finale")
    If wsVerif Is Nothing Then
        Set wsVerif = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsVerif.Name = "Verification_Finale"
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = False
    wsVerif.Cells.Clear

    ' --- ÉTAPE 1 : Lire toutes les données ---
    Dim i As Long, j As Long, eleve As Variant
    ' Lecture Affectations
    For i = 2 To wsResultats.Cells(wsResultats.Rows.Count, "A").End(xlUp).Row
        Dim projetAffecte As String: projetAffecte = Trim(wsResultats.Cells(i, 1).Value)
        Dim listeElevesStr As String: listeElevesStr = Trim(wsResultats.Cells(i, 2).Value)
        If listeElevesStr <> "Aucun" And listeElevesStr <> "" Then
            Dim elevesAffectesArray As Variant: elevesAffectesArray = Split(listeElevesStr, ", ")
            For Each eleve In elevesAffectesArray
                affectations(Trim(eleve)) = projetAffecte
            Next eleve
        End If
    Next i
    
    ' --- LECTURE DES PRÉFÉRENCES (NOUVELLE LOGIQUE FIABLE) ---
    For i = 2 To wsPrefsEleves.Cells(wsPrefsEleves.Rows.Count, "A").End(xlUp).Row
        Dim nomEleve As String: nomEleve = Trim(wsPrefsEleves.Cells(i, 1).Value)
        If nomEleve <> "" Then
            ' On crée la collection DIRECTEMENT dans le dictionnaire pour cet élève
            Set prefsEleves(nomEleve) = New Collection
            
            ' On remplit la collection qui est DÉJÀ dans le dictionnaire
            For j = 2 To wsPrefsEleves.Cells(i, wsPrefsEleves.Columns.Count).End(xlToLeft).Column
                prefsEleves(nomEleve).Add Trim(wsPrefsEleves.Cells(i, j).Value)
            Next j
        End If
    Next i

    ' --- ÉTAPE 2 : Écrire les données brutes dans la feuille de vérification ---
    wsVerif.Range("A1:C1").Value = Array("Élève", "Projet Affecté (lu)", "Rang Calculé")
    wsVerif.Range("D1").Value = "Vœux de l'élève (lus) ->"
    wsVerif.Range("A1:D1").Font.Bold = True
    
    Dim ligneVerif As Long: ligneVerif = 1
    
    For Each eleve In prefsEleves.Keys
        ligneVerif = ligneVerif + 1
        wsVerif.Cells(ligneVerif, 1).Value = eleve
        
        If affectations.Exists(eleve) Then
            wsVerif.Cells(ligneVerif, 2).Value = affectations(eleve)
        Else
            wsVerif.Cells(ligneVerif, 2).Value = "Non affecté"
        End If
        
        Dim listeDeChoix As Collection: Set listeDeChoix = prefsEleves(eleve)
        For j = 1 To listeDeChoix.Count
            wsVerif.Cells(ligneVerif, 3 + j).Value = listeDeChoix(j)
        Next j
    Next eleve
    
    ' --- ÉTAPE 3 : Calculer le rang avec une formule Excel ---
    For i = 2 To ligneVerif
        If wsVerif.Cells(i, 2).Value <> "Non affecté" And wsVerif.Cells(i, 2).Value <> "" Then
            wsVerif.Cells(i, 3).FormulaR1C1 = "=IFERROR(MATCH(RC2, RC4:RC100, 0), ""Non trouvé"")"
        Else
            wsVerif.Cells(i, 3).Value = "N/A"
        End If
    Next i
    
    wsVerif.Columns.AutoFit
    Application.ScreenUpdating = True
    MsgBox "La feuille 'Verification_Finale' a été créée avec la nouvelle méthode de lecture. Veuillez l'inspecter.", vbInformation

End Sub
'****************************************************************************************
' MACRO : Créer un tableau de bord graphique avec étiquettes de données
'****************************************************************************************
Sub CreerDashboard()

    ' --- Déclarations ---
    Dim wsBilan As Worksheet, wsDashboard As Worksheet
    Dim cht As ChartObject
    Dim dataRange As Range, labelsRange As Range

    ' --- Association des feuilles ---
    On Error Resume Next
    Set wsBilan = ThisWorkbook.Sheets("Bilan_Performance")
    ' Crée ou vide la feuille de Dashboard
    Set wsDashboard = ThisWorkbook.Sheets("Dashboard")
    If wsDashboard Is Nothing Then
        Set wsDashboard = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsDashboard.Name = "Dashboard"
    End If
    On Error GoTo 0

    If wsBilan Is Nothing Then
        MsgBox "Erreur : La feuille 'Bilan_Performance' est introuvable. Veuillez d'abord exécuter la macro de bilan.", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False
    wsDashboard.Cells.Clear
    ' Supprime les anciens graphiques pour éviter les doublons
    For Each cht In wsDashboard.ChartObjects
        cht.Delete
    Next cht

    ' --- GRAPHIQUE 1 : Répartition des Élèves (Barres) ---
    Set dataRange = wsBilan.Range("B3:B5")
    Set labelsRange = wsBilan.Range("A3:A5")
    
    Set cht = wsDashboard.ChartObjects.Add(Left:=50, Top:=30, Width:=400, Height:=250)
    With cht.Chart
        .SetSourceData Source:=dataRange
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Répartition des Élèves"
        .SeriesCollection(1).XValues = labelsRange
        .HasLegend = False
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Nombre d'élèves"
        .ApplyDataLabels ' NOUVEAU : Ajout des étiquettes
    End With

    ' --- GRAPHIQUE 2 : Satisfaction des Vœux (Histogramme) ---
    Set dataRange = wsBilan.Range("B11:B13")
    Set labelsRange = wsBilan.Range("A11:A13")

    Set cht = wsDashboard.ChartObjects.Add(Left:=500, Top:=30, Width:=400, Height:=250)
    With cht.Chart
        .SetSourceData Source:=dataRange
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Satisfaction des Vœux"
        .SeriesCollection(1).Name = "Nombre d'élèves"
        .SeriesCollection(1).XValues = labelsRange
        .HasLegend = False
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Nombre d'élèves"
        .ApplyDataLabels ' NOUVEAU : Ajout des étiquettes
    End With

    ' --- GRAPHIQUE 3 : Performance des Projets (Barres Horizontales) ---
    Set dataRange = wsBilan.Range("B18:B19")
    Set labelsRange = wsBilan.Range("A18:A19")

    Set cht = wsDashboard.ChartObjects.Add(Left:=50, Top:=300, Width:=400, Height:=250)
    With cht.Chart
        .SetSourceData Source:=dataRange
        .ChartType = xlBarClustered
        .HasTitle = True
        .ChartTitle.Text = "Performance des Projets"
        .SeriesCollection(1).Name = "Nombre de projets"
        .SeriesCollection(1).XValues = labelsRange
        .HasLegend = False
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Nombre de projets"
        .ApplyDataLabels ' NOUVEAU : Ajout des étiquettes
    End With

    Application.ScreenUpdating = True
    wsDashboard.Activate

End Sub
