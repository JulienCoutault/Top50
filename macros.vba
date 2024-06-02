Dim lenom, Lepath As String

Sub Archivage()
'
' Macro d'archivage des données
' Elaboré à partir d'Import Version du 26 septembre 2005 par rené
' Modifiée le 03/10/2005 par Yok's
' Séparation de l'archivage / import le 05/06/2006 par Yok's
' Adaptation aux nouveaux tops 50 le 20/08/07 par Yok's
'
' *********************************************
' *** on se note le nom du fichier Excel    ***
' *********************************************
    lenom = ActiveWorkbook.Name
    Lepath = ActiveWorkbook.Path
' **********************************************************
' *** On ouvre le fichier archive                        ***
' *** on décale toutes les arhives vers la gauche        ***
' **********************************************************
    Workbooks.Open Filename:=Lepath & "\Archives.xls"
    Range("D1:IU3000").Select
    Selection.Copy
    Range("B1").Select
    ActiveSheet.Paste
' **********************************************************
' *** on ouvre le fichier .csv et on le remet au format  ***
' *** On note sa date d'enregistrement                   ***
' **********************************************************
    Workbooks.Open Filename:=Lepath & "\Copie500.csv"
    For i = 1 To 3000
        If Cells(i, 1).Value <> "" Then
            Cells(i, 1).Activate
            Selection.TextToColumns DataType:=xlDelimited, _
            ConsecutiveDelimiter:=False, Other:=True, Otherchar:=";"
        End If
    Next
'    Date = FileDateTime(Lepath & "\Copie500.csv")
' **********************************************************
' *** on copie les informations dans les archives        ***
' *** on ferme le fichier .csv sans le sauvegarder       ***
' **********************************************************
    Range("B1:C2650").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("Archives.xls").Activate
    Range("IT2").Select
    ActiveSheet.Paste
    Range("IT1").Value = Date
    Windows("Copie500.csv").Close SaveChanges:=False
' **********************************************************
' *** on ferme le fichier des archives en sauvegardant   ***
' **********************************************************
    Windows("Archives.xls").Close SaveChanges:=True
End Sub
    
Sub Transfert()
' Sous marcro d'importation séparée le 05/06/2006 par yok's
' Permet de transférer dans top50 les données à exploiter
' sans modifier le fichier archive

' *********************************************
' *** on se note le nom du fichier Excel    ***
' *********************************************
    lenom = ActiveWorkbook.Name
    Lepath = ActiveWorkbook.Path
' **********************************************************
' *** On ouvre le fichier archive                        ***
' **********************************************************
    Workbooks.Open Filename:=Lepath & "\Archives.xls"
    
' **********************************************************
' ***On copie les informations des archives dans le top50***
' ***    Classements mondiaux et national par billets    ***
' **********************************************************
    Windows("Archives.xls").Activate
    Range("IR627:IS1126").Select
    Selection.Copy
    Windows(lenom).Activate
    Sheets("TOP50 billets").Select
    Range("C5").Select
    ActiveSheet.Paste
    Windows("Archives.xls").Activate
    Range("IT627:IU1126").Select
    Selection.Copy
    Windows(lenom).Activate
    Sheets("TOP50 billets").Select
    Range("G5").Select
    ActiveSheet.Paste
    Windows("Archives.xls").Activate
    Range("IR1127:IS1376").Select
    Selection.Copy
    Windows(lenom).Activate
    Sheets("TOP50 billets").Select
    Range("K5").Select
    ActiveSheet.Paste
    Windows("Archives.xls").Activate
    Range("IT1127:IU1376").Select
    Selection.Copy
    Windows(lenom).Activate
    Sheets("TOP50 billets").Select
    Range("O5").Select
    ActiveSheet.Paste
' **********************************************************
' ***     Classements mondiaux et national des hits      ***
' **********************************************************
    Windows("Archives.xls").Activate
    Range("IR1377:IS1976").Select
    Selection.Copy
    Windows(lenom).Activate
    Sheets("TOP30 hits").Select
    Range("C4").Select
    ActiveSheet.Paste
    Windows("Archives.xls").Activate
    Range("IT1377:IU1976").Select
    Selection.Copy
    Windows(lenom).Activate
    Sheets("TOP30 hits").Select
    Range("G4").Select
    ActiveSheet.Paste
    Windows("Archives.xls").Activate
    Range("IR1977:IS2126").Select
    Selection.Copy
    Windows(lenom).Activate
    Sheets("TOP30 hits").Select
    Range("K4").Select
    ActiveSheet.Paste
    Windows("Archives.xls").Activate
    Range("IT1977:IU2126").Select
    Selection.Copy
    Windows(lenom).Activate
    Sheets("TOP30 hits").Select
    Range("O4").Select
    ActiveSheet.Paste
' **********************************************************
' ***  Classements mondiaux et national des plus actifs  ***
' **********************************************************
    Windows("Archives.xls").Activate
    Range("IR77:IS576").Select
    Selection.Copy
    Windows(lenom).Activate
    Sheets("TOP20 Actifs").Select
    Range("C4").Select
    ActiveSheet.Paste
    Windows("Archives.xls").Activate
    Range("IT77:IU576").Select
    Selection.Copy
    Windows(lenom).Activate
    Sheets("TOP20 Actifs").Select
    Range("G4").Select
    ActiveSheet.Paste
    Windows("Archives.xls").Activate
    Range("IR577:IS626").Select
    Selection.Copy
    Windows(lenom).Activate
    Sheets("TOP20 Actifs").Select
    Range("K4").Select
    ActiveSheet.Paste
    Windows("Archives.xls").Activate
    Range("IT577:IU626").Select
    Selection.Copy
    Windows(lenom).Activate
    Sheets("TOP20 Actifs").Select
    Range("O4").Select
    ActiveSheet.Paste
' **********************************************************
' ***Classements mondiaux et national du meilleur parrain***
' **********************************************************
    Windows("Archives.xls").Activate
    Range("IR2127:IS2226").Select
    Selection.Copy
    Windows(lenom).Activate
    Sheets("TOP20 Parrains").Select
    Range("C4").Select
    ActiveSheet.Paste
    Windows("Archives.xls").Activate
    Range("IT2127:IU2226").Select
    Selection.Copy
    Windows(lenom).Activate
    Sheets("TOP20 Parrains").Select
    Range("G4").Select
    ActiveSheet.Paste
    Windows("Archives.xls").Activate
    Range("IR2227:IS2251").Select
    Selection.Copy
    Windows(lenom).Activate
    Sheets("TOP20 Parrains").Select
    Range("K4").Select
    ActiveSheet.Paste
    Windows("Archives.xls").Activate
    Range("IT2227:IU2251").Select
    Selection.Copy
    Windows(lenom).Activate
    Sheets("TOP20 Parrains").Select
    Range("O4").Select
    ActiveSheet.Paste
' **********************************************************
' ***          Classements des 24 meilleurs pays         ***
' **********************************************************
    Windows("Archives.xls").Activate
    Range("IR2:IS26").Select
    Selection.Copy
    Windows(lenom).Activate
    Sheets("TOP Pays").Select
    Range("C3").Select
    ActiveSheet.Paste
    Windows("Archives.xls").Activate
    Range("IT2:IU26").Select
    Selection.Copy
    Windows(lenom).Activate
    Sheets("TOP Pays").Select
    Range("G3").Select
    ActiveSheet.Paste
    Windows("Archives.xls").Activate
    Range("IR27:IS51").Select
    Selection.Copy
    Windows(lenom).Activate
    Sheets("TOP Pays").Select
    Range("C67").Select
    ActiveSheet.Paste
    Windows("Archives.xls").Activate
    Range("IT27:IU51").Select
    Selection.Copy
    Windows(lenom).Activate
    Sheets("TOP Pays").Select
    Range("G67").Select
    ActiveSheet.Paste
    Windows("Archives.xls").Activate
    Range("IR52:IS76").Select
    Selection.Copy
    Windows(lenom).Activate
    Sheets("TOP Pays").Select
    Range("C35").Select
    ActiveSheet.Paste
    Windows("Archives.xls").Activate
    Range("IT52:IU76").Select
    Selection.Copy
    Windows(lenom).Activate
    Sheets("TOP Pays").Select
    Range("G35").Select
    ActiveSheet.Paste
' **********************************************************
' ***    Classements des meilleurs villes par billets    ***
' **********************************************************
    Windows("Archives.xls").Activate
    Range("IR2252:IS2401").Select
    Selection.Copy
    Windows(lenom).Activate
    Sheets("TOP Villes Billets").Select
    Range("C4").Select
    ActiveSheet.Paste
    Windows("Archives.xls").Activate
    Range("IT2252:IU2401").Select
    Selection.Copy
    Windows(lenom).Activate
    Sheets("TOP Villes Billets").Select
    Range("G4").Select
    ActiveSheet.Paste
    Windows("Archives.xls").Activate
    Range("IR2402:IS2451").Select
    Selection.Copy
    Windows(lenom).Activate
    Sheets("TOP Villes Billets").Select
    Range("K4").Select
    ActiveSheet.Paste
    Windows("Archives.xls").Activate
    Range("IT2402:IU2451").Select
    Selection.Copy
    Windows(lenom).Activate
    Sheets("TOP Villes Billets").Select
    Range("O4").Select
    ActiveSheet.Paste
    Windows("Archives.xls").Activate
    Range("IR2452:IS2476").Select
    Selection.Copy
    Windows(lenom).Activate
    Sheets("TOP Villes Billets").Select
    Range("K55").Select
    ActiveSheet.Paste
    Windows("Archives.xls").Activate
    Range("IT2452:IU2476").Select
    Selection.Copy
    Windows(lenom).Activate
    Sheets("TOP Villes Billets").Select
    Range("O55").Select
    ActiveSheet.Paste
' **********************************************************
' *** Classements des meilleurs villes par utilisateurs  ***
' **********************************************************
    Windows("Archives.xls").Activate
    Range("IR2477:IS2626").Select
    Selection.Copy
    Windows(lenom).Activate
    Sheets("TOP Villes Utilisateurs").Select
    Range("C4").Select
    ActiveSheet.Paste
    Windows("Archives.xls").Activate
    Range("IT2477:IU2626").Select
    Selection.Copy
    Windows(lenom).Activate
    Sheets("TOP Villes Utilisateurs").Select
    Range("G4").Select
    ActiveSheet.Paste
    Windows("Archives.xls").Activate
    Range("IR2627:IS2651").Select
    Selection.Copy
    Windows(lenom).Activate
    Sheets("TOP Villes Utilisateurs").Select
    Range("K4").Select
    ActiveSheet.Paste
    Windows("Archives.xls").Activate
    Range("IT2627:IU2651").Select
    Selection.Copy
    Windows(lenom).Activate
    Sheets("TOP Villes Utilisateurs").Select
    Range("O4").Select
    ActiveSheet.Paste
' **********************************************************
' ***                  Import de la date                 ***
' **********************************************************
    Windows("Archives.xls").Activate
    Range("IT1").Copy
    Windows(lenom).Activate
    Sheets("General").Select
    Range("A5").Select
    ActiveSheet.Paste
    
' **********************************************************
' *** on ferme le fichier des archives sans sauvegarder  ***
' **********************************************************
    Windows("Archives.xls").Close SaveChanges:=False
    
    
End Sub

Sub Calcul()
'
' Macro de mise en place des données et calcul
' Version du 16 septembre 2005 par Michel
' Modifiée le 04/10/2005 par Yok's
' Adapté le 12/06/06 par Yok's pour la nouvelle version
' Adapté le 07/01/18 par yok's pour numéro des tops

' ****************************************************************
' ***    on demande le numéro du top 50 et le rédacteur        ***
' ****************************************************************
Sheets("General").Select
Nbrtop = InputBox("Réalisation du Top 50 N° ?")
Semaine = Cells(5, 5).Value
nom = InputBox("Nom du rédacteur ?")
Cells(5, 6).Value = Nbrtop
Cells(8, 2).Value = "Edition " & Nbrtop - 150 & " - Semaine " & Semaine & " - " & Cells(5, 2).Text & " -[color=blue] par " & nom & "[/color]"
Cells(8, 17).Value = "Edition " & Nbrtop & " - Semaine " & Semaine & " - " & Cells(5, 2).Text & " -[color=blue] par " & nom & "[/color]"
'Cells(71, 20).Value = Nbrtop

' **********************************************************
' *** Classement top 50 billets                          ***
' *** On fait une boucle sur les 60 premiers             ***
' *** On initialise les valeurs                          ***
' **********************************************************
 Sheets("TOP50 billets").Select
 For Top = 1 To 60
    Cells(Top + 4, 28).Value = True
    Cells(Top + 4, 29).Value = True
    Cells(Top + 4, 30).Value = True
    Cells(Top + 4, 31).Value = False
    Cells(Top + 4, 32).Value = 0
    Cells(Top + 4, 33).Value = 0
    Cells(Top + 4, 34).Value = 0
    Cells(Top + 4, 35).Value = 0
    For c250 = 1 To 250
' **********************************************************
' *** On fait une boucle sur les 250 premiers français   ***
' *** On compare les noms et on modifie si correspondance***
' **********************************************************
        If Cells(c250 + 4, 11).Value = Cells(Top + 4, 15).Value Then
            Cells(Top + 4, 26).Value = c250 - Top
            Cells(Top + 4, 27).Value = Cells(Top + 4, 16).Value - Cells(c250 + 4, 12).Value
            Cells(Top + 4, 28).Value = False
            Cells(Top + 4, 29).Value = False
            If c250 <= 50 Then
                Cells(Top + 4, 30).Value = False
                For h150 = 1 To 150
                    If Cells(h150 + 4, 23).Value = Cells(Top + 4, 15).Value Then Cells(Top + 4, 32).Value = Cells(Top + 4, 16).Value / Cells(h150 + 4, 24).Value
                    If Cells(h150 + 4, 19).Value = Cells(Top + 4, 15).Value Then Cells(Top + 4, 33).Value = Cells(Top + 4, 12).Value / Cells(h150 + 4, 20).Value
                Next h150
                For c500 = 1 To 500
                    If Cells(c500 + 4, 7).Value = Cells(Top + 4, 15) Then Cells(Top + 4, 34).Value = c500
                    If Cells(c500 + 4, 3).Value = Cells(Top + 4, 15) Then Cells(Top + 4, 35).Value = c500
                Next c500
            End If
' **********************************************************
' *** Idem avec les inactif de la semaine précédente     ***
' *** qui sont repassé à l'état actif                    ***
' **********************************************************
        Else
            If Cells(c250 + 4, 11).Value & " [inactif depuis 30 jours]" = Cells(Top + 4, 15).Value Then
                Cells(Top + 4, 26).Value = c250 - Top
                Cells(Top + 4, 27).Value = Cells(Top + 4, 16).Value - Cells(c250 + 4, 12).Value
                Cells(Top + 4, 28).Value = True
                Cells(Top + 4, 29).Value = False
                If c250 <= 50 Then
                    Cells(Top + 4, 30).Value = False
                    For h150 = 1 To 150
                        If Cells(h150 + 4, 23).Value = Cells(Top + 4, 15).Value Then Cells(Top + 4, 32).Value = Cells(Top + 4, 16).Value / Cells(h150 + 4, 24).Value
                        If Cells(h150 + 4, 19).Value & " [inactif depuis 30 jours]" = Cells(Top + 4, 15).Value Then Cells(Top + 4, 33).Value = Cells(Top + 4, 12).Value / Cells(h150 + 4, 20).Value
                    Next h150
                    For c500 = 1 To 500
                        If Cells(c500 + 4, 7).Value = Cells(Top + 4, 15) Then Cells(Top + 4, 34).Value = c500
                        If Cells(c500 + 4, 3).Value & " [inactif depuis 30 jours]" = Cells(Top + 4, 15) Then Cells(Top + 4, 35).Value = c500
                    Next c500
                End If
            Else
' **********************************************************
' *** Idem avec les actif de la semaine précédente       ***
' *** qui sont passé à l'état inactif                    ***
' **********************************************************
                If Cells(c250 + 4, 11).Value = Cells(Top + 4, 15).Value & " [inactif depuis 30 jours]" Then
                    Cells(Top + 4, 26).Value = c250 - Top
                    Cells(Top + 4, 27).Value = Cells(Top + 4, 16).Value - Cells(c250 + 4, 12).Value
                    Cells(Top + 4, 28).Value = False
                    Cells(Top + 4, 29).Value = True
                    If c250 <= 50 Then
                        Cells(Top + 4, 30).Value = False
                        For h150 = 1 To 150
                            If Cells(h150 + 4, 23).Value = Cells(Top + 4, 15).Value Then Cells(Top + 4, 32).Value = Cells(Top + 4, 16).Value / Cells(h150 + 4, 24).Value
                            If Cells(h150 + 4, 19).Value = Cells(Top + 4, 15).Value & " [inactif depuis 30 jours]" Then Cells(Top + 4, 33).Value = Cells(Top + 4, 12).Value / Cells(h150 + 4, 20).Value
                        Next h150
                        For c500 = 1 To 500
                            If Cells(c500 + 4, 7).Value = Cells(Top + 4, 15) Then Cells(Top + 4, 34).Value = c500
                            If Cells(c500 + 4, 3).Value = Cells(Top + 4, 15) & " [inactif depuis 30 jours]" Then Cells(Top + 4, 35).Value = c500
                        Next c500
                    End If
                End If
            End If
        End If
    Next c250
    If Cells(Top + 4, 28) And Cells(Top + 4, 29) Then Cells(Top + 4, 31).Value = True
Next Top
    
' **********************************************************
' *** Classement top 30 hits                             ***
' *** On fait une boucle sur les 40 premiers             ***
' *** On initialise les valeurs                          ***
' **********************************************************
 Sheets("TOP30 hits").Select
 For Top = 1 To 40
    Cells(Top + 3, 20).Value = True
    Cells(Top + 3, 21).Value = True
    Cells(Top + 3, 22).Value = True
    Cells(Top + 3, 23).Value = False
    Cells(Top + 3, 24).Value = 0
    Cells(Top + 3, 25).Value = 0
    For c50 = 1 To 150
' **********************************************************
' *** On fait une boucle sur les 150 premiers français   ***
' *** et une boucle sur les 600 hitteurs internationaux  ***
' *** On compare les noms et on modifie si correspondance***
' **********************************************************
        If Cells(c50 + 3, 11).Value = Cells(Top + 3, 15).Value Then
            Cells(Top + 3, 18).Value = c50 - Top
            Cells(Top + 3, 19).Value = Cells(Top + 3, 16).Value - Cells(c50 + 3, 12).Value
            Cells(Top + 3, 20).Value = False
            Cells(Top + 3, 21).Value = False
            If c50 <= 30 Then
                Cells(Top + 3, 22).Value = False
                For h600 = 1 To 600
                    If Cells(h600 + 3, 7).Value = Cells(Top + 3, 15).Value Then Cells(Top + 3, 24).Value = h600
                    If Cells(h600 + 3, 3).Value = Cells(Top + 3, 15).Value Then Cells(Top + 3, 25).Value = h600
                Next h600
            End If
' **********************************************************
' *** Idem avec les inactif de la semaine précédente     ***
' *** qui sont repassé à l'état actif                    ***
' **********************************************************
        Else
            If Cells(c50 + 3, 11).Value & " [inactif depuis 30 jours]" = Cells(Top + 3, 15).Value Then
                Cells(Top + 3, 18).Value = c50 - Top
                Cells(Top + 3, 19).Value = Cells(Top + 3, 16).Value - Cells(c50 + 3, 12).Value
                Cells(Top + 3, 20).Value = True
                Cells(Top + 3, 21).Value = False
                If c50 <= 30 Then
                    Cells(Top + 3, 22).Value = False
                    For h600 = 1 To 600
                        If Cells(h600 + 3, 7).Value = Cells(Top + 3, 15).Value Then Cells(Top + 3, 24).Value = h600
                        If Cells(h600 + 3, 3).Value & " [inactif depuis 30 jours]" = Cells(Top + 3, 15).Value Then Cells(Top + 3, 25).Value = h600
                    Next h600
                End If
            Else
' **********************************************************
' *** Idem avec les actif de la semaine précédente       ***
' *** qui sont passé à l'état inactif                    ***
' **********************************************************
                If Cells(c50 + 3, 11).Value = Cells(Top + 3, 15).Value & " [inactif depuis 30 jours]" Then
                    Cells(Top + 3, 18).Value = c50 - Top
                    Cells(Top + 3, 19).Value = Cells(Top + 3, 16).Value - Cells(c50 + 3, 12).Value
                    Cells(Top + 3, 20).Value = False
                    Cells(Top + 3, 21).Value = True
                    If c50 <= 30 Then
                        Cells(Top + 3, 22).Value = False
                        For h600 = 1 To 600
                            If Cells(h600 + 3, 7).Value = Cells(Top + 3, 15).Value Then Cells(Top + 3, 24).Value = h600
                            If Cells(h600 + 3, 3).Value = Cells(Top + 3, 15).Value & " [inactif depuis 30 jours]" Then Cells(Top + 3, 25).Value = h600
                        Next h600
                    End If
                End If
            End If
        End If
    Next c50
    If Cells(Top + 3, 20) And Cells(Top + 3, 21) Then Cells(Top + 3, 23).Value = True
Next Top


' **********************************************************
' *** Classement top 20 plus actifs                      ***
' *** On fait une boucle sur les 30 premiers             ***
' *** On initialise les valeurs                          ***
' **********************************************************
 Sheets("TOP20 Actifs").Select
 For Top = 1 To 30
    Cells(Top + 3, 18).Value = -51
    Cells(Top + 3, 19).Value = Cells(Top + 3, 16).Value & " ?"
    Cells(Top + 3, 20).Value = True
    Cells(Top + 3, 21).Value = False
    Cells(Top + 3, 22).Value = 0
    Cells(Top + 3, 23).Value = 0
    For c50 = 1 To 50
' **********************************************************
' *** On fait une boucle sur les 50 premiers français    ***
' *** On compare les noms et on modifie si correspondance***
' **********************************************************
        If Cells(c50 + 3, 11).Value = Cells(Top + 3, 15) Then
            Cells(Top + 3, 18).Value = c50 - Top
            Cells(Top + 3, 19).Value = Cells(Top + 3, 16).Value - Cells(c50 + 3, 12).Value
            If c50 <= 20 Then Cells(Top + 3, 20).Value = False
            If Top <= 20 Then
                For c250 = 1 To 500
                    If Cells(c250 + 3, 7).Value = Cells(Top + 3, 15) Then Cells(Top + 3, 22).Value = c250
                    If Cells(c250 + 3, 3).Value = Cells(Top + 3, 15) Then Cells(Top + 3, 23).Value = c250
                Next c250
            End If
        End If
    Next c50
    If Cells(Top + 3, 18).Value = -51 Then Cells(Top + 3, 21).Value = True
Next Top


' **********************************************************
' *** Classement top 20 meilleurs parrains               ***
' *** On fait une boucle sur les 25 premiers             ***
' *** On initialise les valeurs                          ***
' **********************************************************
 Sheets("TOP20 parrains").Select
 For Top = 1 To 25
    Cells(Top + 3, 20).Value = True
    Cells(Top + 3, 21).Value = True
    Cells(Top + 3, 22).Value = True
    Cells(Top + 3, 23).Value = False
    Cells(Top + 3, 24).Value = 0
    Cells(Top + 3, 25).Value = 0
    For c25 = 1 To 25
' **********************************************************
' ***   On fait une boucle sur les 25 premiers français  ***
' *** et une boucle sur les 100 premiers internationaux  ***
' *** On compare les noms et on modifie si correspondance***
' **********************************************************
        If Cells(c25 + 3, 11).Value = Cells(Top + 3, 15).Value Then
            Cells(Top + 3, 18).Value = c25 - Top
            Cells(Top + 3, 19).Value = Cells(Top + 3, 16).Value - Cells(c25 + 3, 12).Value
            Cells(Top + 3, 20).Value = False
            Cells(Top + 3, 21).Value = False
            If c25 <= 20 Then
                Cells(Top + 3, 22).Value = False
                For h100 = 1 To 100
                    If Cells(h100 + 3, 7).Value = Cells(Top + 3, 15).Value Then Cells(Top + 3, 24).Value = h100
                    If Cells(h100 + 3, 3).Value = Cells(Top + 3, 15).Value Then Cells(Top + 3, 25).Value = h100
                Next h100
            End If
' **********************************************************
' *** Idem avec les actif de la semaine précédente       ***
' *** qui sont passé à l'état inactif                    ***
' **********************************************************
        Else
            If Cells(c25 + 3, 11).Value & " [inactif depuis 30 jours]" = Cells(Top + 3, 15) Then
                Cells(Top + 3, 18).Value = c25 - Top
                Cells(Top + 3, 19).Value = Cells(Top + 3, 16).Value - Cells(c25 + 3, 12).Value
                Cells(Top + 3, 20).Value = True
                Cells(Top + 3, 21).Value = False
                If c25 <= 20 Then
                    Cells(Top + 3, 22).Value = False
                    For h100 = 1 To 100
                        If Cells(h100 + 3, 7).Value = Cells(Top + 3, 15).Value Then Cells(Top + 3, 24).Value = h100
                        If Cells(h100 + 3, 3).Value & " [inactif depuis 30 jours]" = Cells(Top + 3, 15).Value Then Cells(Top + 3, 25).Value = h100
                    Next h100
                End If
            Else
' **********************************************************
' *** Idem avec les inactif de la semaine précédente     ***
' *** qui sont repassé à l'état actif                    ***
' **********************************************************
                If Cells(c25 + 3, 11).Value = Cells(Top + 3, 15) & " [inactif depuis 30 jours]" Then
                    Cells(Top + 3, 18).Value = c25 - Top
                    Cells(Top + 3, 19).Value = Cells(Top + 3, 16).Value - Cells(c25 + 3, 12).Value
                    Cells(Top + 3, 20).Value = False
                    Cells(Top + 3, 21).Value = True
                    If c25 <= 20 Then
                        Cells(Top + 3, 22).Value = False
                        For h100 = 1 To 100
                            If Cells(h100 + 3, 7).Value = Cells(Top + 3, 15).Value Then Cells(Top + 3, 24).Value = h100
                            If Cells(h100 + 3, 3).Value = Cells(Top + 3, 15).Value & " [inactif depuis 30 jours]" Then Cells(Top + 3, 25).Value = h100
                        Next h100
                    End If
                End If
            End If
        End If
    Next c25
    If Cells(Top + 3, 12) And Cells(Top + 3, 13) Then Cells(Top + 3, 15).Value = True
Next Top


' **********************************************************
' *** Classement top 13 des pays                         ***
' *** On fait une boucle sur les 20 premiers             ***
' *** On initialise les valeurs                          ***
' **********************************************************
 Sheets("TOP Pays").Select
 For Top = 1 To 20
    Cells(Top + 3, 12).Value = True
    Cells(Top + 35, 12).Value = True
    Cells(Top + 67, 12).Value = True
    For c24 = 1 To 24
' **********************************************************
' *** On fait une boucle sur les 24 premiers pays        ***
' *** On compare les noms et on modifie si correspondance***
' **********************************************************
        If Cells(c24 + 3, 3).Value = Cells(Top + 3, 7) Then
            Cells(Top + 3, 10).Value = c24 - Top
            Cells(Top + 3, 11).Value = Cells(Top + 3, 8).Value - Cells(c24 + 3, 4).Value
            If c24 <= 13 Then Cells(Top + 3, 12).Value = False
        End If
        If Cells(c24 + 35, 3).Value = Cells(Top + 35, 7) Then
            Cells(Top + 35, 10).Value = c24 - Top
            Cells(Top + 35, 11).Value = Cells(Top + 35, 8).Value - Cells(c24 + 35, 4).Value
            If c24 <= 13 Then Cells(Top + 35, 12).Value = False
        End If
        If Cells(c24 + 67, 3).Value = Cells(Top + 67, 7) Then
            Cells(Top + 67, 10).Value = c24 - Top
            Cells(Top + 67, 11).Value = Cells(Top + 67, 8).Value - Cells(c24 + 67, 4).Value
            If c24 <= 13 Then Cells(Top + 67, 12).Value = False
        End If
        If Top <= 13 And Cells(c24 + 3, 3).Value = Cells(Top + 67, 7) Then
            Cells(Top + 67, 23).Value = Cells(c24 + 3, 4).Value \ Cells(Top + 67, 8).Value
        End If
    Next c24
Next Top


' **********************************************************
' *** Classement top 20 des villes / utilisateurs        ***
' *** On fait une boucle sur les 25 premiers             ***
' *** On initialise les valeurs                          ***
' **********************************************************
 Sheets("TOP Villes Utilisateurs").Select
 For Top = 1 To 25
    Cells(Top + 3, 20).Value = True
    Cells(Top + 3, 21).Value = 0
    Cells(Top + 3, 22).Value = 0
    For c25 = 1 To 25
' **********************************************************
' *** On fait une boucle sur les 25 premières villes     ***
' *** On compare les noms et on modifie si correspondance***
' **********************************************************
        If Cells(c25 + 3, 11).Value = Cells(Top + 3, 15) Then
            Cells(Top + 3, 18).Value = c25 - Top
            Cells(Top + 3, 19).Value = Cells(Top + 3, 16).Value - Cells(c25 + 3, 12).Value
            If c25 <= 20 Then
                Cells(Top + 3, 20).Value = False
                For h150 = 1 To 150
                    If Cells(h150 + 3, 7).Value = Cells(Top + 3, 15).Value Then Cells(Top + 3, 21).Value = h150
                    If Cells(h150 + 3, 3).Value = Cells(Top + 3, 15).Value Then Cells(Top + 3, 22).Value = h150
                Next h150
            End If
        End If
    Next c25
Next Top

' **********************************************************
' *** Classement top 40 des villes / billets             ***
' *** On fait une boucle sur les 50 premiers             ***
' *** On initialise les valeurs                          ***
' **********************************************************
 Sheets("TOP Villes Billets").Select
 For Top = 1 To 50
    Cells(Top + 3, 20).Value = True
    Cells(Top + 3, 21).Value = 0
    Cells(Top + 3, 22).Value = 0
    For c50 = 1 To 50
' **********************************************************
' *** On fait une boucle sur les 50 premières villes     ***
' *** On compare les noms et on modifie si correspondance***
' **********************************************************
        If Cells(c50 + 3, 11).Value = Cells(Top + 3, 15) Then
            Cells(Top + 3, 18).Value = c50 - Top
            Cells(Top + 3, 19).Value = Cells(Top + 3, 16).Value - Cells(c50 + 3, 12).Value
            If c50 <= 40 Then
                Cells(Top + 3, 20).Value = False
                For h150 = 1 To 150
                    If Cells(h150 + 3, 7).Value = Cells(Top + 3, 15).Value Then Cells(Top + 3, 21).Value = h150
                    If Cells(h150 + 3, 3).Value = Cells(Top + 3, 15).Value Then Cells(Top + 3, 22).Value = h150
                Next h150
            End If
        End If
    Next c50
Next Top

' **********************************************************
' *** Classement top 20 des villes / hits                ***
' *** On fait une boucle sur les 25 premiers             ***
' *** On initialise les valeurs                          ***
' **********************************************************
 Sheets("TOP Villes Billets").Select
 For Top = 1 To 25
    Cells(Top + 54, 20).Value = True
    Cells(Top + 54, 21).Value = 0
    Cells(Top + 54, 22).Value = 0
    For c25 = 1 To 25
' **********************************************************
' *** On fait une boucle sur les 25 premières villes     ***
' *** On compare les noms et on modifie si correspondance***
' **********************************************************
        If Cells(c25 + 54, 11).Value = Cells(Top + 54, 15) Then
            Cells(Top + 54, 18).Value = c25 - Top
            Cells(Top + 54, 19).Value = Cells(Top + 54, 16).Value - Cells(c25 + 54, 12).Value
            If c25 <= 20 Then
                Cells(Top + 54, 20).Value = False
                For h50 = 1 To 50
                    If Cells(h50 + 3, 15).Value = Cells(Top + 54, 15).Value Then Cells(Top + 54, 21).Value = Cells(h50 + 3, 16).Value
                    If Cells(h50 + 3, 11).Value = Cells(Top + 54, 15).Value Then Cells(Top + 54, 22).Value = Cells(h50 + 3, 12).Value
                Next h50
            End If
        End If
    Next c25
Next Top

Sheets("General").Select
End Sub

Sub Records()
'
' Macro de mise à jour des records
' Version du 17 septembre 2005 par rené
' Modifiée le 09/10/05 par Yok's
' Adapté le 07/08/06 à la nouvelle version par Yok's
' Ajout d'un record le 03/12/06 par Yok's
' Extension du top hebdomadaire le 07/01/18 par Yok's
'
' ********************************************************************
' *** on récupère le numéro du top 50 en court                     ***
' ********************************************************************
Sheets("General").Select
top50 = Cells(5, 6).Value

' *********************************************
' *** on se note le nom du fichier Excel    ***
' *********************************************
lenom = ActiveWorkbook.Name
Lepath = ActiveWorkbook.Path
Workbooks.Open Filename:=Lepath & "\Archives.xls"
Range("IQ3001:IU3099").Copy
Windows(lenom).Activate
Range("AM8").Select
ActiveSheet.Paste

' ********************************************************************
' ***   on compare le numéro du top50 à la dernière mise à jour    ***
' ********************************************************************
If Cells(7, 39).Value = top50 Then
    Msg = MsgBox("La mise a jour a déjà été effectuée, la rééditer fausserait les résultats.", 0, "Attention, MaJ non effectuée")
Else
    Cells(7, 39).Value = top50
' ********************************************************************
' ***Comparaison des lignes de records et mise à jour si nécessaire***
' ********************************************************************
' Ligne 8 (29)
    valeur = Cells(8, 40).Value
    If Cells(29, 6).Value >= valeur Then
        If Cells(29, 6).Value = valeur Then
            Cells(8, 39).Value = ")[b][Color=blue][Record égalé : + "
            Cells(8, 41).Value = " ][/color][/b]"
        Else
            Cells(8, 39).Value = ")[b][Color=green][Nouveau record : + "
            Cells(8, 41).Value = " ][/color][/b] (Préc.:" & valeur & " )"
        End If
        Cells(8, 40).Value = Cells(29, 6).Value
        Cells(8, 42).Value = top50
    Else
        Cells(8, 39).Value = ")[b]_____Record : + "
        Cells(8, 41).Value = "[/b]"
    End If
' Ligne 9 (30)
    valeur = Cells(9, 40).Value
    If Cells(30, 6).Value >= valeur Then
        If Cells(30, 6).Value = valeur Then
            Cells(9, 39).Value = ")[b][Color=blue][Record égalé : + "
            Cells(9, 41).Value = " ][/color][/b]"
        Else
            Cells(9, 39).Value = ")[b][Color=green][Nouveau record : + "
            Cells(9, 41).Value = " ][/color][/b] (Préc.:" & Round(valeur * 100, 2) & "%)"
        End If
        Cells(9, 40).Value = Cells(30, 6).Value
        Cells(9, 42).Value = top50
    Else
        Cells(9, 39).Value = ")[b]_____Record : + "
        Cells(9, 41).Value = "[/b]"
    End If
' Ligne 10 (31)
    valeur = Cells(10, 40).Value
    If Cells(31, 6).Value >= valeur Then
        If Cells(31, 6).Value = valeur Then
            Cells(10, 39).Value = ")[b][Color=blue][Record égalé : + "
            Cells(10, 41).Value = " ][/color][/b]"
        Else
            Cells(10, 39).Value = ")[b][Color=green][Nouveau record : + "
            Cells(10, 41).Value = " ][/color][/b] (Préc.:" & valeur & " )"
        End If
        Cells(10, 40).Value = Cells(31, 6).Value
        Cells(10, 42).Value = top50
    Else
        Cells(10, 39).Value = ")[b]_____Record : + "
        Cells(10, 41).Value = "[/b]"
    End If
' Ligne 11 (32)
    valeur = Cells(11, 40).Value
    If Cells(32, 6).Value >= valeur Then
        If Cells(32, 6).Value = valeur Then
            Cells(11, 39).Value = ")[b][Color=blue][Record égalé : + "
            Cells(11, 41).Value = " ][/color][/b]"
        Else
            Cells(11, 39).Value = ")[b][Color=green][Nouveau record : + "
            Cells(11, 41).Value = " ][/color][/b] (Préc.:" & Round(valeur * 100, 2) & "%)"
        End If
        Cells(11, 40).Value = Cells(32, 6).Value
        Cells(11, 42).Value = top50
    Else
        Cells(11, 39).Value = ")[b]_____Record : + "
        Cells(11, 41).Value = "[/b]"
    End If
' Ligne 12 (53)
    valeur = Cells(12, 40).Value
    If Cells(53, 6).Value >= valeur Then
        If Cells(53, 6).Value = valeur Then
            Cells(12, 39).Value = ")[b][Color=blue][Record égalé : + "
            Cells(12, 41).Value = " ][/color][/b]"
        Else
            Cells(12, 39).Value = ")[b][Color=green][Nouveau record : + "
            Cells(12, 41).Value = " ][/color][/b] (Préc.: " & valeur & " )"
        End If
        Cells(12, 40).Value = Cells(53, 6).Value
        Cells(12, 42).Value = top50
    Else
        Cells(12, 39).Value = ")[b]_____Record : + "
        Cells(12, 41).Value = "[/b]"
    End If
' Ligne 13 (54)
    valeur = Cells(13, 40).Value
    If Cells(54, 6).Value >= valeur Then
        If Cells(54, 6).Value = valeur Then
            Cells(13, 39).Value = ")[b][Color=blue][Record égalé : + "
            Cells(13, 41).Value = " ][/color][/b]"
        Else
            Cells(13, 39).Value = ")[b][Color=green][Nouveau record : + "
            Cells(13, 41).Value = " ][/color][/b] (Préc.:" & Round(valeur * 100, 2) & "%)"
        End If
        Cells(13, 40).Value = Cells(54, 6).Value
        Cells(13, 42).Value = top50
    Else
        Cells(13, 39).Value = ")[b]_____Record : + "
        Cells(13, 41).Value = "[/b]"
    End If
' Ligne 14 (55)
    valeur = Cells(14, 40).Value
    If Cells(55, 6).Value >= valeur Then
        If Cells(55, 6).Value = valeur Then
            Cells(14, 39).Value = ")[b][Color=blue][Record égalé : + "
            Cells(14, 41).Value = " ][/color][/b]"
        Else
            Cells(14, 39).Value = ")[b][Color=green][Nouveau record : + "
            Cells(14, 41).Value = " ][/color][/b] (Préc.:" & valeur & " )"
        End If
        Cells(14, 40).Value = Cells(55, 6).Value
        Cells(14, 42).Value = top50
    Else
        Cells(14, 39).Value = ")[b]_____Record : + "
        Cells(14, 41).Value = "[/b]"
    End If
' Ligne 15 (56)
    valeur = Cells(15, 40).Value
    If Cells(56, 6).Value >= valeur Then
        If Cells(56, 6).Value = valeur Then
            Cells(15, 39).Value = ")[b][Color=blue][Record égalé : + "
            Cells(15, 41).Value = " ][/color][/b]"
        Else
            Cells(15, 39).Value = ")[b][Color=green][Nouveau record : + "
            Cells(15, 41).Value = " ][/color][/b] (Préc.:" & Round(valeur * 100, 2) & "%)"
        End If
        Cells(15, 40).Value = Cells(56, 6).Value
        Cells(15, 42).Value = top50
    Else
        Cells(15, 39).Value = ")[b]_____Record : + "
        Cells(15, 41).Value = "[/b]"
    End If
' Ligne 16 (77)
    valeur = Cells(16, 40).Value
    If Cells(77, 6).Value >= valeur Then
        If Cells(77, 6).Value = valeur Then
            Cells(16, 39).Value = ")[b][Color=blue][Record égalé : + "
            Cells(16, 41).Value = " ][/color][/b]"
        Else
            Cells(16, 39).Value = ")[b][Color=green][Nouveau record : + "
            Cells(16, 41).Value = " ][/color][/b] (Préc.:" & valeur & " )"
        End If
        Cells(16, 40).Value = Cells(77, 6).Value
        Cells(16, 42).Value = top50
    Else
        Cells(16, 39).Value = ")[b]_____Record : + "
        Cells(16, 41).Value = "[/b]"
    End If
' Ligne 17 (78)
    valeur = Cells(17, 40).Value
    If Cells(78, 6).Value >= valeur Then
        If Cells(78, 6).Value = valeur Then
            Cells(17, 39).Value = ")[b][Color=blue][Record égalé : + "
            Cells(17, 41).Value = " ][/color][/b]"
        Else
            Cells(17, 39).Value = ")[b][Color=green][Nouveau record : + "
            Cells(17, 41).Value = " ][/color][/b] (Préc.:" & Round(valeur * 100, 2) & "%)"
        End If
        Cells(17, 40).Value = Cells(78, 6).Value
        Cells(17, 42).Value = top50
    Else
        Cells(17, 39).Value = ")[b]_____Record : + "
        Cells(17, 41).Value = "[/b]"
    End If
' Ligne 18 (79)
    valeur = Cells(18, 40).Value
    If Cells(79, 6).Value >= valeur Then
        If Cells(79, 6).Value = valeur Then
            Cells(18, 39).Value = ")[b][Color=blue][Record égalé : + "
            Cells(18, 41).Value = " ][/color][/b]"
        Else
            Cells(18, 39).Value = ")[b][Color=green][Nouveau record : + "
            Cells(18, 41).Value = " ][/color][/b] (Préc.:" & valeur & " )"
        End If
        Cells(18, 40).Value = Cells(79, 6).Value
        Cells(18, 42).Value = top50
    Else
        Cells(18, 39).Value = ")[b]_____Record : + "
        Cells(18, 41).Value = "[/b]"
    End If
' Ligne 19 (80)
    valeur = Cells(19, 40).Value
    If Cells(80, 6).Value >= valeur Then
        If Cells(80, 6).Value = valeur Then
            Cells(19, 39).Value = ")[b][Color=blue][Record égalé : + "
            Cells(19, 41).Value = " ][/color][/b]"
        Else
            Cells(19, 39).Value = ")[b][Color=green][Nouveau record : + "
            Cells(19, 41).Value = " ][/color][/b] (Préc.:" & Round(valeur * 100, 2) & "%)"
        End If
        Cells(19, 40).Value = Cells(80, 6).Value
        Cells(19, 42).Value = top50
    Else
        Cells(19, 39).Value = ")[b]_____Record : + "
        Cells(19, 41).Value = "[/b]"
    End If
' Ligne 20 (132)
    valeur = Cells(20, 40).Value
    If Cells(132, 4).Value >= valeur Then
        If Cells(132, 4).Value = valeur Then
            Cells(20, 39).Value = "[b][Color=blue][Record égalé : + "
            Cells(20, 41).Value = " ][/color][/b]"
        Else
            Cells(20, 39).Value = "[b][Color=green][Nouveau record : + "
            Cells(20, 41).Value = " ][/color][/b] (Préc.:" & Round(valeur, 2) & " )"
        End If
        Cells(20, 40).Value = Cells(132, 4).Value
        Cells(20, 42).Value = top50
    Else
        Cells(20, 39).Value = "[b]_____Record : + "
        Cells(20, 41).Value = "[/b]"
    End If
' Ligne 21 (133)
    valeur = Cells(21, 40).Value
    If Cells(133, 5).Value >= valeur Then
        If Cells(133, 5).Value = valeur Then
            Cells(21, 39).Value = "[b][Color=blue][Record égalé : + "
            Cells(21, 41).Value = " ][/color][/b]"
        Else
            Cells(21, 39).Value = "[b][Color=green][Nouveau record : + "
            Cells(21, 41).Value = " ][/color][/b] (Préc.:" & valeur & " )"
        End If
        Cells(21, 40).Value = Cells(133, 5).Value
        Cells(21, 42).Value = top50
    Else
        Cells(21, 39).Value = "[b]_____Record : + "
        Cells(21, 41).Value = "[/b]"
    End If
' Ligne 22 (134)
    valeur = Cells(22, 40).Value
    If Cells(134, 5).Value >= valeur Then
        If Cells(134, 5).Value = valeur Then
            Cells(22, 39).Value = "[b][Color=blue][Record égalé : + "
            Cells(22, 41).Value = " ][/color][/b]"
        Else
            Cells(22, 39).Value = "[b][Color=green][Nouveau record : + "
            Cells(22, 41).Value = " ][/color][/b] (Préc.:" & valeur & " )"
        End If
        Cells(22, 40).Value = Cells(134, 5).Value
        Cells(22, 42).Value = top50
    Else
        Cells(22, 39).Value = "[b]_____Record : + "
        Cells(22, 41).Value = "[/b]"
    End If
' Ligne 23 (164)
    valeur = Cells(23, 40).Value
    If Cells(164, 4).Value >= valeur Then
        If Cells(164, 4).Value = valeur Then
            Cells(23, 39).Value = "[b][Color=blue][Record égalé : + "
            Cells(23, 41).Value = " ][/color][/b]"
        Else
            Cells(23, 39).Value = "[b][Color=green][Nouveau record : + "
            Cells(23, 41).Value = " ][/color][/b] (Préc.:" & Round(valeur, 2) & " )"
        End If
        Cells(23, 40).Value = Cells(164, 4).Value
        Cells(23, 42).Value = top50
    Else
        Cells(23, 39).Value = "[b]_____Record : + "
        Cells(23, 41).Value = "[/b]"
    End If
' Ligne 24 (165)
    valeur = Cells(24, 40).Value
    If Cells(165, 5).Value >= valeur Then
        If Cells(165, 5).Value = valeur Then
            Cells(24, 39).Value = "[b][Color=blue][Record égalé : + "
            Cells(24, 41).Value = " ][/color][/b]"
        Else
            Cells(24, 39).Value = "[b][Color=green][Nouveau record : + "
            Cells(24, 41).Value = " ][/color][/b] (Préc.:" & valeur & " )"
        End If
        Cells(24, 40).Value = Cells(165, 5).Value
        Cells(24, 42).Value = top50
    Else
        Cells(24, 39).Value = "[b]_____Record : + "
        Cells(24, 41).Value = "[/b]"
    End If
' Ligne 25 (166)
    valeur = Cells(25, 40).Value
    If Cells(166, 5).Value >= valeur Then
        If Cells(166, 5).Value = valeur Then
            Cells(25, 39).Value = "[b][Color=blue][Record égalé : + "
            Cells(25, 41).Value = " ][/color][/b]"
        Else
            Cells(25, 39).Value = "[b][Color=green][Nouveau record : + "
            Cells(25, 41).Value = " ][/color][/b] (Préc.:" & valeur & " )"
        End If
        Cells(25, 40).Value = Cells(166, 5).Value
        Cells(25, 42).Value = top50
    Else
        Cells(25, 39).Value = "[b]_____Record : + "
        Cells(25, 41).Value = "[/b]"
    End If
' Ligne 26 (196)
    valeur = Cells(26, 40).Value
    If Cells(196, 4).Value >= valeur Then
        If Cells(196, 4).Value = valeur Then
            Cells(26, 39).Value = "[b][Color=blue][Record égalé : + "
            Cells(26, 41).Value = " ][/color][/b]"
        Else
            Cells(26, 39).Value = "[b][Color=green][Nouveau record : + "
            Cells(26, 41).Value = " ][/color][/b] (Préc.:" & Round(valeur, 2) & " )"
        End If
        Cells(26, 40).Value = Cells(196, 4).Value
        Cells(26, 42).Value = top50
    Else
        Cells(26, 39).Value = "[b]_____Record : + "
        Cells(26, 41).Value = "[/b]"
    End If
' Ligne 27 (197)
    valeur = Cells(27, 40).Value
    If Cells(197, 5).Value >= valeur Then
        If Cells(197, 5).Value = valeur Then
            Cells(27, 39).Value = "[b][Color=blue][Record égalé : + "
            Cells(27, 41).Value = " ][/color][/b]"
        Else
            Cells(27, 39).Value = "[b][Color=green][Nouveau record : + "
            Cells(27, 41).Value = " ][/color][/b] (Préc.: " & valeur & " )"
        End If
        Cells(27, 40).Value = Cells(197, 5).Value
        Cells(27, 42).Value = top50
    Else
        Cells(27, 39).Value = "[b]_____Record : + "
        Cells(27, 41).Value = "[/b]"
    End If
' Ligne 28 (198)
    valeur = Cells(28, 41).Value
    qui = Cells(28, 40).Value
    If Cells(198, 5).Value <= valeur Then
        If Cells(198, 5).Value = valeur Then
            Cells(28, 39).Value = "[b][Color=blue][Record égalé : "
            Cells(28, 42).Value = " ][/color][/b]"
        Else
            Cells(28, 39).Value = "[b][Color=green][Nouveau record : "
            Cells(28, 42).Value = " ][/color][/b] (Préc.:" & qui & Round(valeur, 0) & " )"
        End If
        Cells(28, 40).Value = Cells(198, 4).Value & " avec + "
        Cells(28, 41).Value = Cells(198, 5).Value
        Cells(28, 43).Value = top50
    Else
        Cells(28, 39).Value = "[b]_____Record : "
        Cells(28, 42).Value = "[/b]"
    End If
' Ligne 29 (199)
    valeur = Cells(29, 41).Value
    qui = Cells(29, 40).Value
    If Cells(199, 5).Value <= valeur Then
        If Cells(199, 5).Value = valeur Then
            Cells(29, 39).Value = "[b][Color=blue][Record égalé : "
            Cells(29, 42).Value = " ][/color][/b]"
        Else
            Cells(29, 39).Value = "[b][Color=green][Nouveau record : "
            Cells(29, 42).Value = " ][/color][/b] (Préc.:" & qui & Round(valeur, 0) & " )"
        End If
        Cells(29, 40).Value = Cells(199, 4).Value & " avec "
        Cells(29, 41).Value = Cells(199, 5).Value
        Cells(29, 43).Value = top50
    Else
        Cells(29, 39).Value = "[b]_____Record : "
        Cells(29, 42).Value = "[/b]"
    End If
' Ligne 31 (69b)
    valeur = Cells(31, 40).Value
    If Cells(69, 21).Value >= valeur Then
        If Cells(69, 21).Value = valeur Then
            Cells(31, 39).Value = ")[b][Color=blue][Record égalé : + "
            Cells(31, 41).Value = " ][/color][/b]"
        Else
            Cells(31, 39).Value = ")[b][Color=green][Nouveau record : + "
            Cells(31, 41).Value = " ][/color][/b] (Préc.:" & valeur & " )"
        End If
        Cells(31, 40).Value = Cells(69, 21).Value
        Cells(31, 42).Value = top50
    Else
        Cells(31, 39).Value = ")[b]_____Record : + "
        Cells(31, 41).Value = "[/b]"
    End If
' Ligne 32 (70b)
    valeur = Cells(32, 40).Value
    If Cells(70, 21).Value >= valeur Then
        If Cells(70, 21).Value = valeur Then
            Cells(32, 39).Value = ")[b][Color=blue][Record égalé : + "
            Cells(32, 41).Value = " ][/color][/b]"
        Else
            Cells(32, 39).Value = ")[b][Color=green][Nouveau record : + "
            Cells(32, 41).Value = " ][/color][/b] (Préc.:" & Round(valeur * 100, 2) & "%)"
        End If
        Cells(32, 40).Value = Cells(70, 21).Value
        Cells(32, 42).Value = top50
    Else
        Cells(32, 39).Value = ")[b]_____Record : + "
        Cells(32, 41).Value = "[/b]"
    End If
' Ligne 33 (71b)
    valeur = Cells(33, 40).Value
    If Cells(71, 19).Value >= valeur Then
        If Cells(71, 19).Value = valeur Then
            Cells(33, 39).Value = "[b][Color=blue][Record égalé : "
            Cells(33, 41).Value = " ][/color][/b]"
        Else
            Cells(33, 39).Value = "[b][Color=green][Nouveau record : "
            Cells(33, 41).Value = " ][/color][/b] (Préc.:" & Round(valeur, 2) & " )"
        End If
        Cells(33, 40).Value = Cells(71, 19).Value
        Cells(33, 42).Value = top50
    Else
        Cells(33, 39).Value = "[b]_____Record : "
        Cells(33, 41).Value = "[/b]"
    End If
' Ligne 34 (73b)
    valeur = Cells(34, 41).Value
    qui = Cells(34, 40).Value
    If Cells(73, 20).Value >= valeur Then
        If Cells(73, 20).Value = valeur Then
            Cells(34, 39).Value = "[b][Color=blue][Record égalé : "
            Cells(34, 42).Value = " ][/color][/b]"
        Else
            Cells(34, 39).Value = "[b][Color=green][Nouveau record : "
            Cells(34, 42).Value = " ][/color][/b] (Préc.:" & qui & valeur & " )"
        End If
        Cells(34, 40).Value = Cells(73, 18).Value & " avec + "
        Cells(34, 41).Value = Cells(73, 20).Value
        Cells(34, 43).Value = top50
    Else
        Cells(34, 39).Value = "[b]_____Record : "
        Cells(34, 42).Value = "[/b]"
    End If
' Ligne 35 (74b)
    valeur = Cells(35, 41).Value
    qui = Cells(35, 40).Value
    If Cells(74, 20).Value >= valeur Then
        If Cells(74, 20).Value = valeur Then
            Cells(35, 39).Value = "[b][Color=blue][Record égalé : "
            Cells(35, 42).Value = " ][/color][/b]"
        Else
            Cells(35, 39).Value = "[b][Color=green][Nouveau record : "
            Cells(35, 42).Value = " ][/color][/b] (Préc.:" & qui & valeur & " )"
        End If
        Cells(35, 40).Value = Cells(74, 18).Value & " avec + "
        Cells(35, 41).Value = Cells(74, 20).Value
        Cells(35, 43).Value = top50
    Else
        Cells(35, 39).Value = "[b]_____Record : "
        Cells(35, 42).Value = "[/b]"
    End If
' Ligne 36 (75b)
    valeur = Cells(36, 41).Value
    qui = Cells(36, 40).Value
    If Cells(75, 20).Value >= valeur Then
        If Cells(75, 20).Value = valeur Then
            Cells(36, 39).Value = "[b][Color=blue][Record égalé : "
            Cells(36, 42).Value = " ][/color][/b]"
        Else
            Cells(36, 39).Value = "[b][Color=green][Nouveau record : "
            Cells(36, 42).Value = " ][/color][/b] (Préc.:" & qui & valeur & " )"
        End If
        Cells(36, 40).Value = Cells(75, 18).Value & " avec + "
        Cells(36, 41).Value = Cells(75, 20).Value
        Cells(36, 43).Value = top50
    Else
        Cells(36, 39).Value = "[b]_____Record : "
        Cells(36, 42).Value = "[/b]"
    End If
' Ligne 37 (76b)
    valeur = Cells(37, 40).Value
    qui = Cells(37, 42).Value
    If Cells(76, 20).Value >= valeur Then
        If Cells(76, 20).Value = valeur Then
            Cells(37, 39).Value = "[b][Color=blue][Record égalé : "
            Cells(37, 43).Value = top50
        Else
            Cells(37, 39).Value = "[b][Color=green][Nouveau record : "
            Cells(37, 43).Value = top50 & " (Préc.:" & valeur & qui & " )"
        End If
        Cells(37, 40).Value = Cells(76, 20).Value
        Cells(37, 41).Value = " places ][/color][/b]"
        Cells(37, 42).Value = Cells(76, 21).Value
    Else
        Cells(37, 39).Value = "[b]_____Record : "
        Cells(37, 41).Value = " places [/b]"
    End If
' Ligne 38 (77b)
    valeur = Cells(38, 40).Value
    qui = Cells(38, 42).Value
    If Cells(77, 20).Value >= valeur Then
        If Cells(77, 20).Value = valeur Then
            Cells(38, 39).Value = "[b][Color=blue][Record égalé : "
            Cells(38, 43).Value = top50
        Else
            Cells(38, 39).Value = "[b][Color=green][Nouveau record : "
            Cells(38, 43).Value = top50 & " (Préc.:" & valeur & qui & " )"
        End If
        Cells(38, 40).Value = Cells(77, 20).Value
        Cells(38, 41).Value = " places ][/color][/b]"
        Cells(38, 42).Value = Cells(77, 21).Value
    Else
        Cells(38, 39).Value = "[b]_____Record : "
        Cells(38, 41).Value = " places [/b]"
    End If
' Ligne 39 (78b)
    valeur = Cells(39, 40).Value
    If Cells(78, 20).Value >= valeur Then
        If Cells(78, 20).Value = valeur Then
            Cells(39, 39).Value = "[b][Color=blue][Record égalé : "
            Cells(39, 41).Value = " ][/color][/b]"
        Else
            Cells(39, 39).Value = "[b][Color=green][Nouveau record : "
            Cells(39, 41).Value = " ][/color][/b] (Préc.:" & valeur & " )"
        End If
        Cells(39, 40).Value = Cells(78, 20).Value
        Cells(39, 42).Value = top50
    Else
        Cells(39, 39).Value = "[b]_____Record : "
        Cells(39, 41).Value = "[/b]"
    End If
' Ligne 40 (79b)
    valeur = Cells(40, 40).Value
    If Cells(79, 20).Value >= valeur Then
        If Cells(79, 20).Value = valeur Then
            Cells(40, 39).Value = "[b][Color=blue][Record égalé : "
            Cells(40, 41).Value = " ][/color][/b]"
        Else
            Cells(40, 39).Value = "[b][Color=green][Nouveau record : "
            Cells(40, 41).Value = " ][/color][/b] (Préc.:" & valeur & " )"
        End If
        Cells(40, 40).Value = Cells(79, 20).Value
        Cells(40, 42).Value = top50
    Else
        Cells(40, 39).Value = "[b]_____Record : "
        Cells(40, 41).Value = "[/b]"
    End If
' Ligne 41 (80b)
    valeur = Cells(41, 40).Value
    If Cells(80, 20).Value >= valeur Then
        If Cells(80, 20).Value = valeur Then
            Cells(41, 39).Value = "[b][Color=blue][Record égalé : "
            Cells(41, 41).Value = " ][/color][/b]"
        Else
            Cells(41, 39).Value = "[b][Color=green][Nouveau record : "
            Cells(41, 41).Value = " ][/color][/b] (Préc.:" & valeur & " )"
        End If
        Cells(41, 40).Value = Cells(80, 20).Value
        Cells(41, 42).Value = top50
    Else
        Cells(41, 39).Value = "[b]_____Record : "
        Cells(41, 41).Value = "[/b]"
    End If
' Ligne 42 (81b)
    valeur = Cells(42, 40).Value
    If Cells(81, 20).Value <= valeur Then
        If Cells(81, 20).Value = valeur Then
            Cells(42, 39).Value = "[b][Color=blue][Record égalé, Min : "
            Cells(42, 41).Value = " ][/color][/b]"
        Else
            Cells(42, 39).Value = "[b][Color=green][Nouveau record, Min : "
            Cells(42, 41).Value = " ][/color][/b] (Préc.:" & valeur & " )"
        End If
        Cells(42, 40).Value = Cells(81, 20).Value
        Cells(42, 42).Value = top50
    Else
        Cells(42, 39).Value = "[b]_____Min : "
        Cells(42, 41).Value = "[/b]"
    End If
' Ligne 43 (82b)
    valeur = Cells(43, 40).Value
    If Cells(81, 20).Value >= valeur Then
        If Cells(81, 20).Value = valeur Then
            Cells(43, 39).Value = "[b][Color=blue][Record égalé, Max : "
            Cells(43, 41).Value = " ][/color][/b]"
        Else
            Cells(43, 39).Value = "[b][Color=green][Nouveau record, Max : "
            Cells(43, 41).Value = " ][/color][/b] (Préc.:" & valeur & " )"
        End If
        Cells(43, 40).Value = Cells(82, 20).Value
        Cells(43, 42).Value = top50
    Else
        Cells(43, 39).Value = "[b]_____Max : "
        Cells(43, 41).Value = "[/b]"
    End If
' Ligne 44 (83b)
    valeur = Cells(44, 40).Value
    If Cells(83, 20).Value >= valeur Then
        If Cells(83, 20).Value = valeur Then
            Cells(44, 39).Value = "[b][Color=blue][Record égalé : "
            Cells(44, 41).Value = " ][/color][/b]"
        Else
            Cells(44, 39).Value = "[b][Color=green][Nouveau record : "
            Cells(44, 41).Value = " ][/color][/b] (Préc.:" & valeur & " )"
        End If
        Cells(44, 40).Value = Cells(83, 20).Value
        Cells(44, 42).Value = top50
    Else
        Cells(44, 39).Value = "[b]_____Record : "
        Cells(44, 41).Value = "[/b]"
    End If
' Ligne 45 (84b)
    valeur = Cells(45, 40).Value
    If Cells(84, 20).Value >= valeur Then
        If Cells(84, 20).Value = valeur Then
            Cells(45, 39).Value = "[b][Color=blue][Record égalé : "
            Cells(45, 41).Value = " ][/color][/b]"
        Else
            Cells(45, 39).Value = "[b][Color=green][Nouveau record : "
            Cells(45, 41).Value = " ][/color][/b] (Préc.:" & valeur & " )"
        End If
        Cells(45, 40).Value = Cells(84, 20).Value
        Cells(45, 42).Value = top50
    Else
        Cells(45, 39).Value = "[b]_____Record : "
        Cells(45, 41).Value = "[/b]"
    End If
' Ligne 46 (85b)
    valeur = Cells(46, 40).Value
    If Cells(85, 20).Value >= valeur Then
        If Cells(85, 20).Value = valeur Then
            Cells(46, 39).Value = "[b][Color=blue][Record égalé : "
            Cells(46, 41).Value = " ][/color][/b]"
        Else
            Cells(46, 39).Value = "[b][Color=green][Nouveau record : "
            Cells(46, 41).Value = " ][/color][/b] (Préc.:" & valeur & " )"
        End If
        Cells(46, 40).Value = Cells(85, 20).Value
        Cells(46, 42).Value = top50
    Else
        Cells(46, 39).Value = "[b]_____Record : "
        Cells(46, 41).Value = "[/b]"
    End If
' Ligne 47 (88b)
    valeur = Cells(47, 40).Value
    If Cells(88, 19).Value <= valeur Then
        If Cells(88, 19).Value = valeur Then
            Cells(47, 39).Value = "[b][Color=blue][Record égalé : "
            Cells(47, 41).Value = " ][/color][/b]"
        Else
            Cells(47, 39).Value = "[b][Color=green][Nouveau record : "
            Cells(47, 41).Value = " ][/color][/b] (Préc.:" & Round(valeur, 2) & " )"
        End If
        Cells(47, 40).Value = Cells(88, 19).Value
        Cells(47, 42).Value = top50
    Else
        Cells(47, 39).Value = "[b]_____Record : "
        Cells(47, 41).Value = "[/b]"
    End If
' Ligne 48 (89b)
    valeur = Cells(48, 41).Value
    qui = Cells(48, 40).Value
    If Cells(89, 20).Value <= valeur Then
        If Cells(89, 20).Value = valeur Then
            Cells(48, 39).Value = "[b][Color=blue][Record égalé : "
            Cells(48, 42).Value = " ][/color][/b]"
        Else
            Cells(48, 39).Value = "[b][Color=green][Nouveau record : "
            Cells(48, 42).Value = " ][/color][/b] (Préc.:" & qui & Round(valeur, 2) & " )"
        End If
        Cells(48, 40).Value = Cells(89, 19).Value & " avec + "
        Cells(48, 41).Value = Cells(89, 20).Value
        Cells(48, 43).Value = top50
    Else
        Cells(48, 39).Value = "[b]_____Record : "
        Cells(48, 42).Value = "[/b]"
    End If
' Ligne 49 (90b)
    valeur = Cells(49, 41).Value
    qui = Cells(49, 40).Value
    If Cells(90, 20).Value <= valeur Then
        If Cells(90, 20).Value = valeur Then
            Cells(49, 39).Value = "[b][Color=blue][Record égalé : "
            Cells(49, 42).Value = " ][/color][/b]"
        Else
            Cells(49, 39).Value = "[b][Color=green][Nouveau record : "
            Cells(49, 42).Value = " ][/color][/b] (Préc.:" & qui & Round(valeur, 2) & " )"
        End If
        Cells(49, 40).Value = Cells(90, 19).Value & " avec "
        Cells(49, 41).Value = Cells(90, 20).Value
        Cells(49, 43).Value = top50
    Else
        Cells(49, 39).Value = "[b]_____Record : "
        Cells(49, 42).Value = "[/b]"
    End If
' Ligne 50 (129b)
    valeur = Cells(50, 40).Value
    If Cells(129, 21).Value >= valeur Then
        If Cells(129, 21).Value = valeur Then
            Cells(50, 39).Value = ")[b][Color=blue][Record égalé : + "
            Cells(50, 41).Value = " ][/color][/b]"
        Else
            Cells(50, 39).Value = ")[b][Color=green][Nouveau record : + "
            Cells(50, 41).Value = " ][/color][/b] (Préc.:" & valeur & " )"
        End If
        Cells(50, 40).Value = Cells(129, 21).Value
        Cells(50, 42).Value = top50
    Else
        Cells(50, 39).Value = ")[b]_____Record : + "
        Cells(50, 41).Value = "[/b]"
    End If
' Ligne 51 (130b)
    valeur = Cells(51, 40).Value
    If Cells(130, 21).Value >= valeur Then
        If Cells(130, 21).Value = valeur Then
            Cells(51, 39).Value = ")[b][Color=blue][Record égalé : + "
            Cells(51, 41).Value = " ][/color][/b]"
        Else
            Cells(51, 39).Value = ")[b][Color=green][Nouveau record : + "
            Cells(51, 41).Value = " ][/color][/b] (Préc.:" & Round(valeur * 100, 2) & "%)"
        End If
        Cells(51, 40).Value = Cells(130, 21).Value
        Cells(51, 42).Value = top50
    Else
        Cells(51, 39).Value = ")[b]_____Record : + "
        Cells(51, 41).Value = "[/b]"
    End If
' Ligne 52 (131b)
    valeur = Cells(52, 40).Value
    If Cells(131, 21).Value >= valeur Then
        If Cells(131, 21).Value = valeur Then
            Cells(52, 39).Value = ")[b][Color=blue][Record égalé : + "
            Cells(52, 41).Value = " ][/color][/b]"
        Else
            Cells(52, 39).Value = ")[b][Color=green][Nouveau record : + "
            Cells(52, 41).Value = " ][/color][/b] (Préc.:" & valeur & " )"
        End If
        Cells(52, 40).Value = Cells(131, 21).Value
        Cells(52, 42).Value = top50
    Else
        Cells(52, 39).Value = ")[b]_____Record : + "
        Cells(52, 41).Value = "[/b]"
    End If
' Ligne 53 (132b)
    valeur = Cells(53, 40).Value
    If Cells(132, 21).Value >= valeur Then
        If Cells(132, 21).Value = valeur Then
            Cells(53, 39).Value = ")[b][Color=blue][Record égalé : + "
            Cells(53, 41).Value = " ][/color][/b]"
        Else
            Cells(53, 39).Value = ")[b][Color=green][Nouveau record : + "
            Cells(53, 41).Value = " ][/color][/b] (Préc.:" & Round(valeur * 100, 2) & " )"
        End If
        Cells(53, 40).Value = Cells(132, 21).Value
        Cells(53, 42).Value = top50
    Else
        Cells(53, 39).Value = ")[b]_____Record : + "
        Cells(53, 41).Value = "[/b]"
    End If
' Ligne 54 (133b)
    valeur = Cells(54, 40).Value
    If Cells(133, 20).Value >= valeur Then
        If Cells(133, 20).Value = valeur Then
            Cells(54, 39).Value = "[b][Color=blue][Record égalé : "
            Cells(54, 41).Value = " ][/color][/b]"
        Else
            Cells(54, 39).Value = "[b][Color=green][Nouveau record : "
            Cells(54, 41).Value = " ][/color][/b] (Préc.:" & valeur & " )"
        End If
        Cells(54, 40).Value = Cells(133, 20).Value
        Cells(54, 42).Value = top50
    Else
        Cells(54, 39).Value = "[b]_____Record : "
        Cells(54, 41).Value = "[/b]"
    End If
' Ligne 55 (134b)
    valeur = Cells(55, 40).Value
    qui = Cells(55, 42).Value
    If Cells(134, 20).Value >= valeur Then
        If Cells(134, 20).Value = valeur Then
            Cells(55, 39).Value = "[b][Color=blue][Record égalé : "
            Cells(55, 41).Value = " hits ][/color][/b](Préc." & qui & ")"
            Else
            Cells(55, 39).Value = "[b][Color=green][Nouveau record : "
            Cells(55, 41).Value = " hits ][/color][/b] (Préc.:" & valeur & " hits " & qui & " )"
        End If
        Cells(55, 40).Value = Cells(134, 20).Value
        Cells(55, 42).Value = Cells(134, 21).Value
        Cells(55, 43).Value = top50
        Else
        Cells(55, 39).Value = "[b]_____Record : "
        Cells(55, 41).Value = " hits [/b]"
    End If
' Ligne 56 (135b)
    valeur = Cells(56, 40).Value
    If Cells(135, 20).Value >= valeur Then
        If Cells(135, 20).Value = valeur Then
            Cells(56, 39).Value = "[b][Color=blue][Record égalé : "
            Cells(56, 41).Value = " ][/color][/b]"
        Else
            Cells(56, 39).Value = "[b][Color=green][Nouveau record : "
            Cells(56, 41).Value = " ][/color][/b] (Préc.:" & valeur & " )"
        End If
        Cells(56, 40).Value = Cells(135, 20).Value
        Cells(56, 42).Value = top50
    Else
        Cells(56, 39).Value = "[b]_____Record : "
        Cells(56, 41).Value = "[/b]"
    End If
' Ligne 57 (165b)
    valeur = Cells(57, 40).Value
    If Cells(165, 20).Value >= valeur Then
        If Cells(165, 20).Value = valeur Then
            Cells(57, 39).Value = "[b][Color=blue][Record égalé : "
            Cells(57, 41).Value = " ][/color][/b]"
        Else
            Cells(57, 39).Value = "[b][Color=green][Nouveau record : "
            Cells(57, 41).Value = " ][/color][/b] (Préc.:" & valeur & " )"
        End If
        Cells(57, 40).Value = Cells(165, 20).Value
        Cells(57, 42).Value = top50
    Else
        Cells(57, 39).Value = "[b]_____Record : "
        Cells(57, 41).Value = "[/b]"
    End If
' Ligne 58 (166b)
    valeur = Cells(58, 40).Value
    If Cells(166, 20).Value >= valeur Then
        If Cells(166, 20).Value = valeur Then
            Cells(58, 39).Value = "[b][Color=blue][Record égalé : "
            Cells(58, 41).Value = " ][/color][/b]"
        Else
            Cells(58, 39).Value = "[b][Color=green][Nouveau record : "
            Cells(58, 41).Value = " ][/color][/b] (Préc.:" & valeur & " )"
        End If
        Cells(58, 40).Value = Cells(166, 20).Value
        Cells(58, 42).Value = top50
    Else
        Cells(58, 39).Value = "[b]_____Record : "
        Cells(58, 41).Value = "[/b]"
    End If
' Ligne 59 (167b)
    valeur = Cells(59, 40).Value
    If Cells(167, 20).Value >= valeur Then
        If Cells(167, 20).Value = valeur Then
            Cells(59, 39).Value = "[b][Color=blue][Record égalé : "
            Cells(59, 41).Value = " ][/color][/b]"
        Else
            Cells(59, 39).Value = "[b][Color=green][Nouveau record : "
            Cells(59, 41).Value = " ][/color][/b] (Préc.:" & valeur & " )"
        End If
        Cells(59, 40).Value = Cells(167, 20).Value
        Cells(59, 42).Value = top50
    Else
        Cells(59, 39).Value = "[b]_____Record : "
        Cells(59, 41).Value = "[/b]"
    End If
' Ligne 60 (168b)
    valeur = Cells(60, 40).Value
    If Cells(168, 20).Value >= valeur Then
        If Cells(168, 20).Value = valeur Then
            Cells(60, 39).Value = "[b][Color=blue][Record égalé : "
            Cells(60, 41).Value = " ][/color][/b]"
        Else
            Cells(60, 39).Value = "[b][Color=green][Nouveau record : "
            Cells(60, 41).Value = " ][/color][/b] (Préc.:" & valeur & " )"
        End If
        Cells(60, 40).Value = Cells(168, 20).Value
        Cells(60, 42).Value = top50
    Else
        Cells(60, 39).Value = "[b]_____Record : "
        Cells(60, 41).Value = "[/b]"
    End If
' Ligne 61 (241b)
    valeur = Cells(61, 40).Value
    If Cells(241, 21).Value >= valeur Then
        If Cells(241, 21).Value = valeur Then
            Cells(61, 39).Value = ")[b][Color=blue][Record égalé : + "
            Cells(61, 41).Value = " ][/color][/b]"
        Else
            Cells(61, 39).Value = ")[b][Color=green][Nouveau record : + "
            Cells(61, 41).Value = " ][/color][/b] (Préc.:" & valeur & " )"
        End If
        Cells(61, 40).Value = Cells(231, 21).Value
        Cells(61, 42).Value = top50
    Else
        Cells(61, 39).Value = ")[b]_____Record : + "
        Cells(61, 41).Value = "[/b]"
    End If
' Ligne 62 (242b)
    valeur = Cells(62, 40).Value
    If Cells(242, 21).Value >= valeur Then
        If Cells(242, 21).Value = valeur Then
            Cells(62, 39).Value = ")[b][Color=blue][Record égalé : + "
            Cells(62, 41).Value = " ][/color][/b]"
        Else
            Cells(62, 39).Value = ")[b][Color=green][Nouveau record : + "
            Cells(62, 41).Value = " ][/color][/b] (Préc.:" & Round(valeur * 100, 2) & ")"
        End If
        Cells(62, 40).Value = Cells(232, 21).Value
        Cells(62, 42).Value = top50
    Else
        Cells(62, 39).Value = ")[b]_____Record : + "
        Cells(62, 41).Value = "[/b]"
    End If
' Ligne 63 (243b)
    valeur = Cells(63, 40).Value
    If Cells(243, 21).Value >= valeur Then
        If Cells(243, 21).Value = valeur Then
            Cells(63, 39).Value = ")[b][Color=blue][Record égalé : + "
            Cells(63, 41).Value = " ][/color][/b]"
        Else
            Cells(63, 39).Value = ")[b][Color=green][Nouveau record : + "
            Cells(63, 41).Value = " ][/color][/b] (Préc.:" & valeur & " )"
        End If
        Cells(63, 40).Value = Cells(233, 21).Value
        Cells(63, 42).Value = top50
    Else
        Cells(63, 39).Value = ")[b]_____Record : + "
        Cells(63, 41).Value = "[/b]"
    End If
' Ligne 64 (244b)
    valeur = Cells(64, 40).Value
    If Cells(244, 21).Value >= valeur Then
        If Cells(244, 21).Value = valeur Then
            Cells(64, 39).Value = ")[b][Color=blue][Record égalé : + "
            Cells(64, 41).Value = " ][/color][/b]"
        Else
            Cells(64, 39).Value = ")[b][Color=green][Nouveau record : + "
            Cells(64, 41).Value = " ][/color][/b] (Préc.:" & Round(valeur * 100, 2) & " )"
        End If
        Cells(64, 40).Value = Cells(234, 21).Value
        Cells(64, 42).Value = top50
    Else
        Cells(64, 39).Value = ")[b]_____Record : + "
        Cells(64, 41).Value = "[/b]"
    End If
' Ligne 65 (245b)
    valeur = Cells(65, 40).Value
    qui = Cells(65, 42).Value
    If Cells(245, 20).Value >= valeur Then
        If Cells(245, 20).Value = valeur Then
            Cells(65, 39).Value = "[b][Color=blue][Record égalé : "
            Cells(65, 41).Value = " filleuls ][/color][/b](Préc." & qui & ")"
            Else
            Cells(65, 39).Value = "[b][Color=green][Nouveau record : "
            Cells(65, 41).Value = " filleuls ][/color][/b] (Préc.:" & valeur & " hits " & qui & ")"
        End If
        Cells(65, 40).Value = Cells(235, 20).Value
        Cells(65, 42).Value = Cells(235, 21).Value
        Cells(65, 43).Value = top50
        Else
        Cells(65, 39).Value = "[b]_____Record : "
        Cells(65, 41).Value = " filleuls [/b]"
    End If
' Ligne 66 (246b)
    valeur = Cells(66, 40).Value
    If Cells(246, 20).Value >= valeur Then
        If Cells(246, 20).Value = valeur Then
            Cells(66, 39).Value = "[b][Color=blue][Record égalé : "
            Cells(66, 41).Value = " ][/color][/b]"
        Else
            Cells(66, 39).Value = "[b][Color=green][Nouveau record : "
            Cells(66, 41).Value = " ][/color][/b] (Préc.:" & valeur & " )"
        End If
        Cells(66, 40).Value = Cells(236, 20).Value
        Cells(66, 42).Value = top50
    Else
        Cells(66, 39).Value = "[b]_____Record : "
        Cells(66, 41).Value = "[/b]"
    End If

' ********************************************************************
' *** on compare le nom des trois premiers à la liste              ***
' *** Si le nom fait partie de la liste, on incrémente la valeur   ***
' *** Sinon on ajoute le nom à la liste et on incrémente           ***_
' ********************************************************************
    For i = 1 To 3
        Nomtop = Cells(72 + i, 18).Value
        Lignetop = 0
        For j = 1 To 40
'           On enlève la couleur verte si il y a eu un nouveau dans la liste la semaine dernière
            If i = 1 Then
                Cells(172 + j, 17).Value = ""
                Cells(172 + j, 20).Value = ""
                Cells(172 + j, 23).Value = ""
                Cells(172 + j, 25).Value = ""
            End If
            If LCase(Cells(66 + j, 40).Value) = LCase(Nomtop) Then Lignetop = j
'           Si le nom n'est pas trouvé on met la ligne en vert
            If Cells(66 + j, 40).Value = "" And Lignetop = 0 Then
                Cells(172 + j, 17).Value = "[b][Color=Green]"
                Cells(172 + j, 25).Value = "[/color][/b]"
                Cells(66 + j, 40).Value = Nomtop
                Lignetop = 1
                If i = 1 Then Cells(66 + j, 41).Value = Cells(66 + j, 41).Value + 1
                Cells(66 + j, 42).Value = Cells(66 + j, 42).Value + 1
            End If
'           On incrémente les valeurs sur le nom retrouvé
            If i = 1 And Lignetop = j Then
                Cells(66 + j, 41).Value = Cells(66 + j, 41).Value + 1
                Cells(66 + j, 42).Value = Cells(66 + j, 42).Value + 1
                Cells(172 + j, 20).Value = "[Color=Green]"
                Cells(172 + j, 25).Value = "[/color]"
            End If
            If i <> 1 And Lignetop = j Then
                Cells(66 + j, 42).Value = Cells(66 + j, 42).Value + 1
                Cells(172 + j, 23).Value = "[Color=Green]"
                Cells(172 + j, 25).Value = "[/color]"
            End If
        Next
    Next

' ********************************************************************
' ***        on copie le classement dans la zone daffichage        ***
' ***   on retrie la liste selon le nombre 1 er / Nombre podium    ***
' ********************************************************************
    Range("AN67:AN106").Copy
    Range("R173").Select
    ActiveSheet.Paste
    Range("AO67:AO106").Copy
    Range("U173").Select
    ActiveSheet.Paste
    Range("AP67:AP106").Copy
    Range("X173").Select
    ActiveSheet.Paste
    Range("Q173:Y212").Select
    Selection.Sort Key1:=Range("U173"), Order1:=xlDescending, Key2:=Range( _
        "X173"), Order2:=xlDescending, Header:=xlGuess, OrderCustom:=1, _
        MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, _
        DataOption2:=xlSortNormal
    
' **********************************************************
' *** On sauvegarde les records dans archives            ***
' **********************************************************
    Windows(lenom).Activate
    Range("AM8:AQ106").Select
    Selection.Copy
    Windows("Archives.xls").Activate
    Range("IQ3001").Select
    ActiveSheet.Paste
        
' **********************************************************
' *** on ferme le fichier des archives en sauvegardant   ***
' **********************************************************
    Windows("Archives.xls").Close SaveChanges:=True
    End If
End Sub
Sub Plage()
'
' Plage Macro
' Macro enregistrée le 11/08/2006 par Yok's
'
    Range("B8:N213").Copy
End Sub
Sub Plage2()
'
' Plage Macro
' Macro enregistrée le 11/09/2007 par Yok's
' Modifiée le 07/01/18 par yok's
'
    Range("Q8:AJ260").Copy
End Sub

