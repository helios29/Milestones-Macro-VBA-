Attribute VB_Name = "JalonMacro"
Public Macro, NomFichier, Fichier_Jalon As String
Public intNbOR, i, q, col_Raj As Integer
Public Tableau_Engagement() As String
Public s(1) As String
Public a(1) As String
Public Nombre_d_Engagements As Integer

Sub Milestones_Macro()
    
    Application.StatusBar = False
    Application.ScreenUpdating = False
    Macro = ActiveWorkbook.Name
    
'     Application.DisplayAlerts = False
'        Sheets("PIT").Delete
'        Sheets("CTC").Delete
'        Sheets("Old PIT").Delete
'        Sheets("Old CTC").Delete
'    Application.DisplayAlerts = True
    
    s(0) = "PIT"
    s(1) = "CTC"
    a(0) = Macro
    
    
    Sheets("Interface").Select
'    intNbOR = Sheets("OR").Range("A65536").End(xlUp).Row
    Z = 0
    
    'Ouverture des fichiers
    '***************************
    For q = 1 To 4
    
        NomFichier = Workbooks(Macro).Worksheets("Interface").Cells(9 + Z, 3).Value
        
        If NomFichier = "" Then Exit For
        
        Application.StatusBar = "Ouverture du fichier " & NomFichier & " ..."
        
        Set Classeur = Excel.Workbooks.Open(NomFichier, UpdateLinks:=0)
        Application.StatusBar = "Fichier " & NomFichier & " ouvert..."
        
        Fichier_Jalon = ActiveWorkbook.Name
        a(1) = Fichier_Jalon
        
        Call Date_Format
        
        'Intégration jalon du mois
        '***************************************************
        If q = 3 Or q = 4 Then
        
            Call Analyse_Data
            
            If q = 4 Then
                Call New_Milestones
                Call Old_Milestones
            End If
            
            Workbooks(Fichier_Jalon).Close SaveChanges:=False
            
        ElseIf q = 1 Then
            Call Intégration_Jalons_Du_Mois '=> intégration du fichier dans la macro
        Else
        
            Call Intégration_Jalons_Du_Mois_M '=> intégration des nouveaux mois dans le fichier
            Workbooks(Fichier_Jalon).Close SaveChanges:=False
            
        End If
        
        
        Z = Z + 4
        
        
    Next q
    
        
    Call Mise_En_Forme
    
    Sheets("CTC").Range("A9:AB9").AutoFilter
    Sheets("Interface").Select
    
    MsgBox "Macro completed !"
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
End Sub

Sub Intégration_Jalons_Du_Mois()
    
    With Workbooks(Fichier_Jalon)
        .Sheets.Add.Name = "manouvellefeuille"
        .Sheets("CTC").Move After:=Workbooks(Macro).Sheets(1)
        .Sheets("PIT").Move After:=Workbooks(Macro).Sheets(1)
    End With
    
   Workbooks(Fichier_Jalon).Close SaveChanges:=False
   
End Sub

Sub Intégration_Jalons_Du_Mois_M()

    Dim date_jalon As Date
    
    If q = 2 Then
        
        Workbooks(Macro).Activate
        date_jalon = Workbooks(Fichier_Jalon).Sheets("PIT").Range("F8")
        Mois_Jalon_M = Month(date_jalon)
        date_jalon = Workbooks(Macro).Sheets("PIT").Range("F8")
        Mois_Jalon = Month(date_jalon)
        
        col_Raj = Mois_Jalon - Mois_Jalon_M
        
        For t = 0 To 1
            Sheets(s(t)).Select
             
            For j = 1 To col_Raj
                
                Columns("F:F").Insert Shift:=xlToRight
            
            Next j
            
            Workbooks(Fichier_Jalon).Activate
            Sheets("PIT").Select
            Range(Cells(8, 6), Cells(8, 6 + col_Raj - 1)).Copy
            Workbooks(Macro).Activate
            Sheets(s(t)).Range("F8").PasteSpecial Paste:=xlPasteValues
            
        Next t
        
        Application.CutCopyMode = False
        
        Call Analyse_Data
     
    End If

End Sub

Sub Analyse_Data()
   
   
    'Préparation des fichiers
    '**************************
    
    If q = 2 Then
    
        For t = 0 To 1
            
            Workbooks(a(t)).Activate
            ActiveSheet.AutoFilterMode = False
            Call Préparation_Fichiers
            
        Next t
        
    Else
       
        Call Préparation_Fichiers
        
    End If
    
    Call Recherche_Valeur


End Sub


Sub Préparation_Fichiers()
    
    'Retrait du filtre + préparation clef
    '**********************************************
     'For x = 0 To 1
     
     For x = 0 To 1
       
        Sheets(s(x)).Select
        Range("A8:ZZ8").AutoFilter
        
        Der = Range("A8").End(xlDown).Row
        
        Columns("A:A").Insert Shift:=xlToRight
        Columns("A:A").Insert Shift:=xlToRight
        Range("A8") = "Key"
        Range("B8") = "Status"
        Range("A9").FormulaR1C1 = "=RC[2]&RC[3]&RC[4]&RC[6]"
        Range("A9").AutoFill Destination:=Range("A9:A" & Der)
      
    Next x

End Sub

Sub Recherche_Valeur()
    
    'On cherche les valeurs
    '*************************
    Workbooks(a(0)).Activate
    
    For x = 0 To 1
        
        Sheets(s(x)).Select
        Der = Range("A8").End(xlDown).Row
    
        For p = 9 To Der
        
            jalon = Range("A" & p)
            
            Set Recherche_jalon = Workbooks(a(1)).Sheets(s(x)).Range("A:A").Find(what:=jalon, LookIn:=xlValues, lookat:=xlWhole)
            
            'Si on a trouvé quelque chose
            '**********************************
            If Not Recherche_jalon Is Nothing Then
                
                With Workbooks(a(1)).Sheets(s(x))
                
                    Row_Match = Recherche_jalon.Row
                    Place_valeur = .Range("G" & Row_Match).End(xlToRight).Column
                    Valeur = .Range("G" & Row_Match).End(xlToRight)
                    date_jalon = .Cells(8, Place_valeur)
                     
                End With
                
                Set placement_jalon = Workbooks(a(0)).Sheets(s(x)).Range("H8:BB8").Find(date_jalon, , xlValues, xlWhole, , , False)
                
                If Not placement_jalon Is Nothing Then
                
                    jalon_trouve = placement_jalon.Column
                    
                    If Cells(p, jalon_trouve) = "" Then
                        Cells(p, 2) = "slide"
                        Cells(p, jalon_trouve) = Valeur
                        Cells(p, jalon_trouve).NumberFormat = "#,##0"
                        If q = 2 Then Cells(p, jalon_trouve).Interior.Color = 5287936
                        
                        If q = 3 Then
                            With Cells(p, jalon_trouve).Interior
                                .ThemeColor = xlThemeColorAccent6
                                .TintAndShade = 0.399975585192419
                            End With
                        End If
                        
                        If q = 4 Then
                            
                            With Cells(p, jalon_trouve).Interior
                                .ThemeColor = xlThemeColorAccent6
                                .TintAndShade = 0.799981688894314
                            End With
                        End If
                        
                    End If
                    
                End If
                
                
    
            End If
            
        Next p


    Next x
    
End Sub


Sub Date_Format()

     For R = 0 To 1
        Sheets(s(R)).Rows("8:8").NumberFormat = "General"
    Next R
    
End Sub

Sub Mise_En_Forme()
    
    Dim arr() As String
    
    For R = 0 To 1
        
        Sheets(s(R)).Select
        
        Columns("A:A").Delete Shift:=xlToLeft
        
        Range(Cells(8, 7), Cells(8, 24 + col_Raj)).NumberFormat = "mmm-yy"
        Range(Columns(7), Columns(24 + col_Raj)).ColumnWidth = 8.5
    
        With Range("A8").Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    
        Range("A8").Borders(xlEdgeRight).LineStyle = xlNone
       
        With Range("A7:AQ7").Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With

        
        With Range("B8:C8")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
        
        Columns("A:A").ColumnWidth = 8.73
        Columns("B:B").ColumnWidth = 9.91
        Columns("C:C").ColumnWidth = 10
        Columns("D:D").ColumnWidth = 12.27
        Columns("E:E").ColumnWidth = 14.09
        Columns("F:F").ColumnWidth = 14
        Columns("G:AA").ColumnWidth = 7.73
        
        Range("B2") = "Extraction date:"
        Range("C2") = Date
        
        'Misse en place de la légende
        '**********************************
        
        Range("E2") = "Legend :"
        
        p = 13
        
        For x = 1 To 3
        
            arr = Split(Sheets("Interface").Cells(p, 3), "\")
            Value = arr(UBound(arr))
            p = p + 4
            
            Cells(x + 2, 6) = Value
            
            If x = 1 Then Cells(x + 2, 5).Interior.Color = 5287936
            If x = 2 Then
            
                With Cells(x + 2, 5).Interior
                
                    .ThemeColor = xlThemeColorAccent6
                    .TintAndShade = 0.399975585192419
                
                End With
            End If
            
            If x = 3 Then
                
                With Cells(x + 2, 5).Interior
                
                    .ThemeColor = xlThemeColorAccent6
                    .TintAndShade = 0.799981688894314
                
                End With
            End If
            
        Next x
        
        
        Range("F2") = ""
        Range("C3") = ""
        
        With Range("F3:F5")
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = False
        End With
    
        With Range("E2")
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
        
        Columns("A:A").HorizontalAlignment = xlCenter
        
       Der = Range("B8").End(xlDown).Row
        
        With Range("A9:AA" & Der).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With Range("A9:AA" & Der).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With Range("A9:AA" & Der).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With Range("A9:AA" & Der).Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlHairline
        End With
        
        With Range("H8:I8")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
        
         With Range("B2")
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlBottom
        End With
        
        ActiveWindow.DisplayGridlines = False
    
        Range("B2").VerticalAlignment = xlCenter
        
        Columns("E:E").HorizontalAlignment = xlLeft
    
        With Rows("8:8")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        With Range(Columns(7), Columns(24 + col_Raj))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        Rows("6:6").Insert Shift:=xlDown
        Range("K7").AutoFill Destination:=Range("G7:K7"), Type:=xlFillDefault
        
         Range("A9:AB9").AutoFilter
        
        Range("E6").Interior.Pattern = xlNone
        
        With Range("A9").Interior
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.399975585192419
        End With
        
        With Range("R5:R" & Der + 3).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Color = -16777024
            .Weight = xlMedium
        End With
        
        ActiveWindow.DisplayGridlines = False
        
        Range("A1").Select
        
    Next R
    
    For i = 0 To 1
    
        
        Worksheets("Old " & s(i)).Select
        Cells.EntireColumn.AutoFit
        Rows("8:8").RowHeight = 26.5
        Range("A8").FormulaR1C1 = "Contract number"
    
        Columns("A:A").ColumnWidth = 13.82
        Columns("C:C").ColumnWidth = 26.73
    
        Range("F8:AD8").NumberFormat = "mm/yyyy"
        Columns("G:AD").ColumnWidth = 6.55
        
        ActiveWindow.Zoom = 65
        
         With Range("A8:B8")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
        
        Columns("A:A").ColumnWidth = 8.73
        Columns("A:A").ColumnWidth = 9.91
        Columns("B:B").ColumnWidth = 10
        Columns("C:C").ColumnWidth = 12.27
        Columns("D:D").ColumnWidth = 14.09
        Columns("E:E").ColumnWidth = 14
        Columns("F:WW").ColumnWidth = 7.73
        
        Der = Range("A8").End(xlDown).Row
        
        With Range("A9:X" & Der).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With Range("A9:X" & Der).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With Range("A9:X" & Der).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With Range("A9:X" & Der).Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlHairline
        End With
        
        Range("F3") = "Old - " & s(i)
        
         With Range("F3:J4")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        Range("F3:J4").Merge
        Range("F3:J4").Font.Size = 18
    
        With Range("F3:J4").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
        
        Range("A8").Select
        
        ActiveWindow.DisplayGridlines = False
        
    Next i
    
End Sub

Sub New_Milestones()

    For x = 0 To 1
        
        Workbooks(a(0)).Activate
        Sheets(s(x)).Select
        Der = Range("A8").End(xlDown).Row
    
        For p = 9 To Der
        
            jalon = Range("A" & p)
            
            Set Recherche_jalon = Workbooks(a(1)).Sheets(s(x)).Range("A:A").Find(what:=jalon, LookIn:=xlValues, lookat:=xlWhole)
            
            'Si nouveau jalon
            '*******************
            If Recherche_jalon Is Nothing Then
                
                With Range(Cells(p, 1), Cells(p, 28)).Interior
                    .ThemeColor = xlThemeColorAccent1
                    .TintAndShade = 0.399975585192419
                End With
                
                Workbooks(a(0)).Sheets(s(x)).Range("B" & p) = "New Milst."
    
            End If
            
        Next p


    Next x


End Sub


Sub Old_Milestones()
     
    
    
    For x = 0 To 1
        
        Z = 0
        
        Workbooks(a(0)).Activate
        Sheets.Add(After:=Worksheets(s(x))).Name = "Old " & s(x)
        
        Workbooks(a(1)).Activate
        Sheets(s(x)).Select
        Der = Range("A8").End(xlDown).Row
        Range(Cells(8, 3), Cells(8, 47)).Copy
        
        Workbooks(Macro).Activate
        
        Sheets("Old " & s(x)).Range("A8").PasteSpecial Paste:=xlPasteValues
        Sheets("Old " & s(x)).Range("A8").PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
        Workbooks(a(1)).Activate
        
        For p = 9 To Der
        
            jalon = Range("A" & p)
            
            Set Recherche_jalon = Workbooks(a(0)).Sheets(s(x)).Range("A:A").Find(what:=jalon, LookIn:=xlValues, lookat:=xlWhole)
            
            'Si ancien jalon
            '*******************
            If Recherche_jalon Is Nothing Then
                
                Workbooks(a(1)).Sheets(s(x)).Range(Cells(p, 3), Cells(p, 47)).Copy
                Workbooks(a(0)).Activate
                Sheets("Old " & s(x)).Select
                Cells(Z + 9, 1).PasteSpecial Paste:=xlPasteValues
                Cells(Z + 9, 1).PasteSpecial Paste:=xlPasteFormats
                Application.CutCopyMode = False
                Z = Z + 1
                Workbooks(a(1)).Activate
                  
                
    
            End If
            
            
        Next p
        
        

    Next x

End Sub

Sub Fichier_1()

    Dim fd As FileDialog
    Dim Actionclicked As Boolean
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    'Sélectionner Microsoft Scripting Runtime comme libraire
    
'    Enlever les filtres puis les remettre afin de ne pouvoir intégrer que les fichiers ".csv"
    fd.Filters.Clear
'    fd.Filters.Add "Fichier CSV", "*.csv"
    fd.Filters.Add "Tous les fichiers excel", "*.xl*"
    
'    Sélectionne le filtre 1
    fd.FilterIndex = 1
    
    'Interdire la sélection multiple de fichiers
    fd.AllowMultiSelect = False
    
    'Sélection du disque M, endroit où se trouve les fichiers
    fd.InitialFileName = "C:\"
    
    'Changement du titre
    fd.Title = ""
    
    'Changement du titre du bouton de de validation
    fd.ButtonName = "Intégrer les données"
    
    Actionclicked = fd.Show
    
    If Actionclicked Then
        Range("C9").Value = fd.SelectedItems(1)
    End If

End Sub
