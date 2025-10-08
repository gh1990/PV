VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    On Error Resume Next
    
    ' Activeaza foaia START la deschiderea fi?ierului
    ThisWorkbook.Sheets("START").Activate
    ' Activeaza tabul "Proces Verbal" (id="tabTest") din Ribbon la deschidere
    If Not Module1.ribbonUI Is Nothing Then
        Module1.ribbonUI.ActivateTab "tabTest"
    End If
    
    On Error GoTo 0
    
End Sub

Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
    On Error GoTo SafeExit
    If Sh.Name = "CalculMateriale" Or Sh.Name = "Statistica" Then
        Application.DisplayAlerts = False
        Sh.Delete
        Application.DisplayAlerts = True
    End If
    
SafeExit:
    Application.DisplayAlerts = True
End Sub

' Subrutina Workbook_SheetChange - declansata atunci cand se face o modificare intr-o foaie de lucru
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    
    Dim bdSheets As Variant
    bdSheets = Array("Obiect", "Norma", "Materiale", "Utilaj", "Transport", "Liste")

    Dim iii As Integer
    For iii = LBound(bdSheets) To UBound(bdSheets)
        If Sh.Name = bdSheets(iii) Then
            Module1.FlagUserFormNeedsRefresh = True
            Exit For
        End If
    Next iii
    
        ' Actualizeaza ribbon-ul daca se schimba foaia "Liste"
    If Sh.Name = "Liste" Then
        On Error Resume Next
        Module1.LoadLists
        On Error GoTo 0
    End If
      
        ' Verificam daca s-a modificat direct celula C18 pe o foaie PV_*
    If Left(Sh.Name, 3) = "PV_" And Not Intersect(Target, Sh.Range("C18")) Is Nothing Then
        Dim linie As Long
        linie = Module1.CautaValoareRand("PVModel", "D61")
        If Sh.Range("C18").value = "Servicii prestate" Then
            Sh.Range("G" & linie + 4).value = ""
            Sh.Range("D" & linie + 4).value = ""
        Else
            Sh.Range("G" & linie + 4).value = Module1.CautaSiReturneaza("Liste", "Sector", Module1.cmbEchipaText, "SefSIDTP")
            Sh.Range("D" & linie + 4).value = "Sef SIDTP"
        End If
   End If
      
      
    ' Verificam dacг modificarea afecteaza coloanele E sau F
    If Not Intersect(Target, Sh.Range("E:F")) Is Nothing Then
        
                ' >>>>>> VERIFICARE FOAIE: Ruleaza DOAR pe foi care incep cu "PV_"
        If Left(Sh.Name, 3) <> "PV_" Then Exit Sub
        
        ' IMPORTANT: Dezactivam evenimentele pentru a preveni bucla infinita
        Application.EnableEvents = False
        
        On Error GoTo ErrorHandler
        
        ' Declarati variabilele pentru a stoca informatiile necesare
        Dim ws As Worksheet ' Foaia de lucru curenta
        Dim cell As Range ' Celula curenta
        Dim lastRow As Long ' Ultimul rand ne-gol din coloana C
        Dim i, j, k As Integer ' Contoare pentru bucle
        Dim aManopera, bManopera As String ' Adrese pentru inceputul si sfarsitul sectiunii de "Manopera"
        Dim StartManopera, StopManopera As Boolean 'nr. adresei inceput si sfarsit "Manopera"
        Dim aMateriale, bMateriale As String ' Adrese pentru inceputul si sfarsitul sectiunii de "Materiale"
        Dim iUtilaj, jUtilaj, iTransport, nrOrdine As Integer ' Contoare si variabile pentru sectiuni specifice
        Dim aUtilaj, bUtilaj, cUtilaj, dUtilaj As String ' Adrese pentru sectiunea de "Utilaj"
        Dim aTransport, bTransport As String ' Adrese pentru sectiunea de "Transport"
        Dim nrManopera, nrMateriale, nrUtilaj As Integer ' Numere pentru indexare
        Dim adrUnitateMasura, adrManopera, adrMateriale, adrUtilaj, adrResurse, adrManopera2, adrManopera3 As String ' Formule si adrese pentru diverse sectiuni
        Dim adrNorma, adrCantitate As String ' Formule si adrese norma, cantitate
        Dim mesaj As Boolean ' Variabila pentru mesaje
        Dim valoareK1 As Variant ' Variabila pentru stocarea valorii din K1
        Dim valoareK2 As Variant ' Variabila pentru stocarea valorii din K2
        Dim foundRow As Long ' Variabila pentru randul curent gasit
        Dim rowNum As Long
        Dim ManoperaTotal As Double
        ' Setam foaia activa curenta
        Set ws = Sh
        nrOrdine = 0 ' Initializam numarul de ordine
        iUtilaj = 0 ' Initializam indexul pentru "Utilaj"
        kUtilaj = 0 ' Initializam un alt index pentru "Utilaj"
        ManoperaTotal = 0
        StartManopera = False
        StopManopera = False
        
        ' Verificam daca modificarea afecteaza coloanele E sau F
        adrResurse = "=ROUND(SUM(" ' Initializam formula pentru "Resurse"                             '---------------------
        adrManopera2 = "=ROUND(SUM(" ' Initializam formula pentru "Manopera suplimentara"
        adrManopera3 = "=ROUND(SUM(" ' Initializam formula pentru "Manopera suplimentara"
        ' Obtinem valoarea din K1 - noua celula de referinta
        valoareK1 = ws.Cells(29, "H").value
        
        ' Determinam ultimul rand completat din coloana C
        lastRow = ws.Cells(ws.Rows.count, "C").End(xlUp).Row
        For i = 1 To lastRow ' Parcurgem fiecare rand pana la ultimul
            If ws.Cells(i, "C").value = "Manopera:" Then
                foundRow = i
                aManopera = "G" & foundRow + 1
                nrOrdine = nrOrdine + 1
                ws.Cells(i - 1, "A").value = nrOrdine
                nrManopera = i - 1
                adrResurse = adrResurse & "G" & i - 1 & ","
                adrNorma = "=ROUND(H" & (foundRow - 1)
                adrCantitate = "*E" & (foundRow - 1) & ",3)"
                adrUnitateMasura = "=ROUND(G" & (foundRow - 1) & "/E" & (foundRow - 1) & ",3)"
            End If
            
            If ws.Cells(i, "C").value = "Total manopera" Then
                foundRow = i
                bManopera = "G" & foundRow - 1
                ws.Cells(i, "G").Formula = "=ROUND(SUM(" & aManopera & ":" & bManopera & "),3)"  '---------------------------------
                adrManopera = "=SUM(G" & i & ")"
                ws.Cells(nrManopera, "G").Formula = adrManopera
                ws.Range("A" & nrManopera & ":G" & nrManopera).Interior.Color = RGB(217, 217, 217)
                adrManopera2 = adrManopera2 & "G" & i & ","
                adrManopera3 = adrManopera3 & "E" & i - 1 & ","
                For Each cell In ws.Range(aManopera & ":" & bManopera)
                    rowNum = cell.Row
                    cell.Formula = "=ROUND(E" & rowNum & "*F" & rowNum & ",3)"       '------------------------------------
                    ws.Cells(rowNum, "F").value = valoareK1
                    ws.Cells(rowNum, "E").Formula = adrNorma & adrCantitate
                    ws.Cells(rowNum - 1, "F").Formula = "=ROUND(E" & rowNum & "/E" & (rowNum - 2) & ",3)"
                    ws.Cells(rowNum - 1, "G").Formula = "=ROUND(F" & rowNum & "*F" & (rowNum - 1) & ",3)"
                    ws.Cells(rowNum - 2, "F").Formula = adrUnitateMasura
                    ws.Cells(rowNum, "D").value = "h-om"
                    ws.Cells(rowNum - 1, "D").value = "h-om"
                    ws.Cells(rowNum - 1, "E").value = "1"
                Next cell

            End If
            
            If ws.Cells(i, "C").value = "Materiale:" Then
                foundRow = i
                aMateriale = "G" & foundRow + 1
            End If
            
            If ws.Cells(i, "C").value = "Total materiale" Then
                foundRow = i
                bMateriale = "G" & foundRow - 1
                ws.Cells(i, "G").Formula = "=ROUND(SUM(" & aMateriale & ":" & bMateriale & "),3)"  '---------------------
                
                adrMateriale = Left(adrManopera, Len(adrManopera) - 1) & "," & "G" & foundRow & ")"
                ws.Cells(nrManopera, "G").Formula = adrMateriale
                
                For Each cell In ws.Range(aMateriale & ":" & bMateriale)
                    rowNum = cell.Row
                    cell.Formula = "=ROUND(E" & rowNum & "*F" & rowNum & ",3)"             '------------------------------
                Next cell
            End If
            
            If (ws.Cells(i, "C").value = "Ma" & ChrW(537) & "ini " & ChrW(537) & "i utilaje") Or (ws.Cells(i, "C").value = "Ma" & ChrW(537) & "ini " & ChrW(537) & "i utilaje:") Then
                If ws.Cells(i - 1, "C").value = "Total materiale" Then
                    foundRow = i
                    iUtilaj = i
                    aUtilaj = "G" & foundRow + 1
                    For j = i To lastRow
                        If aUtilaj <> Empty Then
                            If ws.Cells(j, "C").value = "Manopera:" Then
                                foundRow = j
                                bUtilaj = "G" & foundRow - 2
                                ws.Cells(iUtilaj, "G").Formula = "=ROUND(SUM(" & aUtilaj & ":" & bUtilaj & "),3)" '---------------------------
                                adrMateriale = Left(adrMateriale, Len(adrMateriale) - 1) & "," & "G" & iUtilaj & ")"
                                ws.Cells(nrManopera, "G").Formula = adrMateriale
                                Exit For
                            Else
                                If ws.Cells(j, "C").value = "Transportare echipei " & ChrW(537) & "i materialelor" Then
                                    foundRow = j
                                    bUtilaj = "G" & foundRow - 1
                                    ws.Cells(iUtilaj, "G").Formula = "=ROUND(SUM(" & aUtilaj & ":" & bUtilaj & "),3)" '--------------------------
                                    adrMateriale = Left(adrMateriale, Len(adrMateriale) - 1) & "," & "G" & iUtilaj & ")"
                                    ws.Cells(nrManopera, "G").Formula = adrMateriale
                                    Exit For
                                End If
                            End If
                        End If
                    Next j
                End If
                
                For Each cell In ws.Range(aUtilaj & ":" & bUtilaj)
                    rowNum = cell.Row
                    cell.Formula = "=ROUND(E" & rowNum & "*F" & rowNum & ",3)"              '--------------------------------------------
                Next cell
            End If
            
            If ws.Cells(i, "C").value = "Transportare echipei " & ChrW(537) & "i materialelor" Then
                If ws.Cells(i + 1, "C").value = "Ma" & ChrW(537) & "ini " & ChrW(537) & "i utilaje:" Then
                    foundRow = i
                    iTransport = i + 1
                    aTransport = "G" & foundRow + 2
                    nrOrdine = nrOrdine + 1
                    ws.Cells(i, "A").value = nrOrdine
                    ws.Range("A" & i & ":G" & i).Interior.Color = RGB(217, 217, 217)
                    adrResurse = adrResurse & "G" & i & ","
                    ws.Cells(i, "G").Formula = "=G" & i + 1
                End If
            End If
            
            If ws.Cells(i, "C").value = "TOTAL RESURSE" Then
                foundRow = i
                bTransport = "G" & foundRow - 1
                ws.Cells(iTransport, "G").Formula = "=ROUND(SUM(" & aTransport & ":" & bTransport & "),3)"  '----------------
                
                For Each cell In ws.Range(aTransport & ":" & bTransport)
                    rowNum = cell.Row
                    cell.Formula = "=ROUND(E" & rowNum & "*F" & rowNum & ",3)"               '------------------------------
                Next cell
                             
                ws.Cells(i, "G").Formula = Left(adrResurse, Len(adrResurse) - 1) & "),3)"                  '-----------------------
                ws.Cells(i + 1, "G").Formula = Left(adrManopera2, Len(adrManopera2) - 1) & "),3)"
                ws.Cells(i + 1, "E").Formula = Left(adrManopera3, Len(adrManopera3) - 1) & "),3)"
                ws.Cells(i + 2, "G").Formula = "=ROUND(G" & i + 1 & "*0.24,3)"                             '------------------------
                ws.Cells(i + 3, "G").Formula = "=ROUND((G" & i + 1 & "+ G" & i + 2 & ")*0.1,3)"            '------------------------
                ws.Cells(i + 4, "G").Formula = "=ROUND(SUM(G" & i & ",G" & i + 2 & ":G" & i + 3 & "),3)"   '------------------------
                ws.Cells(i + 5, "G").Formula = "=ROUND(G" & i + 4 & "* 0.145,3)"                           '------------------------
                ws.Cells(i + 6, "G").Formula = "=ROUND(SUM(G" & i + 4 & ":G" & i + 5 & "),3)"              '------------------------
                k = i + 1
            End If
            
        Next i
ExitSub:
        Application.EnableEvents = True
        Exit Sub
        
ErrorHandler:
        MsgBox "A apгrut o eroare: " & Err.Description, vbExclamation
        Resume ExitSub
    End If
End Sub

'--------------
'
'--------------
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    On Error GoTo SafeExit
    ' Daca avem o foaie deschisa temporar ?i tocmai am activat o foaie "mereu vizibila",
    ' atunci restauram vizibilitatea foii temporare (o reascundem cum era оnainte)
    If Len(TempVisibleSheetName) > 0 Then
        If StrComp(Sh.Name, TempVisibleSheetName, vbTextCompare) <> 0 Then
            If IsAlwaysVisibleSheetName(Sh.Name) Then
                Application.EnableEvents = False
                RestoreTempVisibility
            End If
        End If
    End If
SafeExit:
    Application.EnableEvents = True
End Sub
