Attribute VB_Name = "Module1"
Option Explicit
' Stare globala pentru o singura foaie vizualizata temporar
Public TempVisibleSheetName As String
Public TempPrevVisibility As XlSheetVisibility
' Ribbon global
Public ribbonUI As IRibbonUI
Public Rindul As Long
' ARRAY-URI stocate ca Variant (pentru a primi dict.Keys)
Private echipeArr As Variant
Private tipArr As Variant
' Text curent selectat (pentru getText)
Private cmbEchipaText As String
Private cmbTipText As String
' Foi-model/foi date care pot fi ascunse
Private Const MODEL_SHEETS As String = "PVModel|FisaModel|Liste|Obiect|Norma|Materiale|Utilaj|Transport"
' =============================================
' Utilitare: acces sigur la foi + protectie structura
' =============================================
Private Function SafeGetSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set SafeGetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If SafeGetSheet Is Nothing Then
        Err.Raise vbObjectError + 513, "SafeGetSheet", "Foaia '" & sheetName & "' nu exista in acest registru."
    End If
End Function
' Deprotejeaza structura, returneaza True daca era protejata (ca sa o restauram la final)
Private Function UnprotectStructureIfNeeded() As Boolean
    On Error Resume Next
    If ThisWorkbook.ProtectStructure Then
        ThisWorkbook.Unprotect
        UnprotectStructureIfNeeded = True
    Else
        UnprotectStructureIfNeeded = False
    End If
    On Error GoTo 0
End Function
' Restaureaza structura daca era protejata anterior
Private Sub RestoreStructureIfNeeded(ByVal wasProtected As Boolean)
    On Error Resume Next
    If wasProtected Then
        ThisWorkbook.Protect Structure:=True, Windows:=False
    End If
    On Error GoTo 0
End Sub
Private Sub WithTempVisibility(ByVal ws As Worksheet, ByRef prevVis As XlSheetVisibility)
    prevVis = ws.Visible
    If prevVis <> xlSheetVisible Then
        ws.Visible = xlSheetVisible
    End If
End Sub
Private Sub RestoreVisibility(ByVal ws As Worksheet, ByVal prevVis As XlSheetVisibility)
    If Not ws Is Nothing Then
        ws.Visible = prevVis
    End If
End Sub
' E foaie model/BD?
Private Function IsModelOrDataSheet(ByVal n As String) As Boolean
    IsModelOrDataSheet = (InStr(1, "|" & MODEL_SHEETS & "|", "|" & n & "|", vbTextCompare) > 0)
End Function
' Exista un sheet (Worksheets sau Chart) cu numele dat?
Private Function SheetExistsAll(ByVal sheetName As String) As Boolean
    Dim s As Object
    For Each s In ThisWorkbook.Sheets
        If StrComp(s.Name, sheetName, vbTextCompare) = 0 Then
            SheetExistsAll = True
            Exit Function
        End If
    Next s
    SheetExistsAll = False
End Function
' Ultima foaie "mereu vizibila" (in ordinea tab-urilor)
Private Function FindLastAlwaysVisibleSheet() As Worksheet
    Dim ws As Worksheet, lastWS As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If IsAlwaysVisibleSheetName(ws.Name) Then
            Set lastWS = ws
        End If
    Next ws
    Set FindLastAlwaysVisibleSheet = lastWS
End Function
' Definim "foaie mereu vizibila"
Public Function IsAlwaysVisibleSheetName(ByVal n As String) As Boolean
    If Left$(n, 3) = "PV_" Then
        IsAlwaysVisibleSheetName = True
    ElseIf Left$(n, 2) = "F_" Then
        IsAlwaysVisibleSheetName = True
    ElseIf StrComp(n, "CalculMateriale", vbTextCompare) = 0 Then
        IsAlwaysVisibleSheetName = True
    Else
        IsAlwaysVisibleSheetName = False
    End If
End Function
Public Sub AscundeFoiModel()
    Dim ws As Worksheet
    Dim wasProt As Boolean
    wasProt = UnprotectStructureIfNeeded()
    On Error GoTo Clean
    For Each ws In ThisWorkbook.Worksheets
        Select Case True
            Case Left(ws.Name, 3) = "PV_"
                ws.Visible = xlSheetVisible
            Case Left(ws.Name, 2) = "F_"
                ws.Visible = xlSheetVisible
            Case ws.Name = "CalculMateriale"
                ws.Visible = xlSheetVisible
            Case InStr(1, "|" & MODEL_SHEETS & "|", "|" & ws.Name & "|", vbTextCompare) > 0
                ws.Visible = xlSheetVeryHidden
            Case Else
                ' nu atingem alte foi
        End Select
    Next ws
Clean:
    RestoreStructureIfNeeded wasProt
End Sub
Public Sub ArataFoiModel(Optional ByVal vis As XlSheetVisibility = xlSheetVisible)
    Dim nameArr() As String, i As Long
    Dim wasProt As Boolean
    wasProt = UnprotectStructureIfNeeded()
    On Error GoTo Clean
    nameArr = Split(MODEL_SHEETS, "|")
    For i = LBound(nameArr) To UBound(nameArr)
        On Error Resume Next
        ThisWorkbook.Worksheets(nameArr(i)).Visible = vis
        On Error GoTo 0
    Next i
Clean:
    RestoreStructureIfNeeded wasProt
End Sub
' =============================================
' Ribbon
' =============================================
Public Sub RibbonOnLoad(ribbon As IRibbonUI)
    Set ribbonUI = ribbon
    LoadLists
    AscundeFoiModel
End Sub
' =============================================
' Încarca listele din foaia "Liste"
' =============================================
Private Sub LoadLists()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim dictEchipe As Object, dictTip As Object
    Dim colSector As Long, colTip As Long
    Dim headerRow As Long
    On Error GoTo ErrHandler
    Set ws = SafeGetSheet("Liste")
    Set dictEchipe = CreateObject("Scripting.Dictionary")
    Set dictTip = CreateObject("Scripting.Dictionary")
    colSector = 0: colTip = 0: headerRow = 0
    For j = 1 To 10
        For i = 1 To ws.Cells(j, ws.Columns.count).End(xlToLeft).Column
            Select Case Trim(LCase(ws.Cells(j, i).value))
                Case "sector": colSector = i: headerRow = j
                Case "tiplucrari": colTip = i: headerRow = j
            End Select
        Next i
    Next j
    If colSector = 0 Or colTip = 0 Then
        MsgBox "Nu s-au gasit coloanele 'Sector' si/sau 'TipLucrari' in foaia Liste.", vbExclamation
        Exit Sub
    End If
    ' Sector
    lastRow = ws.Cells(ws.Rows.count, colSector).End(xlUp).Row
    For i = headerRow + 1 To lastRow
        If Trim(ws.Cells(i, colSector).value & "") <> "" Then
            dictEchipe(Trim(ws.Cells(i, colSector).value)) = 1
        End If
    Next i
    If dictEchipe.count > 0 Then
        echipeArr = dictEchipe.Keys
        cmbEchipaText = echipeArr(LBound(echipeArr))
    Else
        echipeArr = Array(): cmbEchipaText = ""
    End If
    ' TipLucrari
    lastRow = ws.Cells(ws.Rows.count, colTip).End(xlUp).Row
    For i = headerRow + 1 To lastRow
        If Trim(ws.Cells(i, colTip).value & "") <> "" Then
            dictTip(Trim(ws.Cells(i, colTip).value)) = 1
        End If
    Next i
    If dictTip.count > 0 Then
        tipArr = dictTip.Keys
        cmbTipText = tipArr(LBound(tipArr))
    Else
        tipArr = Array(): cmbTipText = ""
    End If
    ' Reîmprospatare Ribbon
    If Not ribbonUI Is Nothing Then
        ribbonUI.InvalidateControl "cmbEchipa"
        ribbonUI.InvalidateControl "cmbTip"
    End If
    Exit Sub
ErrHandler:
    MsgBox "Eroare la LoadLists: " & Err.Number & " - " & Err.Description, vbExclamation
End Sub
' =============================================
' CALLBACKS Combo Echipa
' =============================================
Public Sub cmbEchipa_getItemCount(control As IRibbonControl, ByRef count)
    If IsArray(echipeArr) Then
        If UBound(echipeArr) >= LBound(echipeArr) Then
            count = UBound(echipeArr) - LBound(echipeArr) + 1
        Else
            count = 0
        End If
    Else
        count = 0
    End If
End Sub
Public Sub cmbEchipa_getItemLabel(control As IRibbonControl, index As Integer, ByRef label)
    If IsArray(echipeArr) Then
        If index >= LBound(echipeArr) And index <= UBound(echipeArr) Then
            label = echipeArr(index): Exit Sub
        End If
    End If
    label = ""
End Sub
Public Sub cmbEchipa_getText(control As IRibbonControl, ByRef text)
    text = cmbEchipaText
End Sub
Public Sub cmbEchipa_onChange(control As IRibbonControl, text As String)
    Dim linie As Long
    cmbEchipaText = text
    If ActiveSheet.Name Like "PV_*" Then
        ActiveSheet.Range("C11").value = CautaSiReturneaza("Liste", "Sector", cmbEchipaText, "NumeSector")
        ActiveSheet.Range("C13").value = CautaSiReturneaza("Liste", "Sector", cmbEchipaText, "NumeSector")
        ActiveSheet.Range("H29").value = CautaSiReturneaza("Liste", "Sector", cmbEchipaText, "MediaSalariala")
        ActiveSheet.Range("E1").value = " "
        linie = CautaValoareRand("PVModel", "A63")
        ActiveSheet.Range("C" & linie).value = CautaSiReturneaza("Liste", "Sector", cmbEchipaText, "Normator") & "             "
        linie = CautaValoareRand("PVModel", "A67")
        ActiveSheet.Range("C" & linie).value = CautaSiReturneaza("Liste", "Sector", cmbEchipaText, "SefSector") & "             "
        linie = CautaValoareRand("PVModel", "D63")
        ActiveSheet.Range("G" & linie).value = CautaSiReturneaza("Liste", "Sector", cmbEchipaText, "SefDepartament")
        linie = CautaValoareRand("PVModel", "D65")
        ActiveSheet.Range("G" & linie).value = CautaSiReturneaza("Liste", "Sector", cmbEchipaText, "SefSIDTP")
    Else
        MsgBox "Foia data nu este un Proces Verbal", vbExclamation
    End If
End Sub
' =============================================
' CALLBACKS Combo Tip lucrare
' =============================================
Public Sub cmbTip_getItemCount(control As IRibbonControl, ByRef count)
    If IsArray(tipArr) Then
        If UBound(tipArr) >= LBound(tipArr) Then
            count = UBound(tipArr) - LBound(tipArr) + 1
        Else
            count = 0
        End If
    Else
        count = 0
    End If
End Sub
Public Sub cmbTip_getItemLabel(control As IRibbonControl, index As Integer, ByRef label)
    If IsArray(tipArr) Then
        If index >= LBound(tipArr) And index <= UBound(tipArr) Then
            label = tipArr(index): Exit Sub
        End If
    End If
    label = ""
End Sub
Public Sub cmbTip_getText(control As IRibbonControl, ByRef text)
    text = cmbTipText
End Sub
Public Sub cmbTip_onChange(control As IRibbonControl, text As String)
    cmbTipText = text
    If ActiveSheet.Name Like "PV_*" Then
        ActiveSheet.Range("C18").value = cmbTipText
    Else
        MsgBox "Foia data nu este un Proces Verbal", vbExclamation
    End If
End Sub
' =============================================
' Buton Refresh
' =============================================
Public Sub btnRefresh_onAction(control As IRibbonControl)
    LoadLists
    AscundeFoiModel
    FormeazaFisa Nothing
    CalcMateriale Nothing
    MsgBox "Listele au fost reincarcate din foaia 'Liste'.", vbInformation
End Sub
' =============================================
' MACRO-URI PENTRU BUTOANE MENIU
' =============================================
Public Sub AddProcesVerbalNou(control As IRibbonControl)
    Dim linie As Long
    'MsgBox "Buton apasat: AddProcesVerbalNou", vbInformation
    CopiazaFoaie "PVModel", "PV_"
    ActiveSheet.Range("C18").value = cmbTipText
    ActiveSheet.Range("C11").value = CautaSiReturneaza("Liste", "Sector", cmbEchipaText, "NumeSector")
    ActiveSheet.Range("C13").value = CautaSiReturneaza("Liste", "Sector", cmbEchipaText, "NumeSector")
    ActiveSheet.Range("H29").value = CautaSiReturneaza("Liste", "Sector", cmbEchipaText, "MediaSalariala")
    ActiveSheet.Range("E1").value = " "
    linie = CautaValoareRand("PVModel", "A63")
    ActiveSheet.Range("C" & linie).value = CautaSiReturneaza("Liste", "Sector", cmbEchipaText, "Normator") & "             "
    linie = CautaValoareRand("PVModel", "A67")
    ActiveSheet.Range("C" & linie).value = CautaSiReturneaza("Liste", "Sector", cmbEchipaText, "SefSector") & "             "
    linie = CautaValoareRand("PVModel", "D63")
    ActiveSheet.Range("G" & linie).value = CautaSiReturneaza("Liste", "Sector", cmbEchipaText, "SefDepartament")
    linie = CautaValoareRand("PVModel", "D65")
    ActiveSheet.Range("G" & linie).value = CautaSiReturneaza("Liste", "Sector", cmbEchipaText, "SefSIDTP")
    ActiveSheet.Range("C11").Select
End Sub
Public Sub AddFormularNorma(control As IRibbonControl)
    If ActiveSheet.Name Like "PV_*" Then
        InsereazaRanduriDinPVModel "31:39"
    Else
        MsgBox "Foia data nu este un Proces Verbal", vbExclamation
    End If
End Sub
Public Sub AddFormularTransport(control As IRibbonControl)
    If ActiveSheet.Name Like "PV_*" Then
        InsereazaRanduriDinPVModel "51:52"
    Else
        MsgBox "Foia data nu este un Proces Verbal", vbExclamation
    End If
End Sub
Public Sub AddObiectPV(control As IRibbonControl)
    If ActiveSheet.Name Like "PV_*" Then
        AddObiect
    Else
        MsgBox "Foia data nu este un Proces Verbal", vbExclamation
    End If
End Sub
Public Sub AddNormaPV(control As IRibbonControl)
    If ActiveSheet.Name Like "PV_*" Then
        AddNorma
    Else
        MsgBox "Foia data nu este un Proces Verbal", vbExclamation
    End If
End Sub
Public Sub AddMaterialePV(control As IRibbonControl)
    If ActiveSheet.Name Like "PV_*" Then
        AddMateriale
    Else
        MsgBox "Foia data nu este un Proces Verbal", vbExclamation
    End If
End Sub
Public Sub AddTransportPV(control As IRibbonControl)
    If ActiveSheet.Name Like "PV_*" Then
        AddTransport
    Else
        MsgBox "Foia data nu este un Proces Verbal", vbExclamation
    End If
End Sub
Public Sub AddUtilajPV(control As IRibbonControl)
    If ActiveSheet.Name Like "PV_*" Then
        AddUtilaj
    Else
        MsgBox "Foia data nu este un Proces Verbal", vbExclamation
    End If
End Sub
Public Sub AddObiectBD(control As IRibbonControl)
    ShowSheetTemporarily "Obiect"
End Sub
Public Sub AddNormaBD(control As IRibbonControl)
    ShowSheetTemporarily "Norma"
End Sub
Public Sub AddMaterialeBD(control As IRibbonControl)
    ShowSheetTemporarily "Materiale"
End Sub
Public Sub AddUtilajBD(control As IRibbonControl)
    ShowSheetTemporarily "Utilaj"
End Sub
Public Sub AddTransportBD(control As IRibbonControl)
    ShowSheetTemporarily "Transport"
End Sub
Public Sub AddListeBD(control As IRibbonControl)
    ShowSheetTemporarily "Liste"
End Sub
Public Sub FormularPV(control As IRibbonControl)
    ShowSheetTemporarily "PVModel"
End Sub
Public Sub FormularFise(control As IRibbonControl)
    ShowSheetTemporarily "FisaModel"
End Sub
Public Sub CalcMateriale(control As IRibbonControl)
    Dim ws As Worksheet
    Dim wasProt As Boolean
    Dim prevScreen As Boolean, prevEvents As Boolean
    
    ' --- Salvare stari ?i pregatire ---
    prevScreen = Application.ScreenUpdating: Application.ScreenUpdating = False
    prevEvents = Application.EnableEvents: Application.EnableEvents = False
    wasProt = UnprotectStructureIfNeeded()
    
    On Error GoTo Finalize
    
    ' --- ?terge foaia "CalculMateriale" daca exista ---
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("CalculMateriale").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' --- Creeaza o noua foaie "CalculMateriale" ---
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
    ws.Name = "CalculMateriale"
    
    ' --- Format ?i antet ---
    With ws
        .Range("C1").value = "Nr"
        .Range("D1").value = "Cantitatea"
        .Range("E1").value = "Nume material"
        .Range("F1").value = "Cantitatea Disponibila"
        .Range("G1").value = "Diferenta"
        .Range("K1").value = "Manopera"
        .Range("M1").value = "h-om"
        
        ' Aliniere antet
        .Range("C1:M1").HorizontalAlignment = xlCenter
        .Range("C1:M1").VerticalAlignment = xlCenter
        .Range("C1:M1").Font.Bold = True
        .Range("C1:M1").Interior.Color = RGB(255, 204, 0)  ' galben
        
        ' Fundal pentru date
        .Range("C2:M900").Interior.Color = RGB(255, 242, 204)  ' bej deschis
        .Range("C2:M900").HorizontalAlignment = xlCenter
        .Range("C2:M900").VerticalAlignment = xlCenter
        
        ' Grupare coloane A:B
        .Columns("A:B").Group
        .Outline.ShowLevels ColumnLevels:=1
        
        ' La?imi fixe
        .Columns("C").ColumnWidth = 12
        .Columns("D").ColumnWidth = 10
        .Columns("E").ColumnWidth = 42
        .Columns("F").ColumnWidth = 20
        .Columns("G").ColumnWidth = 12
        .Columns("K").ColumnWidth = 10
        .Columns("L").ColumnWidth = 16
        .Columns("M").ColumnWidth = 6
    End With
    
    ' --- Activeaza foaia ---
    ws.Visible = xlSheetVisible
    ws.Activate
    
    ' --- Ruleaza logica de calcul ---
    If ExistaPrefix("PV_") Then
        cautaValoarea
        totalManopera
    End If

Finalize:
    RestoreStructureIfNeeded wasProt
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = prevEvents
    Application.DisplayAlerts = True
    If Err.Number <> 0 Then
        MsgBox "Eroare în CalcMateriale: " & Err.Description, vbExclamation
    End If
End Sub

Public Sub FormeazaFisa(control As IRibbonControl)
    If ExistaPrefix("PV_") Then
        ActiveSheet.Range("E1").value = " "
        CreeazaFiseDinPV
    Else
        MsgBox "Nu egzista nici un Proces verbal", vbInformation
    End If
End Sub
Public Sub StergePVFise(control As IRibbonControl)
    StergePVsiFise
    'CalcMateriale Nothing
End Sub
Public Sub Statistica(control As IRibbonControl)
    GenereazaStatistica
End Sub
Public Sub ImportPV(control As IRibbonControl)
    ImportaToateFoileDinAltWorkbook
End Sub
Public Sub ExportPV(control As IRibbonControl)
    ExportSheetsByPrefix ("PV_")
End Sub
Public Sub ExportF(control As IRibbonControl)
    ExportSheetsByPrefix ("F_")
End Sub
Public Sub Autor(control As IRibbonControl)
    MsgBox "Sef SPAS, DSPPAS Brinza Ghenadie", vbInformation
End Sub
'================================
'
'================================
Public Sub ExportSheetsByPrefix(Prefix As String)
    Dim wbSource As Workbook
    Dim wbNew As Workbook
    Dim ws As Worksheet
    Dim matched As Collection
    Dim i As Long
    Dim sheetNames() As Variant
    Dim filePath As String, fileName As String
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Set wbSource = ThisWorkbook
    ' Verificam ca fisierul sursa e salvat (pentru a avea o cale valida)
    If Len(wbSource.Path) = 0 Then
        MsgBox "Fisierul curent nu este salvat. Te rog salveaza-l inainte de export (File > Save).", vbExclamation
        GoTo CleanExit
    End If
    ' Colectam foile care incep cu prefixul
    Set matched = New Collection
    For Each ws In wbSource.Worksheets
        If ws.Name Like Prefix & "*" Then
            matched.Add ws.Name
        End If
    Next ws
    If matched.count = 0 Then
        MsgBox "Nu am gasit foi care incep cu '" & Prefix & "'.", vbInformation
        GoTo CleanExit
    End If
    ' Construim un array cu numele foilor potrivite
    ReDim sheetNames(1 To matched.count)
    For i = 1 To matched.count
        sheetNames(i) = matched(i)
    Next i
    ' Copiem TOATE foile intr-o singura operatiune intr-un workbook nou.
    ' Acest mod pastreaza stilurile/tema si previne schimbarile de font.
    wbSource.Worksheets(sheetNames).Copy
    Set wbNew = ActiveWorkbook
    ' (Optional) Fortam fontul stilului "Normal" sa fie cel din sursa, ca asigurare suplimentara.
    ' Daca nu vrei aceasta fortare, poti comenta blocul de mai jos.
    On Error Resume Next
    With wbNew.Styles("Normal").Font
        .Name = wbSource.Styles("Normal").Font.Name
        .Size = wbSource.Styles("Normal").Font.Size
        ' Proprietatea ThemeFont poate sa nu fie disponibila in toate versiunile.
        ' Daca exista, va fi setata; daca nu, instructiunea e ignorata.
        .ThemeFont = wbSource.Styles("Normal").Font.ThemeFont
    End With
    On Error GoTo ErrHandler
    ' Salvam in acelasi folder cu workbook-ul sursa
    filePath = wbSource.Path & Application.PathSeparator
    fileName = SanitizeFileName(Prefix) & Format(Now, "yyyymmdd_HHMM") & ".xlsx"
    wbNew.SaveAs fileName:=filePath & fileName, FileFormat:=xlOpenXMLWorkbook
    wbNew.Close SaveChanges:=True
    MsgBox "Fisier creat: " & filePath & fileName, vbInformation
CleanExit:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    MsgBox "Eroare: " & Err.Number & " - " & Err.Description, vbCritical
    Resume CleanExit
End Sub
Private Function SanitizeFileName(ByVal s As String) As String
    Dim bad As Variant, ch As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each ch In bad
        s = Replace(s, ch, "_")
    Next
    SanitizeFileName = s
End Function
Public Sub Buton_ExportCuInput()
    Dim pfx As String
    pfx = InputBox("Prefixul foilor de export (ex: PV_):", "Export foi")
    If Trim(pfx) <> "" Then ExportSheetsByPrefix pfx
End Sub
'================================
Public Sub StergePVsiFise(Optional ByVal cereConfirmare As Boolean = True)
    Dim i As Long
    Dim deletableCount As Long
    Dim wasProt As Boolean
    Dim prevEvents As Boolean
    Dim prevScreen As Boolean
    Dim prevAlerts As Boolean
    Dim raspuns As VbMsgBoxResult
    On Error GoTo ErrHandler
    ' 1) Determina cate foi se vor sterge
    deletableCount = 0
    For i = 1 To ThisWorkbook.Worksheets.count
        With ThisWorkbook.Worksheets(i)
            If Left$(.Name, 3) = "PV_" Or Left$(.Name, 2) = "F_" Then
                deletableCount = deletableCount + 1
            End If
        End With
    Next i
    If deletableCount = 0 Then
        MsgBox "Nu s-au gasit foi de sters cu prefixele 'PV_' sau 'F_'.", vbInformation
        Exit Sub
    End If
    ' Siguranta: nu permite stergerea tuturor foilor (Excel necesita cel putin o foaie)
    If deletableCount >= ThisWorkbook.Worksheets.count Then
        MsgBox "Operatia a fost oprita: ar rezulta 0 foi in registru. Adauga/retine cel putin o foaie care nu incepe cu 'PV_' sau 'F_'.", vbExclamation
        Exit Sub
    End If
    ' 2) Confirmare (optional)
    If cereConfirmare Then
        raspuns = MsgBox("Se vor sterge " & deletableCount & " foi (PV_* si F_*)." & vbCrLf & _
                         "Esti sigur ca vrei sa continui?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmare stergere")
        If raspuns <> vbYes Then Exit Sub
    End If
    ' 3) Salveaza stari si pregateste mediul
    prevEvents = Application.EnableEvents: Application.EnableEvents = False
    prevScreen = Application.ScreenUpdating: Application.ScreenUpdating = False
    prevAlerts = Application.DisplayAlerts: Application.DisplayAlerts = False
    wasProt = ThisWorkbook.ProtectStructure
    If wasProt Then ThisWorkbook.Unprotect ' fara parola (conform contextului tau)
    ' 4) Sterge foile tinta in ordine inversa
    For i = ThisWorkbook.Worksheets.count To 1 Step -1
        With ThisWorkbook.Worksheets(i)
            If Left$(.Name, 3) = "PV_" Or Left$(.Name, 2) = "F_" Then
                .Delete
            End If
        End With
    Next i
    ' 5) Restaurare stari
    If wasProt Then ThisWorkbook.Protect Structure:=True, Windows:=False
    Application.DisplayAlerts = prevAlerts
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = prevEvents
    MsgBox "Au fost sterse " & deletableCount & " foi (PV_* si F_*).", vbInformation
    Exit Sub
ErrHandler:
    ' Restaurare stari in caz de eroare
    On Error Resume Next
    If wasProt Then ThisWorkbook.Protect Structure:=True, Windows:=False
    Application.DisplayAlerts = prevAlerts
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = prevEvents
    On Error GoTo 0
    MsgBox "Eroare la stergerea foilor PV_/F_: " & Err.Description, vbExclamation
End Sub
' ============================================
' Functii modul
' ============================================
Private Function ExistaPrefix(ByVal Prefix As String) As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, Len(Prefix)) = Prefix Then
            ExistaPrefix = True
            Exit Function
        End If
    Next ws
    ExistaPrefix = False
End Function
Public Function CautaValoareRand(numeFoaie As String, NumeCelula As String) As Long
    Dim wsSursa As Worksheet, wsActiv As Worksheet
    Dim ValoareCautata As String
    Dim i As Long, lastRow As Long
    Set wsSursa = SafeGetSheet(numeFoaie)
    Set wsActiv = ActiveSheet
    ValoareCautata = Trim$(CStr(wsSursa.Range(NumeCelula).value))
    ' Cauta in coloana A, de la randul 61 pana la ultimul rand cu date
    lastRow = wsActiv.Cells(wsActiv.Rows.count, "A").End(xlUp).Row
    If lastRow < 61 Then lastRow = 61
    For i = 61 To lastRow
        If Trim$(CStr(wsActiv.Cells(i, "A").value)) = ValoareCautata Then
            CautaValoareRand = i
            Exit Function
        End If
    Next i
    ' Cauta in coloana D, de la randul 61 in jos
    lastRow = wsActiv.Cells(wsActiv.Rows.count, "D").End(xlUp).Row
    If lastRow < 61 Then lastRow = 61
    For i = 61 To lastRow
        If Trim$(CStr(wsActiv.Cells(i, "D").value)) = ValoareCautata Then
            CautaValoareRand = i
            Exit Function
        End If
    Next i
    ' Nu s-a gasit ? returneaza 0 (fara mesaj, pentru compatibilitate)
    CautaValoareRand = 0
End Function
Public Function CautaSiReturneaza( _
    ByVal NumeleFoii As String, _
    ByVal NumeleColoaneiCautare As String, _
    ByVal ValoareaCautata As String, _
    ByVal NumeleColoaneiRezultat As String) As String
    On Error GoTo Eroare
    Dim ws As Worksheet
    Dim colIndexCautare As Variant, colIndexRezultat As Variant
    Dim targetRow As Long, valoareRezultat As Variant
    Dim rngCautare As Range
    Dim lastRow As Long, i As Long
    Set ws = SafeGetSheet(NumeleFoii)
    colIndexCautare = Application.Match(Trim(NumeleColoaneiCautare), ws.Rows(1), 0)
    If IsError(colIndexCautare) Then GoTo Eroare
    colIndexRezultat = Application.Match(Trim(NumeleColoaneiRezultat), ws.Rows(1), 0)
    If IsError(colIndexRezultat) Then GoTo Eroare
    lastRow = ws.Cells(ws.Rows.count, CLng(colIndexCautare)).End(xlUp).Row
    If lastRow < 2 Then GoTo Eroare
    Set rngCautare = ws.Range(ws.Cells(2, CLng(colIndexCautare)), ws.Cells(lastRow, CLng(colIndexCautare)))
    For i = 1 To rngCautare.Rows.count
        If Trim(CStr(rngCautare.Cells(i, 1).value)) = Trim(CStr(ValoareaCautata)) Then
            targetRow = rngCautare.Cells(i, 1).Row
            valoareRezultat = ws.Cells(targetRow, CLng(colIndexRezultat)).value
            CautaSiReturneaza = IIf(IsEmpty(valoareRezultat), "", CStr(valoareRezultat))
            Exit Function
        End If
    Next i
Eroare:
    CautaSiReturneaza = ""
End Function
Public Sub InsereazaRanduriDinPVModel(Diapazon As String)
    Dim wsSursa As Worksheet, wsDest As Worksheet
    Dim rngTarget As Range
    Dim targetRow As Long, nrRanduri As Long
    Dim startRow As Long, endRow As Long
    On Error Resume Next
    Set rngTarget = Application.InputBox("Selecteaza o celula deasupra careia sa se insereze rindurile:", "Alege locatia", Type:=8)
    On Error GoTo 0
    If rngTarget Is Nothing Then
        MsgBox "Operatie anulata.", vbExclamation
        Exit Sub
    End If
    Set wsDest = rngTarget.Worksheet
    targetRow = rngTarget.Row
    startRow = CLng(Split(Diapazon, ":")(0))
    endRow = CLng(Split(Diapazon, ":")(1))
    nrRanduri = endRow - startRow + 1
    wsDest.Rows(targetRow & ":" & (targetRow + nrRanduri - 1)).Insert Shift:=xlDown
    Set wsSursa = SafeGetSheet("PVModel")
    wsSursa.Rows(Diapazon).Copy Destination:=wsDest.Rows(targetRow)
    'MsgBox "Rindurile " & Diapazon & " din foaia 'PVModel' au fost inserate deasupra rindului " & targetRow & ".", vbInformation
End Sub
Public Sub InserareRandCopy(Valori As String, Coloane As String, Optional InserareRand As Variant)
    Dim rng As Range
    Dim arrValori() As String, arrColoane() As String
    Dim i As Long
    If IsMissing(InserareRand) Then
        Dim raspuns As VbMsgBoxResult
        raspuns = MsgBox("Doriti sa inserati un rind nou deasupra selectiei?" & vbCrLf & _
                         "— Da: se insereaza rind nou" & vbCrLf & _
                         "— Nu: se scrie in rindul selectat", _
                         vbQuestion + vbYesNo, "Alegeti modul de inserare")
        InserareRand = (raspuns = vbYes)
    End If
    On Error Resume Next
    Set rng = Application.InputBox("Selecteaza o celula pe foaia activa:", "Alege celula", Type:=8)
    On Error GoTo 0
    If rng Is Nothing Then
        MsgBox "Operatie anulata.", vbExclamation
        Exit Sub
    End If
    If InserareRand Then
        Rindul = rng.Row
        rng.EntireRow.Insert Shift:=xlDown
    Else
        Rindul = rng.Row
    End If
    arrValori = Split(Valori, ";")
    arrColoane = Split(Coloane, ";")
    If UBound(arrValori) <> UBound(arrColoane) Then
        MsgBox "Numarul valorilor nu corespunde cu numarul coloanelor.", vbCritical
        Exit Sub
    End If
    For i = LBound(arrValori) To UBound(arrValori)
        On Error Resume Next
        Dim colIndex As Long
        colIndex = rng.Worksheet.Columns(Trim(arrColoane(i))).Column
        If Err.Number <> 0 Then
            MsgBox "Coloana invalida: " & arrColoane(i), vbCritical
            Exit Sub
        End If
        On Error GoTo 0
        rng.Worksheet.Cells(Rindul, colIndex).value = arrValori(i)
    Next i
    'MsgBox "Valorile au fost adaugate in randul " & Rindul & ".", vbInformation
End Sub
' ============================================
' Populare UI (UserForm)
' ============================================
Public Sub PopuleazaComboBoxCuColoaneDinFoaiaCurenta(ByVal sheetName As String)
    Dim ws As Worksheet
    Dim i As Integer, ultimaColoana As Integer
    Dim comboBox As MSForms.comboBox
    Set comboBox = UserForm1.Controls("ComboBox1")
    On Error GoTo ErrHandler
    Set ws = SafeGetSheet(sheetName)
    comboBox.Clear
    ultimaColoana = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    For i = 1 To ultimaColoana
        comboBox.AddItem ws.Cells(1, i).value
    Next i
    If comboBox.ListCount > 0 Then comboBox.ListIndex = 0
    Exit Sub
ErrHandler:
    MsgBox "Eroare la PopuleazaComboBoxCuColoaneDinFoaiaCurenta pentru '" & sheetName & "': " & Err.Description, vbExclamation
End Sub
Public Sub PopuleazaListBoxCuDateDinFoaiaCurenta(ByVal sheetName As String, Optional ByVal latimeColoane As String = "")
    Dim ws As Worksheet
    Dim listBox As MSForms.listBox
    Dim i As Long, j As Integer, k As Long
    Dim ultimaColoana As Integer, ultimulRand As Long
    Dim matriceDate As Variant
    Dim latimeMaxima() As Single
    Set listBox = UserForm1.Controls("ListBox1")
    On Error GoTo ErrHandler
    Set ws = SafeGetSheet(sheetName)
    ultimulRand = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    ultimaColoana = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    If ultimulRand < 2 Or ultimaColoana < 1 Then
        MsgBox "Foaia '" & sheetName & "' nu contine date suficiente!", vbExclamation
        Exit Sub
    End If
    ReDim latimeMaxima(1 To ultimaColoana)
    matriceDate = ws.Range(ws.Cells(1, 1), ws.Cells(ultimulRand, ultimaColoana)).value
    listBox.Clear
    listBox.ColumnCount = ultimaColoana
    Dim latimi As String
    If latimeColoane = "" Then
        For j = 1 To ultimaColoana
            latimeMaxima(j) = Len(CStr(matriceDate(1, j))) * 8 + 20
            For i = 2 To UBound(matriceDate, 1)
                Dim lungimeText As Single
                lungimeText = Len(CStr(matriceDate(i, j))) * 8 + 20
                If lungimeText > latimeMaxima(j) Then latimeMaxima(j) = lungimeText
            Next i
        Next j
        For j = 1 To ultimaColoana
            If j > 1 Then latimi = latimi & ";"
            latimi = latimi & CStr(latimeMaxima(j)) & " pt"
        Next j
    Else
        Dim liniiLatime() As String
        liniiLatime = Split(latimeColoane, ";")
        For j = 1 To ultimaColoana
            If j > 1 Then latimi = latimi & ";"
            If j <= UBound(liniiLatime) + 1 Then
                latimi = latimi & liniiLatime(j - 1) & " pt"
            Else
                latimi = latimi & "100 pt"
            End If
        Next j
    End If
    listBox.ColumnWidths = latimi
    For i = 2 To ultimulRand
        listBox.AddItem ""
        k = i - 2
        For j = 1 To ultimaColoana
            listBox.List(k, j - 1) = CStr(matriceDate(i, j))
        Next j
    Next i
    listBox.Tag = sheetName & "|" & ultimaColoana
    Dim numeColoane As String
    For j = 1 To ultimaColoana
        If j > 1 Then numeColoane = numeColoane & "|"
        numeColoane = numeColoane & CStr(matriceDate(1, j))
    Next j
    If InStr(UserForm1.Tag, "CautareObiecte") > 0 Then
        UserForm1.Tag = "CautareObiecte|" & numeColoane
    ElseIf InStr(UserForm1.Tag, "CautareNorme") > 0 Then
        UserForm1.Tag = "CautareNorme|" & numeColoane
    End If
    Exit Sub
ErrHandler:
    MsgBox "Eroare la PopuleazaListBoxCuDateDinFoaiaCurenta pentru '" & sheetName & "': " & Err.Description, vbExclamation
End Sub
Public Sub FiltreazaListBoxDinFoaiaCurenta(ByVal sheetName As String, ByVal Criteriu As String, ByVal Coloana As String)
    Static ultimaColoanaSelectata As String
    Static dateCacheCompleteMatrix As Variant
    Dim ws As Worksheet
    Dim i As Long, j As Integer
    Dim ultimaColoana As Integer, ultimulRand As Long
    Dim indexColoana As Integer
    Dim listBox As MSForms.listBox
    Dim latimeColoane As String
    Set listBox = UserForm1.Controls("ListBox1")
    If IsEmpty(dateCacheCompleteMatrix) Or ultimaColoanaSelectata <> Coloana Then
        On Error GoTo ErrHandler
        Set ws = SafeGetSheet(sheetName)
        ultimulRand = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
        ultimaColoana = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
        dateCacheCompleteMatrix = ws.Range(ws.Cells(1, 1), ws.Cells(ultimulRand, ultimaColoana)).value
        indexColoana = 0
        For j = 1 To ultimaColoana
            If CStr(dateCacheCompleteMatrix(1, j)) = Coloana Then indexColoana = j: Exit For
        Next j
        If indexColoana = 0 Then
            MsgBox "Coloana selectata nu exista!", vbCritical
            Exit Sub
        End If
        ultimaColoanaSelectata = Coloana
    Else
        ultimulRand = UBound(dateCacheCompleteMatrix, 1)
        ultimaColoana = UBound(dateCacheCompleteMatrix, 2)
        indexColoana = 0
        For j = 1 To ultimaColoana
            If CStr(dateCacheCompleteMatrix(1, j)) = Coloana Then indexColoana = j: Exit For
        Next j
    End If
    latimeColoane = listBox.ColumnWidths
    listBox.Clear
    listBox.ColumnCount = ultimaColoana
    If latimeColoane <> "" Then listBox.ColumnWidths = latimeColoane
    Dim filteredData() As Variant, filteredCount As Long
    Dim criteriu_lcase As String
    criteriu_lcase = LCase(Trim(Criteriu))
    filteredCount = 0
    If criteriu_lcase = "" Then
        filteredCount = ultimulRand - 1
        ReDim filteredData(1 To filteredCount, 1 To ultimaColoana)
        For i = 2 To ultimulRand
            For j = 1 To ultimaColoana
                filteredData(i - 1, j) = dateCacheCompleteMatrix(i, j)
            Next j
        Next i
    Else
        Dim tempArray() As Long
        ReDim tempArray(1 To ultimulRand - 1)
        For i = 2 To ultimulRand
            If InStr(1, LCase(CStr(dateCacheCompleteMatrix(i, indexColoana))), criteriu_lcase) > 0 Then
                filteredCount = filteredCount + 1
                tempArray(filteredCount) = i
            End If
        Next i
        If filteredCount > 0 Then
            ReDim filteredData(1 To filteredCount, 1 To ultimaColoana)
            For i = 1 To filteredCount
                For j = 1 To ultimaColoana
                    filteredData(i, j) = dateCacheCompleteMatrix(tempArray(i), j)
                Next j
            Next i
        End If
    End If
    If filteredCount > 0 Then
        Dim r As Long, c As Long
        For r = 1 To filteredCount
            listBox.AddItem ""
        Next r
        For r = 1 To filteredCount
            For c = 1 To ultimaColoana
                listBox.List(r - 1, c - 1) = CStr(filteredData(r, c))
            Next c
        Next r
    End If
    Exit Sub
ErrHandler:
    MsgBox "Eroare la FiltreazaListBoxDinFoaiaCurenta pentru '" & sheetName & "': " & Err.Description, vbExclamation
End Sub
' ============================================
' Add Obiecte/Norma/Materiale/Utilaj/Transport (UserForm)
' ============================================
Sub AddObiect()
    PopuleazaComboBoxCuColoaneDinFoaiaCurenta "Obiect"
    UserForm1.Tag = "CautareObiect"
    PopuleazaListBoxCuDateDinFoaiaCurenta "Obiect", "45;170"
    UserForm1.Show
End Sub
Sub AddNorma()
    PopuleazaComboBoxCuColoaneDinFoaiaCurenta "Norma"
    UserForm1.Tag = "CautareNorma"
    PopuleazaListBoxCuDateDinFoaiaCurenta "Norma", "60;350;40;50"
    UserForm1.Show
End Sub
Sub AddMateriale()
    PopuleazaComboBoxCuColoaneDinFoaiaCurenta "Materiale"
    UserForm1.Tag = "CautareMateriale"
    PopuleazaListBoxCuDateDinFoaiaCurenta "Materiale", "90;285;50;80"
    UserForm1.Show
End Sub
Sub AddUtilaj()
    PopuleazaComboBoxCuColoaneDinFoaiaCurenta "Utilaj"
    UserForm1.Tag = "CautareUtilaj"
    PopuleazaListBoxCuDateDinFoaiaCurenta "Utilaj", "90;285;50"
    UserForm1.Show
End Sub
Sub AddTransport()
    PopuleazaComboBoxCuColoaneDinFoaiaCurenta "Transport"
    UserForm1.Tag = "CautareTransport"
    PopuleazaListBoxCuDateDinFoaiaCurenta "Transport", "50;50;200;40;50;40;50"
    UserForm1.Show
End Sub
' ============================================
' Creare PV nou – Copiere robusta + redenumire corecta + pozitionare corecta
' ============================================
Public Sub CopiazaFoaie(ByVal numeSursa As String, ByVal Prefix As String)
    Dim wb As Workbook
    Dim wsSursa As Worksheet
    Dim wsNou As Worksheet
    Dim ws As Worksheet
    Dim maxNr As Long, nrCurent As Long
    Dim parteNumar As String
    Dim numeNou As String
    Dim prevVis As XlSheetVisibility
    Dim wasProt As Boolean
    Dim prevEvents As Boolean
    Dim prevScreen As Boolean
    Set wb = ThisWorkbook
    On Error GoTo ErrHandler
    prevEvents = Application.EnableEvents: Application.EnableEvents = False
    prevScreen = Application.ScreenUpdating: Application.ScreenUpdating = False
    wasProt = UnprotectStructureIfNeeded()
    Set wsSursa = SafeGetSheet(numeSursa)
    ' Copiaza modelul la final si leaga-te de copia ACTIVA (Excel o activeaza automat)
    WithTempVisibility wsSursa, prevVis
    wsSursa.Copy After:=wb.Worksheets(wb.Worksheets.count)
    RestoreVisibility wsSursa, prevVis
    Set wsNou = ActiveSheet   ' esential: referinta exact pe copia nou creata
    ' ?? ?? Dezactiveaza protectia pe foaia noua (daca exista)
    On Error Resume Next
    wsNou.Unprotect
    On Error GoTo 0
    ' Determina maxNr din foile care respecta strict pattern-ul PV_<numar>
    maxNr = 0
    For Each ws In wb.Worksheets
        If Left(ws.Name, Len(Prefix)) = Prefix Then
            parteNumar = Trim$(Mid(ws.Name, Len(Prefix) + 1))
            If Len(parteNumar) > 0 And IsNumeric(parteNumar) Then
                nrCurent = CLng(Val(parteNumar))
                If nrCurent > maxNr Then maxNr = nrCurent
            End If
        End If
    Next ws
    ' Generator de nume robust: cauta primul liber (verificare pe Sheets – include ChartSheets)
    nrCurent = maxNr + 1
    numeNou = Prefix & CStr(nrCurent)
    Do While SheetExistsAll(numeNou)
        nrCurent = nrCurent + 1
        numeNou = Prefix & CStr(nrCurent)
    Loop
    ' Redenumeste doar COPIA (wsNou)
    wsNou.Name = numeNou
    ' Pozitionare:
    ' - daca exista PV-uri: dupa ultimul PV_ (dupa numarul max gasit initial, daca exista inca)
    ' - daca NU exista PV_: dupa CalculMateriale (daca exista), altfel lasa la final
    If maxNr > 0 Then
        Dim anchorPV As Worksheet
        Set anchorPV = Nothing
        For Each ws In wb.Worksheets
            If StrComp(ws.Name, Prefix & CStr(maxNr), vbTextCompare) = 0 Then
                Set anchorPV = ws: Exit For
            End If
        Next ws
        If Not anchorPV Is Nothing Then
            If Not wsNou Is anchorPV Then wsNou.Move After:=anchorPV
        End If
    Else
        If SheetExistsAll("CalculMateriale") Then
            wsNou.Move After:=wb.Worksheets("CalculMateriale")
        Else
            ' deja la final
        End If
    End If
    wsNou.Activate
Clean:
    RestoreStructureIfNeeded wasProt
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    Exit Sub
ErrHandler:
    RestoreStructureIfNeeded wasProt
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    MsgBox "Eroare la copierea foii '" & numeSursa & "': " & Err.Description, vbExclamation
End Sub
' ============================================
' Creare Fise din PV
' ============================================
Public Sub CreeazaFiseDinPV()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsModel As Worksheet
    Dim wsPV As Worksheet, wsF As Worksheet
    Dim nrCurent As Long
    Dim numePV As String, numeF As String
    Dim i As Long, j As Long
    Dim maxNr As Long
    Dim cel As Range, cel2 As Range
    Dim rF As Long
    Dim listaNr() As Long
    Dim countNr As Long
    Dim ultimaPV As Worksheet
    Dim prevVis As XlSheetVisibility
    Dim wasProt As Boolean
    Dim prevEvents As Boolean
    Set wb = ThisWorkbook
    On Error GoTo ErrHandler
    Set wsModel = SafeGetSheet("FisaModel")
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    prevEvents = Application.EnableEvents: Application.EnableEvents = False
    wasProt = UnprotectStructureIfNeeded()
    ' Sterge F_* existente
    For i = wb.Worksheets.count To 1 Step -1
        If Left(wb.Worksheets(i).Name, 2) = "F_" Then
            wb.Worksheets(i).Delete
        End If
    Next i
    ' Gaseste ultima foaie PV_
    For Each ws In wb.Worksheets
        If Left(ws.Name, 3) = "PV_" Then Set ultimaPV = ws
    Next ws
    ' Creeaza F_ pentru fiecare PV_
    For Each ws In wb.Worksheets
        If Left(ws.Name, 3) = "PV_" Then
            numePV = ws.Name
            nrCurent = Val(Mid(numePV, 4))
            If nrCurent > 0 Then
                numeF = "F_" & nrCurent
                WithTempVisibility wsModel, prevVis
                wsModel.Copy After:=ultimaPV
                RestoreVisibility wsModel, prevVis
                Set wsF = ultimaPV.Next
                wsF.Name = numeF
                ' Transfer date
                On Error Resume Next
                wsF.Range("A1").value = wb.Worksheets(numePV).Range("C18").value
                wsF.Range("C3").value = CautaSiReturneaza("Liste", "NumeSector", wb.Worksheets(numePV).Range("C11").value, "Sector")
                wsF.Range("H5").value = CautaSiReturneaza("Liste", "Nr", Month(DateAdd("m", -1, Date)), "Luna")
                wsF.Range("C6").value = wb.Worksheets(numePV).Range("C15").value
                wsF.Range("E11").value = CautaSiReturneaza("Liste", "NumeSector", wb.Worksheets(numePV).Range("C11").value, "Brigadir")
                wsF.Range("N11").value = CautaSiReturneaza("Liste", "NumeSector", wb.Worksheets(numePV).Range("C11").value, "NrMatricol")
                wsF.Range("M6").value = wb.Worksheets(numePV).Range("C17").value
                wsF.Range("E13").value = wb.Worksheets(numePV).Range("H29").value
                wsF.Range("B18").value = wb.Worksheets(numePV).Range("C25").value
                wsF.Range("H22").value = CautaSiReturneaza("Liste", "NumeSector", wb.Worksheets(numePV).Range("C11").value, "SefSector")
                wsF.Range("H26").value = CautaSiReturneaza("Liste", "NumeSector", wb.Worksheets(numePV).Range("C11").value, "SefSector")
                wsF.Range("F28").value = CautaSiReturneaza("Liste", "NumeSector", wb.Worksheets(numePV).Range("C11").value, "BrigadirScurt")
                wsF.Range("V22").value = CautaSiReturneaza("Liste", "NumeSector", wb.Worksheets(numePV).Range("C11").value, "BrigadirScurt")
                wsF.Range("V24").value = CautaSiReturneaza("Liste", "NumeSector", wb.Worksheets(numePV).Range("C11").value, "Normator")
                wsF.Range("V28").value = CautaSiReturneaza("Liste", "NumeSector", wb.Worksheets(numePV).Range("C11").value, "SefDepartament")
                wsF.Range("V31").value = CautaSiReturneaza("Liste", "NumeSector", wb.Worksheets(numePV).Range("C11").value, "SefSIDTP")
                On Error GoTo 0
                ' Colecteaza markeri numerici pe A1:A200
                countNr = 0
                Set wsPV = wb.Worksheets(numePV)
                For Each cel In wsPV.Range("A1:A200")
                    If cel.Interior.Color = RGB(217, 217, 217) Then
                        If IsNumeric(cel.value) Then
                            countNr = countNr + 1
                            ReDim Preserve listaNr(1 To countNr)
                            listaNr(countNr) = cel.value
                        End If
                    End If
                Next cel
                If countNr >= 2 Then
                    maxNr = listaNr(countNr - 1)
                ElseIf countNr = 1 Then
                    maxNr = listaNr(1)
                Else
                    maxNr = 0
                End If
                If maxNr > 1 Then
                    For i = 1 To maxNr - 1
                        wsF.Rows(19).Copy
                        wsF.Rows(20).Insert Shift:=xlDown
                    Next i
                End If
                rF = 19
                For j = 1 To countNr - 1
                    For Each cel2 In wsPV.Range("A1:A200")
                        If cel2.Interior.Color = RGB(217, 217, 217) Then
                            If IsNumeric(cel2.value) And cel2.value = listaNr(j) Then
                                wsF.Cells(rF, "A").value = wsPV.Cells(cel2.Row, "B").value
                                wsF.Cells(rF, "B").value = wsPV.Cells(cel2.Row, "C").value
                                wsF.Cells(rF, "H").value = wsPV.Cells(cel2.Row, "D").value
                                wsF.Cells(rF, "I").value = wsPV.Cells(cel2.Row, "E").value
                                wsF.Cells(rF, "K").value = wsPV.Cells(cel2.Row, "H").value
                                rF = rF + 1
                                Exit For
                            End If
                        End If
                    Next cel2
                Next j
                If maxNr >= 1 Then
                    wsF.Cells(20 + (maxNr - 1), "R").Formula = "=SUM(R19:R" & 19 + (maxNr - 1) & ")"
                    wsF.Cells(20 + (maxNr - 1), "U").Formula = "=SUM(U19:U" & 19 + (maxNr - 1) & ")"
                End If
            End If
        End If
    Next ws
    RestoreStructureIfNeeded wasProt
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = prevEvents
    MsgBox "Fisele au fost recreate si completate conform Proceselor Verbale.", vbInformation
    Exit Sub
ErrHandler:
    RestoreStructureIfNeeded wasProt
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = prevEvents
    MsgBox "Eroare la CreeazaFiseDinPV: " & Err.Description, vbExclamation
End Sub
' ============================================
' Calcul sectiuni materiale
' ============================================
Public Sub CompleteazaDisponibilSiDiferentaMateriale()
    Dim wsCalc As Worksheet, wsMat As Worksheet
    Dim colCodMat As Variant, colQtyMat As Variant
    Dim lastRowCalc As Long, lastRowMat As Long
    Dim dict As Object
    Dim i As Long
    Dim code As String
    Dim v As Variant
    Dim disponibil As Double, necesar As Double, diff As Double
    On Error GoTo ErrHandler
    Set wsCalc = ThisWorkbook.Worksheets("CalculMateriale")
    Set wsMat = ThisWorkbook.Worksheets("Materiale")
    colCodMat = Application.Match("CodMaterial", wsMat.Rows(1), 0)
    colQtyMat = Application.Match("Cantitate", wsMat.Rows(1), 0)
    If IsError(colCodMat) Or IsError(colQtyMat) Then
        MsgBox "Nu am gasit antetele 'CodMaterial' si/sau 'Cantitate' in foaia 'Materiale'.", vbExclamation
        Exit Sub
    End If
    lastRowCalc = wsCalc.Cells(wsCalc.Rows.count, "C").End(xlUp).Row
    If lastRowCalc < 2 Then Exit Sub
    wsCalc.Range("F2:G" & lastRowCalc).ClearContents
    Set dict = CreateObject("Scripting.Dictionary")
    lastRowMat = wsMat.Cells(wsMat.Rows.count, CLng(colCodMat)).End(xlUp).Row
    For i = 2 To lastRowMat
        code = Trim$(CStr(wsMat.Cells(i, CLng(colCodMat)).value))
        If code <> "" Then
            v = wsMat.Cells(i, CLng(colQtyMat)).value
            If IsNumeric(v) Then
                If dict.Exists(code) Then
                    dict(code) = CDbl(dict(code)) + CDbl(v)
                Else
                    dict.Add code, CDbl(v)
                End If
            End If
        End If
    Next i
    For i = 2 To lastRowCalc
        code = Trim$(CStr(wsCalc.Cells(i, "C").value))
        If code <> "" Then
            v = wsCalc.Cells(i, "D").value
            If IsNumeric(v) Then necesar = CDbl(v) Else necesar = 0
            If dict.Exists(code) Then disponibil = CDbl(dict(code)) Else disponibil = 0
            diff = disponibil - necesar
            wsCalc.Cells(i, "F").value = Application.WorksheetFunction.Round(disponibil, 3)
            wsCalc.Cells(i, "G").value = Application.WorksheetFunction.Round(diff, 3)
        End If
    Next i
    Exit Sub
ErrHandler:
    MsgBox "Eroare la CompleteazaDisponibilSiDiferentaMateriale: " & Err.Description, vbExclamation
End Sub
Public Sub totalManopera()
    Dim ws As Worksheet
    Dim wsCalc As Worksheet
    Dim totalOre As Double
    Dim rngC As Range, f As Range
    Dim firstAddress As String
    On Error GoTo ErrHandler
    Set wsCalc = ThisWorkbook.Worksheets("CalculMateriale")
    totalOre = 0
    For Each ws In ThisWorkbook.Worksheets
        If Left$(ws.Name, 3) = "PV_" Then
            Set rngC = ws.Columns("C")
            Set f = rngC.Find(What:="inclusiv manopera", LookIn:=xlValues, LookAt:=xlWhole, _
                              SearchOrder:=xlByRows, MatchCase:=False)
            If Not f Is Nothing Then
                firstAddress = f.Address
                Do
                    If IsNumeric(ws.Cells(f.Row, "E").value) Then
                        totalOre = totalOre + CDbl(ws.Cells(f.Row, "E").value)
                    End If
                    Set f = rngC.FindNext(f)
                    If f Is Nothing Then Exit Do
                Loop While f.Address <> firstAddress
            End If
        End If
    Next ws
    wsCalc.Range("L1").value = Application.WorksheetFunction.Round(totalOre, 3)
    With wsCalc
        If .Visible <> xlSheetVisible Then .Visible = xlSheetVisible
        .Activate
        'Application.Goto .Range("L1"), True
    End With
    Exit Sub
ErrHandler:
    MsgBox "Eroare la TotalManopera: " & Err.Description, vbExclamation
End Sub
' ============================================
' Flux CalculMateriale
' ============================================
Sub cautaValoarea()
    Dim ws As Worksheet
    Dim firstPage As Worksheet
    Dim cell As Range
    Dim lastRowFirstPage As Long
    Dim sum As Long
    sum = 0
    Set firstPage = ThisWorkbook.Sheets("CalculMateriale")
    lastRowFirstPage = firstPage.Cells(firstPage.Rows.count, "B").End(xlUp).Row + 1
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> firstPage.Name And ws.Name Like "PV_*" Then
            For Each cell In ws.Range("B1:B" & ws.Cells(ws.Rows.count, "B").End(xlUp).Row)
                If Left(cell.value, 3) = "211" Or Left(cell.value, 3) = "921" Then
                    firstPage.Cells(lastRowFirstPage, "B").value = cell.value
                    firstPage.Cells(lastRowFirstPage, "A").value = cell.Offset(0, 3).value
                    lastRowFirstPage = lastRowFirstPage + 1
                    sum = sum + 1
                End If
            Next cell
        End If
    Next ws
    If sum = 0 Then
        MsgBox "Nu a fost gasit nr. de nomenclator cu cifrele de inceput 211 ori 9211"
        Exit Sub
    End If
    CautaValoareaUnica
    CompareAndSum2
    SumaMat
    CompleteazaDisponibilSiDiferentaMateriale
    totalManopera
End Sub
Sub CautaValoareaUnica()
    Dim firstPage As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    Dim uniqueValues As Collection
    Dim value As Variant
    Dim i As Long
    Set firstPage = ThisWorkbook.Sheets("CalculMateriale")
    lastRow = firstPage.Cells(firstPage.Rows.count, "B").End(xlUp).Row
    Set uniqueValues = New Collection
    For Each cell In firstPage.Range("B1:B" & lastRow)
        value = cell.value
        On Error Resume Next
        uniqueValues.Add value, CStr(value)
        On Error GoTo 0
    Next cell
    firstPage.Range("C2:C" & lastRow).ClearContents
    For i = 2 To uniqueValues.count
        firstPage.Cells(i, "C").value = uniqueValues.Item(i)
    Next i
End Sub
Sub SumaMat()
    Dim firstPage As Worksheet
    Dim ws As Worksheet
    Dim cellC As Range
    Dim cellB As Range
    Dim lastRowC As Long
    Dim Denumire As String
    Set firstPage = ThisWorkbook.Sheets("CalculMateriale")
    lastRowC = firstPage.Cells(firstPage.Rows.count, "C").End(xlUp).Row
    For Each cellC In firstPage.Range("C2:C" & lastRowC)
        Denumire = ""
        For Each ws In ThisWorkbook.Sheets
            If ws.Name <> firstPage.Name And ws.Name Like "PV_*" Then
                Set cellB = ws.Range("B:B").Find(cellC.value, LookIn:=xlValues, LookAt:=xlWhole)
                If Not cellB Is Nothing Then
                    Denumire = ws.Cells(cellB.Row, "C").value
                End If
            End If
        Next ws
        cellC.Offset(0, 2).value = Denumire
    Next cellC
End Sub
Sub CompareAndSum2()
    Dim ws As Worksheet
    Dim lastRowC As Long, lastRowB As Long
    Dim i As Long, j As Long
    Dim valueToFind As Variant
    Dim sumValue As Double
    Set ws = ThisWorkbook.Sheets("CalculMateriale")
    lastRowC = ws.Cells(ws.Rows.count, "C").End(xlUp).Row
    lastRowB = ws.Cells(ws.Rows.count, "B").End(xlUp).Row
    For i = 2 To lastRowC
        valueToFind = ws.Cells(i, "C").value
        sumValue = 0
        For j = 2 To lastRowB
            If ws.Cells(j, "B").value = valueToFind Then
                sumValue = sumValue + ws.Cells(j, "A").value
            End If
        Next j
        ws.Cells(i, "D").value = sumValue
    Next i
End Sub
' ============================================
' Afisare temporara foaie (fara mutare pentru model/BD)
' ============================================
Public Sub ShowSheetTemporarily(ByVal sheetName As String)
    Dim ws As Worksheet
    Dim anchor As Worksheet
    Dim wasProt As Boolean
    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Worksheets(sheetName)
    ' Daca avem alta foaie temporara deschisa, o restauram
    If Len(TempVisibleSheetName) > 0 Then
        If StrComp(TempVisibleSheetName, ws.Name, vbTextCompare) <> 0 Then
            RestoreTempVisibility
        End If
    End If
    wasProt = UnprotectStructureIfNeeded()
    TempPrevVisibility = ws.Visible
    TempVisibleSheetName = ws.Name
    If ws.Visible <> xlSheetVisible Then
        ws.Visible = xlSheetVisible
    End If
    ' Nu mutam foile din MODEL_SHEETS (doar vizibil + activate)
    If Not IsModelOrDataSheet(ws.Name) Then
        Set anchor = FindLastAlwaysVisibleSheet()
        If Not anchor Is Nothing Then
            If StrComp(ws.Name, anchor.Name, vbTextCompare) <> 0 Then
                ws.Move After:=anchor
            End If
        End If
    End If
    ws.Activate
Clean:
    RestoreStructureIfNeeded wasProt
    Exit Sub
ErrHandler:
    RestoreStructureIfNeeded wasProt
    MsgBox "Nu pot deschide foaia '" & sheetName & "': " & Err.Description, vbExclamation
    TempVisibleSheetName = ""
End Sub
Public Sub RestoreTempVisibility()
    If Len(TempVisibleSheetName) = 0 Then Exit Sub
    Dim ws As Worksheet
    Dim wasProt As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(TempVisibleSheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        TempVisibleSheetName = ""
        Exit Sub
    End If
    wasProt = UnprotectStructureIfNeeded()
    On Error Resume Next
    ws.Visible = TempPrevVisibility
    On Error GoTo 0
    RestoreStructureIfNeeded wasProt
    TempVisibleSheetName = ""
End Sub
Public Sub ArataListeTemporar()
    ShowSheetTemporarily "Liste"
End Sub

Public Sub GenereazaStatistica()
    Dim ws As Worksheet
    Dim wsStat As Worksheet
    Dim dictManopera As Object
    Dim dictTransport As Object ' Cheie: TipLucrare & "|" & CodTransport
    Dim tipLucrare As String
    Dim codTransport As String
    Dim denumireTransport As String
    Dim motoare As Double, km As Double
    Dim totalManopera As Double
    Dim rngC As Range, f As Range, cellB As Range
    Dim firstAddress As String
    Dim i As Long, lastRowB As Long
    Dim key As String
    Dim wasProt As Boolean
    Dim prevScreen As Boolean, prevEvents As Boolean
    
    ' --- Pregatire mediu ---
    prevScreen = Application.ScreenUpdating: Application.ScreenUpdating = False
    prevEvents = Application.EnableEvents: Application.EnableEvents = False
    wasProt = UnprotectStructureIfNeeded()
    
    On Error GoTo Finalize
    
    ' --- Creeaza/ob?ine foaia Statistica ---
    Set wsStat = Nothing
    On Error Resume Next
    Set wsStat = ThisWorkbook.Worksheets("Statistica")
    On Error GoTo 0
    
    If wsStat Is Nothing Then
        Set wsStat = ThisWorkbook.Worksheets.Add
        wsStat.Name = "Statistica"
    Else
        wsStat.Cells.Clear
    End If
    
    ' --- Ini?ializare dic?ionare ---
    Set dictManopera = CreateObject("Scripting.Dictionary")
    Set dictTransport = CreateObject("Scripting.Dictionary")
    
    ' --- Parcurge toate PV_* ---
    For Each ws In ThisWorkbook.Worksheets
        If Left$(ws.Name, 3) = "PV_" Then
            ' Citim tipul lucrarii
            tipLucrare = Trim$(CStr(ws.Range("C18").value))
            If tipLucrare = "" Then tipLucrare = "(Fara tip)"
            
            ' --- Colectam manopera ---
            Set rngC = ws.Columns("C")
            Set f = rngC.Find(What:="inclusiv manopera", LookIn:=xlValues, LookAt:=xlWhole, _
                              SearchOrder:=xlByRows, MatchCase:=False)
            totalManopera = 0
            If Not f Is Nothing Then
                firstAddress = f.Address
                Do
                    If IsNumeric(ws.Cells(f.Row, "E").value) Then
                        totalManopera = totalManopera + CDbl(ws.Cells(f.Row, "E").value)
                    End If
                    Set f = rngC.FindNext(f)
                    If f Is Nothing Then Exit Do
                Loop While f.Address <> firstAddress
            End If
            
            ' Adaugam la dic?ionarul de manopera
            If dictManopera.Exists(tipLucrare) Then
                dictManopera(tipLucrare) = dictManopera(tipLucrare) + totalManopera
            Else
                dictManopera.Add tipLucrare, totalManopera
            End If
            
' --- Colectam transporturi ---
lastRowB = ws.Cells(ws.Rows.count, "B").End(xlUp).Row
For i = 1 To lastRowB
    Set cellB = ws.Cells(i, "B")
    If Not IsEmpty(cellB.value) Then
        Dim valB As String: valB = Trim$(CStr(cellB.value))
        If Left$(valB, 2) = "50" Or Left$(valB, 2) = "45" Or Left$(valB, 2) = "42" Then
            codTransport = valB
            denumireTransport = Trim$(CStr(ws.Cells(i, "C").value))
            motoare = 0: km = 0
            
            ' Motoarele (F, acela?i rând)
            If IsNumeric(ws.Cells(i, "E").value) Then
                motoare = CDbl(ws.Cells(i, "E").value)
            End If
            
            ' Km (F, rândul urmator)
            If i + 1 <= ws.Rows.count Then
                If IsNumeric(ws.Cells(i + 1, "E").value) Then
                    km = CDbl(ws.Cells(i + 1, "E").value)
                End If
            End If
            
            key = tipLucrare & "|" & codTransport
            
            If dictTransport.Exists(key) Then
                ' Extragem valorile existente
                Dim arr() As Variant
                arr = dictTransport(key)
                ' Actualizam
                arr(0) = arr(0) + motoare
                arr(1) = arr(1) + km
                ' Reatribuim (denumirea ramâne din prima apari?ie ? nu o schimbam)
                dictTransport(key) = arr
            Else
                ' Stocam: (motoare, km, denumire)
                dictTransport.Add key, Array(motoare, km, denumireTransport)
            End If
        End If
    End If
Next i
        End If
    Next ws
    
    ' --- Scriem în foaia Statistica ---
    With wsStat
        .Range("A1:E1").value = Array("Tip Lucrare", "Manopera (h-om)", "Nume Transport", "M-ht", "Km")
        .Range("A1:E1").Font.Bold = True
        
        Dim r As Long: r = 2
        Dim k As Variant
        
        ' Parcurgem tipurile de lucrari (sortate alfabetic)
        Dim tipuri() As String
        ReDim tipuri(0 To dictManopera.count - 1)
        i = 0
        For Each k In dictManopera.Keys
            tipuri(i) = k: i = i + 1
        Next k
        ' Sortare simpla (bubble sort pentru claritate)
        Dim swapped As Boolean, temp As String
        Do
            swapped = False
            For i = 0 To UBound(tipuri) - 1
                If StrComp(tipuri(i), tipuri(i + 1), vbTextCompare) > 0 Then
                    temp = tipuri(i)
                    tipuri(i) = tipuri(i + 1)
                    tipuri(i + 1) = temp
                    swapped = True
                End If
            Next i
        Loop While swapped
        
        ' Scriem fiecare grup
        For i = 0 To UBound(tipuri)
            tipLucrare = tipuri(i)
            totalManopera = dictManopera(tipLucrare)
            
            ' Gasim toate transporturile pentru acest tip
            Dim transportKeys As Collection
            Set transportKeys = New Collection
            For Each k In dictTransport.Keys
                If Split(k, "|")(0) = tipLucrare Then
                    transportKeys.Add k
                End If
            Next k
            
            If transportKeys.count = 0 Then
                ' Daca nu exista transporturi, tot scriem tipul + manopera
                .Cells(r, 1).value = tipLucrare
                .Cells(r, 2).value = Application.WorksheetFunction.Round(totalManopera, 3)
                r = r + 1
            Else
                Dim j As Long
                For j = 1 To transportKeys.count
                    k = transportKeys(j)
                    denumireTransport = dictTransport(k)(2)
                    motoare = dictTransport(k)(0)
                    km = dictTransport(k)(1)
                    
                    If j = 1 Then
                        .Cells(r, 1).value = tipLucrare
                        .Cells(r, 2).value = Application.WorksheetFunction.Round(totalManopera, 3)
                    End If
                    .Cells(r, 3).value = denumireTransport
                    .Cells(r, 4).value = Application.WorksheetFunction.Round(motoare, 3)
                    .Cells(r, 5).value = Application.WorksheetFunction.Round(km, 3)
                    r = r + 1
                Next j
            End If
        Next i
        
        .Columns.AutoFit
    End With
    
        wsStat.Move After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count)
        ' --- Activeaza foaia ---
        wsStat.Visible = xlSheetVisible
        wsStat.Activate
    
    'MsgBox "Statistica a fost generata cu succes în foaia 'Statistica'.", vbInformation

Finalize:
    RestoreStructureIfNeeded wasProt
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = prevEvents
    If Err.Number <> 0 Then
        MsgBox "Eroare în GenereazaStatistica: " & Err.Description, vbExclamation
    End If
End Sub

Sub ImportaToateFoileDinAltWorkbook()
    Dim fd As FileDialog
    Dim FisierSelectat As String
    Dim wbSursa As Workbook
    Dim wbDestinatie As Workbook
    Dim ws As Worksheet
    
    ' Setam workbook-ul curent (destina?ia)
    Set wbDestinatie = ThisWorkbook
    
    ' Deschidem dialogul pentru alegerea fi?ierului sursa
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Selecteaza fi?ierul Excel pentru import"
        .Filters.Clear
        .Filters.Add "Fi?iere Excel", "*.xls; *.xlsx; *.xlsm"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            FisierSelectat = .SelectedItems(1)
        Else
            MsgBox "Opera?iunea a fost anulata.", vbInformation
            Exit Sub
        End If
    End With
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Deschidem workbook-ul sursa în background
    Set wbSursa = Workbooks.Open(FisierSelectat, ReadOnly:=True)
    
    ' Copiem fiecare foaie în workbook-ul curent
    For Each ws In wbSursa.Worksheets
        ws.Copy After:=wbDestinatie.Sheets(wbDestinatie.Sheets.count)
    Next ws
    
    ' Închidem workbook-ul sursa fara a salva
    wbSursa.Close SaveChanges:=False
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    MsgBox "Toate foile au fost importate cu succes!", vbInformation
End Sub

Public Sub ProtejeazaFoiaSTART()
    Dim ws As Worksheet
    Dim wasProtWb As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("START")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Foaia 'START' nu exista.", vbExclamation
        Exit Sub
    End If
    
    ' 1. Protejeaza workbook-ul (fara parola) pentru a bloca ?tergerea foilor
    wasProtWb = UnprotectStructureIfNeeded()
    ThisWorkbook.Protect Structure:=True, Windows:=False
    RestoreStructureIfNeeded wasProtWb ' Asigura ca ramâne protejat
    
    ' 2. Protejeaza foaia START (fara parola, toate op?iunile blocate)
    ws.Unprotect ' Deblocheaza temporar (daca era protejata)
    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                AllowFormattingCells:=False, _
                AllowFormattingColumns:=False, _
                AllowFormattingRows:=False, _
                AllowInsertingColumns:=False, _
                AllowInsertingRows:=False, _
                AllowDeletingColumns:=False, _
                AllowDeletingRows:=False, _
                AllowSorting:=False, _
                AllowFiltering:=False, _
                AllowUsingPivotTables:=False
    ' Nu se seteaza parola ? oricine poate vedea ca e protejata, dar nu poate edita
    
    ' Op?ional: ascunde bara de formule pentru START (daca dore?ti)
    ' ws.EnableSelection = xlUnlockedCells ' doar daca ai celule deblocate
    
    MsgBox "Foaia 'START' este protejata împotriva ?tergerii ?i modificarii (fara parola).", vbInformation
End Sub

