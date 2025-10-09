Private Sub CommandButton2_Click()
Unload UserForm1
End Sub

Private Sub ListBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        ' Simuleaza click pe butonul Adauga
        Call CommandButton1_Click
        ' (op?ional) opre?te propagarea tastei Enter
        KeyCode = 0
    End If
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ' Simuleaza ac?iunea butonului Adauga
    Call CommandButton1_Click
End Sub

Private Sub UserForm_Activate()
    Dim sheetName As String

    ' Mapeaza Tag-ul UserForm la foaia reala de baza de date (poate folose?ti la validare sau mesaje)
    Select Case Me.Tag
        Case "CautareNorma":      sheetName = "Norma"
        Case "CautareObiect":     sheetName = "Obiect"
        Case "CautareMateriale":  sheetName = "Materiale"
        Case "CautareUtilaj":     sheetName = "Utilaj"
        Case "CautareTransport":  sheetName = "Transport"
        Case "CautareListe":      sheetName = "Liste"
        Case Else:                sheetName = ""
    End Select

    If sheetName <> "" Then
        If Not Module1.SheetExistsAll(sheetName) Then
            MsgBox "Foaia '" & sheetName & "' nu exista оn registru!", vbExclamation
        End If
    Else
        ' Po?i afi?a un mesaj sau sa la?i ListBox-ul nepopulat daca Tag-ul nu e corect.
        'MsgBox "Tag-ul UserForm nu este setat corect!", vbExclamation
    End If

    ' Pozi?ionare ?i alte ini?ializari (po?i pastra)
    Dim latimeEcran As Long
    Dim inaltimeEcran As Long
    Dim latimeForma As Long
    Dim inaltimeForma As Long

    latimeEcran = Application.UsableWidth
    inaltimeEcran = Application.UsableHeight

    latimeForma = Me.Width
    inaltimeForma = Me.Height

    With Me
        .StartUpPosition = 0
        .Left = Application.Max(0, (latimeEcran - latimeForma))
        .Top = Application.Max(0, (inaltimeEcran - inaltimeForma) / 2)
    End With
End Sub



Private Sub CommandButton1_Click()

    ' Revenire de pe Page7 la pagina anterioara
    
    ' Buton de transfer date de la ListBox1 la TextBox-urile din Page1
    If ListBox1.ListIndex = -1 Then
        MsgBox "Selectati o linie din lista!", vbExclamation
        Exit Sub
    End If

    Dim sursaDate As String
    sursaDate = Me.Tag
    
    Dim valoare1 As String
    Dim valoare2 As String
    Dim valoare3 As String
    Dim valoare4 As String
    Dim valoare5 As String
    Dim valoare6 As String
    Dim valoare7 As String

    Select Case sursaDate
        Case "CautareObiect"
            valoare1 = ListBox1.List(ListBox1.ListIndex, 1) ' NumeObiect
            valoare2 = ListBox1.List(ListBox1.ListIndex, 0) ' NrInventar
            ActiveSheet.Range("C15").value = valoare1
            ActiveSheet.Range("C17").value = valoare2
            ActiveSheet.Range("C15").Select
            Unload UserForm1

        Case "CautareNorma"
            valoare1 = ListBox1.List(ListBox1.ListIndex, 0)
            valoare2 = ListBox1.List(ListBox1.ListIndex, 1)
            valoare3 = ListBox1.List(ListBox1.ListIndex, 2)
            valoare4 = ListBox1.List(ListBox1.ListIndex, 3)
            InserareRandCopy valoare1 & ";" & valoare2 & ";" & valoare3 & ";" & valoare4, "B;C;D;H", False
            Unload UserForm1

        Case "CautareMateriale"
            valoare1 = ListBox1.List(ListBox1.ListIndex, 0)
            valoare2 = ListBox1.List(ListBox1.ListIndex, 1)
            valoare3 = ListBox1.List(ListBox1.ListIndex, 2)
            valoare5 = ListBox1.List(ListBox1.ListIndex, 4)
            InserareRandCopy valoare1 & ";" & valoare2 & ";" & valoare3 & ";" & valoare5, "B;C;D;F"
            
        Case "CautareUtilaj"
            valoare1 = ListBox1.List(ListBox1.ListIndex, 0)
            valoare2 = ListBox1.List(ListBox1.ListIndex, 1)
            valoare3 = ListBox1.List(ListBox1.ListIndex, 2)
            valoare4 = ListBox1.List(ListBox1.ListIndex, 3)
            InserareRandCopy valoare1 & ";" & valoare2 & ";" & valoare3 & ";" & valoare4, "C;D;E;F"
            Unload UserForm1

        Case "CautareTransport"
            valoare1 = ListBox1.List(ListBox1.ListIndex, 0)
            valoare2 = ListBox1.List(ListBox1.ListIndex, 1)
            valoare3 = ListBox1.List(ListBox1.ListIndex, 2)
            valoare4 = ListBox1.List(ListBox1.ListIndex, 3)
            valoare5 = ListBox1.List(ListBox1.ListIndex, 4)
            valoare6 = ListBox1.List(ListBox1.ListIndex, 5)
            valoare7 = ListBox1.List(ListBox1.ListIndex, 6)
            InserareRandCopy valoare1 & ";" & valoare3 & ",   " & valoare2 & ";" & valoare4 & ";" & valoare5, "B;C;D;F", False
            Rindul = Rindul + 1
            ActiveSheet.Range("D" & Rindul).value = valoare6
            ActiveSheet.Range("F" & Rindul).value = valoare7
            Rindul = Rindul
            Unload UserForm1

        Case Else
            MsgBox "Sursa de date necunoscuta!", vbExclamation
            Exit Sub
    End Select
End Sub

Private Sub TextBox1_Change()

    ' Cautare in timp real cand se modifica textul
    Dim Criteriu As String
    Dim Coloana As String
    Dim sheetName As String
    
    Criteriu = TextBox1.text
    
    If ComboBox1.ListIndex = -1 Then
        Exit Sub
    End If
    
    Coloana = ComboBox1.text
    
    If InStr(Me.Tag, "Cautare") > 0 Then
        sheetName = Replace(Me.Tag, "Cautare", "")
    Else
        Exit Sub
    End If
    
    FiltreazaListBoxDinFoaiaCurenta sheetName, Criteriu, Coloana

End Sub
